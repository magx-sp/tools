const ns = 'http://www.w3.org/2000/svg';
const canvas = document.getElementById('canvas');
const paperType = document.getElementById('paperType');
const itemW = document.getElementById('itemWidth');
const itemH = document.getElementById('itemHeight');
const rotate = document.getElementById('rotate');
const bleedLong = document.getElementById('bleedLong');
const bleedShort = document.getElementById('bleedShort');
const mbottom = document.getElementById('marginBottom');
const marginRight = document.getElementById('marginRight');
const quantityInput = document.getElementById('quantity');
const sheetsInput = document.getElementById('sheetsInput'); // New: 原紙枚数入力フィールド
const overrideX = document.getElementById('overrideX');
const overrideY = document.getElementById('overrideY');
const calcBtn = document.getElementById('calcSheetsBtn');
const multiDesign = document.getElementById('multiDesign');
const designCount = document.getElementById('designCount');
const designList = document.getElementById('designList');
const designInputs = document.getElementById('designInputs');
const orderType = document.getElementById('orderType');
const marginsEl = document.getElementById('margins');
const resultBody = document.getElementById('resultBody');
const downloadSvgBtn = document.getElementById('downloadSvgBtn');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');

// 計算結果を保持するためのグローバル変数
let currentLayoutResults = null;

function findOptimalLayout(req, total) {
  let best = null;
  function gen(i, rem, arr) {
    if (i === req.length - 1) { if (rem >= 1) evaluate(arr.concat(rem)); return; }
    for (let x = 1; x <= rem - (req.length - i - 1); x++) gen(i+1, rem-x, arr.concat(x));
  }
  function evaluate(layout) {
    const plates = layout.map((f,i) => Math.ceil(req[i]/f));
    const mp = Math.max(...plates);
    const actual = layout.map(f => f*mp);
    const over = actual.map((a,i) => a-req[i]);
    const totalOver = over.reduce((s,v)=>s+v,0);
    if (!best || totalOver < best.totalOver) best={layout,plates:mp,actual,over,totalOver};
  }
  gen(0, total, []);
  return best;
}

function setupDesignInputs() {
  designInputs.style.display=multiDesign.checked ? 'block' : 'none';
  designList.innerHTML='';
  const n = parseInt(designCount.value,10)||1;
  for (let i=1; i<=n; i++) {
    const div = document.createElement('div');
    div.innerHTML = `<label>デザイン${i}／印刷数量<br><input type=\"number\" id=\"dq${i}\" value=\"0\" min=\"0\"></label>`;
    designList.appendChild(div);
    document.getElementById(`dq${i}`).addEventListener('input', render);
  }
  quantityInput.disabled = multiDesign.checked;
  // マルチデザインのチェック状態が変わったら、計算結果をクリアして再計算を促す
  currentLayoutResults = null;
  calculate(); // setupDesignInputsが呼ばれるたびにcalculateも呼ぶ
}

function calcLayout() {
  const [sw, sh] = paperType.value.split(',').map(Number);
  let w = +itemW.value, h = +itemH.value;
  if (rotate.checked) [w,h]=[h,w];
  const bl = +bleedLong.value, bs = +bleedShort.value, mb = +mbottom.value, hr = +marginRight.value;
  const tw = w+bl, th = h+bs;
  const autoCx = Math.max(1, Math.floor((sw-hr)/tw));
  const autoCy = Math.max(1, Math.floor((sh-mb)/th));

  // overrideの値がない場合は自動計算値を設定
  const currentOverrideX = overrideX.value ? +overrideX.value : null;
  const currentOverrideY = overrideY.value ? +overrideY.value : null;

  const cx = currentOverrideX !== null ? currentOverrideX : autoCx;
  const cy = currentOverrideY !== null ? currentOverrideY : autoCy;

  overrideX.disabled=false;
  overrideY.disabled=false;

  if (overrideX.value === '' || isNaN(currentOverrideX)) {
    overrideX.value = autoCx;
  }
  if (overrideY.value === '' || isNaN(currentOverrideY)) {
    overrideY.value = autoCy;
  }

  return {sw,sh,w,h,bl,bs,mb,hr,tw,th,cx,cy,total:cx*cy};
}

function render() {
  const L=calcLayout();
  canvas.setAttribute('viewBox',`0 0 ${L.sw} ${L.sh}`);
  canvas.innerHTML='';
  const frame=document.createElementNS(ns,'rect');
  frame.setAttribute('x',0); frame.setAttribute('y',0);
  frame.setAttribute('width',L.sw); frame.setAttribute('height',L.sh);
  frame.setAttribute('fill','none'); frame.setAttribute('stroke','#333');
  frame.setAttribute('stroke-width','2'); // 紙枠の太さを指定
  canvas.appendChild(frame);
  const startX=L.sw - L.hr - L.cx * L.tw + L.bl;
  const startY=L.sh - L.mb - L.cy * L.th + L.bs;
  const cols=['#add8e6','#ffb6c1','#90ee90','#ffa500','#dda0dd','#87ceeb','#98fb98','#f08080','#e0ffff','#f5deb3'];
  const n=parseInt(designCount.value,10)||1;

  // 面付イメージに反映するためのデザイン順序配列を生成
  let effectiveSeq = [];
  if (multiDesign.checked && currentLayoutResults && currentLayoutResults.layout) {
    const layoutCounts = currentLayoutResults.layout; // 例: [4, 8]
    const totalCells = L.cx * L.cy;
    let designIndex = 0;
    let currentDesignAccumulatedCount = 0;

    const targetDesignCounts = [];
    for (let i = 0; i < layoutCounts.length; i++) {
        targetDesignCounts.push(layoutCounts[i]);
    }

    for (let k = 0; k < totalCells; k++) {
      if (designIndex < targetDesignCounts.length) {
        effectiveSeq.push(designIndex);
        currentDesignAccumulatedCount++;
        if (currentDesignAccumulatedCount >= targetDesignCounts[designIndex]) {
          designIndex++;
          currentDesignAccumulatedCount = 0;
        }
      } else {
        effectiveSeq.push(0);
      }
    }
  } else {
      for (let k = 0; k < L.cx * L.cy; k++) {
          effectiveSeq.push(0);
      }
  }


  for(let i=0;i<L.cy;i++){
    for(let j=0;j<L.cx;j++){
      let idx=0;
      if(multiDesign.checked) {
        if (orderType.value === 'row') {
          idx = effectiveSeq[i*L.cx+j] !== undefined ? effectiveSeq[i*L.cx+j] : 0;
        } else { // 'col'
          idx = effectiveSeq[j*L.cy+i] !== undefined ? effectiveSeq[j*L.cy+i] : 0;
        }
      }
      const r=document.createElementNS(ns,'rect');
      r.setAttribute('x',startX+j*L.tw); r.setAttribute('y',startY+i*L.th);
      r.setAttribute('width',L.w); r.setAttribute('height',L.h); r.setAttribute('fill',cols[idx % cols.length]);
      r.setAttribute('stroke','#666');
      canvas.appendChild(r);
      if(multiDesign.checked){
        const t=document.createElementNS(ns,'text');
        t.setAttribute('x',startX+j*L.tw+L.w/2); t.setAttribute('y',startY+i*L.th+L.h/2);
        t.setAttribute('text-anchor','middle'); t.setAttribute('dominant-baseline','middle');
        t.setAttribute('font-size',Math.min(L.w,L.h)/4); t.textContent=idx+1;
        canvas.appendChild(t);
      }
    }
  }
  const displayLeftMargin = startX;
  const displayTopMargin = startY;
  const displayRightMargin = L.hr; // ハリ
  const displayBottomMargin = L.mb; // クワエ

  marginsEl.textContent=`[余白]　左：${displayLeftMargin.toFixed(1)}mm　右：${displayRightMargin.toFixed(1)}mm　上：${displayTopMargin.toFixed(1)}mm　下：${displayBottomMargin.toFixed(1)}mm`;
}

function calculate() {
  const L=calcLayout();
  let tableRows = ''; // HTML文字列を構築するための変数

  // Enable sheetsInput
  sheetsInput.disabled = false;

  if(multiDesign.checked){
    const req=[];
    for(let i=1;i<=parseInt(designCount.value,10);i++) {
      req.push(+document.getElementById(`dq${i}`).value||0);
    }
    const best=findOptimalLayout(req,L.total);
    currentLayoutResults = best; // 計算結果を保存

    let sheets;
    // If sheetsInput has a value, use it, otherwise calculate based on best.plates
    if (sheetsInput.value !== '' && !isNaN(parseInt(sheetsInput.value, 10)) && parseInt(sheetsInput.value, 10) >= 0) {
      sheets = parseInt(sheetsInput.value, 10);
    } else {
      sheets = best.plates;
      sheetsInput.value = sheets; // Update sheetsInput with the calculated value
    }

    // 原紙寸法と製品サイズ
    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;

    // 面付数
    tableRows += `<tr><td>面付数</td><td>`;
    req.forEach((_,i)=>{
      const f=best.layout[i];
      const rows=f/L.cx;
      // 小数点以下2桁表示
      tableRows += `<div>デザイン${i+1}／${L.cx}×${rows.toFixed(2)}＝${f}面</div>`;
    });
    tableRows += `</td></tr>`;

    // 必要原紙
    tableRows += `<tr><td>必要原紙</td><td>${sheets}枚</td></tr>`;

    // 予備 (各デザインの面付数 × 原紙枚数 − 各印刷数量)
    tableRows += `<tr><td>予備</td><td>`;
    req.forEach((qty,i)=>{
      const f=best.layout[i]; // 面付数
      const spare = (f * sheets) - qty;
      const totalPrinted = f * sheets;
      tableRows += `<div>デザイン${i+1}／${spare}枚（総印刷枚数：${totalPrinted}枚）</div>`;
    });
    tableRows += `</td></tr>`;

    // ドブ・クワエ・ハリ
    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;

  } else { // シングルデザインの場合
    const qty=+quantityInput.value||0;
    let sheets;
    // If sheetsInput has a value, use it, otherwise calculate based on quantity
    if (sheetsInput.value !== '' && !isNaN(parseInt(sheetsInput.value, 10)) && parseInt(sheetsInput.value, 10) >= 0) {
      sheets = parseInt(sheetsInput.value, 10);
    } else {
      sheets = Math.ceil(qty/L.total);
      sheetsInput.value = sheets; // Update sheetsInput with the calculated value
    }

    const res = (sheets * L.total) - qty; // 予備の計算

    currentLayoutResults = null; // シングルデザインの場合はクリア

    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;
    tableRows += `<tr><td>面付数</td><td>${L.cx}×${L.cy}＝${L.total}面</td></tr>`;
    tableRows += `<tr><td>必要原紙</td><td>${sheets}枚</td></tr>`;
    tableRows += `<tr><td>予備</td><td>${res}枚（総印刷枚数：${sheets*L.total}枚）</td></tr>`;
    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;
  }
  resultBody.innerHTML=tableRows;

  // 計算が終わったらSVGを再描画して結果を反映
  render();
}

[paperType,itemW,itemH,rotate,bleedLong,bleedShort,mbottom,marginRight,quantityInput].forEach(el=>el.addEventListener('input',()=>{ sheetsInput.value=''; overrideX.value=''; overrideY.value=''; render(); }));
overrideX.addEventListener('input',render);
overrideY.addEventListener('input',render);
calcBtn.addEventListener('click',calculate);
sheetsInput.addEventListener('input', calculate); // New: Add event listener for sheetsInput

downloadSvgBtn.addEventListener('click',()=>{ const vb=canvas.viewBox.baseVal, cl=canvas.cloneNode(true); cl.setAttribute('xmlns',ns); cl.setAttribute('width',vb.width+'mm'); cl.setAttribute('height',vb.height+'mm'); const xml=new XMLSerializer().serializeToString(cl), blob=new Blob([xml],{type:'image/svg+xml'}), url=URL.createObjectURL(blob), a=document.createElement('a'); a.href=url; a.download='layout.svg'; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url); });
downloadPdfBtn.addEventListener('click',()=>{
  const { jsPDF }=window.jspdf;
  const pdf=new jsPDF({unit:'mm',format:'a4',orientation:'landscape'});

  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  const margin = 10;
  const drawableWidth = pdfWidth - (margin * 2);
  const drawableHeight = pdfHeight - (margin * 2);

  html2canvas(document.getElementById('pdfArea'),{scale:2}).then(imgCan=>{
    const imgData=imgCan.toDataURL('image/jpeg',0.7);

    const capturedWidthPx = imgCan.width;
    const capturedHeightPx = imgCan.height;

    let imgWidthMM;
    let imgHeightMM;

    const imgAspectRatio = capturedWidthPx / capturedHeightPx;
    const drawableAspectRatio = drawableWidth / drawableHeight;

    if (imgAspectRatio > drawableAspectRatio) {
      imgWidthMM = drawableWidth;
      imgHeightMM = drawableWidth / imgAspectRatio;
    } else {
      imgHeightMM = drawableHeight;
      imgWidthMM = drawableHeight * imgAspectRatio;
    }

    const xOffset = margin + (drawableWidth - imgWidthMM) / 2;
    const yOffset = margin + (drawableHeight - imgHeightMM) / 2;

    pdf.addImage(imgData,'JPEG',xOffset,yOffset,imgWidthMM,imgHeightMM);
    pdf.save('layout.pdf');
  });
});
multiDesign.addEventListener('change',() => {
    setupDesignInputs();
    calculate(); // マルチデザインのチェックボックスが変更されたら計算も実行
});
designCount.addEventListener('input',setupDesignInputs);
orderType.addEventListener('change',render);
window.addEventListener('DOMContentLoaded',()=>{
    setupDesignInputs(); // 初期ロード時にもデザイン入力欄をセットアップ
    calculate(); // 初期ロード時に計算結果も表示
});