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

function findOptimalLayout(reqQuantities, totalFacesPerSheet) {
  let best = null;
  let minVariance = Infinity;

  function generateCombinations(designIndex, remainingFaces, currentLayout) {
    if (designIndex === reqQuantities.length - 1) {
      if (remainingFaces >= 1) {
        evaluateLayout(currentLayout.concat(remainingFaces));
      }
      return;
    }
    for (let facesForCurrentDesign = 1;
         facesForCurrentDesign <= remainingFaces - (reqQuantities.length - designIndex - 1);
         facesForCurrentDesign++) {
      generateCombinations(
        designIndex + 1,
        remainingFaces - facesForCurrentDesign,
        currentLayout.concat(facesForCurrentDesign)
      );
    }
  }

  function evaluateLayout(layout) {
    const sheetsPerDesign = layout.map((facesPerSheet, i) =>
        (facesPerSheet > 0) ? Math.ceil(reqQuantities[i] / facesPerSheet) : 0
    );

    const maxSheets = Math.max(...sheetsPerDesign);
    if (maxSheets === 0 && reqQuantities.some(q => q > 0)) {
        return;
    }
    if (maxSheets === 0 && reqQuantities.every(q => q === 0)) {
        best = {
            layout: layout,
            plates: 0,
            actualQuantities: reqQuantities.map(() => 0),
            overQuantities: reqQuantities.map(() => 0),
            overRates: reqQuantities.map(() => 0),
            variance: 0
        };
        minVariance = 0;
        return;
    }

    const actualQuantities = layout.map(facesPerSheet => facesPerSheet * maxSheets);
    const overQuantities = actualQuantities.map((actual, i) => actual - reqQuantities[i]);
    const overRates = overQuantities.map((over, i) =>
        reqQuantities[i] > 0 ? (over / reqQuantities[i]) * 100 : 0
    );

    const validOverRates = overRates.filter((rate, i) => reqQuantities[i] > 0);
    const mean = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + rate, 0) / validOverRates.length : 0;
    const variance = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + Math.pow(rate - mean, 2), 0) / validOverRates.length : 0;

    if (variance < minVariance) {
      minVariance = variance;
      best = {
        layout: layout,
        plates: maxSheets,
        actualQuantities: actualQuantities,
        overQuantities: overQuantities,
        overRates: overRates,
        variance: variance
      };
    }
  }

  if (reqQuantities.length > 0 && totalFacesPerSheet > 0) {
    generateCombinations(0, totalFacesPerSheet, []);
  } else {
      return {
          layout: reqQuantities.map(() => 0),
          plates: 0,
          actualQuantities: reqQuantities.map(() => 0),
          overQuantities: reqQuantities.map(() => 0),
          overRates: reqQuantities.map(() => 0),
          variance: 0
      };
  }

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
    document.getElementById(`dq${i}`).addEventListener('input', calculate);
  }
  quantityInput.disabled = multiDesign.checked;
  currentLayoutResults = null;
  calculate();
}

function calcLayout() {
  const [sw, sh] = paperType.value.split(',').map(Number);
  let w = +itemW.value, h = +itemH.value;
  if (rotate.checked) [w,h]=[h,w];
  const bl = +bleedLong.value, bs = +bleedShort.value, mb = +mbottom.value, hr = +marginRight.value;
  const tw = w+bl, th = h+bs;
  const autoCx = Math.max(1, Math.floor((sw-hr)/tw));
  const autoCy = Math.max(1, Math.floor((sh-mb)/th));

  const currentOverrideX = overrideX.value ? +overrideX.value : null;
  const currentOverrideY = overrideY.value ? +overrideY.value : null;

  const cx = currentOverrideX !== null ? currentOverrideX : autoCx;
  const cy = currentOverrideY !== null ? currentOverrideY : autoCy;

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
  frame.setAttribute('stroke-width','2');
  canvas.appendChild(frame);
  const startX=L.sw - L.hr - L.cx * L.tw + L.bl;
  const startY=L.sh - L.mb - L.cy * L.th + L.bs;
  const cols=['#add8e6','#ffb6c1','#90ee90','#ffa500','#dda0dd','#87ceeb','#98fb98','#f08080','#e0ffff','#f5deb3'];
  const n=parseInt(designCount.value,10)||1;

  let effectiveSeq = [];
  if (multiDesign.checked && currentLayoutResults && currentLayoutResults.layout) {
    const layoutCounts = currentLayoutResults.layout;
    const totalCells = L.cx * L.cy;
    let designIndex = 0;
    let currentDesignAccumulatedCount = 0;

    const targetDesignCounts = [];
    for (let i = 0; i < layoutCounts.length; i++) {
        targetDesignCounts.push(layoutCounts[i]);
    }

    for (let k = 0; k < totalCells; k++) {
      if (designIndex < targetDesignCounts.length) {
        if (currentDesignAccumulatedCount < targetDesignCounts[designIndex]) {
          effectiveSeq.push(designIndex);
          currentDesignAccumulatedCount++;
        } else {
          designIndex++;
          currentDesignAccumulatedCount = 0;
          if (designIndex < targetDesignCounts.length) {
            effectiveSeq.push(designIndex);
            currentDesignAccumulatedCount++;
          } else {
            effectiveSeq.push(0);
          }
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
      idx = effectiveSeq[i*L.cx+j] !== undefined ? effectiveSeq[i*L.cx+j] : 0;

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
  const displayRightMargin = L.hr;
  const displayBottomMargin = L.mb;

  marginsEl.textContent=`[余白]　左：${displayLeftMargin.toFixed(1)}mm　右：${displayRightMargin.toFixed(1)}mm　上：${displayTopMargin.toFixed(1)}mm　下：${displayBottomMargin.toFixed(1)}mm`;
}

function calculate() {
  const L=calcLayout();
  let tableRows = '';

  overrideX.disabled=false;
  overrideY.disabled=false;
  sheetsInput.disabled = false;


  if(multiDesign.checked){
    const req=[];
    for(let i=1;i<=parseInt(designCount.value,10);i++) {
      req.push(+document.getElementById(`dq${i}`).value||0);
    }
    const best=findOptimalLayout(req,L.total);
    currentLayoutResults = best;

    let sheets;
    if (sheetsInput.value !== '' && !isNaN(parseInt(sheetsInput.value, 10)) && parseInt(sheetsInput.value, 10) >= 0) {
      sheets = parseInt(sheetsInput.value, 10);
      if (best) {
          best.plates = sheets;
          best.actualQuantities = best.layout.map(facesPerSheet => facesPerSheet * sheets);
          best.overQuantities = best.actualQuantities.map((actual, i) => actual - req[i]);
          best.overRates = best.overQuantities.map((over, i) =>
              req[i] > 0 ? (over / req[i]) * 100 : 0
          );
          const validOverRates = best.overRates.filter((rate, i) => req[i] > 0);
          const mean = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + rate, 0) / validOverRates.length : 0;
          best.variance = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + Math.pow(rate - mean, 2), 0) / validOverRates.length : 0;
      }
    } else {
      sheets = best.plates;
      sheetsInput.value = sheets;
    }

    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;
    tableRows += `<tr><td>面付数</td><td>`;
    if (best && best.layout) {
        req.forEach((_,i)=>{
            const f=best.layout[i];
            tableRows += `<div>デザイン${i+1}／${f}面</div>`;
        });
    } else {
        tableRows += `<div>計算できません</div>`;
    }
    tableRows += `</td></tr>`;

    // --- 変更点：印刷数量の項目を追加 ---
    tableRows += `<tr><td>印刷数量</td><td>`;
    if (best && req) {
        req.forEach((qty, i) => {
            tableRows += `<div>デザイン${i+1}／${qty}枚</div>`;
        });
    } else {
        tableRows += `<div>-</div>`;
    }
    tableRows += `</td></tr>`;
    // --- 変更点ここまで ---

    // --- 変更点：必要原紙枚数の表記変更 ---
    tableRows += `<tr><td>必要原紙枚数</td><td>${sheets}枚</td></tr>`;
    // --- 変更点ここまで ---

    tableRows += `<tr><td>予備</td><td>`;
    if (best && best.layout) {
        req.forEach((qty,i)=>{
            const spare = best.overQuantities[i];
            const totalPrinted = best.actualQuantities[i];
            const overRate = best.overRates[i];
            tableRows += `<div>デザイン${i+1}／${spare}枚（予備率：${overRate.toFixed(2)}%）</div>`;
        });
        tableRows += `<div>予備率分散：${best.variance.toFixed(4)}</div>`;
    } else {
        tableRows += `<div>計算できません</div>`;
    }
    tableRows += `</td></tr>`;

    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;

  } else { // シングルデザインの場合
    const qty=+quantityInput.value||0;
    let sheets;
    if (sheetsInput.value !== '' && !isNaN(parseInt(sheetsInput.value, 10)) && parseInt(sheetsInput.value, 10) >= 0) {
      sheets = parseInt(sheetsInput.value, 10);
    } else {
      sheets = L.total > 0 ? Math.ceil(qty/L.total) : 0;
      sheetsInput.value = sheets;
    }

    const totalPrinted = sheets * L.total;
    const res = totalPrinted - qty;
    const overRate = qty > 0 ? (res / qty) * 100 : 0;

    currentLayoutResults = null;

    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;
    tableRows += `<tr><td>面付数</td><td>${L.cx}×${L.cy}＝${L.total}面</td></tr>`;

    // --- 変更点：印刷数量の項目を追加 ---
    tableRows += `<tr><td>印刷数量</td><td>${qty}枚</td></tr>`;
    // --- 変更点ここまで ---

    // --- 変更点：必要原紙枚数の表記変更 ---
    tableRows += `<tr><td>必要原紙枚数</td><td>${sheets}枚</td></tr>`;
    // --- 変更点ここまで ---

    tableRows += `<tr><td>予備</td><td>${res}枚（予備率：${overRate.toFixed(2)}%）</td></tr>`;
    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;
  }
  resultBody.innerHTML=tableRows;

  render();
}

[paperType,itemW,itemH,rotate,bleedLong,bleedShort,mbottom,marginRight,quantityInput].forEach(el=>el.addEventListener('input',()=>{ sheetsInput.value=''; overrideX.value=''; overrideY.value=''; calculate(); }));
overrideX.addEventListener('input',calculate);
overrideY.addEventListener('input',calculate);
calcBtn.addEventListener('click',calculate);
sheetsInput.addEventListener('input', calculate);

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
    sheetsInput.value = '';
    setupDesignInputs();
    calculate();
});
designCount.addEventListener('input',setupDesignInputs);
orderType.addEventListener('change',render);
window.addEventListener('DOMContentLoaded',()=>{
    setupDesignInputs();
    calculate();
});