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
}

function calcLayout() {
  const [sw, sh] = paperType.value.split(',').map(Number);
  let w = +itemW.value, h = +itemH.value;
  if (rotate.checked) [w,h]=[h,w];
  const bl = +bleedLong.value, bs = +bleedShort.value, mb = +mbottom.value, hr = +marginRight.value;
  const tw = w+bl, th = h+bs;
  const autoCx = Math.max(1, Math.floor((sw-hr)/tw));
  const autoCy = Math.max(1, Math.floor((sh-mb)/th));

  // Override values are only applied if they are valid numbers and not empty.
  const currentOverrideX = overrideX.value ? +overrideX.value : null;
  const currentOverrideY = overrideY.value ? +overrideY.value : null;

  const cx = currentOverrideX !== null ? currentOverrideX : autoCx;
  const cy = currentOverrideY !== null ? currentOverrideY : autoCy;

  overrideX.disabled=false;
  overrideY.disabled=false;

  // Update override input fields if they were initially empty or invalid
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
  canvas.appendChild(frame);
  const startX=L.sw - L.hr - L.cx * L.tw + L.bl;
  const startY=L.sh - L.mb - L.cy * L.th + L.bs;
  const cols=['#add8e6','#ffb6c1','#90ee90','#ffa500','#dda0dd','#87ceeb','#98fb98','#f08080','#e0ffff','#f5deb3'];
  const n=parseInt(designCount.value,10)||1;
  const seq=[];
  if(multiDesign.checked) {
    const totalCells = L.cx * L.cy;
    for(let k=0;k<n;k++) {
      for(let m=0;m<Math.ceil(totalCells/n);m++) {
        seq.push(k);
      }
    }
  }

  for(let i=0;i<L.cy;i++){
    for(let j=0;j<L.cx;j++){
      let idx=0;
      if(multiDesign.checked) {
        if (orderType.value === 'row') {
          idx = seq[i*L.cx+j] !== undefined ? seq[i*L.cx+j] : 0; // Fallback to 0 if out of bounds
        } else { // 'col'
          idx = seq[j*L.cy+i] !== undefined ? seq[j*L.cy+i] : 0; // Fallback to 0 if out of bounds
        }
      }
      const r=document.createElementNS(ns,'rect');
      r.setAttribute('x',startX+j*L.tw); r.setAttribute('y',startY+i*L.th);
      r.setAttribute('width',L.w); r.setAttribute('height',L.h); r.setAttribute('fill',cols[idx % cols.length]); r.setAttribute('stroke','#666'); // Use modulo for color
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
  marginsEl.textContent=`左:${startX.toFixed(1)}mm  右:${(L.sw-(startX+L.cx*L.tw)+L.bl).toFixed(1)}mm  上:${startY.toFixed(1)}mm  下:${(L.sh-(startY+L.cy*L.th)+L.bs).toFixed(1)}mm`;
}

function calculate() {
  const L=calcLayout();
  const lines=[];
  if(multiDesign.checked){
    const req=[];
    for(let i=1;i<=parseInt(designCount.value,10);i++) {
      req.push(+document.getElementById(`dq${i}`).value||0);
    }
    const best=findOptimalLayout(req,L.total);
    lines.push(`原紙寸法：${L.sw}×${L.sh}mm`,`製品サイズ：${L.w}×${L.h}mm`);
    req.forEach((_,i)=>{
      const f=best.layout[i];
      const rows=f/L.cx;
      const p=i===0?'面付数：':'&emsp;&emsp;';
      lines.push(`${p}デザイン${i+1}／${L.cx}×${rows}＝${f}面`);
    });
    lines.push(`必要原紙：${best.plates}枚`);
    req.forEach((_,i)=>{
      const o=best.over[i];
      const a=best.actual[i];
      const p=i===0?'予備：':'&emsp;&emsp;';
      lines.push(`${p}デザイン${i+1}／${o}枚（総印刷枚数：${a}枚）`);
    });
    lines.push(`ドブ：${L.bl}・${L.bs}（長辺・短辺）／クワエ：${L.mb}／ハリ：${L.hr}（mm）`);
  } else {
    const qty=+quantityInput.value||0;
    const sheets=Math.ceil(qty/L.total);
    const res=sheets*L.total-qty;
    lines.push(`原紙寸法：${L.sw}×${L.sh}mm`,`製品サイズ：${L.w}×${L.h}mm`,`面付数：${L.cx}×${L.cy}＝${L.total}面`,`必要原紙：${sheets}枚`,`予備：${res}枚（総印刷枚数：${sheets*L.total}枚）`,`ドブ：${L.bl}・${L.bs}（長辺・短辺）／クワエ：${L.mb}／ハリ：${L.hr}（mm）`);
  }
  resultBody.innerHTML=`<tr><td colspan=\"2\">${lines.join('<br>')}</td></tr>`;
}

// Event Listeners
[paperType,itemW,itemH,rotate,bleedLong,bleedShort,mbottom,marginRight].forEach(el=>el.addEventListener('input',()=>{ overrideX.value=''; overrideY.value=''; render(); }));
overrideX.addEventListener('input',render);
overrideY.addEventListener('input',render);
calcBtn.addEventListener('click',calculate);

downloadSvgBtn.addEventListener('click',()=>{
  const vb=canvas.viewBox.baseVal;
  const cl=canvas.cloneNode(true);
  cl.setAttribute('xmlns',ns);
  cl.setAttribute('width',vb.width+'mm');
  cl.setAttribute('height',vb.height+'mm');
  const xml=new XMLSerializer().serializeToString(cl);
  const blob=new Blob([xml],{type:'image/svg+xml'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download='layout.svg';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

downloadPdfBtn.addEventListener('click',()=>{
  const { jsPDF }=window.jspdf;
  const pdf=new jsPDF({unit:'mm',format:'a4',orientation:'landscape'});
  const m=15;
  html2canvas(document.getElementById('pdfArea'),{scale:2}).then(imgCan=>{
    const imgData=imgCan.toDataURL('image/jpeg',0.7);
    const pw=pdf.internal.pageSize.getWidth()-m*2;
    const phMax=pdf.internal.pageSize.getHeight()-m*2;
    let ih=(imgCan.height*pw)/imgCan.width;
    let iw=pw;
    if(ih>phMax){
      ih=phMax;
      iw=(imgCan.width*phMax)/imgCan.height;
    }
    pdf.addImage(imgData,'JPEG',m,m,iw,ih);
    pdf.save('layout.pdf');
  });
});

multiDesign.addEventListener('change',setupDesignInputs);
designCount.addEventListener('input',setupDesignInputs);
orderType.addEventListener('change',render);

window.addEventListener('DOMContentLoaded',()=>{
  setupDesignInputs(); // Call this initially to set up the correct state based on checkbox
  render();
});