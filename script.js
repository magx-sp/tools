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
const sheetsInput = document.getElementById('sheetsInput');
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

// エラーメッセージを表示するヘルパー関数
function displayErrorMessage(message) {
    resultBody.innerHTML = `<tr><td colspan="2" style="color: red; font-weight: bold;">${message}</td></tr>`;
    sheetsInput.value = '0';
    currentLayoutResults = null;
    render();
}

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
    const sheetsPerDesign = layout.map((facesPerSheet, i) => {
        if (reqQuantities[i] > 0 && facesPerSheet === 0) {
            return Infinity;
        }
        return (facesPerSheet > 0) ? Math.ceil(reqQuantities[i] / facesPerSheet) : 0;
    });

    const maxSheets = Math.max(...sheetsPerDesign);

    if (maxSheets === Infinity) {
        return;
    }

    if (maxSheets === 0 && reqQuantities.some(q => q > 0)) {
        return;
    }

    if (reqQuantities.every(q => q === 0)) {
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
  designInputs.style.display = multiDesign.checked ? 'block' : 'none';
  designList.innerHTML = '';
  const n = parseInt(designCount.value, 10) || 1;
  for (let i = 1; i <= n; i++) {
    const div = document.createElement('div');
    div.innerHTML = `<label>デザイン${i}／印刷数量<br><input type="number" id="dq${i}" value="0" min="0"></label>`;
    designList.appendChild(div);
    // [FIX START] Event listener for design quantity inputs
    // Add logic to clear sheetsInput to ensure a fresh calculation
    document.getElementById(`dq${i}`).addEventListener('input', () => {
        sheetsInput.value = '';
        calculate();
    });
    // [FIX END]
  }
  quantityInput.disabled = multiDesign.checked;
  currentLayoutResults = null;
  calculate();
}

function calcLayout() {
  const [sw, sh] = paperType.value.split(',').map(Number);
  let w = +itemW.value, h = +itemH.value;
  if (rotate.checked) [w, h] = [h, w];
  const bl = +bleedLong.value, bs = +bleedShort.value, mb = +mbottom.value, hr = +marginRight.value;
  const tw = w + bl, th = h + bs;
  const autoCx = Math.max(1, Math.floor((sw - hr) / tw));
  const autoCy = Math.max(1, Math.floor((sh - mb) / th));

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

  return { sw, sh, w, h, bl, bs, mb, hr, tw, th, cx, cy, total: cx * cy };
}

function generateDesignSequence(layoutCounts, totalCells, orderType) {
  const sequence = [];
  
  if (!layoutCounts || layoutCounts.length === 0) {
    // シングルデザインの場合
    for (let i = 0; i < totalCells; i++) {
      sequence.push(0);
    }
    return sequence;
  }

  // 各デザインの割り当て面数を配列で保持
  const designAssignments = [];
  for (let designIndex = 0; designIndex < layoutCounts.length; designIndex++) {
    for (let count = 0; count < layoutCounts[designIndex]; count++) {
      designAssignments.push(designIndex);
    }
  }

  // 足りない分は最初のデザインで埋める
  while (designAssignments.length < totalCells) {
    designAssignments.push(0);
  }

  // 多すぎる場合は切り詰める
  if (designAssignments.length > totalCells) {
    designAssignments.length = totalCells;
  }

  return designAssignments;
}

function render() {
  const L = calcLayout();
  canvas.setAttribute('viewBox', `0 0 ${L.sw} ${L.sh}`);
  canvas.innerHTML = '';
  
  const frame = document.createElementNS(ns, 'rect');
  frame.setAttribute('x', 0);
  frame.setAttribute('y', 0);
  frame.setAttribute('width', L.sw);
  frame.setAttribute('height', L.sh);
  frame.setAttribute('fill', 'none');
  frame.setAttribute('stroke', '#333');
  frame.setAttribute('stroke-width', '2');
  canvas.appendChild(frame);
  
  const startX = L.sw - L.hr - L.cx * L.tw + L.bl;
  const startY = L.sh - L.mb - L.cy * L.th + L.bs;
  const cols = ['#add8e6', '#ffb6c1', '#90ee90', '#ffa500', '#dda0dd', '#87ceeb', '#98fb98', '#f08080', '#e0ffff', '#f5deb3'];
  
  // デザイン配列を生成
  let effectiveSeq = [];
  if (multiDesign.checked && currentLayoutResults && currentLayoutResults.layout) {
    effectiveSeq = generateDesignSequence(currentLayoutResults.layout, L.cx * L.cy, orderType.value);
  } else {
    effectiveSeq = generateDesignSequence(null, L.cx * L.cy, orderType.value);
  }

  // 面付を描画
  for (let i = 0; i < L.cy; i++) {
    for (let j = 0; j < L.cx; j++) {
      const cellIndex = i * L.cx + j;
      const designIndex = effectiveSeq[cellIndex] || 0;
      
      const r = document.createElementNS(ns, 'rect');
      r.setAttribute('x', startX + j * L.tw);
      r.setAttribute('y', startY + i * L.th);
      r.setAttribute('width', L.w);
      r.setAttribute('height', L.h);
      r.setAttribute('fill', cols[designIndex % cols.length]);
      r.setAttribute('stroke', '#666');
      canvas.appendChild(r);
      
      if (multiDesign.checked) {
        const t = document.createElementNS(ns, 'text');
        t.setAttribute('x', startX + j * L.tw + L.w / 2);
        t.setAttribute('y', startY + i * L.th + L.h / 2);
        t.setAttribute('text-anchor', 'middle');
        t.setAttribute('dominant-baseline', 'middle');
        t.setAttribute('font-size', Math.min(L.w, L.h) / 4);
        t.textContent = multiDesign.checked ? (effectiveSeq[cellIndex] !== undefined ? effectiveSeq[cellIndex] + 1 : '1') : '1';
        canvas.appendChild(t);
      }
    }
  }
  
  const displayLeftMargin = startX;
  const displayTopMargin = startY;
  const displayRightMargin = L.hr;
  const displayBottomMargin = L.mb;

  marginsEl.textContent = `[余白]　左：${displayLeftMargin.toFixed(1)}mm　右：${displayRightMargin.toFixed(1)}mm　上：${displayTopMargin.toFixed(1)}mm　下：${displayBottomMargin.toFixed(1)}mm`;
}

function calculate() {
  const L = calcLayout();
  let tableRows = '';

  overrideX.disabled = false;
  overrideY.disabled = false;
  sheetsInput.disabled = false;

  if (multiDesign.checked) {
    const req = [];
    const n = parseInt(designCount.value, 10) || 1;
    for (let i = 1; i <= n; i++) {
      const element = document.getElementById(`dq${i}`);
      req.push(element ? (+element.value || 0) : 0);
    }

    if (L.total < req.length && req.some(q => q > 0)) {
      displayErrorMessage(`エラー: 面付数 (${L.total}面) がデザイン数 (${req.length}個) より少ないため、各デザインに1面ずつ割り当てる事ができません。`);
      return;
    }

    const best = findOptimalLayout(req, L.total);
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
      sheets = best ? best.plates : 0;
      sheetsInput.value = sheets;
    }

    // 結果テーブルの構築
    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;

    tableRows += `<tr><td>面付数</td><td>`;
    if (best && best.layout) {
      req.forEach((_, i) => {
        const f = best.layout[i];
        tableRows += `<div>デザイン${i + 1}／${f}面</div>`;
      });
    } else {
      tableRows += `<div>計算できません</div>`;
    }
    tableRows += `</td></tr>`;

    tableRows += `<tr><td>印刷数量</td><td>`;
    req.forEach((qty, i) => {
      tableRows += `<div>デザイン${i + 1}／${qty}枚</div>`;
    });
    tableRows += `</td></tr>`;

    tableRows += `<tr><td>必要原紙枚数</td><td>${sheets}枚</td></tr>`;

    tableRows += `<tr><td>予備</td><td>`;
    if (best && best.layout) {
      req.forEach((qty, i) => {
        const spare = best.overQuantities[i];
        const overRate = best.overRates[i];
        tableRows += `<div>デザイン${i + 1}／${spare}枚（予備率：${overRate.toFixed(2)}%）</div>`;
      });
      tableRows += `<div>予備率分散：${best.variance.toFixed(4)}</div>`;
    } else {
      tableRows += `<div>計算できません</div>`;
    }
    tableRows += `</td></tr>`;

    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;

  } else {
    // シングルデザインの場合
    const qty = +quantityInput.value || 0;
    let sheets;
    if (sheetsInput.value !== '' && !isNaN(parseInt(sheetsInput.value, 10)) && parseInt(sheetsInput.value, 10) >= 0) {
      sheets = parseInt(sheetsInput.value, 10);
    } else {
      sheets = L.total > 0 ? Math.ceil(qty / L.total) : 0;
      sheetsInput.value = sheets;
    }

    const totalPrinted = sheets * L.total;
    const res = totalPrinted - qty;
    const overRate = qty > 0 ? (res / qty) * 100 : 0;

    currentLayoutResults = null;

    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;
    tableRows += `<tr><td>面付数</td><td>${L.cx}×${L.cy}＝${L.total}面</td></tr>`;
    tableRows += `<tr><td>印刷数量</td><td>${qty}枚</td></tr>`;
    tableRows += `<tr><td>必要原紙枚数</td><td>${sheets}枚</td></tr>`;
    tableRows += `<tr><td>予備</td><td>${res}枚（予備率：${overRate.toFixed(2)}%）</td></tr>`;
    tableRows += `<tr><td>ドブ・余白</td><td>ドブ：${L.bl}・${L.bs}（長辺・短辺）<br>クワエ：${L.mb}／ハリ：${L.hr}（mm）</td></tr>`;
  }
  
  resultBody.innerHTML = tableRows;
  render();
}

// イベントリスナーの設定
[paperType, itemW, itemH, rotate, bleedLong, bleedShort, mbottom, marginRight, quantityInput].forEach(el => 
  el.addEventListener('input', () => { 
    sheetsInput.value = ''; 
    overrideX.value = ''; 
    overrideY.value = ''; 
    calculate(); 
  })
);

overrideX.addEventListener('input', calculate);
overrideY.addEventListener('input', calculate);
calcBtn.addEventListener('click', calculate);
sheetsInput.addEventListener('input', calculate);

downloadSvgBtn.addEventListener('click', () => { 
  const vb = canvas.viewBox.baseVal, cl = canvas.cloneNode(true); 
  cl.setAttribute('xmlns', ns); 
  cl.setAttribute('width', vb.width + 'mm'); 
  cl.setAttribute('height', vb.height + 'mm'); 
  const xml = new XMLSerializer().serializeToString(cl), 
        blob = new Blob([xml], {type: 'image/svg+xml'}), 
        url = URL.createObjectURL(blob), 
        a = document.createElement('a'); 
  a.href = url; 
  a.download = 'layout.svg'; 
  document.body.appendChild(a); 
  a.click(); 
  document.body.removeChild(a); 
  URL.revokeObjectURL(url); 
});

downloadPdfBtn.addEventListener('click', () => {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({unit: 'mm', format: 'a4', orientation: 'landscape'});

  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  const margin = 10;
  const drawableWidth = pdfWidth - (margin * 2);
  const drawableHeight = pdfHeight - (margin * 2);

  html2canvas(document.getElementById('pdfArea'), {scale: 2}).then(imgCan => {
    const imgData = imgCan.toDataURL('image/jpeg', 0.7);

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

    pdf.addImage(imgData, 'JPEG', xOffset, yOffset, imgWidthMM, imgHeightMM);
    pdf.save('layout.pdf');
  });
});

multiDesign.addEventListener('change', () => {
  sheetsInput.value = '';
  setupDesignInputs();
});

// [FIX START] Event listener for design count input
// Add logic to clear sheetsInput to ensure a fresh calculation
designCount.addEventListener('input', () => {
    sheetsInput.value = '';
    setupDesignInputs();
});
// [FIX END]

orderType.addEventListener('change', render);

window.addEventListener('DOMContentLoaded', () => {
  setupDesignInputs();
});