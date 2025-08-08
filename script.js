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
const multiDesign = document.getElementById('multiDesign');
const manualLayout = document.getElementById('manualLayout');
const designCount = document.getElementById('designCount');
const designList = document.getElementById('designList');
const designInputs = document.getElementById('designInputs');
const marginsEl = document.getElementById('margins');
const resultBody = document.getElementById('resultBody');
const downloadSvgBtn = document.getElementById('downloadSvgBtn');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');

// ニュースバーの閉じる（イベント委任で確実に反応）
document.addEventListener('click', (e) => {
  const btn = e.target.closest('.news-close');
  if (!btn) return;
  const bar = document.getElementById('newsBar');
  if (bar) bar.style.display = 'none';
});

// 計算結果を保持するためのグローバル変数
let currentLayoutResults = null;

// エラーメッセージを表示するヘルパー関数
function displayErrorMessage(message) {
    resultBody.innerHTML = `<tr><td colspan="2" style="color: red; font-weight: bold;">${message}</td></tr>`;
    sheetsInput.value = '0';
    if (multiDesign.checked) {
        const n = parseInt(designCount.value, 10) || 1;
        for (let i = 1; i <= n; i++) {
            const faceInput = document.getElementById(`df${i}`);
            if(faceInput) faceInput.value = 0;
        }
    }
    currentLayoutResults = null;
    render();
}

function findOptimalLayout(reqQuantities, totalFacesPerSheet) {
  // すべて0なら即返す
  if (!reqQuantities || reqQuantities.length === 0 || reqQuantities.every(q => q === 0)) {
    return {
      layout: reqQuantities.map(() => 0),
      plates: 0,
      actualQuantities: reqQuantities.map(() => 0),
      overQuantities: reqQuantities.map(() => 0),
      overRates: reqQuantities.map(() => 0),
      variance: 0
    };
  }

  const n = reqQuantities.length;
  const active = reqQuantities.map(q => q > 0);
  const activeCount = active.filter(Boolean).length;

  if (totalFacesPerSheet <= 0) return null;
  if (activeCount > totalFacesPerSheet) return null; // 1面ずつすら置けない

  // 二分探索で原紙枚数 p を最小化
  const maxQ = Math.max(...reqQuantities);
  let lo = 1, hi = Math.max(1, maxQ), bestP = null, bestLayout = null;

  const feasible = (p) => {
    // 各デザインに必要な面数 f_i = ceil(q_i / p)
    const faces = reqQuantities.map(q => (q > 0 ? Math.ceil(q / p) : 0));
    const sumFaces = faces.reduce((a, b) => a + b, 0);
    return { ok: sumFaces <= totalFacesPerSheet, faces };
  };

  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    const { ok, faces } = feasible(mid);
    if (ok) {
      bestP = mid;
      bestLayout = faces;
      hi = mid - 1;          // さらに小さい枚数を探す
    } else {
      lo = mid + 1;          // もっと枚数が必要
    }
  }

  if (bestP == null) return null;

  // 余ったセルは未使用（-1 埋めで点線表示）にしてムダ刷りを増やさない
  const plates = bestP;
  const actualQuantities = bestLayout.map(f => f * plates);
  const overQuantities = actualQuantities.map((act, i) => act - reqQuantities[i]);
  const overRates = overQuantities.map((ov, i) => reqQuantities[i] > 0 ? (ov / reqQuantities[i]) * 100 : 0);
  const validOver = overRates.filter((_, i) => reqQuantities[i] > 0);
  const mean = validOver.length ? validOver.reduce((a, b) => a + b, 0) / validOver.length : 0;
  const variance = validOver.length ? validOver.reduce((s, r) => s + Math.pow(r - mean, 2), 0) / validOver.length : 0;

  return {
    layout: bestLayout,
    plates,
    actualQuantities,
    overQuantities,
    overRates,
    variance
  };
}

function setupDesignInputs() {
  designInputs.style.display = multiDesign.checked ? 'block' : 'none';
  designList.innerHTML = '';
  const n = parseInt(designCount.value, 10) || 1;
  const isManual = manualLayout.checked;

  for (let i = 1; i <= n; i++) {
    const div = document.createElement('div');
    div.className = 'pair';
    // ▼▼▼ 表示崩れ対策としてラベルを2行にし、高さを揃える ▼▼▼
    div.innerHTML = `
        <label>|デザイン${i}<br>印刷数量
          <input type="number" id="dq${i}" value="0" min="0">
        </label>
        <label>面付数<br>&nbsp;
          <input type="number" id="df${i}" min="0" ${!isManual ? 'disabled' : ''}>
        </label>`;
    // ▲▲▲ 修正ここまで ▲▲▲
    designList.appendChild(div);
    
    document.getElementById(`dq${i}`).addEventListener('input', () => {
        if (!isManual) sheetsInput.value = '';
        calculate();
    });
    document.getElementById(`df${i}`).addEventListener('input', calculate);
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

// ▼▼▼ orderType引数を削除 ▼▼▼
function generateDesignSequence(layoutCounts, totalCells) {
  const sequence = [];
  
  if (!layoutCounts || layoutCounts.length === 0) {
    for (let i = 0; i < totalCells; i++) sequence.push(0);
    return sequence;
  }

  const designAssignments = [];
  for (let designIndex = 0; designIndex < layoutCounts.length; designIndex++) {
    for (let count = 0; count < layoutCounts[designIndex]; count++) {
      designAssignments.push(designIndex);
    }
  }
  
  while (designAssignments.length < totalCells) designAssignments.push(-1); // 余白は-1
  if (designAssignments.length > totalCells) designAssignments.length = totalCells;

  return designAssignments;
}
// ▲▲▲ 修正ここまで ▲▲▲


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
  
  let effectiveSeq = [];
  if (multiDesign.checked && currentLayoutResults && currentLayoutResults.layout) {
    // ▼▼▼ orderType.valueを削除 ▼▼▼
    effectiveSeq = generateDesignSequence(currentLayoutResults.layout, L.cx * L.cy);
  } else {
    effectiveSeq = generateDesignSequence(null, L.cx * L.cy);
  }

  // ▼▼▼ 描画ロジックを修正 ▼▼▼
  for (let i = 0; i < L.cy; i++) {
    for (let j = 0; j < L.cx; j++) {
      const cellIndex = i * L.cx + j;
      const designIndex = effectiveSeq[cellIndex];
      
      // ▼▼▼ 面付数が最大に満たない場合、余白を点線で描画 ▼▼▼
      if (designIndex === -1) {
          const r = document.createElementNS(ns, 'rect');
          r.setAttribute('x', startX + j * L.tw);
          r.setAttribute('y', startY + i * L.th);
          r.setAttribute('width', L.w);
          r.setAttribute('height', L.h);
          r.setAttribute('fill', 'none');
          r.setAttribute('stroke', '#ccc');
          r.setAttribute('stroke-width', '1');
          r.setAttribute('stroke-dasharray', '4 2');
          canvas.appendChild(r);
          continue;
      }
      
      const r = document.createElementNS(ns, 'rect');
      r.setAttribute('x', startX + j * L.tw);
      r.setAttribute('y', startY + i * L.th);
      r.setAttribute('width', L.w);
      r.setAttribute('height', L.h);
      r.setAttribute('fill', multiDesign.checked ? cols[designIndex % cols.length] : '#add8e6');
      r.setAttribute('stroke', '#666');
      canvas.appendChild(r);
      
      // ▼▼▼ 複数デザインの時だけ番号を描画 ▼▼▼
      if (multiDesign.checked) {
        const t = document.createElementNS(ns, 'text');
        t.setAttribute('x', startX + j * L.tw + L.w / 2);
        t.setAttribute('y', startY + i * L.th + L.h / 2);
        t.setAttribute('text-anchor', 'middle');
        t.setAttribute('dominant-baseline', 'middle');
        t.setAttribute('font-size', Math.min(L.w, L.h) / 4);
        t.textContent = designIndex + 1;
        canvas.appendChild(t);
      }
    }
  }
  // ▲▲▲ 描画ロジック修正ここまで ▲▲▲
  
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
    const isManual = manualLayout.checked;
    const req = [];
    const n = parseInt(designCount.value, 10) || 1;
    for (let i = 1; i <= n; i++) {
      const element = document.getElementById(`dq${i}`);
      req.push(element ? (+element.value || 0) : 0);
    }
    
    let result = null;

    if (isManual) {
        // --- 手動モード ---
        const manualLayoutCounts = [];
        let layoutSum = 0;
        for (let i = 1; i <= n; i++) {
            const faceEl = document.getElementById(`df${i}`);
            const faces = faceEl ? (+faceEl.value || 0) : 0;
            manualLayoutCounts.push(faces);
            layoutSum += faces;
        }

        // ▼▼▼ エラー発生時、描画を更新せず処理を中断 ▼▼▼
        if (layoutSum > L.total) {
            resultBody.innerHTML = `<tr><td colspan="2" style="color: red; font-weight: bold;">エラー: 面付数の合計 (${layoutSum}面) が、用紙全体の面付数 (${L.total}面) を超えています。</td></tr>`;
            currentLayoutResults = null; // 結果をリセット
            return;
        }
        // ▲▲▲ 修正ここまで ▲▲▲

        const sheetsPerDesign = req.map((qty, i) => 
            (manualLayoutCounts[i] > 0) ? Math.ceil(qty / manualLayoutCounts[i]) : 0
        );
        const sheets = Math.max(0, ...sheetsPerDesign.filter(s => isFinite(s)));

        if (sheetsInput.value === '' || isNaN(parseInt(sheetsInput.value, 10)) || parseInt(sheetsInput.value, 10) < 0 || sheetsInput.value != sheets) {
            sheetsInput.value = sheets;
        }

        const finalSheets = +sheetsInput.value || 0;
        const actualQuantities = manualLayoutCounts.map(faces => faces * finalSheets);
        const overQuantities = actualQuantities.map((actual, i) => actual - req[i]);
        const overRates = overQuantities.map((over, i) => req[i] > 0 ? (over / req[i]) * 100 : 0);
        const validOverRates = overRates.filter((rate, i) => req[i] > 0);
        const mean = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + rate, 0) / validOverRates.length : 0;
        const variance = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + Math.pow(rate - mean, 2), 0) / validOverRates.length : 0;

        result = {
            layout: manualLayoutCounts,
            plates: finalSheets,
            actualQuantities,
            overQuantities,
            overRates,
            variance
        };

    } else {
        // --- 自動モード ---
        if (L.total < req.length && req.some(q => q > 0)) {
            displayErrorMessage(`エラー: 面付数 (${L.total}面) がデザイン数 (${req.length}個) より少ないため、各デザインに1面ずつ割り当てる事ができません。`);
            return;
        }

        const best = findOptimalLayout(req, L.total);
        result = best;

        if (best) {
            if (sheetsInput.value === '' || isNaN(parseInt(sheetsInput.value, 10)) || parseInt(sheetsInput.value, 10) < 0) {
                sheetsInput.value = best.plates;
            }
            
            best.layout.forEach((faces, i) => {
                const faceInput = document.getElementById(`df${i+1}`);
                if (faceInput) faceInput.value = faces;
            });

            const finalSheets = +sheetsInput.value || 0;
            if (finalSheets !== best.plates) {
                result.plates = finalSheets;
                result.actualQuantities = result.layout.map(faces => faces * finalSheets);
                result.overQuantities = result.actualQuantities.map((actual, i) => actual - req[i]);
                result.overRates = result.overQuantities.map((over, i) => req[i] > 0 ? (over / req[i]) * 100 : 0);
                const validOverRates = result.overRates.filter((rate, i) => req[i] > 0);
                const mean = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + rate, 0) / validOverRates.length : 0;
                result.variance = validOverRates.length > 0 ? validOverRates.reduce((sum, rate) => sum + Math.pow(rate - mean, 2), 0) / validOverRates.length : 0;
            }
        } else {
            req.forEach((_, i) => {
                const faceInput = document.getElementById(`df${i+1}`);
                if (faceInput) faceInput.value = 0;
            });
        }
    }
    
    currentLayoutResults = result;

    tableRows += `<tr><td>原紙寸法</td><td>${L.sw}×${L.sh}mm</td></tr>`;
    tableRows += `<tr><td>製品サイズ</td><td>${L.w}×${L.h}mm</td></tr>`;

    if (result) {
        tableRows += `<tr><td>面付数</td><td>`;
        req.forEach((_, i) => {
            const f = result.layout[i] || 0;
            tableRows += `<div>デザイン${i + 1}／${f}面</div>`;
        });
        tableRows += `</td></tr>`;

        tableRows += `<tr><td>印刷数量</td><td>`;
        req.forEach((qty, i) => {
            tableRows += `<div>デザイン${i + 1}／${qty.toLocaleString()}枚</div>`;
        });
        tableRows += `</td></tr>`;

        tableRows += `<tr><td>必要原紙枚数</td><td>${result.plates.toLocaleString()}枚</td></tr>`;
        tableRows += `<tr><td>予備</td><td>`;
        req.forEach((_, i) => {
            const spare = result.overQuantities[i];
            const overRate = result.overRates[i];
            tableRows += `<div>デザイン${i + 1}／${spare.toLocaleString()}枚（予備率：${overRate.toFixed(2)}%）</div>`;
        });
        tableRows += `<div>予備率分散：${result.variance.toFixed(4)}</div>`;
        tableRows += `</td></tr>`;
    } else {
         tableRows += `<tr><td colspan="2">計算できません</td></tr>`;
    }
    
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
    tableRows += `<tr><td>印刷数量</td><td>${qty.toLocaleString()}枚</td></tr>`;
    tableRows += `<tr><td>必要原紙枚数</td><td>${sheets.toLocaleString()}枚</td></tr>`;
    tableRows += `<tr><td>予備</td><td>${res.toLocaleString()}枚（予備率：${overRate.toFixed(2)}%）</td></tr>`;
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

manualLayout.addEventListener('change', () => {
  const isManual = manualLayout.checked;
  const n = parseInt(designCount.value, 10) || 1;
  for (let i = 1; i <= n; i++) {
      const faceInput = document.getElementById(`df${i}`);
      if (faceInput) {
        faceInput.disabled = !isManual;
      }
  }
  if (!isManual) {
    sheetsInput.value = ''; 
  }
  calculate();
});

designCount.addEventListener('input', () => {
    if (!manualLayout.checked) sheetsInput.value = '';
    setupDesignInputs();
});

// ▼▼▼ orderTypeのイベントリスナーを削除 ▼▼▼
// ▲▲▲ 削除 ▲▲▲

// --- ニュース表示ロジック ---
function setNews(items) {
  const bar = document.getElementById('newsBar');
  if (!bar) return;

  // items: [{ date: 'YYYY/MM/DD', text: 'コメント', href?: 'リンク' }]
  if (!Array.isArray(items) || items.length === 0) {
    bar.style.display = 'none';
    return;
  }

  // フォールバック（noscript相当）を消して作り直し
  bar.innerHTML = '';

  // 中身生成（複数件あれば「・」区切りで並べます）
  const span = document.createElement('span');
  span.innerHTML = items.map(it => {
    const label = `${it.date}　${it.text}`;
    return it.href ? `<a href="${it.href}" target="_blank" rel="noopener noreferrer">${label}</a>` : label;
  }).join('　・　');
  bar.appendChild(span);

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'news-close';
  closeBtn.setAttribute('aria-label', 'このお知らせを閉じる');
  closeBtn.textContent = '×';
  closeBtn.addEventListener('click', () => { bar.style.display = 'none'; });
  bar.appendChild(closeBtn);
}
// --- /ニュース表示ロジック ---


window.addEventListener('DOMContentLoaded', () => {
  setupDesignInputs();
});