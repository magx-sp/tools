// ===== Core logic =====
function findOptimalLayout(reqQuantities, totalFacesPerSheet) {
  if (!reqQuantities || reqQuantities.length === 0 || reqQuantities.every(q => q === 0)) {
    return { layout: reqQuantities.map(() => 0), plates: 0, actualQuantities: reqQuantities.map(() => 0), overQuantities: reqQuantities.map(() => 0), overRates: reqQuantities.map(() => 0) };
  }
  const n = reqQuantities.length;
  const activeCount = reqQuantities.filter(q => q > 0).length;
  if (totalFacesPerSheet <= 0 || activeCount > totalFacesPerSheet) return null;

  const maxQ = Math.max(...reqQuantities);
  let lo = 1, hi = Math.max(1, maxQ), bestP = null, bestLayout = null;

  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    const faces = reqQuantities.map(q => (q > 0 ? Math.ceil(q / mid) : 0));
    if (faces.reduce((a, b) => a + b, 0) <= totalFacesPerSheet) {
      bestP = mid; bestLayout = faces; hi = mid - 1;
    } else {
      lo = mid + 1;
    }
  }

  if (bestP == null) return null;
  const actualQuantities = bestLayout.map(f => f * bestP);
  return {
    layout: bestLayout,
    plates: bestP,
    actualQuantities,
    overQuantities: actualQuantities.map((act, i) => act - reqQuantities[i]),
    overRates: actualQuantities.map((act, i) => reqQuantities[i] > 0 ? ((act - reqQuantities[i]) / reqQuantities[i]) * 100 : 0)
  };
}

function optimizeMagnetSplit(req, totalFacesPerSheet) {
  const n = req.length;
  const activeIdx = Array.from({length:n}, (_,i)=>i).filter(i => req[i] > 0);
  if (activeIdx.length === 0) return { mode: 'single', base: findOptimalLayout(req, totalFacesPerSheet) };
  if (totalFacesPerSheet % 2 !== 0) return { error: '半裁最適化は面付数が偶数である必要があります。' };
  
  const H = totalFacesPerSheet / 2;
  if (activeIdx.length > totalFacesPerSheet) return { error: 'デザイン数が面付数を上回っています。' };

  const order = activeIdx.slice().sort((a,b) => req[b]-req[a]);
  let group = new Array(n).fill('B');
  let aList = [], bList = [];
  
  function estPlates(list, nextIdx) {
    const count = list.length + (nextIdx != null ? 1 : 0);
    if (count > H) return Infinity;
    const facesEach = Math.max(1, Math.floor(H / count));
    let maxSheets = 0;
    const all = list.concat(nextIdx != null ? [nextIdx] : []);
    for (const i of all) { maxSheets = Math.max(maxSheets, Math.ceil(req[i] / facesEach)); }
    return maxSheets;
  }
  
  for (const i of order) {
    const aEst = estPlates(aList, i), bEst = estPlates(bList, i);
    if (aEst < Infinity && aEst <= bEst) { aList.push(i); group[i] = 'A'; } 
    else if (bEst < Infinity) { bList.push(i); group[i] = 'B'; } 
    else { aList.push(i); group[i] = 'A'; }
  }

  function costFromGroup(g) {
    const aReq = new Array(n).fill(0), bReq = new Array(n).fill(0);
    for (let i=0;i<n;i++) {
      if (req[i] <= 0) continue;
      if (g[i] === 'A') aReq[i] = req[i]; else bReq[i] = req[i];
    }
    const aRes = findOptimalLayout(aReq, H), bRes = findOptimalLayout(bReq, H);
    if (!aRes || !bRes) return { magnets: Infinity, sheets: Infinity, aRes, bRes };
    const sheets = Math.max(aRes.plates, bRes.plates);
    return { aRes, bRes, platesA: aRes.plates, platesB: bRes.plates, sheets, magnets: aRes.plates + bRes.plates };
  }

  let best = costFromGroup(group), improved = true, guard = 0;
  while (improved && guard < 100) {
    improved = false; guard++;
    for (const i of activeIdx) {
      const old = group[i], nextG = (old === 'A') ? 'B' : 'A';
      if (nextG === 'A' && group.filter((g, idx) => req[idx] > 0 && g === 'A').length >= H) continue;
      if (nextG === 'B' && group.filter((g, idx) => req[idx] > 0 && g === 'B').length >= H) continue;
      group[i] = nextG;
      const cand = costFromGroup(group);
      if (cand.aRes && cand.bRes && (cand.magnets < best.magnets || (cand.magnets === best.magnets && cand.sheets < best.sheets))) {
        best = cand; improved = true;
      } else { group[i] = old; }
    }
  }

  if (!best.aRes || !best.bRes) return { error: '半裁割り当てエラーです。' };
  const detail = new Array(n).fill(null).map(()=>({side:'-', faces:0, actual:0, spare:0}));
  for (let i=0;i<n;i++) {
    if (req[i]>0 && group[i]==='A') detail[i] = { side:'A', faces: best.aRes.layout[i], actual: best.aRes.actualQuantities[i], spare: best.aRes.overQuantities[i] };
    if (req[i]>0 && group[i]==='B') detail[i] = { side:'B', faces: best.bRes.layout[i], actual: best.bRes.actualQuantities[i], spare: best.bRes.overQuantities[i] };
  }
  return { H, group, detail, ...best };
}

function partitionDesignsDP(activeIndices, reqArray, plateCount, facesPerPlate, isMagnet) {
  const n = activeIndices.length;
  const items = activeIndices.map(idx => ({ q: reqArray[idx], idx })).sort((a, b) => b.q - a.q);

  const dp = Array.from({ length: n + 1 }, () => new Array(plateCount + 1).fill(Infinity));
  const choice = Array.from({ length: n + 1 }, () => new Array(plateCount + 1).fill(0));
  const bestRes = Array.from({ length: n + 1 }, () => new Array(plateCount + 1).fill(null));

  dp[0][0] = 0;

  for (let p = 1; p <= plateCount; p++) {
    for (let i = 1; i <= n; i++) {
      for (let k = 1; k <= Math.min(i, facesPerPlate); k++) {
        const prev = dp[i - k][p - 1];
        if (prev !== Infinity) {
          const groupReqs = items.slice(i - k, i).map(x => x.q);
          let cost = Infinity;
          let res = null;

          if (!isMagnet) {
            res = findOptimalLayout(groupReqs, facesPerPlate);
            if (res) cost = res.plates;
          } else {
            if (groupReqs.length <= facesPerPlate) {
              res = optimizeMagnetSplit(groupReqs, facesPerPlate);
              if (res && !res.error) cost = res.sheets;
            }
          }

          if (cost !== Infinity && prev + cost < dp[i][p]) {
            dp[i][p] = prev + cost;
            choice[i][p] = k;
            bestRes[i][p] = res;
          }
        }
      }
    }
  }

  let bestP = -1;
  let minTotal = Infinity;
  for (let p = 1; p <= plateCount; p++) {
    if (dp[n][p] < minTotal) {
      minTotal = dp[n][p];
      bestP = p;
    }
  }

  if (bestP === -1) return null;

  const filledPlates = [];
  let currI = n;
  let currP = bestP;
  while (currI > 0 && currP > 0) {
    const k = choice[currI][currP];
    const pData = { designs: [], reqs: [], result: bestRes[currI][currP] };
    for (let j = currI - k; j < currI; j++) {
      pData.designs.push(items[j].idx);
      pData.reqs.push(items[j].q);
    }
    filledPlates.unshift(pData);
    currI -= k;
    currP -= 1;
  }
  while (filledPlates.length < plateCount) filledPlates.push({ designs: [], reqs: [], result: null });
  return { plates: filledPlates, totalSheets: minTotal };
}

// ===== Excel/CSV Helpers =====
function stripBOM(s){ return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }
function parseCSV(text) {
  text = stripBOM(String(text || ""));
  const rows = []; let cur = '', row = [], inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], next = text[i+1];
    if (inQuotes) {
      if (c === '"' && next === '"') { cur += '"'; i++; continue; }
      if (c === '"') { inQuotes = false; continue; }
      cur += c; continue;
    }
    if (c === '"') { inQuotes = true; continue; }
    if (c === ',') { row.push(cur); cur = ''; continue; }
    if (c === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    if (c === '\r') { if (next === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    cur += c;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows.filter(r => r.some(v => String(v).trim() !== ''));
}
function normalizeCsvRows(rows) {
  if (!rows || rows.length === 0) return [];
  const first = rows[0].map(v => String(v).trim().toLowerCase());
  const hasHeader = first.includes('name') || first.includes('デザイン名') || first.includes('qty') || first.includes('数量') || first.includes('count');
  return (hasHeader ? rows.slice(1) : rows).map((cols, idx) => {
    const name = (cols[0] ?? '').toString().trim() || `デザイン${idx+1}`;
    const qty = parseInt((cols[1] ?? '').toString().replace(/,/g, '').trim(), 10);
    return { name, qty: Number.isFinite(qty) ? qty : 0 };
  }).filter(r => r.name && Number.isFinite(r.qty));
}

function uint32LE(n){ const b=new Uint8Array(4); new DataView(b.buffer).setUint32(0,n,true); return b; }
function uint16LE(n){ const b=new Uint8Array(2); new DataView(b.buffer).setUint16(0,n,true); return b; }
function strToUint8(s){ return new TextEncoder().encode(s); }
function concatUint8(arrs){ const out=new Uint8Array(arrs.reduce((s,a)=>s+a.length,0)); let o=0; for(const a of arrs){ out.set(a,o); o+=a.length; } return out; }
function makeZip(files){
  const localParts=[], centralParts=[]; let offset=0;
  for(const f of files){
    const nameBytes = strToUint8(f.name), data = f.data;
    const lh = concatUint8([uint32LE(0x04034b50), uint16LE(20), uint16LE(0), uint16LE(0), uint16LE(0), uint16LE(0), uint32LE(0), uint32LE(data.length), uint32LE(data.length), uint16LE(nameBytes.length), uint16LE(0)]);
    localParts.push(lh, nameBytes, data);
    const ch = concatUint8([uint32LE(0x02014b50), uint16LE(20), uint16LE(20), uint16LE(0), uint16LE(0), uint16LE(0), uint16LE(0), uint32LE(0), uint32LE(data.length), uint32LE(data.length), uint16LE(nameBytes.length), uint16LE(0), uint16LE(0), uint16LE(0), uint16LE(0), uint32LE(0), uint32LE(offset), nameBytes]);
    centralParts.push(ch); offset += lh.length + nameBytes.length + data.length;
  }
  const ca = concatUint8(centralParts), la = concatUint8(localParts);
  const endCD = concatUint8([uint32LE(0x06054b50), uint16LE(0), uint16LE(0), uint16LE(files.length), uint16LE(files.length), uint32LE(ca.length), uint32LE(la.length), uint16LE(0)]);
  return new Blob([la, ca, endCD], {type:'application/zip'});
}
function makeXlsxFromRows(rows, sheetName='Sheet1'){
  function escXml(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function cell(v){ if(v===''||v==null) return '<c/>'; if(!isNaN(Number(v)) && v!=='') return `<c t="n"><v>${Number(v)}</v></c>`; return `<c t="inlineStr"><is><t>${escXml(v)}</t></is></c>`; }
  const sheetRows = rows.map(r => '<row>' + r.map(cell).join('') + '</row>').join('');
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`;
  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${escXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;
  return new Blob([makeZip([{name: "[Content_Types].xml", data: strToUint8(contentTypes)},{name: "_rels/.rels", data: strToUint8(rootRels)},{name: "xl/workbook.xml", data: strToUint8(wbXml)},{name: "xl/_rels/workbook.xml.rels", data: strToUint8(relsXml)},{name: "xl/worksheets/sheet1.xml", data: strToUint8(sheetXml)}])], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function saveXlsx(filename, rows){
  const blob = makeXlsxFromRows(rows, 'result'), url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
}

// ===== UI Elements & Events =====
const facesPerSheetEl = document.getElementById('facesPerSheet');
const plateCountEl = document.getElementById('plateCount');
const designCountEl = document.getElementById('designCount');
const designGrid = document.getElementById('designGrid');
const applyCountBtn = document.getElementById('applyCount');
const resultTable = document.getElementById('resultTable');
const tbody = resultTable.querySelector('tbody');
const resultHeader = document.getElementById('resultHeader');
const summaryCell = document.getElementById('summaryCell');
const magnetModeEl = document.getElementById('magnetMode');
const csvFileEl = document.getElementById('csvFile');
const dropzone = document.getElementById('dropzone');

function buildDesignInputs() {
  const n = Math.max(1, parseInt(designCountEl.value || '1', 10));
  designGrid.innerHTML = '<header>デザイン名</header><header>必要数量</header><header></header>';
  for (let i = 0; i < n; i++) {
    const name = document.createElement('input'); name.type = 'text'; name.value = 'デザイン' + (i+1); name.dataset.role = 'name'; name.addEventListener('change', runCalc);
    const qty = document.createElement('input'); qty.type = 'number'; qty.min = '0'; qty.step = '1'; qty.value = i===0?'1000':(i===1?'2000':'3000'); qty.dataset.role = 'qty'; qty.addEventListener('change', runCalc);
    const calcBtn = document.createElement('button'); calcBtn.textContent = '計算'; calcBtn.className = 'btn'; calcBtn.addEventListener('click', runCalc);
    designGrid.append(name, qty, calcBtn);
  }
}

function runCalc() {
  const totalFaces = Math.max(0, parseInt(facesPerSheetEl.value || '0', 10));
  const plateCount = Math.max(1, parseInt(plateCountEl.value || '1', 10));
  const req = Array.from(designGrid.querySelectorAll('input[data-role="qty"]')).map(el => Math.max(0, parseInt(el.value || '0', 10)));
  const names = Array.from(designGrid.querySelectorAll('input[data-role="name"]')).map(el => el.value || '');
  const n = req.length;
  
  if (totalFaces <= 0) { alert('面付数は1以上で入力してね。'); return; }
  const activeIdx = Array.from({length:n}, (_,i)=>i).filter(i => req[i] > 0);
  if (activeIdx.length > totalFaces * plateCount) { alert(`版数が足りないよ！設定だと最大 ${totalFaces * plateCount} 種までしか置けません。版数を増やしてね。`); return; }

  const dpResult = partitionDesignsDP(activeIdx, req, plateCount, totalFaces, magnetModeEl.checked);
  if (!dpResult) { alert('指定の版数・面付数では全デザインを最適に配置できませんでした。版数か面付数を増やしてみてね。'); return; }

  // 空き枠がある場合、予備の割合が低く必要数量が多いデザインで枠を埋める
  for (let p = 0; p < plateCount; p++) {
    const plate = dpResult.plates[p];
    if (plate.designs.length === 0 || !plate.result) continue;

    if (!magnetModeEl.checked) {
      const res = plate.result;
      const currentFaces = res.layout.reduce((a, b) => a + b, 0);
      let rem = totalFaces - currentFaces;
      
      while (rem > 0) {
        let bestIdx = -1, minRate = Infinity, maxQ = -1;
        for (let j = 0; j < plate.reqs.length; j++) {
          const q = plate.reqs[j]; if (q <= 0) continue;
          const actual = res.layout[j] * res.plates;
          const rate = (actual - q) / q;
          if (rate < minRate || (rate === minRate && q > maxQ)) { minRate = rate; maxQ = q; bestIdx = j; }
        }
        if (bestIdx !== -1) {
          res.layout[bestIdx]++;
          res.actualQuantities[bestIdx] = res.layout[bestIdx] * res.plates;
          res.overQuantities[bestIdx] = res.actualQuantities[bestIdx] - plate.reqs[bestIdx];
          rem--;
        } else { break; }
      }
    } else {
      const res = plate.result;
      const H = totalFaces / 2;
      for (const side of ['A', 'B']) {
        const platesSide = (side === 'A') ? res.platesA : res.platesB;
        let currentFaces = 0;
        for (let j = 0; j < plate.reqs.length; j++) {
          if (res.detail[j].side === side) currentFaces += res.detail[j].faces;
        }
        let rem = H - currentFaces;
        
        while (rem > 0) {
          let bestIdx = -1, minRate = Infinity, maxQ = -1;
          for (let j = 0; j < plate.reqs.length; j++) {
            if (res.detail[j].side !== side) continue;
            const q = plate.reqs[j]; if (q <= 0) continue;
            const actual = res.detail[j].faces * platesSide;
            const rate = Math.max(0, actual - q) / q;
            if (rate < minRate || (rate === minRate && q > maxQ)) { minRate = rate; maxQ = q; bestIdx = j; }
          }
          if (bestIdx !== -1) {
            res.detail[bestIdx].faces++;
            res.detail[bestIdx].actual = res.detail[bestIdx].faces * platesSide;
            res.detail[bestIdx].spare = res.detail[bestIdx].actual - plate.reqs[bestIdx];
            rem--;
          } else { break; }
        }
      }
    }
  }

  const assignedPlates = new Array(n).fill(null);
  const allRes = { layout: new Array(n).fill(0), actual: new Array(n).fill(0), spare: new Array(n).fill(0), sides: new Array(n).fill('-') };
  let totalPlatesNeeded = 0;
  let totalMagnets = 0;
  let plateBreakdown = [];

  for (let p = 0; p < plateCount; p++) {
    const plate = dpResult.plates[p];
    if (plate.designs.length === 0) continue;

    plate.designs.forEach((globalIdx, localIdx) => {
      assignedPlates[globalIdx] = p + 1;
      if (!magnetModeEl.checked) {
        const res = plate.result;
        allRes.layout[globalIdx] = res.layout[localIdx];
        allRes.actual[globalIdx] = res.actualQuantities[localIdx];
        allRes.spare[globalIdx] = res.overQuantities[localIdx];
      } else {
        const res = plate.result;
        const d = res.detail[localIdx];
        allRes.sides[globalIdx] = d.side;
        allRes.layout[globalIdx] = d.faces;
        allRes.actual[globalIdx] = d.actual;
        allRes.spare[globalIdx] = d.spare;
      }
    });

    if (!magnetModeEl.checked) {
      totalPlatesNeeded += plate.result.plates;
      plateBreakdown.push(`版${p+1}: ${plate.result.plates.toLocaleString()}枚`);
    } else {
      totalPlatesNeeded += plate.result.sheets;
      totalMagnets += plate.result.magnets;
      plateBreakdown.push(`版${p+1}: ${plate.result.sheets.toLocaleString()}枚`);
    }
  }

  tbody.innerHTML = '';

  if (!magnetModeEl.checked) {
    resultHeader.innerHTML = '<th>デザイン名</th><th>版</th><th>面付数（編集可）</th><th>必要数量</th><th>総印刷数</th><th>予備枚数</th>';
    for (let i = 0; i < n; i++) {
      if (!assignedPlates[i]) continue;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td style="text-align:left;">${names[i]||`デザイン${i+1}`}</td><td style="text-align:center;">版${assignedPlates[i]}</td><td><input type="number" class="faces-edit" min="0" step="1" value="${allRes.layout[i]}"></td><td>${req[i].toLocaleString()}</td><td data-col="actual">${allRes.actual[i].toLocaleString()}</td><td data-col="spare">${allRes.spare[i].toLocaleString()}</td>`;
      tbody.appendChild(tr);
    }
    summaryCell.innerHTML = `【刷了原紙】 ${plateBreakdown.join(' ＋ ')} ＝ <strong style="font-size:1.1em; color:#1971ff;">総計: ${totalPlatesNeeded.toLocaleString()} 枚</strong>`;
    resultTable.dataset.mode = 'single';
    resultTable._state = { mode:'single', totalFaces, req, names, assignedPlates };
    tbody.querySelectorAll('input.faces-edit').forEach(inp=> inp.addEventListener('input', ()=> recomputeSingleFromInputs(resultTable._state)));
    
  } else {
    const H = totalFaces / 2;
    resultHeader.innerHTML = '<th>デザイン名</th><th>版</th><th>側</th><th>片側面付（編集可）</th><th>必要数量</th><th>総印刷数（片側）</th><th>予備枚数（片側）</th>';
    for (let i = 0; i < n; i++) {
      if (!assignedPlates[i]) continue;
      const tr = document.createElement('tr');
      tr.innerHTML = `<td style="text-align:left;">${names[i]||`デザイン${i+1}`}</td><td style="text-align:center;">版${assignedPlates[i]}</td><td style="text-align:center;">${allRes.sides[i]}</td><td><input type="number" class="faces-edit" min="0" step="1" value="${allRes.layout[i]}"></td><td>${req[i].toLocaleString()}</td><td data-col="actual">${allRes.actual[i].toLocaleString()}</td><td data-col="spare">${allRes.spare[i].toLocaleString()}</td>`;
      tbody.appendChild(tr);
    }
    summaryCell.innerHTML = `片側面付: ${H}面<br>【刷了原紙】 ${plateBreakdown.join(' ＋ ')} ＝ <strong style="font-size:1.1em; color:#1971ff;">総計: ${totalPlatesNeeded.toLocaleString()} 枚</strong><br>総マグネット: ${totalMagnets.toLocaleString()} 枚`;
    resultTable.dataset.mode = 'magnet';
    resultTable._state = { mode:'magnet', H, req, names, assignedPlates, sides: allRes.sides };
    tbody.querySelectorAll('input.faces-edit').forEach(inp=> inp.addEventListener('input', ()=> recomputeMagnetFromInputs(resultTable._state)));
  }
  resultTable.style.display = 'table';
}

function recomputeSingleFromInputs(state) {
  const {totalFaces, req, assignedPlates} = state, rows = Array.from(tbody.querySelectorAll('tr'));
  const faces = rows.map(tr => Math.max(0, parseInt(tr.querySelector('input.faces-edit').value || '0', 10)));
  const plateFaces = {}, plateSheets = {};
  assignedPlates.forEach(p => { if(p){ plateFaces[p]=0; plateSheets[p]=0; } });
  
  let rowIdx = 0;
  for (let i=0; i<req.length; i++) {
    const p = assignedPlates[i];
    if (!p) continue;
    plateFaces[p] += faces[rowIdx];
    if (faces[rowIdx] > 0) plateSheets[p] = Math.max(plateSheets[p], Math.ceil(req[i]/faces[rowIdx]));
    else if (req[i] > 0) plateSheets[p] = Infinity;
    rowIdx++;
  }
  
  let errorMsg = '', total = 0, breakdown = [];
  for (let p in plateFaces) {
    if (plateFaces[p] > totalFaces) errorMsg += `版${p}の面付合計（${plateFaces[p]}）が超過してます。 `;
    if (!isFinite(plateSheets[p])) errorMsg = '面付数0のデザインがあります。';
    total += plateSheets[p] || 0;
    breakdown.push(`版${p}: ${(plateSheets[p]||0).toLocaleString()}枚`);
  }
  
  rowIdx = 0;
  for (let i=0; i<req.length; i++) {
    const p = assignedPlates[i];
    if (!p) continue;
    const actual = faces[rowIdx] * plateSheets[p], spare = Math.max(0, actual - req[i]);
    rows[rowIdx].querySelector('[data-col="actual"]').textContent = isFinite(actual) ? actual.toLocaleString() : '—';
    rows[rowIdx].querySelector('[data-col="spare"]').textContent = isFinite(spare) ? spare.toLocaleString() : '—';
    rowIdx++;
  }
  summaryCell.style.color = errorMsg ? '#c00' : '';
  summaryCell.innerHTML = errorMsg || `【刷了原紙】 ${breakdown.join(' ＋ ')} ＝ <strong style="font-size:1.1em; color:#1971ff;">総計: ${total.toLocaleString()} 枚</strong>`;
}

function recomputeMagnetFromInputs(state) {
  const {H, req, assignedPlates, sides} = state, rows = Array.from(tbody.querySelectorAll('tr'));
  const faces = rows.map(tr => Math.max(0, parseInt(tr.querySelector('input.faces-edit').value || '0', 10)));
  const pFA = {}, pFB = {}, pSA = {}, pSB = {};
  assignedPlates.forEach(p => { if(p){ pFA[p]=0; pFB[p]=0; pSA[p]=0; pSB[p]=0; } });
  
  let rowIdx = 0;
  for (let i=0; i<req.length; i++) {
    const p = assignedPlates[i];
    if (!p) continue;
    if(sides[i]==='A') { pFA[p] += faces[rowIdx]; pSA[p] = Math.max(pSA[p], faces[rowIdx]>0 ? Math.ceil(req[i]/faces[rowIdx]) : Infinity); }
    if(sides[i]==='B') { pFB[p] += faces[rowIdx]; pSB[p] = Math.max(pSB[p], faces[rowIdx]>0 ? Math.ceil(req[i]/faces[rowIdx]) : Infinity); }
    rowIdx++;
  }
  
  let warn = [], tS = 0, tM = 0, breakdown = [];
  for (let p in pFA) {
    if(pFA[p]>H) warn.push(`版${p}A側が超過`);
    if(pFB[p]>H) warn.push(`版${p}B側が超過`);
    const s = Math.max(pSA[p]||0, pSB[p]||0);
    tS += s;
    tM += (pSA[p]||0) + (pSB[p]||0);
    breakdown.push(`版${p}: ${s.toLocaleString()}枚`);
  }
  
  rowIdx = 0;
  for (let i=0; i<req.length; i++) {
    const p = assignedPlates[i];
    if (!p) continue;
    const actual = faces[rowIdx] * (sides[i]==='A' ? pSA[p] : pSB[p]), spare = Math.max(0, actual - req[i]);
    rows[rowIdx].querySelector('[data-col="actual"]').textContent = isFinite(actual) ? actual.toLocaleString() : '—';
    rows[rowIdx].querySelector('[data-col="spare"]').textContent = isFinite(spare) ? spare.toLocaleString() : '—';
    rowIdx++;
  }
  summaryCell.style.color = warn.length ? '#c00' : '';
  const m = `【刷了原紙】 ${breakdown.join(' ＋ ')} ＝ <strong style="font-size:1.1em; color:#1971ff;">総計: ${isFinite(tS)?tS.toLocaleString():'—'} 枚</strong><br>総マグネット: ${isFinite(tM)?tM.toLocaleString():'—'} 枚`;
  summaryCell.innerHTML = warn.length ? (warn.join(' / ') + '<br>' + m) : m;
}

// ===== CSV Load Handler =====
async function handleCsvFile(file) {
  if (!file) return;
  const text = await file.text();
  const items = normalizeCsvRows(parseCSV(text));
  if (!items.length) {
    alert('CSVの内容が読み取れませんでした。列は「name,qty」形式（ヘッダー有無どちらでも可）で用意してね。');
    return;
  }
  
  designCountEl.value = items.length;
  const totalFaces = Math.max(1, parseInt(facesPerSheetEl.value || '1', 10));
  plateCountEl.value = Math.max(1, Math.ceil(items.length / totalFaces));
  
  buildDesignInputs();
  const nameInputs = Array.from(designGrid.querySelectorAll('input[data-role="name"]'));
  const qtyInputs = Array.from(designGrid.querySelectorAll('input[data-role="qty"]'));
  
  for (let i = 0; i < items.length; i++) {
    if (nameInputs[i]) nameInputs[i].value = items[i].name;
    if (qtyInputs[i]) qtyInputs[i].value = items[i].qty;
  }
  runCalc();
}

// Events
document.getElementById('pickCsv').addEventListener('click', () => { csvFileEl.value = ''; csvFileEl.click(); });
csvFileEl.addEventListener('change', (e) => handleCsvFile(e.target.files[0]));

['dragenter','dragover'].forEach(ev => dropzone.addEventListener(ev, (e) => { 
  e.preventDefault(); e.stopPropagation(); dropzone.classList.add('dragover'); 
}));
['dragleave','drop'].forEach(ev => dropzone.addEventListener(ev, (e) => { 
  e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('dragover'); 
}));
dropzone.addEventListener('drop', (e) => { 
  const file = e.dataTransfer.files && e.dataTransfer.files[0]; 
  if (file) handleCsvFile(file); 
});

applyCountBtn.addEventListener('click', buildDesignInputs);
[magnetModeEl, facesPerSheetEl, plateCountEl].forEach(el => el.addEventListener('change', runCalc));

document.getElementById('exportXlsx').addEventListener('click', () => {
  if (resultTable.style.display === 'none') { alert('先に計算してね。'); return; }
  const rows = [Array.from(resultHeader.children).map(th => th.textContent.trim())];
  rows.push(...Array.from(tbody.querySelectorAll('tr')).map(tr => Array.from(tr.children).map(td => td.querySelector('input') ? td.querySelector('input').value : td.textContent.trim())));
  rows.push([], ['Summary', summaryCell.innerText.replace(/\n/g, '  |  ')]);
  saveXlsx(`result_${(resultTable.dataset.mode||'single')}_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xlsx`, rows);
});

buildDesignInputs();