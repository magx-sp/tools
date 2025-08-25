// ===== Core logic (Updated) =====
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

  // 余ったセルは未使用（-1 埋め）
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

// 元のコードにあった半裁最適化ロジックはそのまま残します
function optimizeMagnetSplit(req, totalFacesPerSheet) {
  const n = req.length;
  const activeIdx = Array.from({length:n}, (_,i)=>i).filter(i => req[i] > 0);
  if (activeIdx.length === 0) { return { mode: 'single', base: findOptimalLayout(req, totalFacesPerSheet) }; }
  if (totalFacesPerSheet % 2 !== 0) { return { error: '面付数は偶数が必要です（半分に割り切れません）。' }; }
  const H = totalFacesPerSheet / 2;

  const order = activeIdx.slice().sort((a,b) => req[b]-req[a]);
  let group = new Array(n).fill('B');
  let aList = [], bList = [];
  function estPlates(list, nextIdx) {
    const count = list.length + (nextIdx != null ? 1 : 0);
    const facesEach = Math.max(1, Math.floor(H / count));
    let maxSheets = 0;
    const all = list.concat(nextIdx != null ? [nextIdx] : []);
    for (const i of all) {
      const sheets = Math.ceil(req[i] / facesEach);
      if (sheets > maxSheets) maxSheets = sheets;
    }
    return maxSheets;
  }
  for (const i of order) {
    const aEst = estPlates(aList, i);
    const bEst = estPlates(bList, i);
    if (aEst <= bEst) { aList.push(i); group[i] = 'A'; } else { bList.push(i); group[i] = 'B'; }
  }

  function costFromGroup(g) {
    const aReq = [], aMap = [];
    const bReq = [], bMap = [];
    for (let i=0;i<n;i++) {
      if (req[i] <= 0) continue;
      if (g[i] === 'A') { aMap.push(i); aReq.push(req[i]); }
      else { bMap.push(i); bReq.push(req[i]); }
    }
    const aRes = findOptimalLayout(aReq, H);
    const bRes = findOptimalLayout(bReq, H);
    const platesA = aRes ? aRes.plates : 0;
    const platesB = bRes ? bRes.plates : 0;
    const sheets = Math.max(platesA, platesB);
    const magnets = platesA + platesB;
    return { aRes, bRes, aMap, bMap, platesA, platesB, sheets, magnets };
  }

  let best = costFromGroup(group);
  let improved = true; let guard = 0;
  while (improved && guard < 100) {
    improved = false; guard++;
    for (const i of activeIdx) {
      const old = group[i];
      group[i] = (old === 'A') ? 'B' : 'A';
      const cand = costFromGroup(group);
      if (cand.magnets < best.magnets || (cand.magnets === best.magnets && cand.sheets < best.sheets)) {
        best = cand; improved = true;
      } else { group[i] = old; }
    }
  }

  const detail = new Array(n).fill(null).map(()=>({side:'-', faces:0, actual:0, spare:0, rate:0}));
  if (best.aRes) {
    for (let k=0;k<best.aMap.length;k++) {
      const i = best.aMap[k];
      detail[i] = { side:'A', faces: best.aRes.layout[k]||0, actual: best.aRes.actualQuantities[k]||0, spare: best.aRes.overQuantities[k]||0, rate: best.aRes.overRates[k]||0 };
    }
  }
  if (best.bRes) {
    for (let k=0;k<best.bMap.length;k++) {
      const i = best.bMap[k];
      detail[i] = { side:'B', faces: best.bRes.layout[k]||0, actual: best.bRes.actualQuantities[k]||0, spare: best.bRes.overQuantities[k]||0, rate: best.bRes.overRates[k]||0 };
    }
  }
  return { H, group, detail, ...best };
}

// ===== CSV helpers =====
function stripBOM(s){ return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }
function parseCSV(text) {
  text = stripBOM(String(text || ""));
  const rows = []; let cur = ''; let row = []; let inQuotes = false;
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
  const body = hasHeader ? rows.slice(1) : rows;
  return body.map((cols, idx) => {
    const name = (cols[0] ?? '').toString().trim() || `デザイン${idx+1}`;
    const qtyRaw = (cols[1] ?? '').toString().replace(/,/g, '').trim();
    const qty = Number.isFinite(parseInt(qtyRaw, 10)) ? parseInt(qtyRaw, 10) : 0;
    return { name, qty };
  }).filter(r => r.name && Number.isFinite(r.qty));
}

// ===== XLSX minimal builder (no compression) =====
function uint32LE(n){ const b=new Uint8Array(4); new DataView(b.buffer).setUint32(0,n,true); return b; }
function uint16LE(n){ const b=new Uint8Array(2); new DataView(b.buffer).setUint16(0,n,true); return b; }
function strToUint8(s){ const enc = new TextEncoder(); return enc.encode(s); }
function concatUint8(arrs){ let len=0; for(const a of arrs) len+=a.length; const out=new Uint8Array(len); let o=0; for(const a of arrs){ out.set(a,o); o+=a.length; } return out; }

function makeZip(files){
  // files: [{name:string, data:Uint8Array}]
  const localParts=[]; const centralParts=[];
  let offset=0;
  for(const f of files){
    const nameBytes = strToUint8(f.name);
    const data = f.data;
    const localHeader = concatUint8([
      uint32LE(0x04034b50), // local file header sig
      uint16LE(20),         // version needed
      uint16LE(0),          // flags
      uint16LE(0),          // compression (0=stored)
      uint16LE(0), uint16LE(0), // time/date (ignored)
      uint32LE(0),          // crc32 (0 OK for no compression w/ flag 0? we'll set 0; Excel tolerates)
      uint32LE(data.length),// comp size
      uint32LE(data.length),// uncomp size
      uint16LE(nameBytes.length),
      uint16LE(0)           // extra length
    ]);
    localParts.push(localHeader, nameBytes, data);
    // central directory
    const centralHeader = concatUint8([
      uint32LE(0x02014b50),
      uint16LE(20), uint16LE(20), // version made by / needed
      uint16LE(0), uint16LE(0),   // flags / compression
      uint16LE(0), uint16LE(0),   // time/date
      uint32LE(0),                // crc32
      uint32LE(data.length), uint32LE(data.length),
      uint16LE(nameBytes.length), uint16LE(0), uint16LE(0), // name/extra/comment
      uint16LE(0), uint16LE(0),   // disk/start attr
      uint32LE(0),                // ext attr
      uint32LE(offset),           // local header offset
      nameBytes
    ]);
    centralParts.push(centralHeader);
    offset += localHeader.length + nameBytes.length + data.length;
  }
  const centralAll = concatUint8(centralParts);
  const localAll = concatUint8(localParts);
  const endCD = concatUint8([
    uint32LE(0x06054b50),
    uint16LE(0), uint16LE(0), // disk numbers
    uint16LE(files.length), uint16LE(files.length),
    uint32LE(centralAll.length),
    uint32LE(localAll.length),
    uint16LE(0) // comment len
  ]);
  return new Blob([localAll, centralAll, endCD], {type:'application/zip'});
}

function makeXlsxFromRows(rows, sheetName='Sheet1'){
  // Build minimal OOXML
  function escXml(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function cell(v){ if(v===''||v==null) return '<c/>'; if(!isNaN(Number(v)) && v!=='') return `<c t="n"><v>${Number(v)}</v></c>`; return `<c t="inlineStr"><is><t>${escXml(v)}</t></is></c>`; }
  const sheetRows = rows.map(r => '<row>' + r.map(cell).join('') + '</row>').join('');
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetRows}</sheetData>
</worksheet>`;

  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${escXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

  const files = [
    {name: "[Content_Types].xml", data: strToUint8(contentTypes)},
    {name: "_rels/.rels", data: strToUint8(rootRels)},
    {name: "xl/workbook.xml", data: strToUint8(wbXml)},
    {name: "xl/_rels/workbook.xml.rels", data: strToUint8(relsXml)},
    {name: "xl/worksheets/sheet1.xml", data: strToUint8(sheetXml)},
  ];
  const zipBlob = makeZip(files);
  return new Blob([zipBlob], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

function saveXlsx(filename, rows){
  const blob = makeXlsxFromRows(rows, 'result');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
}

// ===== UI =====
const facesPerSheetEl = document.getElementById('facesPerSheet');
const designCountEl = document.getElementById('designCount');
const designGrid = document.getElementById('designGrid');
const designCountBadge = document.getElementById('designCountBadge');
const applyCountBtn = document.getElementById('applyCount');
const resultTable = document.getElementById('resultTable');
const tbody = resultTable.querySelector('tbody');
const resultHeader = document.getElementById('resultHeader');
const summaryCell = document.getElementById('summaryCell');
const pickCsvBtn = document.getElementById('pickCsv');
const csvFileEl = document.getElementById('csvFile');
const dropzone = document.getElementById('dropzone');
const sampleBtn = document.getElementById('sampleCsv');
const magnetModeEl = document.getElementById('magnetMode');
const exportXlsxBtn = document.getElementById('exportXlsx');

function buildDesignInputs() {
  const n = Math.max(1, parseInt(designCountEl.value || '1', 10));
  designCountBadge.textContent = n;
  designGrid.innerHTML = '<header>デザイン名</header><header>必要数量</header><header></header>';
  for (let i = 0; i < n; i++) {
    const name = document.createElement('input');
    name.type = 'text'; name.placeholder = 'デザイン' + (i+1); name.value = 'デザイン' + (i+1); name.dataset.role = 'name';
    const qty = document.createElement('input');
    qty.type = 'number'; qty.min = '0'; qty.step = '1'; qty.value = i === 0 ? '1000' : (i === 1 ? '2000' : '3000'); qty.dataset.role = 'qty';
    qty.addEventListener('change', runCalc); name.addEventListener('change', runCalc);
    const calcBtn = document.createElement('button'); calcBtn.textContent = '計算'; calcBtn.className = 'btn'; calcBtn.addEventListener('click', runCalc);
    designGrid.appendChild(name); designGrid.appendChild(qty); designGrid.appendChild(calcBtn);
  }
}

function gatherInputs(){
  const qInputs = Array.from(designGrid.querySelectorAll('input[data-role="qty"]'));
  const nInputs  = Array.from(designGrid.querySelectorAll('input[data-role="name"]'));
  const req = qInputs.map(el => Math.max(0, parseInt(el.value || '0', 10)));
  const names = nInputs.map(el => el.value || '');
  return {req, names};
}

function runCalc() {
  const totalFaces = Math.max(0, parseInt(facesPerSheetEl.value || '0', 10));
  const {req, names} = gatherInputs();
  if (totalFaces <= 0) { alert('面付数は1以上で入力してください。'); return; }
  const activeCount = req.filter(q => q > 0).length;
  if (!magnetModeEl.checked && activeCount > totalFaces) { alert('アクティブなデザイン数（数量>0）が面付数を上回っています。'); return; }
  tbody.innerHTML = '';

  if (!magnetModeEl.checked) {
    const res = findOptimalLayout(req, totalFaces);
    if (!res) { alert('計算に失敗しました。'); return; }
    resultHeader.innerHTML = '<th>デザイン名</th><th>面付数</th><th>総印刷数</th><th>予備数</th><th>予備率</th>';
    for (let i = 0; i < req.length; i++) {
      const tr = document.createElement('tr');
      const faces = res.layout[i] || 0;
      const actual = res.actualQuantities[i] || 0;
      const spare = res.overQuantities[i] || 0;
      const rate  = res.overRates[i] || 0;
      tr.innerHTML =
        `<td style="text-align:left;">${names[i]||`デザイン${i+1}`}</td>` +
        `<td>${faces.toLocaleString()}</td>` +
        `<td>${actual.toLocaleString()}</td>` +
        `<td>${spare.toLocaleString()}</td>` +
        `<td>${rate.toFixed(2)}%</td>`;
      tbody.appendChild(tr);
    }
    summaryCell.textContent = `必要原紙枚数（刷了）: ${res.plates.toLocaleString()} 枚`;
    resultTable.style.display = 'table';
    resultTable.dataset.mode = 'single';
    // 保存
    resultTable._save = { names, req };
    return;
  }

  if (totalFaces % 2 !== 0) { alert('半裁最適化は面付数が偶数の時のみ利用できます。'); return; }
  const opt = optimizeMagnetSplit(req, totalFaces);必要数量
  if (opt.error) { alert(opt.error); return; }

  // 追加要望: 「片側面付」と「総印刷数（片側）」の間に「必要数量」を表示
  resultHeader.innerHTML = '<th>デザイン名</th><th>側</th><th>片側面付</th><th>必要数量</th><th>総印刷数（片側）</th><th>予備率（片側）</th>';
  for (let i=0;i<req.length;i++) {
    const d = opt.detail[i];
    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td style="text-align:left;">${names[i]||`デザイン${i+1}`}</td>` +
      `<td style="text-align:center;">${d.side}</td>` +
      `<td>${(d.faces||0).toLocaleString()}</td>` +
      `<td>${(req[i]||0).toLocaleString()}</td>` +  // new column
      `<td>${(d.actual||0).toLocaleString()}</td>` +
      `<td>${(d.rate||0).toFixed(2)}%</td>`;
    tbody.appendChild(tr);
  }
  const msg = [
    `片側面付数: ${opt.H}`,
    `刷了（原紙）: ${opt.sheets.toLocaleString()} 枚`,
    `使用マグネット: ${opt.magnets.toLocaleString()} 枚（A: ${opt.platesA.toLocaleString()} / B: ${opt.platesB.toLocaleString()}）`,
    `マグネット不使用で廃棄する半裁: ${(opt.sheets - opt.platesA) + (opt.sheets - opt.platesB)} 半裁`,
  ].join('　|　');
  summaryCell.innerHTML = msg;
  resultTable.style.display = 'table';
  resultTable.dataset.mode = 'magnet';
  // 保存用
  resultTable._save = { opt, names, req };
}

// Build rows array from current table for export
function buildRowsForExport(){
  if (resultTable.style.display === 'none') return null;
  const mode = resultTable.dataset.mode || 'single';
  const rows = [];
  // header
  const headers = Array.from(resultTable.querySelectorAll('thead th')).map(th => th.textContent.trim());
  rows.push(headers);
  // body
  rows.push(...Array.from(tbody.querySelectorAll('tr')).map(tr => Array.from(tr.children).map(td => td.textContent.trim())));
  // summary
  const save = resultTable._save || {};
  if (mode === 'single') {
    rows.push([]);
    rows.push(['Summary','必要原紙枚数（刷了）', summaryCell.textContent.replace(/[^\d]/g,'')]);
  } else {
    const meta = save.opt;
    if (meta) {
      rows.push([]);
      rows.push(['Summary','片側面付数', meta.H]);
      rows.push(['Summary','刷了（原紙）', meta.sheets]);
      rows.push(['Summary','使用マグネット（合計）', meta.magnets]);
      rows.push(['Summary','A側半裁', meta.platesA]);
      rows.push(['Summary','B側半裁', meta.platesB]);
      rows.push(['Summary','マグネット未使用で廃棄する半裁', (meta.sheets - meta.platesA) + (meta.sheets - meta.platesB)]);
    }
  }
  return rows;
}

// Events
document.getElementById('pickCsv').addEventListener('click', () => csvFileEl.click());
csvFileEl.addEventListener('change', (e) => handleCsvFile(e.target.files[0]));
['dragenter','dragover'].forEach(ev => dropzone.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.add('dragover'); }));
['dragleave','drop'].forEach(ev => dropzone.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('dragover'); }));
dropzone.addEventListener('drop', (e) => {
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) handleCsvFile(file);
});
document.getElementById('sampleCsv').addEventListener('click', () => {
  const sample = 'name,qty\\n' + Array.from({length:30},(_,i)=>`デザイン${i+1},${(i+1)*100}`).join('\\n');
  const w = window.open('', '_blank'); w.document.write(`<pre>${sample.replace(/</g,'&lt;')}</pre>`);
});
applyCountBtn.addEventListener('click', buildDesignInputs);
document.getElementById('magnetMode').addEventListener('change', runCalc);
facesPerSheetEl.addEventListener('change', runCalc);
const exportXlsx = () => {
  const rows = buildRowsForExport();
  if (!rows) { alert('まず計算を実行してください。'); return; }
  const dt = new Date();
  const stamp = dt.toISOString().slice(0,19).replace(/[:T]/g,'-');
  saveXlsx(`result_${(resultTable.dataset.mode||'single')}_${stamp}.xlsx`, rows);
};
exportXlsxBtn.addEventListener('click', exportXlsx);

// CSV load
async function handleCsvFile(file) {
  if (!file) return;
  const text = await file.text();
  const rows = parseCSV(text);
  const items = normalizeCsvRows(rows);
  if (!items.length) {
    alert('CSVの内容が読み取れませんでした。列は「name,qty」形式（ヘッダー有無どちらでも可）で用意してください。');
    return;
  }
  designCountEl.value = items.length;
  buildDesignInputs();
  const nameInputs  = Array.from(designGrid.querySelectorAll('input[data-role="name"]'));
  const qtyInputs   = Array.from(designGrid.querySelectorAll('input[data-role="qty"]'));
  for (let i = 0; i < items.length; i++) {
    if (nameInputs[i]) nameInputs[i].value = items[i].name;
    if (qtyInputs[i])  qtyInputs[i].value  = items[i].qty;
  }
  runCalc();
}

buildDesignInputs();