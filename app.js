// ===== TOAST SYSTEM =====
function toast(msg, type = 'info') {
  const container = document.getElementById('toastContainer');
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => el.remove(), 4200);
}

// ===== CONFIRM DIALOG =====
let confirmResolve = null;
function showConfirm(title, msg) {
  return new Promise(resolve => {
    confirmResolve = resolve;
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmMsg').textContent = msg;
    document.getElementById('confirmOverlay').classList.add('visible');
  });
}
function closeConfirm(result) {
  document.getElementById('confirmOverlay').classList.remove('visible');
  if (confirmResolve) { confirmResolve(result); confirmResolve = null; }
}

// ===== STATE =====
let tipoProducto = 'juguetes';
let maestroData = null;
let maestroColumns = [];
let maestroColumnOrder = [];
let provData = null;
let provColumns = [];
let provRawBuffer = null;
let provFileName = '';
let resultados = [];
let filtroActual = null;
let provHasEANGlobal = true;
let tnPreviewRows = [];

// ===== TIPO SELECTOR =====
function setTipo(tipo) {
  tipoProducto = tipo;
  document.getElementById('btnJuguetes').className = 'tipo-btn' + (tipo === 'juguetes' ? ' active-juguetes' : '');
  document.getElementById('btnLibros').className = 'tipo-btn' + (tipo === 'libros' ? ' active-libros' : '');
  document.getElementById('ivaCard').style.display = tipo === 'libros' ? 'none' : 'block';
}

// ===== DRAG & DROP =====
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.dropzone').forEach(dz => {
    dz.addEventListener('dragover', e => {
      e.preventDefault();
      dz.classList.add('dragover');
    });
    dz.addEventListener('dragleave', () => {
      dz.classList.remove('dragover');
    });
    dz.addEventListener('drop', e => {
      e.preventDefault();
      dz.classList.remove('dragover');
      const input = dz.querySelector('input[type="file"]');
      if (e.dataTransfer.files.length && input) {
        input.files = e.dataTransfer.files;
        input.dispatchEvent(new Event('change'));
      }
    });
  });
});

// ===== NORMALIZATION =====
function normEAN(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number') {
    val = val.toFixed(0);
  } else {
    val = String(val).trim();
    if (/^\d+\.?\d*[eE][+\-]?\d+$/.test(val)) {
      try { val = BigInt(Math.round(parseFloat(val))).toString(); } catch (e) { val = parseFloat(val).toFixed(0); }
    }
    if (/^\d+\.0+$/.test(val)) val = val.replace(/\.0+$/, '');
  }
  val = val.replace(/[\s\-\.]/g, '');
  val = val.replace(/[^0-9]/g, '');
  return val;
}

function normDesc(val) {
  if (!val) return '';
  return String(val).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normDescCompact(val) {
  if (!val) return '';
  return String(val).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

// ===== SIMILARITY FUNCTIONS =====
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  let curr = new Array(n + 1);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      curr[j] = a[i - 1] === b[j - 1]
        ? prev[j - 1]
        : 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n];
}

function tokenOverlap(a, b) {
  const ta = new Set(normDesc(a).split(' ').filter(t => t.length > 1));
  const tb = new Set(normDesc(b).split(' ').filter(t => t.length > 1));
  if (ta.size === 0 || tb.size === 0) return 0;
  let overlap = 0;
  for (const t of ta) if (tb.has(t)) overlap++;
  return overlap / Math.max(ta.size, tb.size);
}

function combinedSimilarity(a, b) {
  const na = normDesc(a), nb = normDesc(b);
  if (!na || !nb) return 0;
  const maxLen = Math.max(na.length, nb.length);
  const levSim = maxLen > 0 ? 1 - levenshtein(na, nb) / maxLen : 0;
  const tokSim = tokenOverlap(a, b);
  return levSim * 0.4 + tokSim * 0.6;
}

// ===== PRICE PARSER =====
function parsePrice(val) {
  if (val == null) return 0;
  if (typeof val === 'number') return val;
  let s = String(val).replace(/[^0-9.,\-]/g, '');
  if (!s) return 0;
  
  const hasComma = s.includes(',');
  const hasDot = s.includes('.');

  if (hasComma && hasDot) {
    const lastComma = s.lastIndexOf(',');
    const lastDot = s.lastIndexOf('.');
    if (lastComma > lastDot) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
  } 
  else if (hasComma) {
    const match = s.match(/,(\d+)$/);
    if (match && match[1].length === 3) {
      s = s.replace(/,/g, ''); 
    } else {
      s = s.replace(',', '.');
    }
  } 
  else if (hasDot) {
    const match = s.match(/\.(\d+)$/);
    if (match && match[1].length === 3) {
      s = s.replace(/\./g, '');
    }
  }
  
  return parseFloat(s) || 0;
}

// ===== EXCEL READING (CORREGIDO PARA EVITAR QUE ROMPA CSVs LOCALES) =====
// ===== HTML-XLS DETECTION & PARSER =====
// Algunos sistemas de gestión (ej: Tango, etc.) exportan archivos .xls que en realidad
// son HTML con extensión .xls. xlsx.js no puede leerlos; los parseamos con DOMParser.
function isHtmlXls(buffer) {
  try {
    const bytes = new Uint8Array(buffer.slice ? buffer.slice(0, 20) : buffer, 0, 20);
    const sig = String.fromCharCode(...bytes).toLowerCase();
    return sig.startsWith('<') || sig.startsWith('\ufeff<');
  } catch (e) { return false; }
}

function parseHtmlXls(buffer, headerRow) {
  const decoder = new TextDecoder('utf-8');
  const html = decoder.decode(buffer);
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const rows = Array.from(doc.querySelectorAll('table tr'));
  if (rows.length === 0) return { columns: [], data: [], sheetNames: ['Sheet0'] };

  const hIdx = (headerRow || 1) - 1;
  const headerCells = Array.from(rows[hIdx]?.querySelectorAll('td,th') || []);
  const headers = headerCells.map((td, i) => {
    const s = td.textContent.trim();
    return s || `Col_${i + 1}`;
  });

  const data = [];
  for (let i = hIdx + 1; i < rows.length; i++) {
    const cells = Array.from(rows[i].querySelectorAll('td,th'));
    const row = {};
    let hasData = false;
    headers.forEach((h, j) => {
      const val = cells[j] ? cells[j].textContent.trim() : '';
      row[h] = val;
      if (val !== '') hasData = true;
    });
    if (hasData) data.push(row);
  }
  return { columns: headers, data, sheetNames: ['Sheet0'] };
}

function readExcelBuffer(buffer, fileName, headerRow, sheetName) {
  const isXLS = /\.xls$/i.test(fileName) && !/\.xlsx$/i.test(fileName);

  // Detectar HTML-XLS antes de intentar xlsx.js
  if (isXLS && isHtmlXls(buffer)) {
    return parseHtmlXls(buffer, headerRow || 1);
  }

  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    // raw: true previene la lectura incorrecta de CSVs
    wb = XLSX.read(binary, { type: 'binary', raw: true });
  } else {
    wb = XLSX.read(buffer, { type: 'array', raw: true });
  }
  const wsName = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[wsName];
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const hIdx = (headerRow || 1) - 1;
  if (hIdx >= allRows.length) return { columns: [], data: [], sheetNames: wb.SheetNames };

  const headers = allRows[hIdx].map((h, i) => {
    const s = String(h).trim();
    return s || `Col_${i + 1}`;
  });

  const rows = [];
  for (let i = hIdx + 1; i < allRows.length; i++) {
    const row = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      const val = allRows[i][j];
      row[headers[j]] = val != null ? val : '';
      if (val != null && String(val).trim() !== '') hasData = true;
    }
    if (hasData) rows.push(row);
  }
  return { columns: headers, data: rows, sheetNames: wb.SheetNames };
}

function getSheetNames(buffer, fileName) {
  const isXLS = /\.xls$/i.test(fileName) && !/\.xlsx$/i.test(fileName);
  // HTML-XLS siempre tiene una sola hoja
  if (isXLS && isHtmlXls(buffer)) return ['Sheet0'];

  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    wb = XLSX.read(binary, { type: 'binary', raw: true });
  } else {
    wb = XLSX.read(buffer, { type: 'array', raw: true });
  }
  return wb.SheetNames;
}

function autoDetectHeaderRow(buffer, fileName, sheetName) {
  const isXLS = /\.xls$/i.test(fileName) && !/\.xlsx$/i.test(fileName);

  let allRows;
  if (isXLS && isHtmlXls(buffer)) {
    // Para HTML-XLS, parseamos directo y asumimos fila 1
    const result = parseHtmlXls(buffer, 1);
    // Verificar que la primera fila tiene texto (no números)
    return 1;
  }

  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    wb = XLSX.read(binary, { type: 'binary', raw: true });
  } else {
    wb = XLSX.read(buffer, { type: 'array', raw: true });
  }
  const wsName = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[wsName];
  allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  for (let i = 0; i < Math.min(allRows.length - 1, 25); i++) {
    const row = allRows[i];
    if (!row) continue;
    const nonEmpty = row.filter(c => c != null && String(c).trim().length > 0);
    if (nonEmpty.length < 3) continue;
    const textCells = nonEmpty.filter(c => isNaN(Number(c)));
    if (textCells.length < 3 || textCells.length < nonEmpty.length * 0.6) continue;
    const nextRow = allRows[i + 1];
    if (!nextRow) continue;
    const nextNonEmpty = nextRow.filter(c => c != null && String(c).trim().length > 0);
    if (nextNonEmpty.length >= 2) return i + 1;
  }
  return 1;
}

function autoDetectColumn(columns, keywords, prioritizeFirst = false) {
  const kws = keywords.map(k => k.toLowerCase());
  for (const kw of kws) {
    for (let i = 0; i < columns.length; i++) {
      if (columns[i].toLowerCase() === kw) return i;
    }
  }
  for (let i = 0; i < columns.length; i++) {
    const col = columns[i].toLowerCase();
    for (const kw of kws) {
      if (col.includes(kw) || kw.includes(col)) return i;
    }
  }
  return 0;
}

function populateSelect(selectId, columns, selectedIdx) {
  const sel = document.getElementById(selectId);
  sel.innerHTML = '';
  columns.forEach((c, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = c;
    if (i === selectedIdx) opt.selected = true;
    sel.appendChild(opt);
  });
}

// ===== MAESTRO LOAD =====
function onMaestroFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const buffer = e.target.result;
      const result = readExcelBuffer(buffer, file.name, 1);
      maestroData = result.data;
      maestroColumns = result.columns;
      maestroColumnOrder = [...result.columns];

      document.getElementById('dzMaestro').classList.add('loaded');
      document.getElementById('dzMaestroFile').textContent = `✓ ${file.name} — ${maestroData.length} productos`;
      document.getElementById('maestroFields').style.display = 'grid';

      const eanIdx = autoDetectColumn(maestroColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar'], true);
      const descIdx = autoDetectColumn(maestroColumns, ['descripcion', 'descripción', 'titulo', 'titulos', 'nombre', 'producto', 'detalle']);
      const precioIdx = autoDetectColumn(maestroColumns, ['precio compra', 'precio_compra', 'costo', 'precio costo', 'lista1', 'precio', 'pvp']);
      const codIntIdx = autoDetectColumn(maestroColumns, ['codigoprov', 'codigo_prov', 'cod_prov', 'sku', 'codigo interno', 'artículo', 'articulo']);

      populateSelect('selMaestroEAN', maestroColumns, eanIdx);
      populateSelect('selMaestroDesc', maestroColumns, descIdx);
      populateSelect('selMaestroPrecio', maestroColumns, precioIdx);
      populateSelect('selMaestroCodInt', maestroColumns, codIntIdx);

      checkReady();
      toast(`Maestro cargado: ${maestroData.length} productos`, 'success');
    } catch (err) {
      console.error('Error loading maestro file:', err);
      toast('Error al leer el archivo maestro: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ===== PROVEEDOR LOAD =====
function onProvFile(file) {
  if (!file) return;
  provFileName = file.name;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      provRawBuffer = e.target.result;

      const sheets = getSheetNames(provRawBuffer, provFileName);
      const sheetSelect = document.getElementById('selProvSheet');

      if (sheets.length > 1) {
        sheetSelect.innerHTML = '';
        sheets.forEach((name) => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          sheetSelect.appendChild(opt);
        });
        document.getElementById('provSheetConfig').style.display = 'flex';
      } else {
        document.getElementById('provSheetConfig').style.display = 'none';
      }

      const selectedSheet = sheets.length > 1 ? sheets[0] : null;
      const detectedRow = autoDetectHeaderRow(provRawBuffer, provFileName, selectedSheet);
      document.getElementById('provHeaderRow').value = detectedRow;
      loadProvFromBuffer(detectedRow, selectedSheet);
      document.getElementById('provHeaderConfig').style.display = 'flex';
    } catch (err) {
      console.error('Error loading provider file:', err);
      toast('Error al leer el archivo del proveedor: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function onProvSheetChange() {
  const sheetName = document.getElementById('selProvSheet').value;
  const detectedRow = autoDetectHeaderRow(provRawBuffer, provFileName, sheetName);
  document.getElementById('provHeaderRow').value = detectedRow;
  loadProvFromBuffer(detectedRow, sheetName);
}

function reloadProv() {
  const row = parseInt(document.getElementById('provHeaderRow').value) || 1;
  const sheetSelect = document.getElementById('selProvSheet');
  const sheetName = sheetSelect.options.length > 0 ? sheetSelect.value : null;
  loadProvFromBuffer(row, sheetName);
}

function loadProvFromBuffer(headerRow, sheetName) {
  let result = readExcelBuffer(provRawBuffer, provFileName, headerRow, sheetName);

  const realHeaders = result.columns.filter(h => !h.startsWith('Col_'));
  if (realHeaders.length < 2 && result.columns.length > 0) {
    for (let tryRow = headerRow + 1; tryRow <= headerRow + 5 && tryRow <= 25; tryRow++) {
      const retry = readExcelBuffer(provRawBuffer, provFileName, tryRow, sheetName);
      const retryReal = retry.columns.filter(h => !h.startsWith('Col_'));
      if (retryReal.length >= 2) {
        headerRow = tryRow;
        result = retry;
        document.getElementById('provHeaderRow').value = tryRow;
        break;
      }
    }
  }

  provData = result.data;
  provColumns = result.columns;

  const sheetLabel = sheetName ? ` [${sheetName}]` : '';
  document.getElementById('dzProv').classList.add('loaded');
  document.getElementById('dzProvFile').textContent = `✓ ${provFileName}${sheetLabel} — ${provData.length} productos (header fila ${headerRow})`;
  document.getElementById('provFields').style.display = 'grid';

  const eanIdx = autoDetectColumn(provColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar']);
  const precioIdx = autoDetectColumn(provColumns, ['pvp', 'precio', 'lista', 'precio de lista', 'p.v.p', 'valor']);
  const descIdx = autoDetectColumn(provColumns, ['titulo', 'titulos', 'descripcion', 'descripción', 'nombre', 'producto', 'detalle']);

  populateSelect('selProvEAN', provColumns, eanIdx);
  populateSelect('selProvPrecio', provColumns, precioIdx);
  populateSelect('selProvDesc', provColumns, descIdx);

  checkReady();
  toast(`Proveedor cargado: ${provData.length} productos`, 'success');
}

function checkReady() {
  document.getElementById('btnProcesar').disabled = !(maestroData && provData);
}

// ===== PROCESAR =====
function procesar() {
  const mEanCol = maestroColumns[document.getElementById('selMaestroEAN').value];
  const mDescCol = maestroColumns[document.getElementById('selMaestroDesc').value];
  const mPrecioCol = maestroColumns[document.getElementById('selMaestroPrecio').value];
  const mCodIntCol = maestroColumns[document.getElementById('selMaestroCodInt').value];

  const pEanCol = provColumns[document.getElementById('selProvEAN').value];
  const pPrecioCol = provColumns[document.getElementById('selProvPrecio').value];
  const pDescCol = provColumns[document.getElementById('selProvDesc').value];

  const ivaMode = document.getElementById('selIVA').value;
  const ivaRate = parseFloat(document.getElementById('ivaRate').value) / 100;

  const maestroByEAN = new Map();
  const maestroByDesc = new Map();
  const maestroByDescCompact = new Map();
  const maestroDescList = [];
  maestroData.forEach((row, idx) => {
    const ean = normEAN(row[mEanCol]);
    if (ean) maestroByEAN.set(ean, idx);
    const desc = normDesc(row[mDescCol]);
    const descCompact = normDescCompact(row[mDescCol]);
    if (desc) {
      maestroByDesc.set(desc, idx);
      maestroDescList.push({ desc, descCompact, idx, original: row[mDescCol] });
    }
    if (descCompact) maestroByDescCompact.set(descCompact, idx);
  });

  let provHasEAN = false;
  for (let i = 0; i < Math.min(provData.length, 20); i++) {
    const val = normEAN(provData[i][pEanCol]);
    if (val && /^\d{8,14}$/.test(val)) { provHasEAN = true; break; }
  }
  provHasEANGlobal = provHasEAN;
  resultados = [];

  provData.forEach(pRow => {
    const pEanRaw = pRow[pEanCol];
    const pEan = normEAN(pEanRaw);
    const pDesc = String(pRow[pDescCol] || '');
    const pPrecioRaw = parsePrice(pRow[pPrecioCol]);

    let precioNuevo;
    if (tipoProducto === 'libros') {
      precioNuevo = pPrecioRaw;
    } else {
      precioNuevo = ivaMode === 'neto' ? pPrecioRaw * (1 + ivaRate) : pPrecioRaw;
    }

    let mIdx = -1;
    let matchType = 'notfound';
    let fuzzySim = 0;

    if (provHasEAN && pEan && pEan.length >= 8 && maestroByEAN.has(pEan)) {
      mIdx = maestroByEAN.get(pEan);
      matchType = 'ean';
    }

    if (mIdx < 0 && tipoProducto !== 'libros') {
      const pDescNorm = normDesc(pDesc);
      const pDescCompact = normDescCompact(pDesc);

      if (pDescNorm && maestroByDesc.has(pDescNorm)) {
        mIdx = maestroByDesc.get(pDescNorm);
        matchType = 'desc_exact';
      }

      if (mIdx < 0 && pDescCompact && maestroByDescCompact.has(pDescCompact)) {
        mIdx = maestroByDescCompact.get(pDescCompact);
        matchType = 'desc_exact';
      }
    }

    if (mIdx >= 0) {
      const mRow = maestroData[mIdx];
      const precioAnterior = parsePrice(mRow[mPrecioCol]);
      const targetCol = tipoProducto === 'libros' ? findLista1Col() || mPrecioCol : mPrecioCol;

      let estado;
      if (precioNuevo > precioAnterior) estado = 'up';
      else if (precioNuevo < precioAnterior) estado = 'down';
      else estado = 'same';

      const variacion = precioAnterior > 0 ? ((precioNuevo - precioAnterior) / precioAnterior * 100) : 0;

      let pvp = precioNuevo;
      if (tipoProducto !== 'libros') {
        const margin = parseFloat(document.getElementById('marginPct').value) / 100;
        const mode = document.getElementById('marginMode').value;
        pvp = mode === 'sobre_costo' ? precioNuevo * (1 + margin) : precioNuevo / (1 - margin);
      }

      resultados.push({
        maestroIdx: mIdx,
        maestroRow: { ...mRow },
        ean: normEAN(mRow[mEanCol]) || pEan,
        eanDisplay: mRow[mEanCol] || pEanRaw,
        codInt: mRow[mCodIntCol] || '',
        descMaestro: mRow[mDescCol] || '',
        descProv: pDesc,
        precioAnterior,
        precioNuevo: Math.round(precioNuevo * 100) / 100,
        pvp: Math.round(pvp * 100) / 100,
        estado,
        variacion: Math.round(variacion * 100) / 100,
        matchType,
        fuzzySim: Math.round(fuzzySim * 100),
        targetCol,
        checked: true
      });
    } else {
      resultados.push({
        maestroIdx: -1,
        maestroRow: null,
        ean: pEan,
        eanDisplay: pEanRaw,
        codInt: '',
        descMaestro: '',
        descProv: pDesc,
        precioAnterior: 0,
        precioNuevo: Math.round(precioNuevo * 100) / 100,
        pvp: 0,
        estado: 'notfound',
        variacion: 0,
        matchType: 'notfound',
        fuzzySim: 0,
        targetCol: '',
        checked: false
      });
    }
  });

  const matched = resultados.filter(r => r.estado !== 'notfound').length;
  toast(`Procesado: ${matched} matcheados de ${resultados.length} productos`, 'success');
  goStep(2);
}

function findLista1Col() {
  return maestroColumns.find(c => c.toLowerCase().replace(/\s/g, '') === 'lista1') || null;
}

// ===== NAVIGATION =====
function goStep(n) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel' + n).classList.add('active');

  ['step1ind', 'step2ind', 'step3ind'].forEach((id, i) => {
    const el = document.getElementById(id);
    el.className = 'step-indicator';
    if (i + 1 < n) el.classList.add('done');
    if (i + 1 === n) el.classList.add('active');
  });

  if (n === 2) renderReview();
  if (n === 3) renderExport();
}

function goDirectTN() {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel4').classList.add('active');

  ['step1ind', 'step2ind', 'step3ind'].forEach((id) => {
    const el = document.getElementById(id);
    el.className = 'step-indicator';
  });
  document.getElementById('step3ind').classList.add('active');
}

// ===== SELECTION COUNTER =====
function updateSelectionCounter() {
  const count = resultados.filter(r => r.checked && r.estado !== 'notfound').length;
  const el = document.getElementById('selectionCounter');
  if (el) el.textContent = `${count} seleccionado${count !== 1 ? 's' : ''}`;
}

// ===== RENDER REVIEW =====
function renderReview() {
  const banner = document.getElementById('reviewBanner');
  let bannerHTML = '';
  if (tipoProducto === 'libros') {
    banner.className = 'mode-banner libros';
    bannerHTML = '📚 Modo Libros — PVP directo del proveedor → columna LISTA1';
  } else {
    banner.className = 'mode-banner juguetes';
    bannerHTML = '🧸 Modo Juguetes / General — Precio de costo + margen';
  }
  if (!provHasEANGlobal) {
    bannerHTML += '<br><span style="font-size:12px;opacity:0.85">⚠ Proveedor sin códigos de barras — matching por nombre</span>';
  }
  banner.innerHTML = bannerHTML;

  const counts = { up: 0, down: 0, same: 0, notfound: 0 };
  resultados.forEach(r => { counts[r.estado]++; });

  const statsBar = document.getElementById('statsBar');
  statsBar.innerHTML = `
    <div class="stat-pill up ${filtroActual === 'up' ? 'active' : ''}" onclick="setFiltro('up')">▲ ${counts.up} subieron</div>
    <div class="stat-pill down ${filtroActual === 'down' ? 'active' : ''}" onclick="setFiltro('down')">▼ ${counts.down} bajaron</div>
    <div class="stat-pill same ${filtroActual === 'same' ? 'active' : ''}" onclick="setFiltro('same')">= ${counts.same} sin cambio</div>
    <div class="stat-pill notfound ${filtroActual === 'notfound' ? 'active' : ''}" onclick="setFiltro('notfound')">✗ ${counts.notfound} sin match</div>
    <div class="stat-pill ${filtroActual === null ? 'active' : ''}" style="background:var(--accent-bg);color:var(--accent2)" onclick="setFiltro(null)">Todos (${resultados.length})</div>
  `;

  document.getElementById('reviewHead').innerHTML = `<tr>
    <th style="width:40px"><input type="checkbox" class="cb" checked onchange="toggleAll(this.checked)"></th>
    <th>Estado</th>
    <th>Match</th>
    <th>EAN</th>
    <th>Descripción</th>
    <th>Precio Anterior</th>
    <th>Precio Nuevo</th>
    <th>Variación</th>
    ${tipoProducto !== 'libros' ? '<th>PVP Sugerido</th>' : ''}
  </tr>`;

  renderTableRows();
}

function renderTableRows() {
  const search = document.getElementById('searchInput').value.toLowerCase();
  const tbody = document.getElementById('reviewBody');
  tbody.innerHTML = '';

  resultados.forEach((r, i) => {
    if (filtroActual && r.estado !== filtroActual) return;
    if (search) {
      const s = `${r.ean} ${r.descMaestro} ${r.descProv} ${r.codInt}`.toLowerCase();
      if (!s.includes(search)) return;
    }

    const rowClass = r.estado === 'notfound' ? 'row-notfound' : '';

    let matchBadge = '';
    if (r.matchType === 'ean') matchBadge = '<span class="badge badge-ean">EAN</span>';
    else if (r.matchType === 'desc_exact') matchBadge = '<span class="badge badge-desc">DESC</span>';
    else matchBadge = '<span class="badge badge-notfound">SIN MATCH</span>';

    let estadoBadge = '';
    if (r.estado === 'up') estadoBadge = '<span class="badge badge-up">▲ Subió</span>';
    else if (r.estado === 'down') estadoBadge = '<span class="badge badge-down">▼ Bajó</span>';
    else if (r.estado === 'same') estadoBadge = '<span class="badge badge-same">= Igual</span>';
    else estadoBadge = '<span class="badge badge-notfound">✗</span>';

    const priceClass = r.estado === 'up' ? 'price-up' : (r.estado === 'down' ? 'price-down' : 'price-same');
    const desc = r.descMaestro || r.descProv;

    const tr = document.createElement('tr');
    tr.className = rowClass;
    tr.innerHTML = `
      <td class="td-cb">${r.estado === 'notfound' ? '<span style="color:var(--text3)">—</span>' : `<input type="checkbox" class="cb" ${r.checked ? 'checked' : ''} onchange="resultados[${i}].checked=this.checked;updateSelectionCounter()">`}</td>
      <td class="td-badge">${estadoBadge}</td>
      <td class="td-badge">${matchBadge}</td>
      <td class="td-ean">${r.eanDisplay || '—'}</td>
      <td class="td-desc">${desc}</td>
      <td class="td-price price-change ${priceClass}">$${r.precioAnterior.toLocaleString('es-AR', { minimumFractionDigits: 2 })}</td>
      <td class="td-price price-change ${priceClass}">$${r.precioNuevo.toLocaleString('es-AR', { minimumFractionDigits: 2 })}</td>
      <td class="td-price price-change ${priceClass}">${r.variacion > 0 ? '+' : ''}${r.variacion}%</td>
      ${tipoProducto !== 'libros' ? `<td class="td-price price-change">$${r.pvp.toLocaleString('es-AR', { minimumFractionDigits: 2 })}</td>` : ''}
    `;
    tbody.appendChild(tr);
  });

  updateSelectionCounter();
}

function setFiltro(f) {
  filtroActual = filtroActual === f ? null : f;
  renderReview();
}

function filtrarTabla() { renderTableRows(); }

function toggleAll(checked) {
  resultados.forEach(r => {
    if (r.estado === 'notfound') { r.checked = false; }
    else { r.checked = checked; }
  });
  renderTableRows();
}

function toggleMarginConfig() {
  document.getElementById('marginConfig').classList.toggle('visible');
}

function recalcularMargenes() {
  const margin = parseFloat(document.getElementById('marginPct').value) / 100;
  const mode = document.getElementById('marginMode').value;
  resultados.forEach(r => {
    if (r.estado !== 'notfound') {
      r.pvp = mode === 'sobre_costo'
        ? Math.round(r.precioNuevo * (1 + margin) * 100) / 100
        : Math.round(r.precioNuevo / (1 - margin) * 100) / 100;
    }
  });
  renderTableRows();
  toast('Márgenes recalculados', 'success');
}

// ===== RENDER EXPORT =====
function renderExport() {
  const banner = document.getElementById('exportBanner');
  if (tipoProducto === 'libros') {
    banner.className = 'mode-banner libros';
    banner.innerHTML = '📚 Modo Libros — Exportación con PVP directo en LISTA1';
  } else {
    banner.className = 'mode-banner juguetes';
    banner.innerHTML = '🧸 Modo Juguetes / General — Precio de costo actualizado';
  }

  const selected = resultados.filter(r => r.checked && r.estado !== 'notfound');
  document.getElementById('exportInfo').textContent = `Se exportarán ${selected.length} productos (solo los matcheados y seleccionados).`;

  // Auto-cargar datos de gestión desde resultados para paso 3B
  if (selected.length > 0) {
    gestionData = selected.map(r => ({
      _ean: r.ean,
      _desc: r.descMaestro || r.descProv,
      _pvp: tipoProducto === 'libros' ? r.precioNuevo : r.pvp
    }));
    gestionColumns = ['_ean', '_desc', '_pvp'];

    document.getElementById('dzGestion').classList.add('loaded');
    document.getElementById('dzGestionFile').textContent = `✓ Desde paso 2 — ${selected.length} productos con precios actualizados`;
    document.getElementById('gestionFields').style.display = 'none';

    // Setear selects ocultos para que generarTiendaNube los lea
    populateSelect('selGestionEAN', gestionColumns, 0);
    populateSelect('selGestionDesc', gestionColumns, 1);
    populateSelect('selGestionPVP', gestionColumns, 2);

    checkTNReady(false);
  }
}

// ===== EXPORT INTERNO (CORREGIDO: Tipado seguro de texto y números) =====
async function exportarInterno() {
  const selected = resultados.filter(r => r.checked && r.estado !== 'notfound');
  if (selected.length === 0) {
    toast('No hay productos seleccionados para exportar.', 'warning');
    return;
  }

  const ok = await showConfirm('Exportar Excel Interno', `Se van a exportar ${selected.length} productos. ¿Continuar?`);
  if (!ok) return;

  const mPrecioCol = maestroColumns[document.getElementById('selMaestroPrecio').value];
  const mEanCol = maestroColumns[document.getElementById('selMaestroEAN').value];

  const exportRows = selected.map(r => {
    const row = {};
    maestroColumnOrder.forEach(col => {
      let val = r.maestroRow[col];
      if (val == null) { row[col] = ''; return; }
      
      const colLower = col.toLowerCase();
      
      // 1. Proteger códigos para evitar la notación científica (7.16E+12)
      if (col === mEanCol || colLower.includes('ean') || colLower.includes('cod') || colLower.includes('cód') || colLower.includes('barras') || colLower.includes('sku') || colLower.includes('art')) {
        row[col] = String(val);
      } 
      // 2. Limpiar columnas numéricas (IVA, ganancia, otras listas) que hayan venido sucias
      else if (typeof val === 'string' && val.trim() !== '') {
        // Si tiene pinta de ser dinero o porcentaje (solo números, comas, puntos, signos)
        if (/^[\d\s.,$\-%]+$/.test(val.trim())) {
          let num = parsePrice(val);
          row[col] = isNaN(num) ? val : num;
        } else {
          row[col] = val; // Es texto normal (ej: descripción del producto)
        }
      } 
      // 3. Valores que ya están bien
      else {
        row[col] = val;
      }
    });

    if (tipoProducto === 'libros') {
      const lista1Col = findLista1Col();
      if (lista1Col) row[lista1Col] = r.precioNuevo;
    } else {
      row[mPrecioCol] = r.precioNuevo;
    }

    row['_PrecioAnterior'] = r.precioAnterior;
    row['_PrecioNuevo'] = r.precioNuevo;
    row['_Variacion%'] = r.variacion;
    row['_PVP_Sugerido'] = r.pvp || '';
    row['_Estado'] = r.estado;
    row['_Match'] = r.matchType;
    row['_Fecha'] = new Date().toLocaleDateString('es-AR');

    return row;
  });

  const headers = [...maestroColumnOrder, '_PrecioAnterior', '_PrecioNuevo', '_Variacion%', '_PVP_Sugerido', '_Estado', '_Match', '_Fecha'];

  const ws = XLSX.utils.json_to_sheet(exportRows, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Actualización');
  XLSX.writeFile(wb, `PriceSync_Interno_${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast(`Excel exportado con ${selected.length} productos`, 'success');
}

// ===== TIENDA NUBE — 2-FILE FLOW =====
let gestionData = null;
let gestionColumns = [];
let tnRawText = '';
let tnRows = [];
let tnHeaders = [];
let tnSep = ';';
let tnUpdatedRows = [];

function onGestionFile(file, direct) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const buffer = e.target.result;
    const result = readExcelBuffer(buffer, file.name, 1);
    gestionData = result.data;
    gestionColumns = result.columns;

    const suffix = direct ? 'Direct' : '';
    document.getElementById('dzGestion' + suffix).classList.add('loaded');
    document.getElementById('dzGestion' + suffix + 'File').textContent = `✓ ${file.name} — ${gestionData.length} productos`;
    document.getElementById('gestionFields' + suffix).style.display = 'grid';

    const eanIdx = autoDetectColumn(gestionColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar'], true);
    const descIdx = autoDetectColumn(gestionColumns, ['descripcion', 'descripción', 'nombre', 'producto', 'detalle', 'articulo']);
    const pvpIdx = autoDetectColumn(gestionColumns, ['lista1', 'pvp', 'precio venta', 'precio_venta', 'precio publico', 'precio']);

    populateSelect('selGestionEAN' + suffix, gestionColumns, eanIdx);
    populateSelect('selGestionDesc' + suffix, gestionColumns, descIdx);
    populateSelect('selGestionPVP' + suffix, gestionColumns, pvpIdx);

    checkTNReady(direct);
    toast('Archivo de gestión cargado', 'success');
  };
  reader.readAsArrayBuffer(file);
}

function onTNFile(file, direct) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    tnRawText = e.target.result;

    const firstLine = tnRawText.split('\n')[0];
    tnSep = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';

    tnRows = parseCSV(tnRawText, tnSep);
    if (tnRows.length < 2) { toast('El archivo parece vacío.', 'error'); return; }
    tnHeaders = tnRows[0];

    const suffix = direct ? 'Direct' : '';
    document.getElementById('dzTN' + suffix).classList.add('loaded');
    document.getElementById('dzTN' + suffix + 'File').textContent = `✓ ${file.name} — ${tnRows.length - 1} productos`;

    checkTNReady(direct);
    toast('CSV de Tienda Nube cargado', 'success');
  };
  reader.readAsText(file, 'iso-8859-1');
}

function checkTNReady(direct) {
  const suffix = direct ? 'Direct' : '';
  document.getElementById('btnGenerarTN' + suffix).disabled = !(gestionData && tnRows.length > 1);
}

function generarTiendaNube(direct) {
  const suffix = direct ? 'Direct' : '';
  const gEanCol = gestionColumns[document.getElementById('selGestionEAN' + suffix).value];
  const gDescCol = gestionColumns[document.getElementById('selGestionDesc' + suffix).value];
  const gPvpCol = gestionColumns[document.getElementById('selGestionPVP' + suffix).value];

  const pvpByEAN = new Map();
  const gestionByDesc = new Map();
  const gestionByDescCompact = new Map();
  const gestionDescList = [];

  gestionData.forEach((row, idx) => {
    const ean = normEAN(row[gEanCol]);
    const pvp = parsePrice(row[gPvpCol]);
    const desc = String(row[gDescCol] || '').trim();
    const descNorm = normDesc(desc);
    const descCompact = normDescCompact(desc);

    if (ean && pvp > 0) pvpByEAN.set(ean, { pvp, desc, ean });
    if (descNorm && pvp > 0) {
      const entry = { pvp, desc, idx, ean, descNorm, descCompact };
      gestionByDesc.set(descNorm, entry);
      gestionDescList.push(entry);
    }
    if (descCompact && pvp > 0) gestionByDescCompact.set(descCompact, { pvp, desc, idx, ean });
  });

  const barCodeIdx = tnHeaders.findIndex(h => {
    const l = h.toLowerCase().replace(/[""]/g, '').trim();
    return l.includes('código de barras') || l.includes('codigo de barras') || l === 'barcode';
  });
  const precioIdx = tnHeaders.findIndex(h => {
    const l = h.toLowerCase().replace(/[""]/g, '').trim();
    return l === 'precio';
  });
  const nombreIdx = tnHeaders.findIndex(h => {
    const l = h.toLowerCase().replace(/[""]/g, '').trim();
    return l === 'nombre';
  });

  if (precioIdx === -1) {
    toast('No se encontró la columna "Precio" en el CSV de Tienda Nube.', 'error');
    return;
  }

  tnUpdatedRows = [tnHeaders];
  let matchCount = 0;
  tnPreviewRows = [];
  let previewRows = tnPreviewRows;
  let statsByType = { ean: 0, desc_exact: 0, desc_contains: 0, desc_fuzzy: 0 };

  for (let i = 1; i < tnRows.length; i++) {
    const row = [...tnRows[i]];
    while (row.length < tnHeaders.length) row.push('');

    const ean = barCodeIdx >= 0 ? normEAN(row[barCodeIdx]) : '';
    const tnNombre = nombreIdx >= 0 ? String(row[nombreIdx] || '').replace(/^"|"$/g, '').trim() : '';
    let newPVP = null;
    let matchType = '';
    let matchDesc = '';
    let matchGestionEAN = '';

    if (ean && ean.length >= 8 && pvpByEAN.has(ean)) {
      newPVP = pvpByEAN.get(ean).pvp;
      matchType = 'ean';
      matchDesc = pvpByEAN.get(ean).desc;
      matchGestionEAN = ean;
    }

    if (newPVP !== null) {
      const oldPrecio = row[precioIdx];
      const oldPrecioNum = parsePrice(oldPrecio);
      if (Math.abs(oldPrecioNum - newPVP) > 0.01) {
        row[precioIdx] = formatTNPrice(newPVP);
        matchCount++;
        statsByType[matchType] = (statsByType[matchType] || 0) + 1;
        previewRows.push({
          rowIndex: i,
          checked: true,
          ean: barCodeIdx >= 0 ? row[barCodeIdx] : '',
          nombre: tnNombre,
          oldPrecio,
          newPrecio: row[precioIdx],
          matchType,
          matchDesc,
          matchGestionEAN,
          tnEAN: ean
        });
      }
    }
    tnUpdatedRows.push(row);
  }

  document.getElementById('tnPreview' + suffix).style.display = 'block';

  let statsHTML = `
    <div class="stats-bar">
      <div class="stat-pill up">✎ ${matchCount} modificados</div>
      <div class="stat-pill same">— ${tnRows.length - 1 - matchCount} sin cambio</div>
      <div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)">${tnRows.length - 1} total</div>
    </div>`;
  if (matchCount > 0) {
    const eanMismatches = previewRows.filter(p => p.matchType !== 'ean' && p.matchGestionEAN && p.tnEAN && p.tnEAN.length >= 8 && p.tnEAN !== p.matchGestionEAN).length;
    const eanMissing = previewRows.filter(p => p.matchType !== 'ean' && p.matchGestionEAN && (!p.tnEAN || p.tnEAN.length < 8)).length;

    statsHTML += `<div class="stats-bar" style="margin-top:8px">`;
    if (statsByType.ean) statsHTML += `<div class="stat-pill" style="background:var(--green-bg);color:var(--green)">EAN: ${statsByType.ean}</div>`;
    if (statsByType.desc_exact) statsHTML += `<div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)">Nombre exacto: ${statsByType.desc_exact}</div>`;
    if (statsByType.desc_contains) statsHTML += `<div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)">Contiene: ${statsByType.desc_contains}</div>`;
    if (statsByType.desc_fuzzy) statsHTML += `<div class="stat-pill" style="background:var(--orange-bg);color:var(--orange)">⚠ Similitud: ${statsByType.desc_fuzzy}</div>`;
    statsHTML += `</div>`;
    if (eanMismatches > 0 || eanMissing > 0) {
      statsHTML += `<div class="stats-bar" style="margin-top:8px">`;
      if (eanMismatches > 0) statsHTML += `<div class="stat-pill" style="background:var(--orange-bg);color:var(--orange)">⚠ ${eanMismatches} EAN diferentes</div>`;
      if (eanMissing > 0) statsHTML += `<div class="stat-pill" style="background:var(--blue-bg);color:var(--blue)">ℹ ${eanMissing} sin EAN en TN</div>`;
      statsHTML += `</div>`;
    }
  }
  document.getElementById('tnStats' + suffix).innerHTML = statsHTML;

  const tbody = document.getElementById('tnBody' + suffix);
  tbody.innerHTML = '';
  if (previewRows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:24px">No se encontraron precios diferentes para actualizar</td></tr>';
  } else {
    previewRows.forEach((p, idx) => {
      const tr = document.createElement('tr');
      const badgeColors = {
        ean: 'background:var(--green-bg);color:var(--green)',
        desc_exact: 'background:var(--accent-bg);color:var(--accent2)',
        desc_contains: 'background:var(--accent-bg);color:var(--accent2)',
        desc_fuzzy: 'background:var(--orange-bg);color:var(--orange)'
      };
      const badgeLabels = {
        ean: 'EAN',
        desc_exact: 'NOMBRE',
        desc_contains: 'CONTIENE',
        desc_fuzzy: '⚠ SIMILAR'
      };
      const badge = `<span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;${badgeColors[p.matchType]}">${badgeLabels[p.matchType]}</span>`;
      const descExtra = p.matchType !== 'ean' && p.matchDesc ? `<br><span style="font-size:11px;color:var(--text3)">→ ${p.matchDesc}</span>` : '';

      let eanWarning = '';
      if (p.matchType !== 'ean' && p.matchGestionEAN) {
        const tnHasEAN = p.tnEAN && p.tnEAN.length >= 8;
        if (tnHasEAN && p.tnEAN !== p.matchGestionEAN) {
          eanWarning = `<br><span style="font-size:11px;color:var(--orange);font-weight:600">⚠ EAN diferente — TN: ${p.ean || '(vacío)'} vs Gestión: ${p.matchGestionEAN}</span>`;
        } else if (!tnHasEAN) {
          eanWarning = `<br><span style="font-size:11px;color:var(--blue)">ℹ Sin EAN en TN — Gestión: ${p.matchGestionEAN}</span>`;
        }
      }

      tr.innerHTML = `
        <td class="td-cb"><input type="checkbox" class="cb" ${p.checked ? 'checked' : ''} onchange="tnPreviewRows[${idx}].checked=this.checked;updateTNCounter()"></td>
        <td style="font-family:'Space Mono',monospace;font-size:11px">${p.ean || '<span style="color:var(--text3)">—</span>'}</td>
        <td style="max-width:300px">${p.nombre}${descExtra}${eanWarning}</td>
        <td>${badge}</td>
        <td class="price-change price-same">$${p.oldPrecio}</td>
        <td class="price-change price-up">$${p.newPrecio}</td>
      `;
      if (p.matchType === 'desc_fuzzy') tr.style.background = 'var(--orange-bg)';
      tbody.appendChild(tr);
    });
  }

  document.getElementById('tnPreview' + suffix).scrollIntoView({ behavior: 'smooth', block: 'start' });
  toast(`${matchCount} precios actualizados en Tienda Nube`, 'success');
}

function formatTNPrice(num) {
  if (typeof num !== 'number') num = parseFloat(num) || 0;
  return num.toFixed(2);
}

function parseCSV(text, sep) {
  const rows = [];
  let current = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (c === '"' && next === '"') {
        field += '"';
        i++;
      } else if (c === '"') {
        inQuotes = false;
      } else {
        field += c;
      }
    } else {
      if (c === '"') {
        inQuotes = true;
      } else if (c === sep) {
        current.push(field);
        field = '';
      } else if (c === '\n' || (c === '\r' && next === '\n')) {
        current.push(field);
        field = '';
        if (current.length > 1 || current[0] !== '') rows.push(current);
        current = [];
        if (c === '\r') i++;
      } else {
        field += c;
      }
    }
  }
  if (field || current.length > 0) {
    current.push(field);
    if (current.length > 1 || current[0] !== '') rows.push(current);
  }
  return rows;
}

function escapeCSVField(val, sep) {
  const s = String(val);
  if (s.includes(sep) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function toggleAllTN(checked) {
  tnPreviewRows.forEach(p => p.checked = checked);
  document.querySelectorAll('#tnBody .cb, #tnBodyDirect .cb').forEach(cb => cb.checked = checked);
  updateTNCounter();
}

function updateTNCounter() {
  const count = tnPreviewRows.filter(p => p.checked).length;
  const el = document.getElementById('tnSelectionCounter');
  if (el) el.textContent = `${count} seleccionado${count !== 1 ? 's' : ''} de ${tnPreviewRows.length}`;
}

function descargarTiendaNube(direct) {
  if (tnUpdatedRows.length < 2) { toast('No hay datos para exportar.', 'warning'); return; }

  const selected = new Set();
  tnPreviewRows.forEach(p => {
    if (p.checked) selected.add(p.rowIndex);
  });

  if (selected.size === 0) {
    toast('No hay productos seleccionados para exportar.', 'warning');
    return;
  }

  const csvContent = tnUpdatedRows
    .filter((row, i) => i === 0 || selected.has(i))
    .map(row => row.map(cell => escapeCSVField(cell, tnSep)).join(tnSep))
    .join('\r\n');

  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  const encoded = new TextEncoder().encode(csvContent);
  const blob = new Blob([bom, encoded], { type: 'text/csv;charset=utf-8' });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `TiendaNube_Actualizado_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  toast(`CSV exportado con ${selected.size} cambios de precio`, 'success');
}

// ===== DIAGNOSTICO =====
function mostrarDiagnostico() {
  const mEanCol = maestroColumns[document.getElementById('selMaestroEAN').value];
  const pEanCol = provColumns[document.getElementById('selProvEAN').value];

  let info = `COLUMNAS SELECCIONADAS\n`;
  info += `Maestro EAN → "${mEanCol}" | Proveedor EAN → "${pEanCol}"\n\n`;

  info += `MAESTRO — primeros 8 valores:\n`;
  maestroData.slice(0, 8).forEach((row, i) => {
    const raw = row[mEanCol];
    const norm = normEAN(raw);
    info += `  [${i}] raw: ${JSON.stringify(raw)} (${typeof raw}) → norm: ${norm}\n`;
  });

  info += `\nPROVEEDOR — primeros 8 valores:\n`;
  provData.slice(0, 8).forEach((row, i) => {
    const raw = row[pEanCol];
    const norm = normEAN(raw);
    info += `  [${i}] raw: ${JSON.stringify(raw)} (${typeof raw}) → norm: ${norm}\n`;
  });

  const firstProvEAN = normEAN(provData[0]?.[pEanCol]);
  const found = maestroData.some(r => normEAN(r[mEanCol]) === firstProvEAN);
  info += `\nTEST: Primer EAN proveedor "${firstProvEAN}" → ${found ? '✓ ENCONTRADO' : '✗ NO ENCONTRADO'} en maestro`;

  document.getElementById('diagContent').textContent = info;
  document.getElementById('diagModal').classList.add('visible');
}

// ===== INIT =====
setTipo('juguetes');

// ===== PANEL 5 — PROVEEDORES WEB =====
function goDirectWeb() {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel5').classList.add('active');
  ['step1ind', 'step2ind', 'step3ind'].forEach(id => {
    document.getElementById(id).className = 'step-indicator';
  });
  const warning = document.getElementById('webMaestroWarning');
  if (!maestroData) {
    if (warning) warning.style.display = 'flex';
  } else {
    if (warning) warning.style.display = 'none';
    const colConfig = document.getElementById('webPriceColConfig');
    if (colConfig && maestroColumns.length > 0) {
      colConfig.style.display = 'block';
      const precioIdx = autoDetectColumn(maestroColumns, ['lista1', 'pvp', 'precio venta', 'precio_venta', 'lista 1', 'precio', 'precio compra']);
      populateSelect('selWebPrecio', maestroColumns, precioIdx);
    }
  }
  loadWebSuppliers();
}

// ===== WEB SUPPLIERS =====
let webSuppliers = [];
let webSyncResults = [];
let tnWebRawText = '';
let tnWebRows = [];
let tnWebHeaders = [];
let tnWebSep = ';';
let tnWebUpdatedRows = [];
let tnWebPreviewRows = [];

function loadWebSuppliers() {
  try { webSuppliers = JSON.parse(localStorage.getItem('pricesync_web_suppliers') || '[]'); } catch (e) { webSuppliers = []; }
  renderSupplierList();
}

function saveWebSuppliers() {
  localStorage.setItem('pricesync_web_suppliers', JSON.stringify(webSuppliers));
}

function addSupplier() {
  const name = document.getElementById('webSupplierName').value.trim();
  const url = document.getElementById('webSupplierUrl').value.trim();
  if (!name || !url) { toast('Completá nombre y URL del proveedor', 'warning'); return; }
  if (!/^https?:\/\//i.test(url)) { toast('La URL debe empezar con http:// o https://', 'warning'); return; }
  webSuppliers.push({ id: Date.now(), name, url });
  saveWebSuppliers();
  document.getElementById('webSupplierName').value = '';
  document.getElementById('webSupplierUrl').value = '';
  renderSupplierList();
  toast(`Proveedor "${name}" agregado`, 'success');
}

function removeSupplier(id) {
  webSuppliers = webSuppliers.filter(s => s.id !== id);
  saveWebSuppliers();
  renderSupplierList();
}

function renderSupplierList() {
  const el = document.getElementById('webSupplierList');
  if (!el) return;
  if (webSuppliers.length === 0) {
    el.innerHTML = '<p style="color:var(--text3);font-size:13px">No hay proveedores web registrados todavía.</p>';
    return;
  }
  el.innerHTML = webSuppliers.map(s => `
    <div style="display:flex;align-items:center;gap:10px;padding:10px 14px;background:var(--bg2);border-radius:10px;margin-bottom:8px;border:1px solid var(--border)">
      <div style="flex:1">
        <div style="font-weight:600;font-size:13px;color:var(--text1)">${s.name}</div>
        <div style="font-size:11px;color:var(--text3);word-break:break-all">${s.url}</div>
      </div>
      <button onclick="removeSupplier(${s.id})" style="background:rgba(248,113,113,.1);border:1px solid rgba(248,113,113,.3);color:#f87171;padding:5px 10px;border-radius:7px;cursor:pointer;font-size:12px">✕</button>
    </div>
  `).join('');
}

async function fetchPageHTML(url) {
  const proxies = [
    `https://corsproxy.io/?${encodeURIComponent(url)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
  ];
  for (const proxy of proxies) {
    try {
      const res = await fetch(proxy, { signal: AbortSignal.timeout(12000) });
      if (res.ok) return await res.text();
    } catch (e) {}
  }
  throw new Error('No se pudo acceder a la URL (CORS bloqueado o timeout)');
}

function logLine(msg, type) {
  const el = document.getElementById('webProgressLog');
  if (!el) return;
  const line = document.createElement('div');
  line.style.cssText = `font-size:12px;padding:2px 0;color:${type === 'error' ? '#f87171' : type === 'ok' ? 'var(--green)' : 'var(--text2)'}`;
  line.textContent = msg;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

async function startWebSync() {
  if (!maestroData) { toast('Cargá primero el Archivo Maestro en el Paso 1', 'warning'); return; }
  if (webSuppliers.length === 0) { toast('Agregá al menos un proveedor web', 'warning'); return; }

  document.getElementById('webSyncProgress').style.display = 'block';
  document.getElementById('webProgressLog').innerHTML = '';
  document.getElementById('btnWebSync').disabled = true;
  document.getElementById('webResultsSection').style.display = 'none';

  const webPriceColSel = document.getElementById('selWebPrecio');
  const mPriceCol = webPriceColSel ? maestroColumns[webPriceColSel.value] : null;

  const mDescIdx = autoDetectColumn(maestroColumns, ['descripcion', 'descripción', 'titulo', 'titulos', 'nombre', 'producto', 'detalle']);
  const mDescCol = maestroColumns[mDescIdx];
  const mEanIdx = autoDetectColumn(maestroColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar'], true);
  const mEanCol = maestroColumns[mEanIdx];

  const descMap = new Map();
  maestroData.forEach((row, idx) => {
    const desc = normDesc(row[mDescCol] || '');
    if (desc) descMap.set(desc, idx);
  });

  webSyncResults = [];
  const total = webSuppliers.length;
  let done = 0;

  for (const s of webSuppliers) {
    logLine(`🔄 ${s.name} — ${s.url}`);
    const baseUrl = s.url;
    let allScraped = [];
    try {
      const html = await fetchPageHTML(baseUrl);
      let scraped = parseWebProducts(html, s.name, baseUrl);

      // Paginación Tiendanube
      const countMatch = html.match(/LS\.productsCount\s*=\s*(\d+)/);
      if (countMatch && scraped.length > 0) {
        const totalProds = parseInt(countMatch[1]);
        const pages = Math.ceil(totalProds / scraped.length);
        if (pages > 1) {
          logLine(`   ℹ Tiendanube: ${pages} páginas (~${totalProds} productos)`);
          let emptyStreak = 0;
          for (let page = 2; page <= pages; page++) {
            await new Promise(r => setTimeout(r, 1200));
            try {
              const sep = baseUrl.includes('?') ? '&' : '?';
              const pageUrl = `${baseUrl}${sep}page=${page}`;
              const pageHtml = await fetchPageHTML(pageUrl);
              const pageScraped = parseWebProducts(pageHtml, s.name, pageUrl);
              allScraped.push(...pageScraped);
              logLine(`   📄 Pág ${page}/${pages}: ${pageScraped.length} productos`);
              if (pageScraped.length === 0) { emptyStreak++; if (emptyStreak >= 3) break; } else emptyStreak = 0;
            } catch (e) { logLine(`   ⚠ Pág ${page}: ${e.message}`, 'error'); emptyStreak++; if (emptyStreak >= 3) break; }
          }
        }
      }

      // Paginación PrestaShop
      if (scraped.length > 0 && !countMatch) {
        const psPagePattern = /[?&]page=(\d+)/g;
        let maxPage = 1, psMatch;
        while ((psMatch = psPagePattern.exec(html)) !== null) { const pg = parseInt(psMatch[1]); if (pg > maxPage) maxPage = pg; }
        if (maxPage > 1) {
          logLine(`   ℹ PrestaShop: ${maxPage} páginas`);
          let emptyStreak = 0;
          for (let page = 2; page <= maxPage; page++) {
            await new Promise(r => setTimeout(r, 1500));
            try {
              const sep = baseUrl.includes('?') ? '&' : '?';
              const pageUrl = `${baseUrl}${sep}page=${page}`;
              const pageHtml = await fetchPageHTML(pageUrl);
              const pageScraped = parseWebProducts(pageHtml, s.name, pageUrl);
              allScraped.push(...pageScraped);
              logLine(`   📄 Pág ${page}/${maxPage}: ${pageScraped.length} productos`);
              if (pageScraped.length === 0) { emptyStreak++; if (emptyStreak >= 3) break; } else emptyStreak = 0;
            } catch (e) { logLine(`   ⚠ Pág ${page}: ${e.message}`, 'error'); emptyStreak++; if (emptyStreak >= 3) break; }
          }
        }
      }

      allScraped = [...scraped, ...allScraped];
      logLine(`   ✓ ${allScraped.length} productos encontrados`, 'ok');

      allScraped.forEach(p => {
        const mIdx = descMap.get(normDesc(p.name));
        if (mIdx !== undefined) {
          const mRow = maestroData[mIdx];
          const prevPrice = mPriceCol ? parsePrice(mRow[mPriceCol]) : 0;
          webSyncResults.push({
            supplierName: s.name, supplierUrl: s.url, webName: p.name,
            descMaestro: mRow[mDescCol] || '', ean: normEAN(mRow[mEanCol]) || '',
            prevPrice, webPrice: p.price,
            variation: prevPrice > 0 ? Math.round((p.price - prevPrice) / prevPrice * 10000) / 100 : 0,
            status: p.price > prevPrice ? 'up' : p.price < prevPrice ? 'down' : 'same',
            checked: true, maestroIdx: mIdx, maestroRow: { ...mRow }
          });
        }
      });
    } catch (e) { logLine(`   ✗ Error: ${e.message}`, 'error'); }

    done++;
    document.getElementById('webProgressFill').style.width = Math.round(done / total * 100) + '%';
  }

  document.getElementById('btnWebSync').disabled = false;
  renderWebResults();
  document.getElementById('webResultsSection').style.display = 'block';
  toast(`Sincronización completa: ${webSyncResults.length} matches`, 'success');
}

function renderWebResults() {
  const counts = { up: 0, down: 0, same: 0 };
  webSyncResults.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
  document.getElementById('webStatsBar').innerHTML = `
    <div class="stat-pill up">▲ ${counts.up} subieron</div>
    <div class="stat-pill down">▼ ${counts.down} bajaron</div>
    <div class="stat-pill same">= ${counts.same} igual</div>
    <div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)">Total: ${webSyncResults.length}</div>`;
  const tbody = document.getElementById('webResultsBody');
  tbody.innerHTML = '';
  webSyncResults.forEach((r, idx) => {
    const tr = document.createElement('tr');
    const vc = r.status === 'up' ? 'price-up' : r.status === 'down' ? 'price-down' : 'price-same';
    const badge = r.status === 'up' ? '<span class="badge badge-up">▲ Subió</span>' : r.status === 'down' ? '<span class="badge badge-down">▼ Bajó</span>' : '<span class="badge badge-same">= Igual</span>';
    tr.innerHTML = `
      <td class="td-cb"><input type="checkbox" class="cb" ${r.checked ? 'checked' : ''} onchange="webSyncResults[${idx}].checked=this.checked;updateWebCounter()"></td>
      <td>${badge}</td><td style="font-size:12px;color:var(--text2)">${r.supplierName}</td>
      <td>${r.descMaestro}</td><td style="font-size:11px;color:var(--text3)">${r.webName}</td>
      <td class="td-price price-same">$${r.prevPrice.toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
      <td class="td-price ${vc}">$${r.webPrice.toLocaleString('es-AR',{minimumFractionDigits:2})}</td>
      <td class="td-price ${vc}">${r.variation > 0 ? '+' : ''}${r.variation}%</td>`;
    tbody.appendChild(tr);
  });
  updateWebCounter();
}

function updateWebCounter() {
  const count = webSyncResults.filter(r => r.checked).length;
  const el = document.getElementById('webSelectionCounter');
  if (el) el.textContent = `${count} seleccionado${count !== 1 ? 's' : ''}`;
}

function filtrarWebTabla() {
  const search = document.getElementById('webSearchInput').value.toLowerCase();
  document.getElementById('webResultsBody').querySelectorAll('tr').forEach(tr => {
    tr.style.display = tr.textContent.toLowerCase().includes(search) ? '' : 'none';
  });
}

function toggleAllWeb(checked) {
  webSyncResults.forEach(r => r.checked = checked);
  document.querySelectorAll('#webResultsBody .cb').forEach(cb => cb.checked = checked);
  updateWebCounter();
}

function exportarWebExcel() {
  const selected = webSyncResults.filter(r => r.checked);
  if (selected.length === 0) { toast('No hay productos seleccionados', 'warning'); return; }
  const rows = selected.map(r => ({ 'Proveedor': r.supplierName, 'Descripción (catálogo)': r.descMaestro, 'Nombre en web': r.webName, 'EAN': r.ean, 'Precio anterior': r.prevPrice, 'Precio web': r.webPrice, 'Variación %': r.variation, 'Estado': r.status }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb_out = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb_out, ws, 'Web Sync');
  XLSX.writeFile(wb_out, `PriceSync_Web_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast('Excel exportado', 'success');
}

function aplicarWebResultados() {
  const selected = webSyncResults.filter(r => r.checked);
  if (selected.length === 0) { toast('No hay productos seleccionados', 'warning'); return; }
  resultados = selected.map(r => ({
    maestroIdx: r.maestroIdx, maestroRow: r.maestroRow, ean: r.ean, eanDisplay: r.ean,
    codInt: '', descMaestro: r.descMaestro, descProv: r.webName,
    precioAnterior: r.prevPrice, precioNuevo: r.webPrice, pvp: r.webPrice,
    estado: r.status, variacion: r.variation, matchType: 'desc_exact', fuzzySim: 100,
    targetCol: '', checked: true
  }));
  goStep(2);
  toast(`${selected.length} productos cargados desde web → Paso 2`, 'success');
}

function onTNFileWeb(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    tnWebRawText = e.target.result;
    const firstLine = tnWebRawText.split('\n')[0];
    tnWebSep = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';
    tnWebRows = parseCSV(tnWebRawText, tnWebSep);
    if (tnWebRows.length < 2) { toast('El archivo parece vacío.', 'error'); return; }
    tnWebHeaders = tnWebRows[0];
    document.getElementById('dzTNWeb').classList.add('loaded');
    document.getElementById('dzTNWebFile').textContent = `✓ ${file.name} — ${tnWebRows.length - 1} productos`;
    document.getElementById('btnGenerarTNWeb').disabled = false;
    toast('CSV de Tienda Nube cargado', 'success');
  };
  reader.readAsText(file, 'iso-8859-1');
}

function generarTiendaNubeWeb() {
  const selected = webSyncResults.filter(r => r.checked);
  if (selected.length === 0) { toast('No hay productos seleccionados', 'warning'); return; }
  const pvpByDesc = new Map();
  selected.forEach(r => { pvpByDesc.set(normDesc(r.descMaestro), { pvp: r.webPrice, desc: r.descMaestro }); });
  const nombreIdx = tnWebHeaders.findIndex(h => h.toLowerCase().replace(/[""]/g,'').trim() === 'nombre');
  const precioIdx = tnWebHeaders.findIndex(h => h.toLowerCase().replace(/[""]/g,'').trim() === 'precio');
  if (precioIdx === -1) { toast('No se encontró columna "Precio" en el CSV', 'error'); return; }
  tnWebUpdatedRows = [tnWebHeaders]; tnWebPreviewRows = []; let matchCount = 0;
  for (let i = 1; i < tnWebRows.length; i++) {
    const row = [...tnWebRows[i]];
    while (row.length < tnWebHeaders.length) row.push('');
    const tnNombre = nombreIdx >= 0 ? String(row[nombreIdx]||'').replace(/^"|"$/g,'').trim() : '';
    const match = pvpByDesc.get(normDesc(tnNombre));
    if (match) {
      const oldPrecio = row[precioIdx]; const oldNum = parsePrice(oldPrecio);
      if (Math.abs(oldNum - match.pvp) > 0.01) {
        row[precioIdx] = formatTNPrice(match.pvp); matchCount++;
        tnWebPreviewRows.push({ rowIndex: i, checked: true, nombre: tnNombre, oldPrecio, newPrecio: row[precioIdx] });
      }
    }
    tnWebUpdatedRows.push(row);
  }
  document.getElementById('tnPreviewWeb').style.display = 'block';
  document.getElementById('tnStatsWeb').innerHTML = `<div class="stats-bar"><div class="stat-pill up">✎ ${matchCount} modificados</div><div class="stat-pill same">— ${tnWebRows.length - 1 - matchCount} sin cambio</div></div>`;
  const tbody = document.getElementById('tnBodyWeb'); tbody.innerHTML = '';
  tnWebPreviewRows.forEach((p, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="td-cb"><input type="checkbox" class="cb" ${p.checked?'checked':''} onchange="tnWebPreviewRows[${idx}].checked=this.checked;updateTNCounterWeb()"></td><td>${p.nombre}</td><td><span class="badge badge-desc">NOMBRE</span></td><td class="td-price price-same">$${p.oldPrecio}</td><td class="td-price price-up">$${p.newPrecio}</td>`;
    tbody.appendChild(tr);
  });
  updateTNCounterWeb();
  toast(`${matchCount} precios actualizados`, 'success');
}

function updateTNCounterWeb() {
  const count = tnWebPreviewRows.filter(p => p.checked).length;
  const el = document.getElementById('tnSelectionCounterWeb');
  if (el) el.textContent = `${count} seleccionado${count !== 1 ? 's' : ''}`;
}

function toggleAllTNWeb(checked) {
  tnWebPreviewRows.forEach(p => p.checked = checked);
  document.querySelectorAll('#tnBodyWeb .cb').forEach(cb => cb.checked = checked);
  updateTNCounterWeb();
}

function descargarTiendaNubeWeb() {
  const selected = new Set(tnWebPreviewRows.filter(p => p.checked).map(p => p.rowIndex));
  if (selected.size === 0) { toast('No hay seleccionados', 'warning'); return; }
  const csvContent = tnWebUpdatedRows.filter((row,i) => i===0||selected.has(i)).map(row => row.map(cell => escapeCSVField(cell, tnWebSep)).join(tnWebSep)).join('\r\n');
  const bom = new Uint8Array([0xEF,0xBB,0xBF]);
  const encoded = new TextEncoder().encode(csvContent);
  const blob = new Blob([bom, encoded], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `TiendaNube_Web_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
  toast(`CSV exportado con ${selected.size} cambios`, 'success');
}

// ===== PARSERS MULTI-PLATAFORMA =====
function detectPlatform(html, url) {
  if (html.includes('prestashop')||html.includes('PrestaShop')||html.includes('id_product')||html.includes('product-miniature')||/controller=search/.test(url||'')||html.includes('addToCartUrl')) return 'prestashop';
  if (html.includes('googleItems')||html.includes('js-item-name')||html.includes('data-product-price')||html.includes('LS.productsCount')||html.includes('tiendanube')||html.includes('Tiendanube')) return 'tiendanube';
  if (html.includes('woocommerce')||html.includes('wc-product')) return 'woocommerce';
  return 'generic';
}

function parseTiendanubeProducts(html, supplierName) {
  const products = [];
  const giStart = html.indexOf('googleItems');
  if (giStart !== -1) {
    const bracketStart = html.indexOf('[', giStart);
    if (bracketStart !== -1) {
      let depth=0, end=-1;
      for (let i=bracketStart; i<html.length; i++) { if(html[i]==='[') depth++; else if(html[i]===']'){depth--; if(depth===0){end=i+1;break;}} }
      if (end > bracketStart) { try { const items=JSON.parse(html.substring(bracketStart,end)); for(const item of items){const name=item.info?.item_name;const price=item.info?.price;if(name&&price>0)products.push({name:String(name).trim(),price:parseFloat(price)||0,supplierName});} } catch(e){} }
    }
  }
  if (products.length === 0) {
    const namePattern=/class="[^"]*js-item-name[^"]*"[^>]*>([\s\S]*?)<\/div>/gi;
    const pricePattern=/data-product-price="(\d+)"/gi;
    const names=[],prices=[];let nm;
    while((nm=namePattern.exec(html))!==null){const n=nm[1].replace(/<[^>]+>/g,'').trim();if(n)names.push(n);}
    while((nm=pricePattern.exec(html))!==null)prices.push(parseInt(nm[1])/100);
    const count=Math.min(names.length,prices.length);
    for(let i=0;i<count;i++)if(names[i]&&prices[i]>0)products.push({name:names[i],price:prices[i],supplierName});
  }
  if (products.length === 0) {
    const jlds=html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi)||[];
    for(const s of jlds){try{const j=JSON.parse(s.replace(/<\/?script[^>]*>/gi,''));if(j['@type']==='Product'||(Array.isArray(j)&&j[0]?.['@type']==='Product')){const items=Array.isArray(j)?j:[j];for(const item of items){const name=item.name;const price=item.offers?.price||item.offers?.[0]?.price;if(name&&price)products.push({name:String(name).trim(),price:parseFloat(price)||0,supplierName});}}}catch(e){}}
  }
  return deduplicateProducts(products);
}

function parsePrestaShopProducts(html, supplierName) {
  const products = [];
  const jlds=html.match(/<script[^>]+type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi)||[];
  for(const sc of jlds){try{const j=JSON.parse(sc.replace(/<\/?script[^>]*>/gi,'').trim());let items=Array.isArray(j)?j:j['@type']==='Product'?[j]:j['@graph']?j['@graph'].filter(g=>g['@type']==='Product'):j['@type']==='ItemList'?(j.itemListElement||[]).map(e=>e.item||e):[];for(const item of items){if(!item.name)continue;let price=0;if(item.offers){const offers=Array.isArray(item.offers)?item.offers:[item.offers];for(const o of offers){if(o.price){price=parseFloat(o.price)||0;break;}}}if(item.name&&price>0)products.push({name:String(item.name).trim(),price,supplierName});}}catch(e){}}
  if(products.length>0)return deduplicateProducts(products);
  const mp=/<article[^>]*class="[^"]*product-miniature[^"]*"[^>]*>([\s\S]*?)<\/article>/gi;let m;
  while((m=mp.exec(html))!==null){const block=m[1];const nm=block.match(/<h[23][^>]*class="[^"]*product-title[^"]*"[^>]*>\s*<a[^>]*title="([^"]+)"/i)||block.match(/<h[23][^>]*class="[^"]*product-title[^"]*"[^>]*>\s*<a[^>]*>([^<]+)<\/a>/i);const pm=block.match(/<span[^>]*class="[^"]*price[^"]*"[^>]*>([\s\S]*?)<\/span>/i);if(nm&&pm){const name=nm[1].trim();const price=parsePrice(pm[1].replace(/<[^>]+>/g,'').trim());if(name&&price>0)products.push({name,price,supplierName});}}
  if(products.length>0)return deduplicateProducts(products);
  return parseGenericProducts(html, supplierName);
}

function parseWooCommerceProducts(html, supplierName) {
  const products=[];
  const wp=/<li[^>]*class="[^"]*product[^"]*"[^>]*>([\s\S]*?)<\/li>/gi;let m;
  while((m=wp.exec(html))!==null){const block=m[1];const nm=block.match(/<h[23][^>]*class="[^"]*woocommerce-loop-product__title[^"]*"[^>]*>([^<]+)/i);const pm=block.match(/<span[^>]*class="[^"]*woocommerce-Price-amount[^"]*"[^>]*>([\s\S]*?)<\/span>/i)||block.match(/<span[^>]*class="[^"]*price[^"]*"[^>]*>([\s\S]*?)<\/span>/i);if(nm&&pm){const name=nm[1].trim();const price=parsePrice(pm[1].replace(/<[^>]+>/g,'').trim());if(name&&price>0)products.push({name,price,supplierName});}}
  return deduplicateProducts(products);
}

function parseGenericProducts(html, supplierName) {
  const products=[];
  const clean=html.replace(/<script[\s\S]*?<\/script>/gi,'').replace(/<style[\s\S]*?<\/style>/gi,'').replace(/<nav[\s\S]*?<\/nav>/gi,'').replace(/<header[\s\S]*?<\/header>/gi,'').replace(/<footer[\s\S]*?<\/footer>/gi,'');
  const tap=/<h[2345][^>]*>\s*<a[^>]*(?:title="([^"]+)"|>([^<]+)<\/a>)/gi;
  const found=new Map();let match;
  while((match=tap.exec(clean))!==null){const name=(match[1]||match[2]||'').trim();if(!name||name.length<3||found.has(name.toLowerCase()))continue;const snippet=clean.substring(Math.max(0,match.index-300),Math.min(clean.length,match.index+match[0].length+600));const priceM=snippet.match(/([\d]{1,3}(?:\.[\d]{3})*(?:,\d{2})?)\s*\$/)||snippet.match(/\$\s*([\d]{1,3}(?:\.[\d]{3})*(?:,\d{2})?)/)||snippet.match(/\$([\d.,]+)/);if(priceM){const price=parsePrice(priceM[1]);if(price>0){found.set(name.toLowerCase(),true);products.push({name,price,supplierName});}}}
  return deduplicateProducts(products);
}

function deduplicateProducts(products) {
  const seen=new Set();
  return products.filter(p=>{const key=normDesc(p.name);if(!key||seen.has(key))return false;seen.add(key);return true;});
}

function parseWebProducts(html, supplierName, url) {
  const platform=detectPlatform(html,url);
  let products=[];
  if(platform==='tiendanube') products=parseTiendanubeProducts(html,supplierName);
  else if(platform==='prestashop') products=parsePrestaShopProducts(html,supplierName);
  else if(platform==='woocommerce') products=parseWooCommerceProducts(html,supplierName);
  else{products=parseTiendanubeProducts(html,supplierName);if(products.length===0)products=parsePrestaShopProducts(html,supplierName);if(products.length===0)products=parseWooCommerceProducts(html,supplierName);if(products.length===0)products=parseGenericProducts(html,supplierName);}
  if(products.length===0&&platform!=='generic')products=parseGenericProducts(html,supplierName);
  return products;
}
