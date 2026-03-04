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

// ===== TIPO SELECTOR =====
function setTipo(tipo) {
  tipoProducto = tipo;
  document.getElementById('btnJuguetes').className = 'tipo-btn' + (tipo === 'juguetes' ? ' active-juguetes' : '');
  document.getElementById('btnLibros').className = 'tipo-btn' + (tipo === 'libros' ? ' active-libros' : '');
  // Hide IVA card for libros (PVP is final price)
  document.getElementById('ivaCard').style.display = tipo === 'libros' ? 'none' : 'block';
}

// ===== FILE LOADING =====
function normEAN(val) {
  if (val == null || val === '') return '';
  if (typeof val === 'number') {
    val = val.toFixed(0);
  } else {
    val = String(val).trim();
    // Handle scientific notation: "6.53E+16" → full number
    if (/^\d+\.?\d*[eE][+\-]?\d+$/.test(val)) {
      try { val = BigInt(Math.round(parseFloat(val))).toString(); } catch (e) { val = parseFloat(val).toFixed(0); }
    }
    // Remove trailing .0 from string representation of numbers
    if (/^\d+\.0+$/.test(val)) val = val.replace(/\.0+$/, '');
  }
  // Strip dashes, spaces, dots
  val = val.replace(/[\s\-\.]/g, '');
  // Keep only digits
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


// Compact normalization: remove ALL spaces (catches "Fonoloco" vs "FONO LOCO")
function normDescCompact(val) {
  if (!val) return '';
  return String(val).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}


function readExcelBuffer(buffer, fileName, headerRow, sheetName) {
  const isXLS = /\.xls$/i.test(fileName) && !/\.xlsx$/i.test(fileName);
  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    wb = XLSX.read(binary, { type: 'binary' });
  } else {
    wb = XLSX.read(buffer, { type: 'array' });
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
  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    wb = XLSX.read(binary, { type: 'binary' });
  } else {
    wb = XLSX.read(buffer, { type: 'array' });
  }
  return wb.SheetNames;
}

function autoDetectHeaderRow(buffer, fileName, sheetName) {
  const isXLS = /\.xls$/i.test(fileName) && !/\.xlsx$/i.test(fileName);
  let wb;
  if (isXLS) {
    const data = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
    wb = XLSX.read(binary, { type: 'binary' });
  } else {
    wb = XLSX.read(buffer, { type: 'array' });
  }
  const wsName = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[wsName];
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Find first row with 3+ non-empty text cells (not pure numbers),
  // where ALL non-empty cells are text (to skip data rows),
  // followed by a row with actual data
  for (let i = 0; i < Math.min(allRows.length - 1, 25); i++) {
    const row = allRows[i];
    if (!row) continue;
    const nonEmpty = row.filter(c => c != null && String(c).trim().length > 0);
    if (nonEmpty.length < 3) continue;

    const textCells = nonEmpty.filter(c => isNaN(Number(c)));
    // At least 3 text cells and majority should be text (header row = mostly labels)
    if (textCells.length < 3 || textCells.length < nonEmpty.length * 0.6) continue;

    // Check next row has data (and at least some non-text content like numbers)
    const nextRow = allRows[i + 1];
    if (!nextRow) continue;
    const nextNonEmpty = nextRow.filter(c => c != null && String(c).trim().length > 0);
    if (nextNonEmpty.length >= 2) return i + 1; // 1-based
  }
  return 1;
}

function autoDetectColumn(columns, keywords, prioritizeFirst = false) {
  const kws = keywords.map(k => k.toLowerCase());

  // Exact match first
  for (const kw of kws) {
    for (let i = 0; i < columns.length; i++) {
      if (columns[i].toLowerCase() === kw) return i;
    }
  }

  // Then partial match
  let found = [];
  for (let i = 0; i < columns.length; i++) {
    const col = columns[i].toLowerCase();
    for (const kw of kws) {
      if (col.includes(kw) || kw.includes(col)) {
        found.push(i);
        break;
      }
    }
  }
  if (found.length > 0) {
    return prioritizeFirst ? found[0] : found[0];
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

      // Autodetect columns - prioritize first column for EAN
      const eanIdx = autoDetectColumn(maestroColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar'], true);
      const descIdx = autoDetectColumn(maestroColumns, ['descripcion', 'descripción', 'titulo', 'titulos', 'nombre', 'producto', 'detalle']);
      const precioIdx = autoDetectColumn(maestroColumns, ['precio compra', 'precio_compra', 'costo', 'precio costo', 'lista1', 'precio', 'pvp']);
      const codIntIdx = autoDetectColumn(maestroColumns, ['codigoprov', 'codigo_prov', 'cod_prov', 'sku', 'codigo interno', 'artículo', 'articulo']);

      populateSelect('selMaestroEAN', maestroColumns, eanIdx);
      populateSelect('selMaestroDesc', maestroColumns, descIdx);
      populateSelect('selMaestroPrecio', maestroColumns, precioIdx);
      populateSelect('selMaestroCodInt', maestroColumns, codIntIdx);

      checkReady();
    } catch (err) {
      console.error('Error loading maestro file:', err);
      alert('Error al leer el archivo maestro:\n' + err.message);
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

      // Check for multiple sheets
      const sheets = getSheetNames(provRawBuffer, provFileName);
      const sheetSelect = document.getElementById('selProvSheet');

      if (sheets.length > 1) {
        // Show sheet selector
        sheetSelect.innerHTML = '';
        sheets.forEach((name, i) => {
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
      alert('Error al leer el archivo del proveedor:\n' + err.message);
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
  const result = readExcelBuffer(provRawBuffer, provFileName, headerRow, sheetName);

  // Validate: if most headers are empty Col_X, the header row is probably wrong
  const realHeaders = result.columns.filter(h => !h.startsWith('Col_'));
  if (realHeaders.length < 2 && result.columns.length > 0) {
    // Try next few rows
    for (let tryRow = headerRow + 1; tryRow <= headerRow + 5 && tryRow <= 25; tryRow++) {
      const retry = readExcelBuffer(provRawBuffer, provFileName, tryRow, sheetName);
      const retryReal = retry.columns.filter(h => !h.startsWith('Col_'));
      if (retryReal.length >= 2) {
        headerRow = tryRow;
        result.data = retry.data;
        result.columns = retry.columns;
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

  // Build maestro index by EAN and description
  const maestroByEAN = new Map();
  const maestroByDesc = new Map();
  const maestroByDescCompact = new Map(); // no-space version for "Fonoloco" vs "FONO LOCO"
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

  // Auto-detect if provider has real barcodes (8-14 digit numbers)
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
      if (ivaMode === 'neto') {
        precioNuevo = pPrecioRaw * (1 + ivaRate);
      } else {
        precioNuevo = pPrecioRaw;
      }
    }

    // Matching en 2 niveles estrictos:
    // 1) EAN (código de barras)
    // 2) Descripción exacta

    let mIdx = -1;
    let matchType = 'notfound';
    let fuzzySim = 0;

    // Step 1: EAN
    if (provHasEAN && pEan && maestroByEAN.has(pEan)) {
      mIdx = maestroByEAN.get(pEan);
      matchType = 'ean';
    }

    // Step 2: Exact description
    if (mIdx < 0) {
      const pDescNorm = normDesc(pDesc);
      const pDescCompact = normDescCompact(pDesc);

      // 2a: Exact normalized match
      if (pDescNorm && maestroByDesc.has(pDescNorm)) {
        mIdx = maestroByDesc.get(pDescNorm);
        matchType = 'desc_exact';
      }

      // 2b: Exact compact match (no spaces: "Fonoloco" = "FONO LOCO")
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

      // Calculate PVP
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

// ===== RENDER REVIEW =====
function renderReview() {
  // Banner
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
    bannerHTML += '<br><span style="font-size:12px;opacity:0.85">⚠ Proveedor sin códigos de barras — matching automático por nombre de producto</span>';
  }
  banner.innerHTML = bannerHTML;

  // Stats
  const counts = { up: 0, down: 0, same: 0, notfound: 0 };
  resultados.forEach(r => {
    if (r.estado !== 'notfound') counts[r.estado]++;
    else counts.notfound++;
  });

  const statsBar = document.getElementById('statsBar');
  statsBar.innerHTML = `
    <div class="stat-pill up" onclick="setFiltro('up')">▲ ${counts.up} subieron</div>
    <div class="stat-pill down" onclick="setFiltro('down')">▼ ${counts.down} bajaron</div>
    <div class="stat-pill same" onclick="setFiltro('same')">= ${counts.same} sin cambio</div>
    <div class="stat-pill notfound" onclick="setFiltro('notfound')">✗ ${counts.notfound} sin match</div>
    <div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)" onclick="setFiltro(null)">Todos (${resultados.length})</div>
  `;


  // Table header
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
    // Filter
    // Filter
    if (filtroActual && r.estado !== filtroActual) return;
    if (search) {
      const s = `${r.ean} ${r.descMaestro} ${r.descProv} ${r.codInt}`.toLowerCase();
      if (!s.includes(search)) return;
    }

    const rowClass = r.estado === 'notfound' ? 'row-notfound' : '';

    let matchBadge = '';
    if (r.matchType === 'ean') matchBadge = '<span class="badge badge-ean">EAN</span>';
    else if (r.matchType === 'desc_exact') matchBadge = '<span class="badge badge-desc">DESC. EXACTA</span>';
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
      <td class="td-cb">${r.estado === 'notfound' ? '<span style="color:var(--text3)">—</span>' : `<input type="checkbox" class="cb" ${r.checked ? 'checked' : ''} onchange="resultados[${i}].checked=this.checked">`}</td>
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
}

function setFiltro(f) {
  filtroActual = filtroActual === f ? null : f;
  document.querySelectorAll('.stat-pill').forEach(p => p.classList.remove('active'));
  if (filtroActual) {
    const pills = document.querySelectorAll('.stat-pill');
    pills.forEach(p => {
      if (p.classList.contains(filtroActual) || (filtroActual === null && p.textContent.includes('Todos'))) {
        p.classList.add('active');
      }
    });
  }
  renderTableRows();
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
  document.getElementById('exportInfo').textContent = `Se exportarán ${selected.length} productos (solo los matcheados y seleccionados). Los productos de otros proveedores no se tocan.`;
}

// ===== EXPORT INTERNO =====
function exportarInterno() {
  const selected = resultados.filter(r => r.checked && r.estado !== 'notfound');
  if (selected.length === 0) { alert('No hay productos seleccionados para exportar.'); return; }

  const mPrecioCol = maestroColumns[document.getElementById('selMaestroPrecio').value];

  // Build rows preserving original column order
  const exportRows = selected.map(r => {
    const row = {};
    // Copy all original columns in order
    maestroColumnOrder.forEach(col => {
      row[col] = r.maestroRow[col];
    });

    // Update the target price column
    if (tipoProducto === 'libros') {
      const lista1Col = findLista1Col();
      if (lista1Col) row[lista1Col] = r.precioNuevo;
    } else {
      row[mPrecioCol] = r.precioNuevo;
    }

    // Add audit columns at end
    row['_PrecioAnterior'] = r.precioAnterior;
    row['_PrecioNuevo'] = r.precioNuevo;
    row['_Variacion%'] = r.variacion;
    row['_Estado'] = r.estado;
    row['_Match'] = r.matchType;
    row['_Fecha'] = new Date().toLocaleDateString('es-AR');

    return row;
  });

  // Ensure column order
  const headers = [...maestroColumnOrder, '_PrecioAnterior', '_PrecioNuevo', '_Variacion%', '_Estado', '_Match', '_Fecha'];

  const ws = XLSX.utils.json_to_sheet(exportRows, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Actualización');
  XLSX.writeFile(wb, `PriceSync_Interno_${new Date().toISOString().slice(0, 10)}.xlsx`);
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

    // Autodetect columns
    const eanIdx = autoDetectColumn(gestionColumns, ['codigo', 'ean', 'barras', 'isbn', 'upc', 'código', 'codbar'], true);
    const descIdx = autoDetectColumn(gestionColumns, ['descripcion', 'descripción', 'nombre', 'producto', 'detalle', 'articulo']);
    const pvpIdx = autoDetectColumn(gestionColumns, ['lista1', 'pvp', 'precio venta', 'precio_venta', 'precio publico', 'precio']);

    populateSelect('selGestionEAN' + suffix, gestionColumns, eanIdx);
    populateSelect('selGestionDesc' + suffix, gestionColumns, descIdx);
    populateSelect('selGestionPVP' + suffix, gestionColumns, pvpIdx);

    checkTNReady(direct);
  };
  reader.readAsArrayBuffer(file);
}

function onTNFile(file, direct) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    tnRawText = e.target.result;

    // Detect separator
    const firstLine = tnRawText.split('\n')[0];
    tnSep = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';

    tnRows = parseCSV(tnRawText, tnSep);
    if (tnRows.length < 2) { alert('El archivo parece vacío.'); return; }
    tnHeaders = tnRows[0];

    const suffix = direct ? 'Direct' : '';
    document.getElementById('dzTN' + suffix).classList.add('loaded');
    document.getElementById('dzTN' + suffix + 'File').textContent = `✓ ${file.name} — ${tnRows.length - 1} productos`;

    checkTNReady(direct);
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

  // Build lookup: EAN → {pvp, desc} from gestión file
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

  // Find key columns in TN CSV
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
    alert('No se encontró la columna "Precio" en el CSV de Tienda Nube.');
    return;
  }

  // Process: copy all rows, update Precio where matched
  tnUpdatedRows = [tnHeaders];
  let matchCount = 0;
  let previewRows = [];
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

    // Step 1: Match by EAN
    if (ean && pvpByEAN.has(ean)) {
      newPVP = pvpByEAN.get(ean).pvp;
      matchType = 'ean';
      matchDesc = pvpByEAN.get(ean).desc;
      matchGestionEAN = ean;
    }

    // Step 2: Match by name (fallback)
    if (newPVP === null && tnNombre) {
      const tnDescNorm = normDesc(tnNombre);
      const tnDescCompact = normDescCompact(tnNombre);

      // 2a: Exact normalized
      if (tnDescNorm && gestionByDesc.has(tnDescNorm)) {
        const m = gestionByDesc.get(tnDescNorm);
        newPVP = m.pvp;
        matchType = 'desc_exact';
        matchDesc = m.desc;
        matchGestionEAN = m.ean || '';
      }

      // 2b: Exact compact (no spaces)
      if (newPVP === null && tnDescCompact && gestionByDescCompact.has(tnDescCompact)) {
        const m = gestionByDescCompact.get(tnDescCompact);
        newPVP = m.pvp;
        matchType = 'desc_exact';
        matchDesc = m.desc;
        matchGestionEAN = m.ean || '';
      }

      // 2c: Containment
      if (newPVP === null && tnDescNorm && tnDescNorm.length >= 3) {
        for (const entry of gestionDescList) {
          if (entry.descNorm.includes(tnDescNorm) || tnDescNorm.includes(entry.descNorm)) {
            const shorter = Math.min(entry.descNorm.length, tnDescNorm.length);
            const longer = Math.max(entry.descNorm.length, tnDescNorm.length);
            if (shorter / longer >= 0.4) {
              newPVP = entry.pvp;
              matchType = 'desc_contains';
              matchDesc = entry.desc;
              matchGestionEAN = entry.ean || '';
              break;
            }
          }
          // Compact containment
          if (tnDescCompact && entry.descCompact) {
            if (entry.descCompact.includes(tnDescCompact) || tnDescCompact.includes(entry.descCompact)) {
              const shorter = Math.min(entry.descCompact.length, tnDescCompact.length);
              const longer = Math.max(entry.descCompact.length, tnDescCompact.length);
              if (shorter / longer >= 0.4) {
                newPVP = entry.pvp;
                matchType = 'desc_contains';
                matchDesc = entry.desc;
                matchGestionEAN = entry.ean || '';
                break;
              }
            }
          }
        }
      }

      // 2d: Fuzzy (combined similarity)
      if (newPVP === null && tnDescNorm) {
        let bestSim = 0, bestEntry = null;
        for (const entry of gestionDescList) {
          const sim = combinedSimilarity(tnNombre, entry.desc);
          if (sim > bestSim) { bestSim = sim; bestEntry = entry; }
        }
        if (bestSim >= 0.40 && bestEntry) {
          newPVP = bestEntry.pvp;
          matchType = 'desc_fuzzy';
          matchDesc = bestEntry.desc;
          matchGestionEAN = bestEntry.ean || '';
        }
      }
    }

    // Update price if matched and price changed
    if (newPVP !== null) {
      const oldPrecio = row[precioIdx];
      const oldPrecioNum = parsePrice(oldPrecio);
      if (Math.abs(oldPrecioNum - newPVP) > 0.01) {
        row[precioIdx] = formatTNPrice(newPVP);
        matchCount++;
        statsByType[matchType] = (statsByType[matchType] || 0) + 1;
        previewRows.push({
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

  // Show preview
  document.getElementById('tnPreview' + suffix).style.display = 'block';

  // Stats with match type breakdown
  let statsHTML = `
    <div class="stats-bar">
      <div class="stat-pill up">✎ ${matchCount} precios modificados</div>
      <div class="stat-pill same">— ${tnRows.length - 1 - matchCount} sin cambio</div>
      <div class="stat-pill" style="background:var(--accent-bg);color:var(--accent2)">${tnRows.length - 1} total en CSV</div>
    </div>`;
  if (matchCount > 0) {
    // Count barcode mismatches
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
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--text3);padding:24px">No se encontraron precios diferentes para actualizar</td></tr>';
  } else {
    previewRows.forEach(p => {
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
      const badge = `<span style="display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:600;${badgeColors[p.matchType]}">${badgeLabels[p.matchType]}</span>`;

      // Show description match info
      const descExtra = p.matchType !== 'ean' && p.matchDesc ? `<br><span style="font-size:11px;color:var(--text3)">→ ${p.matchDesc}</span>` : '';

      // Show barcode mismatch warning when matched by name
      let eanWarning = '';
      if (p.matchType !== 'ean' && p.matchGestionEAN) {
        const tnHasEAN = p.tnEAN && p.tnEAN.length >= 8;
        if (tnHasEAN && p.tnEAN !== p.matchGestionEAN) {
          // Both have barcodes but they're different
          eanWarning = `<br><span style="font-size:11px;color:var(--orange);font-weight:600">⚠ EAN diferente — TN: ${p.ean || '(vacío)'} vs Gestión: ${p.matchGestionEAN}</span>`;
        } else if (!tnHasEAN) {
          // TN has no barcode, gestión does
          eanWarning = `<br><span style="font-size:11px;color:var(--blue)">ℹ Sin EAN en TN — Gestión: ${p.matchGestionEAN}</span>`;
        }
      }

      tr.innerHTML = `
        <td style="font-family:'Space Mono',monospace;font-size:12px">${p.ean || '<span style="color:var(--text3)">—</span>'}</td>
        <td style="max-width:300px">${p.nombre}${descExtra}${eanWarning}</td>
        <td>${badge}</td>
        <td class="price-change price-same">$${p.oldPrecio}</td>
        <td class="price-change price-up">$${p.newPrecio}</td>
      `;
      if (p.matchType === 'desc_fuzzy') tr.style.background = 'var(--orange-bg)';
      tbody.appendChild(tr);
    });
  }

  // Scroll to preview
  document.getElementById('tnPreview' + suffix).scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function formatTNPrice(num) {
  if (typeof num !== 'number') num = parseFloat(num) || 0;
  const parts = num.toFixed(2).split('.');
  const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return intPart + '.' + parts[1];
}

// Smart price parser: handles "72,900.00" (comma=thousands) and "1234,56" (comma=decimal)
function parsePrice(val) {
  if (val == null) return 0;
  let s = String(val).replace(/[^0-9.,]/g, '');
  if (!s) return 0;
  // If both comma and dot exist: determine which is thousands vs decimal
  const hasComma = s.includes(',');
  const hasDot = s.includes('.');
  if (hasComma && hasDot) {
    // "72,900.00" → comma is thousands (comes before dot)
    // "72.900,00" → dot is thousands (comes before comma)
    const lastComma = s.lastIndexOf(',');
    const lastDot = s.lastIndexOf('.');
    if (lastComma > lastDot) {
      // comma is decimal: "72.900,00" → remove dots, replace comma with dot
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      // dot is decimal: "72,900.00" → remove commas
      s = s.replace(/,/g, '');
    }
  } else if (hasComma) {
    // Only comma: could be thousands ("72,900") or decimal ("1234,56")
    // If comma is followed by exactly 2 digits at end → decimal
    const match = s.match(/,(\d+)$/);
    if (match && match[1].length === 2) {
      s = s.replace(',', '.');
    } else {
      // Thousands separator: remove it
      s = s.replace(/,/g, '');
    }
  }
  // hasDot only → standard format, parse directly
  return parseFloat(s) || 0;
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

function descargarTiendaNube() {
  if (tnUpdatedRows.length < 2) { alert('No hay datos para exportar.'); return; }

  const csvContent = tnUpdatedRows.map(row =>
    row.map(cell => escapeCSVField(cell, tnSep)).join(tnSep)
  ).join('\r\n');

  // UTF-8 BOM for compatibility
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  const encoded = new TextEncoder().encode(csvContent);
  const blob = new Blob([bom, encoded], { type: 'text/csv;charset=utf-8' });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `TiendaNube_Actualizado_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
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

  // Test first match
  const firstProvEAN = normEAN(provData[0]?.[pEanCol]);
  const found = maestroData.some(r => normEAN(r[mEanCol]) === firstProvEAN);
  info += `\nTEST: Primer EAN proveedor "${firstProvEAN}" → ${found ? '✓ ENCONTRADO' : '✗ NO ENCONTRADO'} en maestro`;

  document.getElementById('diagContent').textContent = info;
  document.getElementById('diagModal').classList.add('visible');
}

// ===== INIT =====
setTipo('juguetes');
