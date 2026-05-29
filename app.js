'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
let allRecords  = [];       // flat array: { reporteN, fecha, fechaISO, empresa, equipo, disponible, standby, averia }
let loadedFiles = new Set();// filenames loaded — duplicates ignored silently
let sortEmpresa = 1;        // 1=asc, -1=desc
let sortEquipo  = 1;
let _toastTimer = null;

// Spanish day/month abbreviations for column headers
const DAYS   = ['dom','lun','mar','mié','jue','vie','sáb'];
const MONTHS = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

// ── File loading ──────────────────────────────────────────────────────────────
async function loadFiles(fileList) {
  if (!fileList || !fileList.length) return;
  const files = Array.from(fileList).filter(f => !loadedFiles.has(f.name));
  if (!files.length) return;  // all duplicates — ignore silently

  showToast(`Procesando ${files.length} archivo(s)…`, false, 60000);
  let ok = 0, skipped = 0, errs = 0;

  for (const file of files) {
    try {
      const recs = await parseDPRFile(file);
      if (recs === null) { skipped++; continue; }
      allRecords.push(...recs);
      loadedFiles.add(file.name);
      ok++;
    } catch (err) {
      console.warn(`[${file.name}]`, err);
      errs++;
    }
  }

  rebuildFilterOptions();
  applyFilters();
  updateKPIs();

  const msg = `${ok} DPR(s) agregado(s) — ${allRecords.length} registros en total` +
    (skipped ? ` · ${skipped} sin hoja Recursos` : '') +
    (errs    ? ` · ${errs} error(es)` : '');
  showToast(msg);

  const chip = document.getElementById('statusChip');
  chip.textContent = `${loadedFiles.size} DPR${loadedFiles.size !== 1 ? 's' : ''}`;
  chip.style.display = '';

  const inp = document.getElementById('fileMore');
  if (inp) inp.value = '';
}

// ── DPR Parser ────────────────────────────────────────────────────────────────
function parseDPRFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result),
                             { type: 'array', cellDates: true });

        const sheetName = wb.SheetNames.find(n => /recursos/i.test(n));
        if (!sheetName) { resolve(null); return; }

        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName],
                                             { header: 1, defval: null });

        // Extract date & report number (search in first 12 rows)
        let reporteN = null, fecha = null;
        for (let i = 0; i < Math.min(12, raw.length); i++) {
          const row = raw[i];
          if (row.some(c => /reporte/i.test(String(c || '')))) {
            for (const c of row) {
              if (typeof c === 'number' && c > 0 && c < 10000) { reporteN = c; break; }
            }
            for (const c of row) {
              if (c instanceof Date) { fecha = c; break; }
              if (typeof c === 'number' && c > 40000) {
                fecha = new Date(Math.round((c - 25569) * 86400000)); break;
              }
            }
            break;
          }
        }
        if (!reporteN) {
          const m = file.name.match(/(\d{3,4})(?:\s*-\s*R\d+)?\s*\.xlsx?$/i);
          if (m) reporteN = parseInt(m[1], 10);
        }

        // Find equipment header row
        let hdrIdx = 58; // fixed position, fallback
        for (let i = 50; i < Math.min(75, raw.length); i++) {
          if (raw[i].some(c => /identificac/i.test(String(c || '')))) { hdrIdx = i; break; }
        }

        // Extract equipment rows (both left B-G and right J-O tables)
        const records = [];
        for (let i = hdrIdx + 1; i < Math.min(hdrIdx + 33, raw.length); i++) {
          const row = raw[i];
          if (!row) continue;
          const emp_l = str(row[2]), eq_l = str(row[3]);
          if (emp_l && eq_l) {
            records.push(makeRec(reporteN, fecha, file.name, emp_l, eq_l, row[4], row[5], row[6]));
          }
          const emp_r = str(row[10]), eq_r = str(row[11]);
          if (emp_r && eq_r) {
            records.push(makeRec(reporteN, fecha, file.name, emp_r, eq_r, row[12], row[13], row[14]));
          }
        }
        resolve(records);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('No se pudo leer el archivo'));
    reader.readAsArrayBuffer(file);
  });
}

function makeRec(reporteN, fecha, fileName, empresa, equipo, disponible, standby, averia) {
  return {
    reporteN: reporteN || '—',
    fecha,
    fechaStr: fecha ? formatDate(fecha) : '—',
    fechaISO: fecha ? toISO(fecha) : '',
    empresa, equipo,
    disponible: toNum(disponible),
    standby:    toNum(standby),
    averia:     toNum(averia),
    fileName,
  };
}

// ── Filter options (dropdowns) ────────────────────────────────────────────────
function rebuildFilterOptions() {
  const empresas = [...new Set(allRecords.map(r => r.empresa))].sort();
  const equipos  = [...new Set(allRecords.map(r => r.equipo))].sort();

  const fE = document.getElementById('fEmpresa');
  const fQ = document.getElementById('fEquipo');
  const selE = fE.value, selQ = fQ.value;

  fE.innerHTML = '<option value="">Todas</option>' +
    empresas.map(v => `<option value="${esc(v)}"${selE === v ? ' selected' : ''}>${v}</option>`).join('');
  fQ.innerHTML = '<option value="">Todos</option>' +
    equipos.map(v => `<option value="${esc(v)}"${selQ === v ? ' selected' : ''}>${v}</option>`).join('');
}

function clearFilters() {
  ['fDateFrom','fDateTo','fSearch'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('fEmpresa').value = '';
  document.getElementById('fEquipo').value  = '';
  document.getElementById('fMostrar').value = 'disponible';
  applyFilters();
}

// ── Main filter + pivot render ────────────────────────────────────────────────
function applyFilters() {
  const dateFrom = document.getElementById('fDateFrom').value;
  const dateTo   = document.getElementById('fDateTo').value;
  const fEmpresa = document.getElementById('fEmpresa').value;
  const fEquipo  = document.getElementById('fEquipo').value;
  const mostrar  = document.getElementById('fMostrar').value || 'disponible';
  const q        = document.getElementById('fSearch').value.toLowerCase().trim();

  // ── 1. Filter records ─────────────────────────────────────────────────────
  const recs = allRecords.filter(r => {
    if (dateFrom && r.fechaISO < dateFrom) return false;
    if (dateTo   && r.fechaISO > dateTo)   return false;
    if (fEmpresa && r.empresa !== fEmpresa) return false;
    if (fEquipo  && r.equipo  !== fEquipo)  return false;
    if (q && !`${r.empresa} ${r.equipo}`.toLowerCase().includes(q)) return false;
    return true;
  });

  // ── 2. Build sorted date list (columns) ───────────────────────────────────
  const datesSet = new Set(recs.map(r => r.fechaISO).filter(Boolean));
  const dates    = [...datesSet].sort();  // ascending ISO → chronological

  // ── 3. Build lookup: empresa+equipo+date → { disponible, standby, averia } ─
  const lookup = Object.create(null);
  for (const r of recs) {
    if (!r.fechaISO) continue;
    const key = `${r.empresa}\x00${r.equipo}\x00${r.fechaISO}`;
    if (!lookup[key]) lookup[key] = { disponible: 0, standby: 0, averia: 0 };
    lookup[key].disponible += r.disponible;
    lookup[key].standby    += r.standby;
    lookup[key].averia     += r.averia;
  }

  // ── 4. Build sorted pair list (rows) ──────────────────────────────────────
  const pairsMap = Object.create(null);
  for (const r of recs) pairsMap[`${r.empresa}\x00${r.equipo}`] = true;

  let pairs = Object.keys(pairsMap).map(k => {
    const [empresa, equipo] = k.split('\x00');
    return { empresa, equipo };
  });

  // Sort by empresa (with current direction) then equipo
  pairs.sort((a, b) => {
    const ec = a.empresa.localeCompare(b.empresa) * sortEmpresa;
    if (ec !== 0) return ec;
    return a.equipo.localeCompare(b.equipo) * sortEquipo;
  });

  // ── 5. Render pivot ───────────────────────────────────────────────────────
  renderPivot(pairs, dates, lookup, mostrar);
  updateRowCount(pairs.length, dates.length);
}

// ── Pivot render ──────────────────────────────────────────────────────────────
function renderPivot(pairs, dates, lookup, mostrar) {
  const thead   = document.getElementById('pivotHead');
  const tbody   = document.getElementById('tableBody');
  const emptyEl = document.getElementById('emptyMsg');
  const emptyTx = document.getElementById('emptyText');
  const table   = document.getElementById('mainTable');

  if (!pairs.length || !dates.length) {
    table.style.display = 'none';
    emptyEl.style.display = '';
    if (emptyTx) {
      emptyTx.innerHTML = allRecords.length === 0
        ? 'Haz clic en <strong>Agregar DPRs</strong> para cargar archivos'
        : 'Sin resultados para los filtros actuales';
    }
    return;
  }
  table.style.display = '';
  emptyEl.style.display = 'none';

  // ── thead ─────────────────────────────────────────────────────────────────
  const dateHdrs = dates.map(d =>
    `<th class="col-date" title="${d}">${fmtDateHdr(d)}</th>`
  ).join('');

  thead.innerHTML = `<tr>
    <th class="col-empresa" onclick="toggleSortEmpresa()">
      Empresa<span class="si ${sortEmpresa === 1 ? 'sort-asc' : 'sort-desc'}"></span>
    </th>
    <th class="col-equipo" onclick="toggleSortEquipo()">
      Listado maquinaría<span class="si ${sortEquipo === 1 ? 'sort-asc' : 'sort-desc'}"></span>
    </th>
    ${dateHdrs}
  </tr>`;

  // ── tbody — group empresa with rowspan + expand/collapse ─────────────────
  const empresaCount = Object.create(null);
  pairs.forEach(({ empresa }) => { empresaCount[empresa] = (empresaCount[empresa] || 0) + 1; });

  const empresaOrder = [];
  const seenE = new Set();
  pairs.forEach(({ empresa }) => { if (!seenE.has(empresa)) { empresaOrder.push(empresa); seenE.add(empresa); } });
  const grpIdx = Object.create(null);
  empresaOrder.forEach((e, i) => { grpIdx[e] = i; });

  const emittedEmpresa = new Set();
  const valCls = mostrar === 'standby' ? 'v-warn' : mostrar === 'averia' ? 'v-err' : 'v-ok';

  tbody.innerHTML = pairs.map(({ empresa, equipo }) => {
    const gid  = grpIdx[empresa];
    const grp  = gid % 2 === 0 ? 'grp-even' : 'grp-odd';

    let empresaTd = '';
    if (!emittedEmpresa.has(empresa)) {
      emittedEmpresa.add(empresa);
      const span = empresaCount[empresa];
      empresaTd = `<td id="eg-${gid}" class="col-empresa empresa-group" rowspan="${span}" data-span="${span}" data-collapsed="0"><button class="toggle-btn" onclick="toggleEmpresa(${gid})" title="Expandir/colapsar"><span class="toggle-icon">▼</span></button>${esc(empresa)}<span class="empresa-count">${span}</span></td>`;
    }

    const cells = dates.map(d => {
      const val = lookup[`${empresa}\x00${equipo}\x00${d}`];
      const n   = val ? val[mostrar] : 0;
      return n > 0
        ? `<td class="date-cell ${valCls}">${n}</td>`
        : `<td class="date-cell"></td>`;
    }).join('');

    return `<tr class="${grp}" data-gid="${gid}">${empresaTd}<td class="col-equipo">${equipo}</td>${cells}</tr>`;
  }).join('');
}

// ── Expand / Collapse ─────────────────────────────────────────────────────────
function toggleEmpresa(gid) {
  const cell = document.getElementById('eg-' + gid);
  if (!cell) return;
  const rows      = [...document.querySelectorAll(`tr[data-gid="${gid}"]`)];
  const collapsed = cell.dataset.collapsed === '1';

  if (collapsed) {
    // EXPAND: show rows first, then restore rowspan
    rows.slice(1).forEach(r => { r.style.display = ''; });
    cell.rowSpan = +cell.dataset.span;
  } else {
    // COLLAPSE: shrink rowspan first, then hide rows
    cell.rowSpan = 1;
    rows.slice(1).forEach(r => { r.style.display = 'none'; });
  }

  cell.dataset.collapsed = collapsed ? '0' : '1';
  const icon = cell.querySelector('.toggle-icon');
  if (icon) icon.textContent = collapsed ? '▼' : '▶';
}

function collapseAll() {
  document.querySelectorAll('.empresa-group[data-collapsed="0"]').forEach(cell => {
    toggleEmpresa(+cell.id.replace('eg-', ''));
  });
}

function expandAll() {
  document.querySelectorAll('.empresa-group[data-collapsed="1"]').forEach(cell => {
    toggleEmpresa(+cell.id.replace('eg-', ''));
  });
}

// ── Sort toggles ──────────────────────────────────────────────────────────────
function toggleSortEmpresa() { sortEmpresa = -sortEmpresa; applyFilters(); }
function toggleSortEquipo()  { sortEquipo  = -sortEquipo;  applyFilters(); }

// ── KPI ───────────────────────────────────────────────────────────────────────
function updateKPIs() {
  const fechas = allRecords.map(r => r.fechaISO).filter(Boolean).sort();
  const range  = fechas.length
    ? `${formatDate(new Date(fechas[0]))} → ${formatDate(new Date(fechas[fechas.length - 1]))}`
    : '—';
  document.getElementById('kDprs').textContent  = loadedFiles.size || '—';
  document.getElementById('kRange').textContent = range;
  document.getElementById('kEquip').textContent = allRecords.length || '—';
  document.getElementById('kDisp').textContent  = allRecords.reduce((s, r) => s + r.disponible, 0) || '—';
  document.getElementById('kStby').textContent  = allRecords.reduce((s, r) => s + r.standby,    0) || '—';
  document.getElementById('kAver').textContent  = allRecords.reduce((s, r) => s + r.averia,     0) || '—';
}

function updateRowCount(rows, cols) {
  document.getElementById('rowCount').textContent =
    allRecords.length ? `${rows} combinações · ${cols} datas` : '';
}

// ── Export ────────────────────────────────────────────────────────────────────
function exportExcel() {
  if (!allRecords.length) { showToast('Sin datos para exportar', true); return; }
  showToast('Generando Excel…', false, 10000);
  try {
    // Re-compute current pivot data for export
    const dateFrom = document.getElementById('fDateFrom').value;
    const dateTo   = document.getElementById('fDateTo').value;
    const fEmpresa = document.getElementById('fEmpresa').value;
    const fEquipo  = document.getElementById('fEquipo').value;
    const mostrar  = document.getElementById('fMostrar').value || 'disponible';
    const q        = document.getElementById('fSearch').value.toLowerCase().trim();

    const recs = allRecords.filter(r => {
      if (dateFrom && r.fechaISO < dateFrom) return false;
      if (dateTo   && r.fechaISO > dateTo)   return false;
      if (fEmpresa && r.empresa !== fEmpresa) return false;
      if (fEquipo  && r.equipo  !== fEquipo)  return false;
      if (q && !`${r.empresa} ${r.equipo}`.toLowerCase().includes(q)) return false;
      return true;
    });

    const datesSet = new Set(recs.map(r => r.fechaISO).filter(Boolean));
    const dates    = [...datesSet].sort();
    const lookup   = Object.create(null);
    for (const r of recs) {
      if (!r.fechaISO) continue;
      const key = `${r.empresa}\x00${r.equipo}\x00${r.fechaISO}`;
      if (!lookup[key]) lookup[key] = { disponible: 0, standby: 0, averia: 0 };
      lookup[key].disponible += r.disponible;
      lookup[key].standby    += r.standby;
      lookup[key].averia     += r.averia;
    }
    const pairsMap = Object.create(null);
    for (const r of recs) pairsMap[`${r.empresa}\x00${r.equipo}`] = true;
    const pairs = Object.keys(pairsMap)
      .map(k => { const [empresa, equipo] = k.split('\x00'); return { empresa, equipo }; })
      .sort((a, b) => a.empresa.localeCompare(b.empresa) || a.equipo.localeCompare(b.equipo));

    // Build AOA
    const header = ['Empresa', 'Listado maquinaría', ...dates.map(fmtDateHdr)];
    const rows   = pairs.map(({ empresa, equipo }) => {
      const cells = dates.map(d => {
        const val = lookup[`${empresa}\x00${equipo}\x00${d}`];
        return val ? (val[mostrar] || '') : '';
      });
      return [empresa, equipo, ...cells];
    });

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    // Column widths
    ws['!cols'] = [{ wch: 20 }, { wch: 26 }, ...dates.map(() => ({ wch: 12 }))];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Equipos_Pivot');
    XLSX.writeFile(wb, 'Equipos_Maquinaria_Pivot.xlsx');
    showToast('Excel exportado con éxito');
  } catch (err) {
    showToast('Error al exportar: ' + err.message, true);
  }
}

// ── Theme ─────────────────────────────────────────────────────────────────────
function toggleTheme() {
  const dark = document.documentElement.getAttribute('data-theme') === 'dark';
  document.documentElement.setAttribute('data-theme', dark ? '' : 'dark');
  document.querySelector('#themeBtn i').className = dark ? 'bi bi-moon-fill' : 'bi bi-sun-fill';
  localStorage.setItem('eq-theme', dark ? '' : 'dark');
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function showToast(msg, isErr, dur) {
  clearTimeout(_toastTimer);
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = 'toast' + (isErr ? ' err' : '');
  el.style.display = '';
  _toastTimer = setTimeout(() => { el.style.display = 'none'; }, dur || 3800);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function str(v)    { return String(v == null ? '' : v).trim(); }
function toNum(v)  { const n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n); }
function esc(s)    { return s.replace(/[&"<>]/g, c => ({'&':'&amp;','"':'&quot;','<':'&lt;','>':'&gt;'}[c])); }

// Use UTC methods to avoid timezone off-by-one (Excel dates are stored as UTC midnight)
function toISO(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const y  = d.getUTCFullYear();
  const m  = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

function formatDate(d) {
  if (!(d instanceof Date) || isNaN(d)) return '—';
  const day   = String(d.getUTCDate()).padStart(2, '0');
  const month = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${day}/${month}/${d.getUTCFullYear()}`;
}

/** Format ISO date as pivot column header: "02-02-2026" */
function fmtDateHdr(iso) {
  const d  = new Date(iso + 'T12:00:00');
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}-${mm}-${d.getFullYear()}`;
}

// ── Init ──────────────────────────────────────────────────────────────────────
(function init() {
  const t = localStorage.getItem('eq-theme') || '';
  if (t) {
    document.documentElement.setAttribute('data-theme', t);
    document.querySelector('#themeBtn i').className = 'bi bi-sun-fill';
  }
  // Show initial empty state
  document.getElementById('mainTable').style.display = 'none';
})();
