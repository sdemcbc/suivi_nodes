/* ============================================================
   SITE DOWN DASHBOARD — app.js
   Logique : import Excel, filtres, tri, KPI, export XLSX
   ============================================================ */

// ── État global ──────────────────────────────────────────────
let allData = [];   // données brutes parsées
let filtered = [];   // données après filtres
let sortCol = '';
let sortDir = 'asc';
let maxSeconds = 0;    // durée max en secondes (pour le slider)

// ── Colonnes attendues (insensible à la casse) ───────────────
const COL_MAP = {
  location: ['location', 'loc', 'ville', 'city', 'region', 'site location'],
  technology: ['technology', 'technologie', 'tech', 'type', 'network type', 'rat'],
  site_id: ['site id', 'siteid', 'id', 'site_id', 'site number', 'code site'],
  site_name: ['site name', 'sitename', 'nom site', 'site_name', 'name'],
  duration: ['duration', 'durée', 'duree', 'outage duration', 'down time', 'downtime', 'coupure'],
  last_occurred: ['last occurred on', 'last occurred', 'dernière occurrence', 'last outage', 'date coupure', 'occurred on'],
  power_type: ['power type', 'powertype', 'power', 'énergie', 'energie', 'type énergie', 'power_type'],
  mail_received: ['mail received time', 'mail received', 'date réception', 'date mail', 'mail received time'],
};

let reportTimestamp = '—'; // Timestamp global du rapport (Mail Received Time)

// ── État global VSWR ──────────────────────────────────────────
let vswrAllData = [];
let vswrFiltered = [];
let vswrReportTimestamp = '—';

const VSWR_COL_MAP = {
  site_id: ['site id', 'siteid', 'id', 'site_id', 'site number', 'code site'],
  site_name: ['site name', 'sitename', 'nom site', 'site_name', 'name'],
  antenna: ['antenna', 'sector', 'antenne', 'secteur', 'cell', 'cellule', 'ant', 'sect'],
  vswr_value: ['vswr', 'vswr value', 'valeur vswr', 'vswr_value', 'ratio'],
  check_date: ['date', 'check date', 'date check', 'date mesure', 'test date'],
  mail_received: ['mail received time', 'mail received', 'date réception', 'date mail'],
};

// ── Utilitaires ──────────────────────────────────────────────

/**
 * Convertit une valeur de durée en SECONDES totales.
 * Reconnaît : hh:mm, hh:mm:ss, nombre décimal (heures), valeur Excel fraction.
 */
function toSeconds(val) {
  if (val === null || val === undefined || val === '') return 0;
  const s = String(val).trim();

  // Format hh:mm:ss (précision complète)
  const hhmmss = s.match(/^(\d+):(\d{2}):(\d{2})$/);
  if (hhmmss)
    return parseInt(hhmmss[1], 10) * 3600
      + parseInt(hhmmss[2], 10) * 60
      + parseInt(hhmmss[3], 10);

  // Format hh:mm (pas de secondes)
  const hhmm = s.match(/^(\d+):(\d{2})$/);
  if (hhmm)
    return parseInt(hhmm[1], 10) * 3600
      + parseInt(hhmm[2], 10) * 60;

  // Valeur décimale :
  //   • 0 < n < 1  → fraction de jour Excel (ex: 0.1528 = 03:40:10) → ×86400
  //   • n >= 1     → nombre d'heures (ex: 3.67 h)                   → ×3600
  const n = parseFloat(s.replace(',', '.'));
  if (!isNaN(n) && n >= 0) {
    return n > 0 && n < 1
      ? Math.round(n * 86400)   // fraction de jour Excel
      : Math.round(n * 3600);   // nombre d'heures
  }

  return 0;
}

/**
 * Formate des secondes totales en HH:MM:SS continu.
 * Les heures dépassant 24h ne sont PAS converties en jours.
 * Ex : 90061 s → 25:01:01
 */
function fmtSeconds(totalSec) {
  const h = Math.floor(totalSec / 3600);
  const m = Math.floor((totalSec % 3600) / 60);
  const sc = totalSec % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sc).padStart(2, '0')}`;
}

/**
 * Formate une valeur brute Excel de durée en HH:MM:SS continu (sans conversion en jours).
 * - Nombre décimal Excel (fraction de jour) : 0.1528 → 03:40:10 / 2.0 → 48:00:00
 * - Chaîne hh:mm ou hh:mm:ss                : conservée telle quelle
 */
function formatRawDuration(rawVal) {
  if (rawVal === null || rawVal === undefined || rawVal === '') return '—';

  // Valeur numérique Excel : fraction de jour (0 < n < 1 ou n >= 1 représente des jours)
  if (typeof rawVal === 'number') {
    const totalSec = Math.round(rawVal * 86400);
    const h = Math.floor(totalSec / 3600);
    const m = Math.floor((totalSec % 3600) / 60);
    const sc = totalSec % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sc).padStart(2, '0')}`;
  }

  // Valeur texte : conserver telle quelle
  return String(rawVal).trim() || '—';
}

/** Formate des MINUTES en HH:MM (pour le slider et les filtres texte) */
function fmtMinutes(m) {
  const h = Math.floor(m / 60);
  const mn = String(m % 60).padStart(2, '0');
  return `${h}:${mn}`;
}

/** Convertit une valeur hh:mm/hh:mm:ss en MINUTES (pour les filtres texte & slider) */
function toMinutes(val) {
  return Math.floor(toSeconds(val) / 60);
}

/** Cherche le nom de colonne dans les headers selon COL_MAP */
function findCol(headers, key, map = COL_MAP) {
  const candidates = map[key] || [];
  return headers.find(h => candidates.includes(h.toLowerCase().trim())) || null;
}

function formatExcelDate(rawVal) {
  if (rawVal === null || rawVal === undefined || rawVal === '') return '—';
  if (typeof rawVal === 'number') {
    const d = new Date(Math.round((rawVal - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      const day = String(d.getUTCDate()).padStart(2, '0');
      const month = String(d.getUTCMonth() + 1).padStart(2, '0');
      const year = d.getUTCFullYear();
      const hh = String(d.getUTCHours()).padStart(2, '0');
      const mm = String(d.getUTCMinutes()).padStart(2, '0');
      return `${day}/${month}/${year} ${hh}:${mm}`;
    }
  }
  return String(rawVal).trim() || '—';
}

/** Classe CSS selon la durée (en secondes) */
function durClass(sec) {
  if (sec < 3600) return 'dur-low';    // < 1h
  if (sec < 10800) return 'dur-medium'; // < 3h
  return 'dur-high';
}

// ── Parsing Excel / CSV ──────────────────────────────────────
function parseWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (raw.length < 2) return alert('Le fichier semble vide ou sans données.');

  const headers = raw[0].map(h => String(h));
  const cols = {
    location: findCol(headers, 'location'),
    technology: findCol(headers, 'technology'),
    site_id: findCol(headers, 'site_id'),
    site_name: findCol(headers, 'site_name'),
    last_occurred: findCol(headers, 'last_occurred'),
    duration: findCol(headers, 'duration'),
    power_type: findCol(headers, 'power_type'),
    mail_received: findCol(headers, 'mail_received'),
  };

  // Extraire le Mail Received Time (on prend la 1ère valeur trouvée dans les données)
  if (cols.mail_received && raw.length > 1) {
    reportTimestamp = formatExcelDate(raw[1][headers.indexOf(cols.mail_received)]);
  }

  // Rapport des colonnes non trouvées
  const missing = Object.entries(cols).filter(([, v]) => !v).map(([k]) => k);
  if (missing.length) {
    console.warn('Colonnes non trouvées :', missing);
  }

  allData = raw.slice(1).map(row => {
    const get = col => col ? String(row[headers.indexOf(col)] ?? '').trim() : '—';

    // Récupérer la valeur brute de la colonne « Last Occurred On »
    // et la formater en date lisible (JJ/MM/AAAA HH:MM si possible)
    const rawOccurred = cols.last_occurred
      ? row[headers.indexOf(cols.last_occurred)]
      : null;
    let dateCoupure = formatExcelDate(rawOccurred);

    // Récupérer la valeur brute de la colonne « Duration »
    const rawDuration = cols.duration ? row[headers.indexOf(cols.duration)] : null;
    // Calculer les secondes directement depuis la valeur brute Excel
    const durSec = (rawDuration !== null && rawDuration !== undefined && rawDuration !== '')
      ? (typeof rawDuration === 'number'
        ? Math.round(rawDuration * 86400)          // fraction de jour Excel
        : toSeconds(String(rawDuration)))           // chaîne texte hh:mm:ss
      : 0;

    return {
      location: get(cols.location),
      technology: get(cols.technology),
      site_id: get(cols.site_id),
      site_name: get(cols.site_name),
      date_coupure: dateCoupure,
      duration: formatRawDuration(rawDuration),      // affichage original fidèle
      power_type: get(cols.power_type),
      _durSec: durSec,
      _durMin: Math.floor(durSec / 60),              // pour slider
    };
  }).filter(r => {
    // Garder uniquement les lignes ayant au moins un identifiant réel (non vide)
    const hasId = r.site_id && r.site_id !== '—' && r.site_id.trim() !== '';
    const hasName = r.site_name && r.site_name !== '—' && r.site_name.trim() !== '';
    return hasId || hasName;
  });

  initApp();
}

// ── Initialisation post-import ───────────────────────────────
function initApp() {
  // Afficher les sections
  ['kpi-section', 'tech-kpi-section', 'filters-section', 'table-section'].forEach(id =>
    document.getElementById(id).classList.remove('hidden'));
  document.getElementById('import-section').classList.add('hidden');


  // Remplir le select Location
  const locations = [...new Set(allData.map(r => r.location).filter(v => v && v !== '—'))].sort();
  const sel = document.getElementById('filter-location');
  sel.innerHTML = '<option value="">Toutes les locations</option>';
  locations.forEach(l => {
    const o = document.createElement('option');
    o.value = o.textContent = l;
    sel.appendChild(o);
  });

  // Remplir le select Power Type
  const powerTypes = [...new Set(allData.map(r => r.power_type).filter(v => v && v !== '—'))].sort();
  const selPower = document.getElementById('filter-power-type');
  selPower.innerHTML = '<option value="">Tous les types d\'énergie</option>';
  powerTypes.forEach(p => {
    const o = document.createElement('option');
    o.value = o.textContent = p;
    selPower.appendChild(o);
  });

  // Configurer le slider (en minutes)
  maxSeconds = Math.max(...allData.map(r => r._durSec), 0);
  const maxMin = Math.ceil(maxSeconds / 60);
  const rMin = document.getElementById('range-min');
  const rMax = document.getElementById('range-max');
  rMin.max = rMax.max = maxMin;
  rMin.value = 0;
  rMax.value = maxMin;
  updateRangeFill();

  applyFilters();
}

// ── Filtres ───────────────────────────────────────────────────
function applyFilters() {
  const loc = document.getElementById('filter-location').value.toLowerCase();
  const sid = document.getElementById('filter-site-id').value.toLowerCase();
  const sname = document.getElementById('filter-site-name').value.toLowerCase();
  const ptype = document.getElementById('filter-power-type').value.toLowerCase();
  const durMinTxt = document.getElementById('filter-dur-min').value.trim();
  const durMaxTxt = document.getElementById('filter-dur-max').value.trim();
  const rMin = parseInt(document.getElementById('range-min').value, 10);
  const rMax = parseInt(document.getElementById('range-max').value, 10);

  // Durée depuis champs texte (priorité sur slider si remplis)
  const durMinM = durMinTxt ? toMinutes(durMinTxt) : rMin;
  const durMaxM = durMaxTxt ? toMinutes(durMaxTxt) : rMax;

  filtered = allData.filter(r => {
    if (loc && !r.location.toLowerCase().includes(loc)) return false;
    if (sid && !r.site_id.toLowerCase().includes(sid)) return false;
    if (sname && !r.site_name.toLowerCase().includes(sname)) return false;
    if (ptype && !r.power_type.toLowerCase().includes(ptype)) return false;
    if (r._durMin < durMinM || r._durMin > durMaxM) return false;
    return true;
  });

  if (sortCol) sortData();
  renderTable();
  renderKPIs();
}

// ── Tri ───────────────────────────────────────────────────────
function sortData() {
  filtered.sort((a, b) => {
    let va = a[sortCol] ?? '';
    let vb = b[sortCol] ?? '';
    if (sortCol === 'duration') { va = a._durMin; vb = b._durMin; }
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    if (va < vb) return sortDir === 'asc' ? -1 : 1;
    if (va > vb) return sortDir === 'asc' ? 1 : -1;
    return 0;
  });
}

// ── Rendu tableau ─────────────────────────────────────────────
function renderTable() {
  const tbody = document.getElementById('table-body');
  const noRes = document.getElementById('no-results');
  const count = document.getElementById('results-count');

  count.textContent = `${filtered.length} résultat(s)`;

  if (!filtered.length) {
    tbody.innerHTML = '';
    noRes.classList.remove('hidden');
    return;
  }
  noRes.classList.add('hidden');

  tbody.innerHTML = filtered.map(r => `
    <tr>
      <td>${r.location}</td>
      <td>${r.technology}</td>
      <td>${r.site_id}</td>
      <td>${r.site_name}</td>
      <td>${r.power_type}</td>
      <td class="col-date-coupure">${r.date_coupure}</td>
      <td class="${durClass(r._durSec)}">${r.duration}</td>
    </tr>
  `).join('');
}

// ── Rendu KPI ─────────────────────────────────────────────────
function renderKPIs() {
  document.getElementById('kpi-now-val').textContent = reportTimestamp;

  document.getElementById('kpi-total-val').textContent = allData.length;
  document.getElementById('kpi-filtered-val').textContent = filtered.length;



  const locs = new Set(filtered.map(r => r.location).filter(v => v && v !== '—'));
  document.getElementById('kpi-locations-val').textContent = locs.size;

  renderTechKPIs();
}

// ── Durée par technologie ─────────────────────────────────────
function renderTechKPIs() {
  const grid = document.getElementById('tech-kpi-grid');
  if (!grid) return;

  // Agréger les secondes par technologie (en lisant _durSec issu de la colonne Duration)
  const techMap = {};
  filtered.forEach(r => {
    const tech = r.technology && r.technology !== '—' ? r.technology : 'N/A';
    if (!techMap[tech]) techMap[tech] = { seconds: 0, count: 0 };
    techMap[tech].seconds += r._durSec;
    techMap[tech].count += 1;
  });

  // Trier par ordre personnalisé : 2G, 3G, 4G
  const techOrder = ['2G', '3G', '4G'];
  const entries = Object.entries(techMap).sort((a, b) => {
    const indexA = techOrder.indexOf(a[0]);
    const indexB = techOrder.indexOf(b[0]);
    // Si aucun n'est dans l'ordre défini
    if (indexA === -1 && indexB === -1) return 0;
    // Si a n'est pas dans l'ordre, le mettre après
    if (indexA === -1) return 1;
    // Si b n'est pas dans l'ordre, le mettre après
    if (indexB === -1) return -1;
    // Sinon, utiliser l'ordre défini
    return indexA - indexB;
  });

  if (!entries.length) {
    grid.innerHTML = '<span style="font-size:12px;color:var(--text-muted)">Aucune donnée</span>';
    return;
  }

  grid.innerHTML = entries.map(([tech, { seconds, count }]) => `
    <div class="tech-badge">
      <span class="tech-name">${tech}</span>
      <span class="tech-duration">${fmtSeconds(seconds)}</span>
      <span class="tech-count">${count} node${count > 1 ? 's' : ''}</span>
    </div>
  `).join('');
}

// ── Range slider ─────────────────────────────────────────────
function updateRangeFill() {
  const rMin = document.getElementById('range-min');
  const rMax = document.getElementById('range-max');
  const fill = document.getElementById('range-fill');
  const label = document.getElementById('range-label');
  const max = parseInt(rMax.max, 10) || 1;
  const lo = parseInt(rMin.value, 10);
  const hi = parseInt(rMax.value, 10);

  fill.style.left = (lo / max * 100) + '%';
  fill.style.right = ((max - hi) / max * 100) + '%';
  label.textContent = fmtMinutes(lo) + ' — ' + fmtMinutes(hi);
}

// ── Export XLSX ────────────────────────────────────────────────
function exportXLSX() {
  const headers = ['Location', 'Technologie', 'Site ID', 'Site Name', 'Power Type', 'Date Coupure', 'Duration'];
  
  // Préparer les données pour XLSX (tableau de tableaux)
  const data = [headers];
  filtered.forEach(r => {
    data.push([r.location, r.technology, r.site_id, r.site_name, r.power_type, r.date_coupure, r.duration]);
  });

  // Créer la feuille et le classeur
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Nodes_Down");
  
  // Télécharger le fichier
  XLSX.writeFile(wb, 'site_down_export.xlsx');
}

// ── Événements ────────────────────────────────────────────────
function handleFile(file) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    parseWorkbook(wb);
  };
  reader.readAsArrayBuffer(file);
}

// Sélection fichier
document.getElementById('file-input').addEventListener('change', e => handleFile(e.target.files[0]));

// Drag & Drop
const dz = document.getElementById('drop-zone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.classList.remove('drag-over');
  handleFile(e.dataTransfer.files[0]);
});

// Filtres texte
['filter-location', 'filter-site-id', 'filter-site-name', 'filter-power-type', 'filter-dur-min', 'filter-dur-max']
  .forEach(id => document.getElementById(id).addEventListener('input', applyFilters));

// Sliders durée
document.getElementById('range-min').addEventListener('input', function () {
  if (parseInt(this.value) > parseInt(document.getElementById('range-max').value))
    this.value = document.getElementById('range-max').value;
  updateRangeFill(); applyFilters();
});
document.getElementById('range-max').addEventListener('input', function () {
  if (parseInt(this.value) < parseInt(document.getElementById('range-min').value))
    this.value = document.getElementById('range-min').value;
  updateRangeFill(); applyFilters();
});

// Sync champs texte → slider
['filter-dur-min', 'filter-dur-max'].forEach(id => {
  document.getElementById(id).addEventListener('input', () => {
    const minM = toMinutes(document.getElementById('filter-dur-min').value);
    const maxM = toMinutes(document.getElementById('filter-dur-max').value) || maxMinutes;
    const rMin = document.getElementById('range-min');
    const rMax = document.getElementById('range-max');
    if (!isNaN(minM)) rMin.value = Math.min(minM, maxMinutes);
    if (!isNaN(maxM)) rMax.value = Math.min(maxM, maxMinutes);
    updateRangeFill();
  });
});

// Reset
document.getElementById('btn-reset').addEventListener('click', () => {
  document.getElementById('filter-location').value = '';
  document.getElementById('filter-site-id').value = '';
  document.getElementById('filter-site-name').value = '';
  document.getElementById('filter-power-type').value = '';
  document.getElementById('filter-dur-min').value = '';
  document.getElementById('filter-dur-max').value = '';
  document.getElementById('range-min').value = 0;
  document.getElementById('range-max').value = Math.ceil(maxSeconds / 60);
  updateRangeFill(); applyFilters();
});

// Export
document.getElementById('btn-export').addEventListener('click', exportXLSX);

// Tri colonnes
document.querySelectorAll('#data-table th').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.col;
    if (sortCol === col) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
    else { sortCol = col; sortDir = 'asc'; }
    document.querySelectorAll('#data-table th').forEach(t => t.classList.remove('sort-asc', 'sort-desc'));
    th.classList.add(sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
    sortData(); renderTable();
  });
});

// ============================================================
//   LOGIQUE VSWR
// ============================================================

function parseVSWRWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (raw.length < 2) return;

  const headers = raw[0].map(h => String(h));
  const cols = {
    site_id: findCol(headers, 'site_id', VSWR_COL_MAP),
    site_name: findCol(headers, 'site_name', VSWR_COL_MAP),
    antenna: findCol(headers, 'antenna', VSWR_COL_MAP),
    vswr_value: findCol(headers, 'vswr_value', VSWR_COL_MAP),
    check_date: findCol(headers, 'check_date', VSWR_COL_MAP),
    mail_received: findCol(headers, 'mail_received', VSWR_COL_MAP),
  };

  // Mail Received Time
  if (cols.mail_received && raw.length > 1) {
    vswrReportTimestamp = formatExcelDate(raw[1][headers.indexOf(cols.mail_received)]);
  }

  vswrAllData = raw.slice(1).map(row => {
    const get = (colName) => {
      const idx = headers.indexOf(cols[colName]);
      return idx !== -1 ? row[idx] : '—';
    };

    const val = parseFloat(get('vswr_value')) || 0;
    return {
      site_id: String(get('site_id')),
      site_name: String(get('site_name')),
      antenna: String(get('antenna')),
      vswr: val,
      check_date: formatExcelDate(get('check_date')),
    };
  });

  vswrFiltered = [...vswrAllData];
  
  // Afficher les sections
  document.getElementById('kpi-section-vswr').classList.remove('hidden');
  document.getElementById('filters-section-vswr').classList.remove('hidden');
  document.getElementById('table-section-vswr').classList.remove('hidden');

  renderVSWRKPIs();
  renderVSWRTable();
}

function renderVSWRKPIs() {
  document.getElementById('kpi-vswr-now-val').textContent = vswrReportTimestamp;
  document.getElementById('kpi-vswr-total-val').textContent = new Set(vswrAllData.map(d => d.site_id)).size;
  
  const highCount = vswrFiltered.filter(d => d.vswr > 1.5).length;
  document.getElementById('kpi-vswr-high-val').textContent = highCount;
  
  const avg = vswrFiltered.length > 0 
    ? (vswrFiltered.reduce((s, r) => s + r.vswr, 0) / vswrFiltered.length).toFixed(2)
    : '—';
  document.getElementById('kpi-vswr-avg-val').textContent = avg;
}

function renderVSWRTable() {
  const tbody = document.getElementById('vswr-table-body');
  document.getElementById('results-count-vswr').textContent = `${vswrFiltered.length} résultat(s)`;
  
  tbody.innerHTML = vswrFiltered.map(r => `
    <tr>
      <td>${r.site_id}</td>
      <td>${r.site_name}</td>
      <td>${r.antenna}</td>
      <td style="font-weight:600; color:${r.vswr > 1.5 ? 'var(--error)' : 'var(--success)'}">${r.vswr.toFixed(2)}</td>
      <td>
        <span class="status-badge ${r.vswr > 1.5 ? 'status-down' : 'status-up'}">
          ${r.vswr > 1.5 ? 'CRITIQUE' : 'NORMAL'}
        </span>
      </td>
      <td>${r.check_date}</td>
    </tr>
  `).join('');
}

function applyVSWRFilters() {
  const sid = document.getElementById('filter-vswr-site-id').value.toLowerCase();
  const status = document.getElementById('filter-vswr-status').value;

  vswrFiltered = vswrAllData.filter(r => {
    if (sid && !r.site_id.toLowerCase().includes(sid)) return false;
    if (status === 'ok' && r.vswr > 1.5) return false;
    if (status === 'critical' && r.vswr <= 1.5) return false;
    return true;
  });

  renderVSWRKPIs();
  renderVSWRTable();
}

// Events VSWR
const vswrInput = document.getElementById('file-input-vswr');
const vswrDz = document.getElementById('drop-zone-vswr');

function handleVSWRFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
    parseVSWRWorkbook(wb);
  };
  reader.readAsArrayBuffer(file);
}

vswrInput.addEventListener('change', (e) => handleVSWRFile(e.target.files[0]));

vswrDz.addEventListener('dragover', e => { e.preventDefault(); vswrDz.classList.add('drag-over'); });
vswrDz.addEventListener('dragleave', () => vswrDz.classList.remove('drag-over'));
vswrDz.addEventListener('drop', e => {
  e.preventDefault();
  vswrDz.classList.remove('drag-over');
  handleVSWRFile(e.dataTransfer.files[0]);
});

document.getElementById('filter-vswr-site-id').addEventListener('input', applyVSWRFilters);
document.getElementById('filter-vswr-status').addEventListener('change', applyVSWRFilters);
document.getElementById('btn-reset-vswr').addEventListener('click', () => {
  document.getElementById('filter-vswr-site-id').value = '';
  document.getElementById('filter-vswr-status').value = '';
  applyVSWRFilters();
});

document.getElementById('btn-export-vswr').addEventListener('click', () => {
  const headers = ['Site ID', 'Site Name', 'Antenna', 'VSWR', 'Status', 'Date Check'];
  const data = [headers];
  vswrFiltered.forEach(r => {
    data.push([r.site_id, r.site_name, r.antenna, r.vswr, r.vswr > 1.5 ? 'CRITIQUE' : 'NORMAL', r.check_date]);
  });
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "VSWR_Report");
  XLSX.writeFile(wb, 'vswr_export.xlsx');
});

// ── Auto-chargement : fichiers locaux ──────────────────────────
async function autoLoad() {
  // 1. Charger suivi_site.xlsx
  const dropTitle = document.querySelector('.drop-title');
  if (dropTitle) dropTitle.textContent = 'Chargement de suivi_site.xlsx…';

  try {
    const res = await fetch('./suivi_site.xlsx');
    if (res.ok) {
      const buffer = await res.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
      parseWorkbook(wb);
    }
  } catch (err) { console.info('Auto-load Node Down non dispo'); }

  // 2. Charger suivi_vswr.xlsx
  const dropTitleVswr = document.querySelector('#drop-zone-vswr .drop-title');
  if (dropTitleVswr) dropTitleVswr.textContent = 'Chargement de suivi_vswr.xlsx…';

  try {
    const resV = await fetch('./suivi_vswr.xlsx');
    if (resV.ok) {
      const bufferV = await resV.arrayBuffer();
      const wbV = XLSX.read(new Uint8Array(bufferV), { type: 'array' });
      parseVSWRWorkbook(wbV);
    } else {
      if (dropTitleVswr) dropTitleVswr.textContent = 'Glissez votre fichier VSWR ici';
    }
  } catch (err) { 
    console.info('Auto-load VSWR non dispo');
    if (dropTitleVswr) dropTitleVswr.textContent = 'Glissez votre fichier VSWR ici';
  }
}

autoLoad();
