/* ============================================================
   SITE DOWN DASHBOARD — app.js
   Logique : import Excel, filtres, tri, KPI, export CSV
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
 * Formate des secondes totales :
 *   < 24h  →  HH:MM:SS          (ex: 03:40:10)
 *   ≥ 24h  →  AJ HH:MM:SS      (ex: 2J 03:40:10)
 */
function fmtSeconds(totalSec) {
  const days = Math.floor(totalSec / 86400);
  const rem = totalSec % 86400;
  const h = Math.floor(rem / 3600);
  const m = Math.floor((rem % 3600) / 60);
  const sc = rem % 60;
  const time = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(sc).padStart(2, '0')}`;
  return days > 0 ? `${days}J ${time}` : time;
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
function findCol(headers, key) {
  const candidates = COL_MAP[key];
  return headers.find(h => candidates.includes(h.toLowerCase().trim())) || null;
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
  };

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
    let dateCoupure = '—';
    if (rawOccurred !== null && rawOccurred !== undefined && rawOccurred !== '') {
      // Valeur numérique Excel (numéro de série de date)
      if (typeof rawOccurred === 'number') {
        const d = new Date(Math.round((rawOccurred - 25569) * 86400 * 1000));
        if (!isNaN(d.getTime())) {
          const day = String(d.getUTCDate()).padStart(2, '0');
          const month = String(d.getUTCMonth() + 1).padStart(2, '0');
          const year = d.getUTCFullYear();
          const hh = String(d.getUTCHours()).padStart(2, '0');
          const mm = String(d.getUTCMinutes()).padStart(2, '0');
          dateCoupure = `${day}/${month}/${year} ${hh}:${mm}`;
        }
      } else {
        // Valeur texte : conserver telle quelle
        dateCoupure = String(rawOccurred).trim() || '—';
      }
    }

    return {
      location: get(cols.location),
      technology: get(cols.technology),
      site_id: get(cols.site_id),
      site_name: get(cols.site_name),
      date_coupure: dateCoupure,
      duration: get(cols.duration),
      _durSec: toSeconds(get(cols.duration)),
      _durMin: Math.floor(toSeconds(get(cols.duration)) / 60), // pour slider
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
      <td class="col-date-coupure">${r.date_coupure}</td>
      <td class="${durClass(r._durSec)}">${r.duration}</td>
    </tr>
  `).join('');
}

// ── Rendu KPI ─────────────────────────────────────────────────
function renderKPIs() {
  document.getElementById('kpi-total-val').textContent = allData.length;
  document.getElementById('kpi-filtered-val').textContent = filtered.length;

  const totalSec = filtered.reduce((s, r) => s + r._durSec, 0);
  document.getElementById('kpi-duration-val').textContent = fmtSeconds(totalSec);

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

  const entries = Object.entries(techMap).sort((a, b) => b[1].seconds - a[1].seconds);

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

// ── Export CSV ────────────────────────────────────────────────
function exportCSV() {
  const headers = ['Location', 'Technologie', 'Site ID', 'Site Name', 'Date Coupure', 'Duration'];
  const rows = filtered.map(r =>
    [r.location, r.technology, r.site_id, r.site_name, r.date_coupure, r.duration]
      .map(v => `"${String(v).replace(/"/g, '""')}"`)
      .join(',')
  );
  const csv = [headers.join(','), ...rows].join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'site_down_export.csv';
  a.click(); URL.revokeObjectURL(url);
}

// ── Événements ────────────────────────────────────────────────
function handleFile(file) {
  if (!file) return;
  document.getElementById('file-badge').textContent = file.name;
  document.getElementById('file-badge').classList.remove('hidden');

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
['filter-location', 'filter-site-id', 'filter-site-name', 'filter-dur-min', 'filter-dur-max']
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
  document.getElementById('filter-dur-min').value = '';
  document.getElementById('filter-dur-max').value = '';
  document.getElementById('range-min').value = 0;
  document.getElementById('range-max').value = Math.ceil(maxSeconds / 60);
  updateRangeFill(); applyFilters();
});

// Export
document.getElementById('btn-export').addEventListener('click', exportCSV);

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

// ── Auto-chargement : suivi_site.xlsx (même dossier) ───────────
async function autoLoad() {
  // Afficher un indicateur de chargement dans la drop-zone
  const dropTitle = document.querySelector('.drop-title');
  const origTitle = dropTitle ? dropTitle.textContent : '';
  if (dropTitle) dropTitle.textContent = 'Chargement de suivi_site.xlsx…';

  try {
    const res = await fetch('./suivi_site.xlsx');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });

    // Mettre à jour le badge fichier dans le header
    const badge = document.getElementById('file-badge');
    badge.textContent = 'suivi_nodes_down';
    badge.classList.remove('hidden');

    parseWorkbook(wb);

  } catch (err) {
    // Fallback : restaurer la zone d'import manuelle
    console.info(
      'Chargement auto non disponible.\n' +
      'Ouvrez la page via un serveur local (ex: VS Code Live Server).\n' +
      'Erreur :', err.message
    );
    if (dropTitle) dropTitle.textContent = origTitle;
  }
}

autoLoad();
