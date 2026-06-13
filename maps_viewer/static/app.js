'use strict';

let allVisits = [];
let filteredVisits = [];
let currentPage = 1;
const PAGE_SIZE = 50;
let map = null;
let markers = [];

// ── DOM refs ──────────────────────────────────────────────────────────────────
const dropZone    = document.getElementById('dropZone');
const fileInput   = document.getElementById('fileInput');
const btnUpload   = document.getElementById('btnUpload');
const spinner     = document.getElementById('spinner');
const uploadError = document.getElementById('uploadError');
const fileList    = document.getElementById('fileList');
const statsBar    = document.getElementById('statsBar');
const tableSection = document.getElementById('tableSection');
const mapSection  = document.getElementById('mapSection');
const tableBody   = document.getElementById('tableBody');
const tableInfo   = document.getElementById('tableInfo');
const searchInput = document.getElementById('searchInput');
const dateFrom    = document.getElementById('dateFrom');
const dateTo      = document.getElementById('dateTo');
const btnTable    = document.getElementById('btnTable');
const btnMap      = document.getElementById('btnMap');
const btnReset    = document.getElementById('btnReset');
const modal       = new bootstrap.Modal(document.getElementById('detailModal'));

// ── Drag-and-drop ─────────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  fileInput.files = e.dataTransfer.files;
  onFilesSelected();
});
fileInput.addEventListener('change', onFilesSelected);

function onFilesSelected() {
  const files = fileInput.files;
  if (!files.length) return;
  const names = Array.from(files).map(f => `<span class="badge bg-secondary me-1">${f.name}</span>`).join('');
  fileList.innerHTML = `File selezionati: ${names}`;
  btnUpload.classList.remove('d-none');
  uploadError.classList.add('d-none');
}

// ── Upload & parse ────────────────────────────────────────────────────────────
btnUpload.addEventListener('click', async () => {
  const files = fileInput.files;
  if (!files.length) return;

  btnUpload.disabled = true;
  spinner.classList.remove('d-none');
  uploadError.classList.add('d-none');

  const formData = new FormData();
  for (const f of files) formData.append('files', f);

  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();

    if (data.error) throw new Error(data.error);

    allVisits = data.visits;
    filteredVisits = [...allVisits];
    currentPage = 1;

    if (data.errors && data.errors.length) {
      uploadError.innerHTML = '<strong>Avvisi:</strong><br>' + data.errors.join('<br>');
      uploadError.classList.remove('d-none');
    }

    renderStats();
    renderTable();
    statsBar.classList.remove('d-none');
    tableSection.classList.remove('d-none');
    document.getElementById('uploadSection').classList.add('d-none');

    if (map) { clearMapMarkers(); }

  } catch (err) {
    uploadError.textContent = 'Errore: ' + err.message;
    uploadError.classList.remove('d-none');
  } finally {
    btnUpload.disabled = false;
    spinner.classList.add('d-none');
  }
});

// ── Stats ─────────────────────────────────────────────────────────────────────
function renderStats() {
  document.getElementById('statTotal').textContent = allVisits.length.toLocaleString('it');

  const days = new Set(allVisits.map(v => v.date)).size;
  document.getElementById('statDays').textContent = days.toLocaleString('it');

  const withDur = allVisits.filter(v => v.duration_minutes > 0);
  if (withDur.length) {
    const longest = withDur.reduce((a, b) => b.duration_minutes > a.duration_minutes ? b : a);
    document.getElementById('statLongest').textContent = longest.duration;
    const avg = Math.round(withDur.reduce((s, v) => s + v.duration_minutes, 0) / withDur.length);
    document.getElementById('statAvg').textContent = fmtMinutes(avg);
  }
}

function fmtMinutes(m) {
  const h = Math.floor(m / 60), min = m % 60;
  return h > 0 ? `${h}h ${min}m` : `${min}m`;
}

// ── Table ─────────────────────────────────────────────────────────────────────
function renderTable() {
  const start = (currentPage - 1) * PAGE_SIZE;
  const page  = filteredVisits.slice(start, start + PAGE_SIZE);

  tableBody.innerHTML = page.map((v, i) => `
    <tr onclick="showDetail(${allVisits.indexOf(v)})">
      <td class="text-muted">${start + i + 1}</td>
      <td title="${escHtml(v.name)}"><strong>${escHtml(v.name)}</strong></td>
      <td class="address-cell text-muted" title="${escHtml(v.address)}">${escHtml(v.address)}</td>
      <td>${v.date}</td>
      <td>${v.time}</td>
      <td>${v.duration !== '—' ? `<span class="dur-badge">${v.duration}</span>` : '—'}</td>
      <td>
        <a href="https://maps.google.com/?q=${v.lat},${v.lng}" target="_blank"
           class="btn btn-outline-primary btn-sm map-btn" onclick="event.stopPropagation()">
          <i class="bi bi-pin-map"></i>
        </a>
      </td>
    </tr>`).join('');

  renderPagination();
  const total = filteredVisits.length;
  tableInfo.textContent = `Mostro ${Math.min(start + PAGE_SIZE, total)} di ${total.toLocaleString('it')} record`;
}

function renderPagination() {
  const totalPages = Math.ceil(filteredVisits.length / PAGE_SIZE);
  let existing = document.getElementById('pagination');
  if (existing) existing.remove();

  if (totalPages <= 1) return;

  const nav = document.createElement('div');
  nav.id = 'pagination';
  nav.className = 'pagination-bar my-3';

  const makeBtn = (label, page, disabled, active) => {
    const btn = document.createElement('button');
    btn.className = `btn btn-sm ${active ? 'btn-primary' : 'btn-outline-secondary'}`;
    btn.innerHTML = label;
    btn.disabled = disabled;
    btn.onclick = () => { currentPage = page; renderTable(); window.scrollTo(0, 0); };
    return btn;
  };

  nav.appendChild(makeBtn('<i class="bi bi-chevron-double-left"></i>', 1, currentPage === 1, false));
  nav.appendChild(makeBtn('<i class="bi bi-chevron-left"></i>', currentPage - 1, currentPage === 1, false));

  let startP = Math.max(1, currentPage - 2);
  let endP   = Math.min(totalPages, currentPage + 2);
  for (let p = startP; p <= endP; p++) {
    nav.appendChild(makeBtn(p, p, false, p === currentPage));
  }

  nav.appendChild(makeBtn('<i class="bi bi-chevron-right"></i>', currentPage + 1, currentPage === totalPages, false));
  nav.appendChild(makeBtn('<i class="bi bi-chevron-double-right"></i>', totalPages, currentPage === totalPages, false));

  tableSection.appendChild(nav);
}

// ── Filters ───────────────────────────────────────────────────────────────────
function applyFilters() {
  const q = searchInput.value.trim().toLowerCase();
  const from = dateFrom.value;   // YYYY-MM-DD
  const to   = dateTo.value;

  filteredVisits = allVisits.filter(v => {
    if (q && !v.name.toLowerCase().includes(q) && !v.address.toLowerCase().includes(q)) return false;
    if (from || to) {
      // v.date is DD/MM/YYYY → convert to YYYY-MM-DD for comparison
      const [d, m, y] = v.date.split('/');
      const isoDate = `${y}-${m}-${d}`;
      if (from && isoDate < from) return false;
      if (to   && isoDate > to)   return false;
    }
    return true;
  });

  currentPage = 1;
  renderTable();
}

searchInput.addEventListener('input', applyFilters);
dateFrom.addEventListener('change', applyFilters);
dateTo.addEventListener('change', applyFilters);
btnReset.addEventListener('click', () => {
  searchInput.value = '';
  dateFrom.value = '';
  dateTo.value = '';
  filteredVisits = [...allVisits];
  currentPage = 1;
  renderTable();
});

// ── Detail modal ──────────────────────────────────────────────────────────────
window.showDetail = function(idx) {
  const v = allVisits[idx];
  document.getElementById('modalTitle').textContent = v.name;
  document.getElementById('modalBody').innerHTML = `
    <div class="detail-row"><span class="label"><i class="bi bi-building me-1"></i>Luogo</span><span class="value">${escHtml(v.name)}</span></div>
    <div class="detail-row"><span class="label"><i class="bi bi-geo me-1"></i>Indirizzo</span><span class="value">${escHtml(v.address)}</span></div>
    <div class="detail-row"><span class="label"><i class="bi bi-calendar3 me-1"></i>Data</span><span class="value">${v.date}</span></div>
    <div class="detail-row"><span class="label"><i class="bi bi-clock me-1"></i>Orario</span><span class="value">${v.time}</span></div>
    <div class="detail-row"><span class="label"><i class="bi bi-hourglass me-1"></i>Permanenza</span><span class="value">${v.duration}</span></div>
    <div class="detail-row"><span class="label"><i class="bi bi-crosshair me-1"></i>Coordinate</span><span class="value">${v.lat}, ${v.lng}</span></div>
    ${v.confidence ? `<div class="detail-row"><span class="label"><i class="bi bi-patch-check me-1"></i>Confidenza</span><span class="value">${Math.round(v.confidence)}%</span></div>` : ''}
  `;
  document.getElementById('modalMapLink').href = `https://maps.google.com/?q=${v.lat},${v.lng}`;
  modal.show();
};

// ── Map view ──────────────────────────────────────────────────────────────────
btnTable.addEventListener('click', () => {
  btnTable.classList.add('active');
  btnMap.classList.remove('active');
  tableSection.classList.remove('d-none');
  mapSection.classList.add('d-none');
  btnTable.classList.replace('btn-outline-light', 'btn-light');
  btnMap.classList.replace('btn-light', 'btn-outline-light');
});

btnMap.addEventListener('click', () => {
  btnMap.classList.add('active');
  btnTable.classList.remove('active');
  mapSection.classList.remove('d-none');
  tableSection.classList.add('d-none');
  btnMap.classList.replace('btn-outline-light', 'btn-light');
  btnTable.classList.replace('btn-light', 'btn-outline-light');
  initMap();
});

function initMap() {
  if (!map) {
    map = L.map('map').setView([41.9, 12.5], 6);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap contributors'
    }).addTo(map);
  }
  clearMapMarkers();
  loadMapMarkers();
}

function clearMapMarkers() {
  markers.forEach(m => m.remove());
  markers = [];
}

function loadMapMarkers() {
  // Use displayed (filtered) visits
  const visits = filteredVisits.filter(v => v.lat !== 0 || v.lng !== 0);
  if (!visits.length) return;

  const bounds = [];
  visits.forEach(v => {
    const m = L.circleMarker([v.lat, v.lng], {
      radius: 7,
      fillColor: '#1a73e8',
      color: '#fff',
      weight: 2,
      fillOpacity: 0.85
    }).addTo(map);

    m.bindPopup(`
      <strong>${escHtml(v.name)}</strong><br/>
      <span class="text-muted">${escHtml(v.address)}</span><br/>
      <small>${v.date} ${v.time} — ${v.duration}</small>
    `);
    markers.push(m);
    bounds.push([v.lat, v.lng]);
  });

  if (bounds.length) map.fitBounds(bounds, { padding: [40, 40] });
}

// ── Util ──────────────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
