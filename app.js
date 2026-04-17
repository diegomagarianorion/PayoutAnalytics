// ─── STATE ───────────────────────────────────────────────────────────────────
let ALL_RECORDS = [];
let FILTERED = [];
let CHARTS = {};
let sortCol = 'date', sortDir = -1;
let currentPage = 1;
const PER_PAGE = 50;

// ─── UTILS ───────────────────────────────────────────────────────────────────
const fmtUSD = v => {
    if (v >= 1e6) return `$${(v / 1e6).toFixed(2)}M`;
    if (v >= 1e3) return `$${(v / 1e3).toFixed(1)}K`;
    return `$${(+v).toFixed(2)}`;
};
const fmtN = v => Number(v).toLocaleString('en-US');
const fmtP = v => `${v.toFixed(1)}%`;

const STATUS_MAP = {
    paid: 'Paid', pagado: 'Paid',
    rejected: 'Rejected', rechazado: 'Rejected',
    pending: 'Pending', pendiente: 'Pending',
    approved: 'Approved', aprobado: 'Approved',
    submitted: 'Submitted', 'rise pending': 'Pending'
};
const METHOD_OK = ['USDT-TRC20', 'USDT-ERC20', 'RISE', 'PAYPAL'];
const SHEET_ORDER = ['APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY-2026', 'FEBRUARY-2026'];

function normStatus(v) {
    const k = String(v || '').trim().toLowerCase();
    return STATUS_MAP[k] || 'Pending';
}
function normMethod(v) {
    const s = String(v || '').trim();
    if (!s) return '';
    if (s.startsWith('0x') || (s.startsWith('T') && s.length > 20)) return '';
    if (/^[^@]+@[^@]+\.[a-z]{2,}$/i.test(s)) return '';
    if (s.startsWith('http')) return '';
    return s;
}
function parseDate(v) {
    if (!v) return '';
    if (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}/)) return v.slice(0, 10);
    if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
    if (typeof v === 'number') {
        const d = XLSX.SSF.parse_date_code(v);
        if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }
    const s = String(v);
    const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return '';
}
function pNum(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

// ─── PARSE XLSX ───────────────────────────────────────────────────────────────
function parseXlsx(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const records = [];
    for (const sn of wb.SheetNames) {
        const ws = wb.Sheets[sn];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
        if (!rows || rows.length < 2) continue;
        const hdr = (rows[0] || []).map(h => String(h || '').toLowerCase());
        const isNew = hdr.includes('account') || hdr.includes('date');
        const data = rows.slice(1);
        for (let i = 0; i < data.length; i++) {
            const r = data[i] || [];
            if (r.every(c => c === null || c === undefined || c === '')) continue;
            let rec;
            if (isNew) {
                rec = {
                    id: `${sn}-${i}`, date: parseDate(r[0]), account: r[1] ? String(r[1]).replace(/\.0$/, '') : '',
                    name: String(r[2] || '').trim(), email: String(r[3] || '').trim().toLowerCase(),
                    method: normMethod(r[4]), wallet: String(r[5] || '').trim(),
                    submitted: pNum(r[6]), deduction: pNum(r[7]),
                    total: r[8] ? pNum(r[8]) : null,
                    condition: normStatus(r[9]), comment: String(r[11] || '').trim(),
                    sheet: sn, month: sn.toUpperCase()
                };
            } else {
                rec = {
                    id: `${sn}-${i}`, date: parseDate(r[0]), account: '',
                    name: String(r[1] || '').trim(), email: String(r[2] || '').trim().toLowerCase(),
                    method: normMethod(r[3]), wallet: String(r[4] || '').trim(),
                    submitted: pNum(r[5]), deduction: pNum(r[6]),
                    total: r[7] ? pNum(r[7]) : null,
                    condition: normStatus(r[8]), comment: String(r[10] || '').trim(),
                    sheet: sn, month: sn.toUpperCase()
                };
            }
            if (!rec.name || rec.submitted === 0) continue;
            records.push(rec);
        }
    }
    return records;
}

// ─── STATS ───────────────────────────────────────────────────────────────────
function computeStats(recs) {
    const paid = recs.filter(r => r.condition === 'Paid');
    const rejected = recs.filter(r => r.condition === 'Rejected');
    const pending = recs.filter(r => !['Paid', 'Rejected'].includes(r.condition));
    const paidAmt = paid.reduce((s, r) => s + (r.total ?? r.submitted), 0);
    const submitAmt = recs.reduce((s, r) => s + r.submitted, 0);
    const dedAmt = recs.reduce((s, r) => s + r.deduction, 0);

    const methodMap = {};
    for (const r of recs) {
        const m = r.method || 'Unknown';
        methodMap[m] = (methodMap[m] || 0) + 1;
    }
    const reasonMap = {};
    for (const r of rejected) {
        if (!r.comment) continue;
        r.comment.split(/\//).map(p => p.trim()).filter(Boolean).forEach(p => {
            const k = p.replace(/\s+/g, ' ').trim();
            if (k) reasonMap[k] = (reasonMap[k] || 0) + 1;
        });
    }
    const monthMap = {};
    for (const r of recs) {
        if (!monthMap[r.sheet]) monthMap[r.sheet] = { month: r.sheet, req: 0, paid: 0, rej: 0, paidAmt: 0, subAmt: 0 };
        const m = monthMap[r.sheet];
        m.req++; m.subAmt += r.submitted;
        if (r.condition === 'Paid') { m.paid++; m.paidAmt += (r.total ?? r.submitted); }
        if (r.condition === 'Rejected') m.rej++;
    }
    const monthlyData = Object.values(monthMap).sort((a, b) => {
        const ai = SHEET_ORDER.findIndex(s => s === a.month.toUpperCase());
        const bi = SHEET_ORDER.findIndex(s => s === b.month.toUpperCase());
        return (ai < 0 ? 99 : ai) - (bi < 0 ? 99 : bi);
    });
    const uniqueEmails = new Set(recs.map(r => r.email).filter(Boolean));
    return {
        total: recs.length, paid: paid.length, rej: rejected.length, pend: pending.length,
        paidAmt, submitAmt, dedAmt,
        apprRate: recs.length ? (paid.length / recs.length) * 100 : 0,
        rejRate: recs.length ? (rejected.length / recs.length) * 100 : 0,
        avgPaid: paid.length ? paidAmt / paid.length : 0,
        uniqueTraders: uniqueEmails.size,
        methodMap, reasonMap, monthlyData
    };
}

// ─── SETUP FILTERS ────────────────────────────────────────────────────────────
function setupFilters(recs) {
    const methods = [...new Set(recs.map(r => r.method).filter(m => METHOD_OK.includes(m)))];
    const months = [...new Set(recs.map(r => r.sheet))];
    const reasons = [...new Set(recs.flatMap(r => r.comment ? r.comment.split(/\//).map(p => p.trim()).filter(Boolean) : []))].sort();

    const fm = document.getElementById('f-method');
    fm.innerHTML = '<option value="">Todos los Métodos</option>' + methods.map(m => `<option>${m}</option>`).join('');
    const fmo = document.getElementById('f-month');
    fmo.innerHTML = '<option value="">Todos los Meses</option>' + months.map(m => `<option>${m}</option>`).join('');
    const fr = document.getElementById('f-reason');
    fr.innerHTML = '<option value="">Todas las Razones</option>' + reasons.slice(0, 25).map(r => `<option>${r}</option>`).join('');
}

// ─── FILTERING ───────────────────────────────────────────────────────────────
function applyFilters() {
    const search = document.getElementById('f-search').value.toLowerCase();
    const status = document.getElementById('f-status').value;
    const method = document.getElementById('f-method').value;
    const month = document.getElementById('f-month').value;
    const reason = document.getElementById('f-reason').value.toLowerCase();
    const from = document.getElementById('f-from').value;
    const to = document.getElementById('f-to').value;
    const minA = parseFloat(document.getElementById('f-min').value) || 0;
    const maxA = parseFloat(document.getElementById('f-max').value) || Infinity;

    FILTERED = ALL_RECORDS.filter(r => {
        if (search && !r.name.toLowerCase().includes(search) && !r.email.includes(search) && !r.account.includes(search) && !r.comment.toLowerCase().includes(search)) return false;
        if (status && r.condition !== status) return false;
        if (method && r.method !== method) return false;
        if (month && r.sheet !== month) return false;
        if (reason && !r.comment.toLowerCase().includes(reason)) return false;
        if (from && r.date && r.date < from) return false;
        if (to && r.date && r.date > to) return false;
        if (r.submitted < minA || r.submitted > maxA) return false;
        return true;
    });

    currentPage = 1;
    renderTable();
    document.getElementById('btn-export').style.display = 'inline-flex';
    const badge = document.getElementById('rec-count-badge');
    if (badge) badge.textContent = fmtN(FILTERED.length);
}

function clearFilters() {
    ['f-search', 'f-from', 'f-to', 'f-min', 'f-max'].forEach(id => document.getElementById(id).value = '');
    ['f-status', 'f-method', 'f-month', 'f-reason'].forEach(id => document.getElementById(id).value = '');
    applyFilters();
}

// ─── SORT ─────────────────────────────────────────────────────────────────────
function sortBy(col) {
    if (sortCol === col) sortDir *= -1; else { sortCol = col; sortDir = -1; }
    FILTERED.sort((a, b) => {
        const av = a[col] ?? '', bv = b[col] ?? '';
        let c = typeof av === 'number' ? av - bv : String(av).localeCompare(String(bv));
        return c * sortDir;
    });
    currentPage = 1;
    renderTable();
    document.querySelectorAll('th span[id^="srt-"]').forEach(el => el.textContent = '');
    const el = document.getElementById(`srt-${col}`);
    if (el) el.textContent = sortDir > 0 ? ' ↑' : ' ↓';
}

// ─── STATUS BADGE ─────────────────────────────────────────────────────────────
function sbadge(s) {
    const map = { Paid: 'sp', Rejected: 'sr', Pending: 'se', Submitted: 'se', Approved: 'sa' };
    const ic = { Paid: '✓', Rejected: '✕', Pending: '◌', Submitted: '◌', Approved: '✓' };
    return `<span class="sb ${map[s] || 'se'}">${ic[s] || '?'} ${s}</span>`;
}
function mbadge(m) {
    const ic = { 'USDT-TRC20': '🟡', 'USDT-ERC20': '🔵', 'RISE': '⚡', 'PAYPAL': '🅿' };
    if (!m || !METHOD_OK.includes(m)) return '—';
    return `<span class="mb">${ic[m] || '💱'} ${m}</span>`;
}
function ctags(comment) {
    if (!comment) return '';
    return comment.split(/\//).map(p => p.trim()).filter(Boolean).map(t => `<span class="ctag">${t}</span>`).join('');
}

// ─── RENDER TABLE ─────────────────────────────────────────────────────────────
function renderTable() {
    const totalPages = Math.ceil(FILTERED.length / PER_PAGE);
    const page = FILTERED.slice((currentPage - 1) * PER_PAGE, currentPage * PER_PAGE);

    const tbody = document.getElementById('rec-tbody');
    if (!page.length) {
        tbody.innerHTML = `<tr><td colspan="10"><div class="empty"><div class="ei">🔍</div><div>Sin resultados</div></div></td></tr>`;
    } else {
        tbody.innerHTML = page.map(r => `
      <tr>
        <td style="color:var(--t1);font-weight:500">${r.date || '—'}</td>
        <td style="font-family:monospace;font-size:11px">${r.account || '—'}</td>
        <td style="color:var(--t1);max-width:170px">${r.name}</td>
        <td>${mbadge(r.method)}</td>
        <td style="color:var(--gold);font-weight:600">$${r.submitted.toFixed(2)}</td>
        <td style="color:${r.deduction > 0 ? 'var(--red)' : 'var(--t3)'}">${r.deduction > 0 ? `-$${r.deduction.toFixed(2)}` : '—'}</td>
        <td style="color:var(--green);font-weight:600">${r.total != null ? `$${r.total.toFixed(2)}` : '—'}</td>
        <td>${sbadge(r.condition)}</td>
        <td style="max-width:200px">${ctags(r.comment)}</td>
        <td style="color:var(--t3);font-size:11px">${r.sheet}</td>
      </tr>`).join('');
    }

    // Summary
    const paid = FILTERED.filter(r => r.condition === 'Paid');
    const rej = FILTERED.filter(r => r.condition === 'Rejected');
    const sumSub = FILTERED.reduce((s, r) => s + r.submitted, 0);
    const sumPaid = paid.reduce((s, r) => s + (r.total ?? r.submitted), 0);
    document.getElementById('sum-row').innerHTML = `
    <span>📊 Total: <strong>${fmtN(FILTERED.length)}</strong></span>
    <span class="sg">✅ Pagados: <strong>${fmtN(paid.length)}</strong></span>
    <span class="sr">❌ Rechazados: <strong>${fmtN(rej.length)}</strong></span>
    <span class="so">💰 Solicitado: <strong>${fmtUSD(sumSub)}</strong></span>
    <span class="sg">✓ Pagado: <strong>${fmtUSD(sumPaid)}</strong></span>`;

    // Pager
    const pager = document.getElementById('pager');
    if (totalPages <= 1) { pager.innerHTML = `<span>${fmtN(FILTERED.length)} registros</span>`; return; }
    let btns = `<button class="pb" onclick="goPage(1)" ${currentPage === 1 ? 'disabled' : ''}>«</button>`;
    btns += `<button class="pb" onclick="goPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>‹</button>`;
    const start = Math.max(1, Math.min(currentPage - 3, totalPages - 6));
    const end = Math.min(totalPages, start + 6);
    for (let p = start; p <= end; p++) btns += `<button class="pb ${p === currentPage ? 'on' : ''}" onclick="goPage(${p})">${p}</button>`;
    btns += `<button class="pb" onclick="goPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>›</button>`;
    btns += `<button class="pb" onclick="goPage(${totalPages})" ${currentPage === totalPages ? 'disabled' : ''}>»</button>`;
    pager.innerHTML = `<span>Mostrando ${(currentPage - 1) * PER_PAGE + 1}–${Math.min(currentPage * PER_PAGE, FILTERED.length)} de ${fmtN(FILTERED.length)}</span><div class="pager-btns">${btns}</div>`;
}
function goPage(p) { currentPage = p; renderTable(); }

// ─── STATS GRID ───────────────────────────────────────────────────────────────
function renderStats(s) {
    document.getElementById('stats-grid').innerHTML = [
        { cls: 'c-gold', ico: '💼', ic: 'ic-gold', lbl: 'Total Solicitudes', val: fmtN(s.total), sub: `${s.monthlyData.length} meses` },
        { cls: 'c-green', ico: '✅', ic: 'ic-green', lbl: 'Pagados', val: fmtN(s.paid), sub: `<span class="badge bg-up">↑ ${fmtP(s.apprRate)}</span>` },
        { cls: 'c-red', ico: '❌', ic: 'ic-red', lbl: 'Rechazados', val: fmtN(s.rej), sub: `<span class="badge bg-dn">↑ ${fmtP(s.rejRate)}</span>` },
        { cls: 'c-blue', ico: '⏳', ic: 'ic-blue', lbl: 'Pendientes', val: fmtN(s.pend), sub: 'En espera' },
        { cls: 'c-gold', ico: '💰', ic: 'ic-gold', lbl: 'Total Pagado', val: fmtUSD(s.paidAmt), sub: `Avg ${fmtUSD(s.avgPaid)}` },
        { cls: 'c-purple', ico: '📨', ic: 'ic-purple', lbl: 'Total Solicitado', val: fmtUSD(s.submitAmt), sub: `Descuentos: ${fmtUSD(s.dedAmt)}` },
        { cls: 'c-blue', ico: '👥', ic: 'ic-blue', lbl: 'Traders Únicos', val: fmtN(s.uniqueTraders), sub: 'por email' },
    ].map(c => `
    <div class="stat ${c.cls}">
      <div class="stat-ico ${c.ic}">${c.ico}</div>
      <div class="stat-lbl">${c.lbl}</div>
      <div class="stat-val">${c.val}</div>
      <div class="stat-sub">${c.sub}</div>
    </div>`).join('');
}

// ─── CHARTS ───────────────────────────────────────────────────────────────────
function destroyChart(id) {
    if (CHARTS[id]) { CHARTS[id].destroy(); delete CHARTS[id]; }
}

const CHART_DEFAULTS = {
    plugins: { legend: { labels: { color: '#8b9ab5', font: { size: 11, family: 'Inter' } } } },
    scales: {
        x: { ticks: { color: '#4a566a', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.04)' } },
        y: { ticks: { color: '#4a566a', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.04)' } }
    }
};

function mkChart(id, cfg) {
    destroyChart(id);
    CHARTS[id] = new Chart(document.getElementById(id), cfg);
}

function renderCharts(s) {
    const labels = s.monthlyData.map(m => m.month.slice(0, 3));

    // Row 1: Monthly area + pie + monthly bar
    document.getElementById('charts-row1').innerHTML = `
    <div class="chart-box" style="grid-column:span 2">
      <div class="chart-title">📈 Volumen Mensual</div>
      <div class="chart-wrap"><canvas id="ch-area" height="90"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">🍩 Distribución de Estado</div>
      <div class="chart-wrap"><canvas id="ch-pie" height="160"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">📊 Solicitudes Mensuales</div>
      <div class="chart-wrap"><canvas id="ch-bar" height="160"></canvas></div>
    </div>`;

    mkChart('ch-area', {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'Monto Pagado', data: s.monthlyData.map(m => m.paidAmt), borderColor: '#2dd99e', backgroundColor: 'rgba(45,217,158,.1)', fill: true, tension: .3, borderWidth: 2, pointRadius: 3 },
                { label: 'Solicitado', data: s.monthlyData.map(m => m.subAmt), borderColor: '#f5c842', backgroundColor: 'rgba(245,200,66,.06)', fill: true, tension: .3, borderWidth: 2, pointRadius: 3 },
            ]
        },
        options: { ...CHART_DEFAULTS, responsive: true, plugins: { ...CHART_DEFAULTS.plugins, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmtUSD(ctx.parsed.y)}` } } }, scales: { ...CHART_DEFAULTS.scales, y: { ...CHART_DEFAULTS.scales.y, ticks: { color: '#4a566a', font: { size: 10 }, callback: v => fmtUSD(v) } } } }
    });

    mkChart('ch-pie', {
        type: 'doughnut',
        data: { labels: ['Pagados', 'Rechazados', 'Pendientes'], datasets: [{ data: [s.paid, s.rej, s.pend], backgroundColor: ['#2dd99e', '#f05c6e', '#f5a623'], borderWidth: 0, hoverOffset: 4 }] },
        options: { responsive: true, cutout: '60%', plugins: { legend: { position: 'bottom', labels: { color: '#8b9ab5', font: { size: 11 }, padding: 12 } } } }
    });

    mkChart('ch-bar', {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'Pagados', data: s.monthlyData.map(m => m.paid), backgroundColor: '#2dd99e', borderRadius: 4 },
                { label: 'Rechazados', data: s.monthlyData.map(m => m.rej), backgroundColor: '#f05c6e', borderRadius: 4 },
            ]
        },
        options: { ...CHART_DEFAULTS, responsive: true }
    });

    // Row 2: Methods + rejection reasons + approval rate
    const topReasons = Object.entries(s.reasonMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const methodEntries = Object.entries(s.methodMap).filter(([m]) => METHOD_OK.includes(m)).sort((a, b) => b[1] - a[1]);

    document.getElementById('charts-row2').innerHTML = `
    <div class="chart-box">
      <div class="chart-title">💳 Métodos de Pago</div>
      <div id="methods-prog" style="padding:4px 0"></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">🚫 Top Razones de Rechazo</div>
      <div class="chart-wrap"><canvas id="ch-reasons" height="220"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">📅 Resumen Mensual</div>
      <div style="overflow-x:auto;max-height:240px;overflow-y:auto">
        <table>
          <thead><tr><th>Mes</th><th>Requests</th><th>Pagados</th><th>Rechazados</th><th>Pagado</th></tr></thead>
          <tbody>${s.monthlyData.map(m => `<tr>
            <td style="color:var(--t1);font-weight:600">${m.month}</td>
            <td>${fmtN(m.req)}</td>
            <td style="color:var(--green)">${fmtN(m.paid)}</td>
            <td style="color:var(--red)">${fmtN(m.rej)}</td>
            <td style="color:var(--gold);font-weight:600">${fmtUSD(m.paidAmt)}</td>
          </tr>`).join('')}</tbody>
        </table>
      </div>
    </div>`;

    // Methods progress bars
    const totalM = methodEntries.reduce((s, [, v]) => s + v, 0);
    const methodColors = { 'USDT-TRC20': '#2dd99e', 'USDT-ERC20': '#4f8ef7', 'RISE': '#f5c842', 'PAYPAL': '#a78bfa' };
    document.getElementById('methods-prog').innerHTML = methodEntries.map(([m, c]) => `
    <div class="prog-row">
      <div class="prog-label">
        <span style="color:var(--t2)">${m}</span>
        <span style="color:${methodColors[m] || '#8b9ab5'};font-weight:700">${fmtN(c)} (${((c / totalM) * 100).toFixed(1)}%)</span>
      </div>
      <div class="prog-track"><div class="prog-fill" style="width:${(c / totalM) * 100}%;background:${methodColors[m] || '#8b9ab5'}"></div></div>
    </div>`).join('');

    mkChart('ch-reasons', {
        type: 'bar',
        data: {
            labels: topReasons.map(([r]) => r.length > 28 ? r.slice(0, 28) + '…' : r),
            datasets: [{ label: 'Rechazos', data: topReasons.map(([, v]) => v), backgroundColor: '#f05c6e', borderRadius: 4 }]
        },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#4a566a', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.04)' } }, y: { ticks: { color: '#8b9ab5', font: { size: 10 } }, grid: { display: false } } } }
    });
}

// ─── ANALYSIS TAB ─────────────────────────────────────────────────────────────
function renderAnalysis(s) {
    const labels = s.monthlyData.map(m => m.month.slice(0, 3));
    document.getElementById('analysis-charts').innerHTML = `
    <div class="chart-box" style="grid-column:span 2">
      <div class="chart-title">📈 Tasa de Aprobación Mensual</div>
      <div class="chart-wrap"><canvas id="ch-appr" height="90"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">💳 Distribución de Métodos</div>
      <div class="chart-wrap"><canvas id="ch-mpie" height="180"></canvas></div>
    </div>
    <div class="chart-box">
      <div class="chart-title">🚫 Razones de Rechazo Completas</div>
      <div class="chart-wrap"><canvas id="ch-allreasons" height="180"></canvas></div>
    </div>`;

    mkChart('ch-appr', {
        type: 'bar',
        data: {
            labels,
            datasets: [{ label: 'Tasa Aprobación %', data: s.monthlyData.map(m => m.req > 0 ? Math.round((m.paid / m.req) * 100) : 0), backgroundColor: '#f5c842', borderRadius: 4 }]
        },
        options: { ...CHART_DEFAULTS, responsive: true, plugins: { legend: { display: false } }, scales: { ...CHART_DEFAULTS.scales, y: { ...CHART_DEFAULTS.scales.y, max: 100, ticks: { color: '#4a566a', font: { size: 10 }, callback: v => `${v}%` } } } }
    });

    const mEntries = Object.entries(s.methodMap).filter(([m]) => METHOD_OK.includes(m));
    mkChart('ch-mpie', {
        type: 'pie',
        data: { labels: mEntries.map(([m]) => m), datasets: [{ data: mEntries.map(([, v]) => v), backgroundColor: ['#2dd99e', '#4f8ef7', '#f5c842', '#a78bfa'], borderWidth: 0 }] },
        options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { color: '#8b9ab5', font: { size: 11 }, padding: 10 } } } }
    });

    const allR = Object.entries(s.reasonMap).sort((a, b) => b[1] - a[1]).slice(0, 16);
    mkChart('ch-allreasons', {
        type: 'bar',
        data: {
            labels: allR.map(([r]) => r.length > 24 ? r.slice(0, 24) + '…' : r),
            datasets: [{ label: 'Rechazos', data: allR.map(([, v]) => v), backgroundColor: 'rgba(240,92,110,.8)', borderRadius: 4 }]
        },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#4a566a', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.04)' } }, y: { ticks: { color: '#8b9ab5', font: { size: 10 } }, grid: { display: false } } } }
    });

    // Top traders
    const byEmail = {};
    for (const r of ALL_RECORDS) {
        if (!r.email) continue;
        if (!byEmail[r.email]) byEmail[r.email] = { name: r.name, email: r.email, req: 0, paid: 0, rej: 0, totalPaid: 0 };
        byEmail[r.email].req++;
        if (r.condition === 'Paid') { byEmail[r.email].paid++; byEmail[r.email].totalPaid += (r.total ?? r.submitted); }
        if (r.condition === 'Rejected') byEmail[r.email].rej++;
    }
    const traders = Object.values(byEmail).sort((a, b) => b.totalPaid - a.totalPaid).slice(0, 20);
    document.getElementById('traders-tbody').innerHTML = traders.map((t, i) => `
    <tr>
      <td class="${i === 0 ? 'rank-gold' : i === 1 ? 'rank-s' : i === 2 ? 'rank-b' : ''}">${i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : i + 1}</td>
      <td style="color:var(--t1);font-weight:600">${t.name}</td>
      <td style="color:var(--t3);font-size:11px">${t.email}</td>
      <td>${t.req}</td>
      <td style="color:var(--green)">${t.paid}</td>
      <td style="color:var(--red)">${t.rej}</td>
      <td style="color:var(--gold);font-weight:700">${fmtUSD(t.totalPaid)}</td>
      <td style="color:var(--t2)">${t.paid > 0 ? fmtUSD(t.totalPaid / t.paid) : '—'}</td>
    </tr>`).join('');
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function exportExcel() {
    const data = FILTERED.map(r => ({
        'Fecha': r.date, 'Cuenta': r.account, 'Nombre': r.name, 'Email': r.email,
        'Método': r.method, 'Wallet': r.wallet,
        'Solicitado ($)': r.submitted, 'Descuento ($)': r.deduction || '',
        'Total ($)': r.total ?? '', 'Status': r.condition,
        'Comentario': r.comment, 'Mes': r.sheet
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [12, 14, 26, 32, 14, 46, 14, 14, 14, 12, 50, 14].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Payouts');
    XLSX.writeFile(wb, `OrionFunded_Payouts_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ─── TAB NAVIGATION ───────────────────────────────────────────────────────────
const TABS = ['dashboard', 'records', 'analysis'];
const TAB_TITLES = { dashboard: '📊 Dashboard', records: '📋 Registros de Payout', analysis: '🔍 Análisis' };

function showTab(tab, btn) {
    TABS.forEach(t => {
        document.getElementById(`tab-${t}`).style.display = t === tab ? '' : 'none';
    });
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    if (btn) btn.classList.add('active');
    document.getElementById('topbar-title').textContent = TAB_TITLES[tab];
    document.getElementById('btn-export').style.display = tab === 'records' ? 'inline-flex' : 'none';
}

// ─── LOAD FILE ────────────────────────────────────────────────────────────────
function loadFile(file) {
    const dropContent = document.getElementById('drop-content');
    if (dropContent) dropContent.innerHTML = '<div class="loading"><div class="spinner"></div><span>Procesando…</span></div>';

    const reader = new FileReader();
    reader.onload = e => {
        try {
            ALL_RECORDS = parseXlsx(e.target.result);
            FILTERED = [...ALL_RECORDS];
            const stats = computeStats(ALL_RECORDS);

            // Switch screens
            document.getElementById('screen-upload').style.display = 'none';
            document.getElementById('screen-dashboard').style.display = '';

            // Sidebar info
            document.getElementById('si-filename').textContent = file.name;
            document.getElementById('si-count').textContent = `${fmtN(ALL_RECORDS.length)} registros`;

            // Topbar date
            document.getElementById('topbar-date').textContent =
                `Orion Funded · Dpto. Riesgo · ${new Date().toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' })}`;

            // Render all
            renderStats(stats);
            renderCharts(stats);
            renderAnalysis(stats);
            setupFilters(ALL_RECORDS);
            renderTable();
            document.getElementById('rec-count-badge').textContent = fmtN(ALL_RECORDS.length);
            showTab('dashboard');
        } catch (err) {
            alert('Error procesando el archivo: ' + err.message);
            if (dropContent) dropContent.innerHTML = `<div class="drop-icon">📊</div><div class="drop-title">Upload PAYOUTS.xlsx</div><div class="drop-sub">Drag & drop o click para seleccionar</div><div style="margin-top:14px"><span class="btn btn-primary">Elegir Archivo</span></div>`;
        }
    };
    reader.readAsArrayBuffer(file);
}

// ─── EVENTS ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // File input 1 (upload screen)
    document.getElementById('file-input').addEventListener('change', e => {
        if (e.target.files[0]) loadFile(e.target.files[0]);
    });
    // File input 2 (sidebar)
    document.getElementById('file-input2').addEventListener('change', e => {
        if (e.target.files[0]) loadFile(e.target.files[0]);
    });

    // Drag & drop
    const dz = document.getElementById('drop-zone');
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
    dz.addEventListener('drop', e => {
        e.preventDefault(); dz.classList.remove('drag');
        const f = e.dataTransfer.files[0];
        if (f) loadFile(f);
    });
});
