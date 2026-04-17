'use client';

import { useState, useCallback, useRef, useMemo } from 'react';
import {
    AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
    XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from 'recharts';
import { parseXlsx, computeStats, filterRecords, exportToXlsx } from '@/lib/parser';
import { PayoutRecord, DashboardStats, FilterState } from '@/lib/types';

const ITEMS_PER_PAGE = 50;

const INITIAL_FILTERS: FilterState = {
    search: '', status: 'all', method: 'all', month: 'all',
    dateFrom: '', dateTo: '', minAmount: '', maxAmount: '', rejectionReason: 'all',
};

function fmtCurrency(v: number) {
    if (v >= 1_000_000) return `$${(v / 1_000_000).toFixed(2)}M`;
    if (v >= 1_000) return `$${(v / 1_000).toFixed(1)}K`;
    return `$${v.toFixed(2)}`;
}

function fmtNum(v: number) {
    return v.toLocaleString('en-US');
}

function StatusBadge({ status }: { status: string }) {
    const classes: Record<string, string> = {
        Paid: 'status-paid', Rejected: 'status-rejected',
        Pending: 'status-pending', Submitted: 'status-pending', Approved: 'status-approved',
    };
    const icons: Record<string, string> = {
        Paid: '✓', Rejected: '✕', Pending: '◌', Submitted: '◌', Approved: '✓',
    };
    return (
        <span className={`status-badge ${classes[status] ?? 'status-pending'}`}>
            {icons[status] ?? '?'} {status}
        </span>
    );
}

function MethodBadge({ method }: { method: string }) {
    const icons: Record<string, string> = {
        'USDT-TRC20': '🟡', 'USDT-ERC20': '🔵', 'RISE': '⚡', 'PAYPAL': '🅿',
    };
    return (
        <span className="method-badge">
            {icons[method] ?? '💱'} {method}
        </span>
    );
}

function CommentTags({ comment }: { comment: string }) {
    if (!comment) return null;
    const tags = comment.split(/\/|,/).map(p => p.trim()).filter(Boolean);
    return (
        <span>
            {tags.map((t, i) => <span key={i} className="comment-tag">{t}</span>)}
        </span>
    );
}

const CUSTOM_TOOLTIP_STYLE = {
    background: '#0f1318',
    border: '1px solid rgba(255,255,255,0.08)',
    borderRadius: '10px',
    color: '#f0f4ff',
    fontSize: '12px',
    padding: '10px 14px',
};

export default function Dashboard() {
    const [records, setRecords] = useState<PayoutRecord[]>([]);
    const [stats, setStats] = useState<DashboardStats | null>(null);
    const [loading, setLoading] = useState(false);
    const [fileName, setFileName] = useState('');
    const [activeTab, setActiveTab] = useState<'dashboard' | 'table' | 'analysis'>('dashboard');
    const [filters, setFilters] = useState<FilterState>(INITIAL_FILTERS);
    const [page, setPage] = useState(1);
    const [sortBy, setSortBy] = useState<keyof PayoutRecord>('date');
    const [sortDir, setSortDir] = useState<'asc' | 'desc'>('desc');
    const [isDrag, setIsDrag] = useState(false);
    const fileRef = useRef<HTMLInputElement>(null);

    const processFile = useCallback(async (file: File) => {
        setLoading(true);
        setFileName(file.name);
        try {
            const buf = await file.arrayBuffer();
            const parsed = parseXlsx(buf);
            const computed = computeStats(parsed);
            setRecords(parsed);
            setStats(computed);
            setActiveTab('dashboard');
            setFilters(INITIAL_FILTERS);
            setPage(1);
        } finally {
            setLoading(false);
        }
    }, []);

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const f = e.target.files?.[0];
        if (f) processFile(f);
    };

    const handleDrop = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDrag(false);
        const f = e.dataTransfer.files?.[0];
        if (f) processFile(f);
    }, [processFile]);

    const handleDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDrag(true); };
    const handleDragLeave = () => setIsDrag(false);

    // Unique values for filters
    const filterOptions = useMemo(() => {
        const months = [...new Set(records.map(r => r.sheet))];
        const methods = [...new Set(records.map(r => r.method).filter(m => m && m !== 'Unknown'))];
        const reasons = [...new Set(
            records.flatMap(r => r.comment ? r.comment.split(/\/|,/).map(p => p.trim()).filter(Boolean) : [])
        )].sort();
        return { months, methods, reasons };
    }, [records]);

    // Filtered + sorted records
    const filtered = useMemo(() => {
        let res = filterRecords(records, filters);
        res = [...res].sort((a, b) => {
            const av = a[sortBy] ?? '';
            const bv = b[sortBy] ?? '';
            let cmp = 0;
            if (typeof av === 'number' && typeof bv === 'number') cmp = av - bv;
            else cmp = String(av).localeCompare(String(bv));
            return sortDir === 'asc' ? cmp : -cmp;
        });
        return res;
    }, [records, filters, sortBy, sortDir]);

    const totalPages = Math.ceil(filtered.length / ITEMS_PER_PAGE);
    const paginated = filtered.slice((page - 1) * ITEMS_PER_PAGE, page * ITEMS_PER_PAGE);

    const handleSort = (col: keyof PayoutRecord) => {
        if (sortBy === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
        else { setSortBy(col); setSortDir('desc'); }
    };

    const applyFilter = (key: keyof FilterState, val: string) => {
        setFilters(f => ({ ...f, [key]: val }));
        setPage(1);
    };

    const handleExport = () => {
        const date = new Date().toISOString().split('T')[0];
        exportToXlsx(filtered, `OrionFunded_Payouts_${date}.xlsx`);
    };

    const SortIcon = ({ col }: { col: keyof PayoutRecord }) => {
        if (sortBy !== col) return <span style={{ opacity: 0.3 }}>⇅</span>;
        return <span style={{ color: 'var(--accent-gold)' }}>{sortDir === 'asc' ? '↑' : '↓'}</span>;
    };

    // ── UPLOAD SCREEN ──
    if (!stats) {
        return (
            <div className="app-container" style={{ alignItems: 'center', justifyContent: 'center', minHeight: '100vh' }}>
                <div style={{ width: '100%', maxWidth: 560, padding: '24px' }}>
                    <div style={{ textAlign: 'center', marginBottom: 40 }}>
                        <div style={{
                            width: 72, height: 72, background: 'linear-gradient(135deg, #f5c842, #e9a20a)',
                            borderRadius: 20, display: 'flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: 36, margin: '0 auto 20px',
                            boxShadow: '0 8px 32px rgba(245,200,66,0.3)',
                        }}>⚡</div>
                        <h1 style={{ fontSize: 32, fontWeight: 900, marginBottom: 8, background: 'linear-gradient(135deg, #f5c842, #f0f4ff)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                            Orion Funded
                        </h1>
                        <p style={{ color: 'var(--text-muted)', fontSize: 14 }}>Risk Department · Payout Dashboard</p>
                    </div>

                    <div
                        className={`upload-zone ${isDrag ? 'drag-over' : ''}`}
                        onClick={() => fileRef.current?.click()}
                        onDrop={handleDrop}
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                    >
                        {loading ? (
                            <div className="loading-overlay" style={{ padding: 20 }}>
                                <div className="spinner" />
                                <span>Processing PAYOUTS.xlsx…</span>
                            </div>
                        ) : (
                            <>
                                <div className="upload-icon">📊</div>
                                <div className="upload-title">Upload PAYOUTS.xlsx</div>
                                <div className="upload-sub">Drag & drop or click to select your file</div>
                                <div style={{ marginTop: 16 }}>
                                    <span className="btn btn-primary">Choose File</span>
                                </div>
                            </>
                        )}
                    </div>
                    <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleFileChange} />

                    <div style={{ display: 'flex', gap: 12, marginTop: 24 }}>
                        {['✓ Multi-sheet support', '✓ Auto-detection', '✓ Export Excel'].map(t => (
                            <div key={t} style={{ flex: 1, background: 'var(--bg-card)', border: '1px solid var(--border)', borderRadius: 10, padding: '10px 12px', fontSize: 11, color: 'var(--text-muted)', textAlign: 'center' }}>{t}</div>
                        ))}
                    </div>
                </div>
            </div>
        );
    }

    // ── MAIN DASHBOARD ──
    return (
        <div className="app-container">
            {/* SIDEBAR */}
            <nav className="sidebar">
                <div className="sidebar-logo">
                    <div className="logo-icon">⚡</div>
                    <div className="logo-text">
                        <div className="logo-title">ORION FUNDED</div>
                        <div className="logo-sub">Risk Dept.</div>
                    </div>
                </div>

                <div className="sidebar-nav">
                    <span className="nav-label">Navigation</span>
                    {[
                        { id: 'dashboard', icon: '📊', label: 'Dashboard' },
                        { id: 'table', icon: '📋', label: 'Payout Records' },
                        { id: 'analysis', icon: '🔍', label: 'Analysis' },
                    ].map(item => (
                        <button
                            key={item.id}
                            className={`nav-item ${activeTab === item.id ? 'active' : ''}`}
                            onClick={() => setActiveTab(item.id as typeof activeTab)}
                        >
                            <span className="nav-icon">{item.icon}</span>
                            {item.label}
                        </button>
                    ))}

                    <span className="nav-label" style={{ marginTop: 16 }}>Data</span>
                    <button className="nav-item" onClick={() => fileRef.current?.click()}>
                        <span className="nav-icon">📁</span>
                        Load New File
                    </button>
                    <button className="nav-item" onClick={handleExport}>
                        <span className="nav-icon">⬇️</span>
                        Export Excel
                    </button>
                </div>

                {/* File info */}
                <div style={{ padding: '12px 16px', margin: '0 12px 0', background: 'var(--bg-glass)', borderRadius: 10, border: '1px solid var(--border)' }}>
                    <div style={{ fontSize: 10, color: 'var(--text-muted)', marginBottom: 4, textTransform: 'uppercase', fontWeight: 600 }}>Loaded File</div>
                    <div style={{ fontSize: 12, color: 'var(--text-secondary)', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{fileName}</div>
                    <div style={{ fontSize: 11, color: 'var(--accent-gold)', marginTop: 4 }}>{fmtNum(records.length)} records</div>
                </div>

                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleFileChange} />
            </nav>

            {/* MAIN */}
            <div className="main-content">
                {/* TOP BAR */}
                <div className="topbar">
                    <div className="topbar-left">
                        <h1>{activeTab === 'dashboard' ? '📊 Dashboard' : activeTab === 'table' ? '📋 Payout Records' : '🔍 Analysis'}</h1>
                        <p>Orion Funded · Risk Department · {new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}</p>
                    </div>
                    <div className="topbar-right">
                        <button className="btn btn-secondary btn-sm" onClick={() => fileRef.current?.click()}>
                            📁 Load File
                        </button>
                        {activeTab === 'table' && (
                            <button className="btn btn-primary btn-sm" onClick={handleExport}>
                                ⬇️ Export {fmtNum(filtered.length)} rows
                            </button>
                        )}
                    </div>
                </div>

                <div className="content-area">
                    {/* ═══════════════════════════════ DASHBOARD TAB */}
                    {activeTab === 'dashboard' && (
                        <>
                            {/* KPI STATS */}
                            <div className="stats-grid">
                                <div className="stat-card gold">
                                    <div className="stat-icon icon-gold">💼</div>
                                    <div className="stat-label">Total Requests</div>
                                    <div className="stat-value">{fmtNum(stats.totalRequests)}</div>
                                    <div className="stat-sub">{filterOptions.months.length} monthly sheets</div>
                                </div>
                                <div className="stat-card green">
                                    <div className="stat-icon icon-green">✅</div>
                                    <div className="stat-label">Paid</div>
                                    <div className="stat-value">{fmtNum(stats.totalPaid)}</div>
                                    <div className="stat-sub">
                                        <span className="stat-badge badge-up">↑ {stats.approvalRate.toFixed(1)}% rate</span>
                                    </div>
                                </div>
                                <div className="stat-card red">
                                    <div className="stat-icon icon-red">❌</div>
                                    <div className="stat-label">Rejected</div>
                                    <div className="stat-value">{fmtNum(stats.totalRejected)}</div>
                                    <div className="stat-sub">
                                        <span className="stat-badge badge-down">↑ {stats.rejectionRate.toFixed(1)}% rate</span>
                                    </div>
                                </div>
                                <div className="stat-card blue">
                                    <div className="stat-icon icon-blue">⏳</div>
                                    <div className="stat-label">Pending/Other</div>
                                    <div className="stat-value">{fmtNum(stats.totalPending)}</div>
                                    <div className="stat-sub">Awaiting decision</div>
                                </div>
                                <div className="stat-card gold">
                                    <div className="stat-icon icon-gold">💰</div>
                                    <div className="stat-label">Total Paid Out</div>
                                    <div className="stat-value">{fmtCurrency(stats.totalPaidAmount)}</div>
                                    <div className="stat-sub">Avg {fmtCurrency(stats.avgPaidAmount)} per payout</div>
                                </div>
                                <div className="stat-card purple">
                                    <div className="stat-icon icon-purple">📨</div>
                                    <div className="stat-label">Total Submitted</div>
                                    <div className="stat-value">{fmtCurrency(stats.totalSubmittedAmount)}</div>
                                    <div className="stat-sub">Deductions: {fmtCurrency(stats.totalDeductions)}</div>
                                </div>
                                <div className="stat-card blue">
                                    <div className="stat-icon icon-blue">👥</div>
                                    <div className="stat-label">Unique Traders</div>
                                    <div className="stat-value">{fmtNum(stats.uniqueTraders)}</div>
                                    <div className="stat-sub">by email</div>
                                </div>
                            </div>

                            {/* CHARTS ROW 1 */}
                            <div className="chart-grid">
                                {/* Monthly Volume */}
                                <div className="chart-container" style={{ gridColumn: 'span 2' }}>
                                    <div className="chart-title">📈 Monthly Volume & Amounts</div>
                                    <ResponsiveContainer width="100%" height={240}>
                                        <AreaChart data={stats.monthlyData} margin={{ top: 4, right: 16, bottom: 4, left: 16 }}>
                                            <defs>
                                                <linearGradient id="gradPaid" x1="0" y1="0" x2="0" y2="1">
                                                    <stop offset="5%" stopColor="#2dd99e" stopOpacity={0.3} />
                                                    <stop offset="95%" stopColor="#2dd99e" stopOpacity={0} />
                                                </linearGradient>
                                                <linearGradient id="gradSubmit" x1="0" y1="0" x2="0" y2="1">
                                                    <stop offset="5%" stopColor="#f5c842" stopOpacity={0.2} />
                                                    <stop offset="95%" stopColor="#f5c842" stopOpacity={0} />
                                                </linearGradient>
                                            </defs>
                                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                                            <XAxis dataKey="month" tick={{ fill: '#4a566a', fontSize: 10 }} />
                                            <YAxis tick={{ fill: '#4a566a', fontSize: 10 }} tickFormatter={v => `$${(v / 1000).toFixed(0)}K`} />
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} formatter={(v: number) => fmtCurrency(v)} />
                                            <Legend wrapperStyle={{ fontSize: 11, color: '#8b9ab5' }} />
                                            <Area type="monotone" dataKey="totalPaidAmount" name="Paid Amount" stroke="#2dd99e" fill="url(#gradPaid)" strokeWidth={2} />
                                            <Area type="monotone" dataKey="totalSubmittedAmount" name="Submitted" stroke="#f5c842" fill="url(#gradSubmit)" strokeWidth={2} />
                                        </AreaChart>
                                    </ResponsiveContainer>
                                </div>

                                {/* Status Pie */}
                                <div className="chart-container">
                                    <div className="chart-title">🍩 Status Distribution</div>
                                    <ResponsiveContainer width="100%" height={240}>
                                        <PieChart>
                                            <Pie
                                                data={stats.statusDistribution}
                                                cx="50%" cy="50%"
                                                innerRadius={60} outerRadius={90}
                                                paddingAngle={3}
                                                dataKey="value"
                                                label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                                                labelLine={false}
                                            >
                                                {stats.statusDistribution.map((entry, i) => (
                                                    <Cell key={i} fill={entry.color} />
                                                ))}
                                            </Pie>
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} />
                                        </PieChart>
                                    </ResponsiveContainer>
                                </div>

                                {/* Monthly requests bar */}
                                <div className="chart-container">
                                    <div className="chart-title">📊 Monthly Requests</div>
                                    <ResponsiveContainer width="100%" height={240}>
                                        <BarChart data={stats.monthlyData} margin={{ top: 4, right: 8, bottom: 4, left: 8 }}>
                                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                                            <XAxis dataKey="month" tick={{ fill: '#4a566a', fontSize: 9 }} />
                                            <YAxis tick={{ fill: '#4a566a', fontSize: 10 }} />
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} />
                                            <Bar dataKey="paid" name="Paid" fill="#2dd99e" radius={[4, 4, 0, 0]} />
                                            <Bar dataKey="rejected" name="Rejected" fill="#f05c6e" radius={[4, 4, 0, 0]} />
                                        </BarChart>
                                    </ResponsiveContainer>
                                </div>
                            </div>

                            {/* CHARTS ROW 2 */}
                            <div className="chart-grid">
                                {/* Payment Methods */}
                                <div className="chart-container">
                                    <div className="chart-title">💳 Payment Methods</div>
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: '4px 0' }}>
                                        {Object.entries(stats.methodBreakdown)
                                            .filter(([m]) => ['USDT-TRC20', 'USDT-ERC20', 'RISE', 'PAYPAL'].includes(m))
                                            .sort(([, a], [, b]) => b - a)
                                            .map(([method, count]) => {
                                                const total = Object.values(stats.methodBreakdown).reduce((a, b) => a + b, 0);
                                                const pct = ((count / total) * 100).toFixed(1);
                                                const colors: Record<string, string> = {
                                                    'USDT-TRC20': '#2dd99e', 'USDT-ERC20': '#4f8ef7', 'RISE': '#f5c842', 'PAYPAL': '#a78bfa',
                                                };
                                                return (
                                                    <div key={method}>
                                                        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4, fontSize: 12 }}>
                                                            <span style={{ color: 'var(--text-secondary)' }}>{method}</span>
                                                            <span style={{ color: colors[method], fontWeight: 700 }}>{fmtNum(count)} ({pct}%)</span>
                                                        </div>
                                                        <div style={{ background: 'var(--bg-secondary)', borderRadius: 4, height: 6, overflow: 'hidden' }}>
                                                            <div style={{ width: `${pct}%`, background: colors[method], height: '100%', borderRadius: 4, transition: 'width 0.5s ease' }} />
                                                        </div>
                                                    </div>
                                                );
                                            })}
                                    </div>
                                </div>

                                {/* Top rejection reasons */}
                                <div className="chart-container">
                                    <div className="chart-title">🚫 Top Rejection Reasons</div>
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 240, overflowY: 'auto' }}>
                                        {Object.entries(stats.rejectionReasons)
                                            .sort(([, a], [, b]) => b - a)
                                            .slice(0, 10)
                                            .map(([reason, count]) => (
                                                <div key={reason} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
                                                    <span style={{ color: 'var(--text-secondary)', fontSize: 12, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{reason}</span>
                                                    <span style={{ background: 'var(--red-dim)', color: 'var(--red)', padding: '2px 8px', borderRadius: 20, fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{count}</span>
                                                </div>
                                            ))}
                                    </div>
                                </div>

                                {/* Monthly table summary */}
                                <div className="chart-container">
                                    <div className="chart-title">📅 Monthly Summary</div>
                                    <div style={{ overflowX: 'auto' }}>
                                        <table className="data-table">
                                            <thead>
                                                <tr>
                                                    <th>Month</th>
                                                    <th>Requests</th>
                                                    <th>Paid</th>
                                                    <th>Rejected</th>
                                                    <th>Paid Amount</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {stats.monthlyData.map(m => (
                                                    <tr key={m.month}>
                                                        <td style={{ color: 'var(--text-primary)', fontWeight: 600 }}>{m.month}</td>
                                                        <td>{fmtNum(m.requests)}</td>
                                                        <td style={{ color: 'var(--green)' }}>{fmtNum(m.paid)}</td>
                                                        <td style={{ color: 'var(--red)' }}>{fmtNum(m.rejected)}</td>
                                                        <td style={{ color: 'var(--accent-gold)', fontWeight: 600 }}>{fmtCurrency(m.totalPaidAmount)}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        </>
                    )}

                    {/* ═══════════════════════════════ TABLE TAB */}
                    {activeTab === 'table' && (
                        <div className="section-card">
                            <div className="section-header">
                                <div className="section-title">
                                    <div className="section-title-icon">📋</div>
                                    Payout Records
                                    <span style={{ background: 'var(--accent-gold-dim)', color: 'var(--accent-gold)', padding: '2px 10px', borderRadius: 20, fontSize: 11, fontWeight: 700 }}>
                                        {fmtNum(filtered.length)} of {fmtNum(records.length)}
                                    </span>
                                </div>
                                <div className="filters-bar">
                                    <input
                                        className="filter-input"
                                        placeholder="🔍 Search name, email, account…"
                                        value={filters.search}
                                        onChange={e => applyFilter('search', e.target.value)}
                                    />
                                    <select className="filter-select" value={filters.status} onChange={e => applyFilter('status', e.target.value)}>
                                        <option value="all">All Status</option>
                                        <option value="Paid">✅ Paid</option>
                                        <option value="Rejected">❌ Rejected</option>
                                        <option value="Pending">⏳ Pending</option>
                                        <option value="Approved">✓ Approved</option>
                                        <option value="Submitted">📩 Submitted</option>
                                    </select>
                                    <select className="filter-select" value={filters.method} onChange={e => applyFilter('method', e.target.value)}>
                                        <option value="all">All Methods</option>
                                        {filterOptions.methods.map(m => <option key={m} value={m}>{m}</option>)}
                                    </select>
                                    <select className="filter-select" value={filters.month} onChange={e => applyFilter('month', e.target.value)}>
                                        <option value="all">All Months</option>
                                        {filterOptions.months.map(m => <option key={m} value={m}>{m}</option>)}
                                    </select>
                                    <select className="filter-select" value={filters.rejectionReason} onChange={e => applyFilter('rejectionReason', e.target.value)}>
                                        <option value="all">All Reasons</option>
                                        {filterOptions.reasons.slice(0, 20).map(r => <option key={r} value={r}>{r}</option>)}
                                    </select>
                                </div>
                                <div className="filters-bar">
                                    <input type="date" className="filter-select" value={filters.dateFrom} onChange={e => applyFilter('dateFrom', e.target.value)} style={{ cursor: 'pointer' }} title="From date" />
                                    <input type="date" className="filter-select" value={filters.dateTo} onChange={e => applyFilter('dateTo', e.target.value)} style={{ cursor: 'pointer' }} title="To date" />
                                    <input type="number" className="filter-input" placeholder="Min $" value={filters.minAmount} onChange={e => applyFilter('minAmount', e.target.value)} style={{ minWidth: 80, maxWidth: 100 }} />
                                    <input type="number" className="filter-input" placeholder="Max $" value={filters.maxAmount} onChange={e => applyFilter('maxAmount', e.target.value)} style={{ minWidth: 80, maxWidth: 100 }} />
                                    <button className="btn btn-secondary btn-sm" onClick={() => { setFilters(INITIAL_FILTERS); setPage(1); }}>
                                        ✕ Clear
                                    </button>
                                    <button className="btn btn-primary btn-sm" onClick={handleExport}>
                                        ⬇️ Export Excel
                                    </button>
                                </div>
                            </div>

                            {/* Summary row */}
                            <div style={{ padding: '10px 22px', borderBottom: '1px solid var(--border)' }}>
                                <div className="summary-row">
                                    <span>📊 Total: <strong>{fmtNum(filtered.length)}</strong></span>
                                    <span style={{ color: 'var(--green)' }}>✅ Paid: <strong>{fmtNum(filtered.filter(r => r.condition === 'Paid').length)}</strong></span>
                                    <span style={{ color: 'var(--red)' }}>❌ Rejected: <strong>{fmtNum(filtered.filter(r => r.condition === 'Rejected').length)}</strong></span>
                                    <span style={{ color: 'var(--accent-gold)' }}>💰 Submitted: <strong>{fmtCurrency(filtered.reduce((s, r) => s + r.submitted, 0))}</strong></span>
                                    <span style={{ color: 'var(--green)' }}>✓ Paid out: <strong>{fmtCurrency(filtered.filter(r => r.condition === 'Paid').reduce((s, r) => s + (r.total ?? r.submitted), 0))}</strong></span>
                                </div>
                            </div>

                            <div className="table-container">
                                <table className="data-table">
                                    <thead>
                                        <tr>
                                            {[
                                                { key: 'date', label: 'Date' },
                                                { key: 'account', label: 'Account' },
                                                { key: 'name', label: 'Name' },
                                                { key: 'method', label: 'Method' },
                                                { key: 'submitted', label: 'Submitted' },
                                                { key: 'deduction', label: 'Deduction' },
                                                { key: 'total', label: 'Total' },
                                                { key: 'condition', label: 'Status' },
                                                { key: 'comment', label: 'Reason / Comment' },
                                                { key: 'sheet', label: 'Month' },
                                            ].map(col => (
                                                <th key={col.key} onClick={() => handleSort(col.key as keyof PayoutRecord)}>
                                                    {col.label} <SortIcon col={col.key as keyof PayoutRecord} />
                                                </th>
                                            ))}
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {paginated.length === 0 ? (
                                            <tr>
                                                <td colSpan={10}>
                                                    <div className="empty-state">
                                                        <div className="empty-state-icon">🔍</div>
                                                        <div className="empty-state-title">No records found</div>
                                                        <div className="empty-state-sub">Try adjusting your filters</div>
                                                    </div>
                                                </td>
                                            </tr>
                                        ) : paginated.map(r => (
                                            <tr key={r.id}>
                                                <td style={{ color: 'var(--text-primary)', fontWeight: 500 }}>{r.date}</td>
                                                <td className="font-mono" style={{ fontSize: 11 }}>{r.account || '—'}</td>
                                                <td style={{ color: 'var(--text-primary)', maxWidth: 180 }}>{r.name}</td>
                                                <td>{r.method && r.method !== 'Unknown' ? <MethodBadge method={r.method} /> : <span style={{ color: 'var(--text-muted)' }}>—</span>}</td>
                                                <td style={{ color: 'var(--accent-gold)', fontWeight: 600 }}>${r.submitted.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                                <td style={{ color: r.deduction > 0 ? 'var(--red)' : 'var(--text-muted)' }}>
                                                    {r.deduction > 0 ? `-$${r.deduction.toFixed(2)}` : '—'}
                                                </td>
                                                <td style={{ color: 'var(--green)', fontWeight: 600 }}>
                                                    {r.total != null ? `$${r.total.toFixed(2)}` : '—'}
                                                </td>
                                                <td><StatusBadge status={r.condition} /></td>
                                                <td style={{ maxWidth: 220 }}><CommentTags comment={r.comment} /></td>
                                                <td style={{ color: 'var(--text-muted)', fontSize: 11 }}>{r.sheet}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            {/* Pagination */}
                            {totalPages > 1 && (
                                <div className="pagination">
                                    <span>Showing {((page - 1) * ITEMS_PER_PAGE) + 1}–{Math.min(page * ITEMS_PER_PAGE, filtered.length)} of {fmtNum(filtered.length)}</span>
                                    <div className="pagination-controls">
                                        <button className="page-btn" onClick={() => setPage(1)} disabled={page === 1}>«</button>
                                        <button className="page-btn" onClick={() => setPage(p => p - 1)} disabled={page === 1}>‹</button>
                                        {Array.from({ length: Math.min(7, totalPages) }, (_, i) => {
                                            let p: number;
                                            if (totalPages <= 7) p = i + 1;
                                            else if (page <= 4) p = i + 1;
                                            else if (page >= totalPages - 3) p = totalPages - 6 + i;
                                            else p = page - 3 + i;
                                            return (
                                                <button key={p} className={`page-btn ${page === p ? 'active' : ''}`} onClick={() => setPage(p)}>{p}</button>
                                            );
                                        })}
                                        <button className="page-btn" onClick={() => setPage(p => p + 1)} disabled={page === totalPages}>›</button>
                                        <button className="page-btn" onClick={() => setPage(totalPages)} disabled={page === totalPages}>»</button>
                                    </div>
                                </div>
                            )}
                        </div>
                    )}

                    {/* ═══════════════════════════════ ANALYSIS TAB */}
                    {activeTab === 'analysis' && (
                        <>
                            <div className="chart-grid">
                                {/* Rejection reasons bar chart */}
                                <div className="chart-container" style={{ gridColumn: 'span 2' }}>
                                    <div className="chart-title">🚫 Rejection Reasons Breakdown</div>
                                    <ResponsiveContainer width="100%" height={280}>
                                        <BarChart
                                            layout="vertical"
                                            data={Object.entries(stats.rejectionReasons).sort(([, a], [, b]) => b - a).slice(0, 12).map(([name, value]) => ({ name: name.length > 30 ? name.slice(0, 30) + '…' : name, value }))}
                                            margin={{ top: 4, right: 24, bottom: 4, left: 180 }}
                                        >
                                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                                            <XAxis type="number" tick={{ fill: '#4a566a', fontSize: 10 }} />
                                            <YAxis type="category" dataKey="name" tick={{ fill: '#8b9ab5', fontSize: 10 }} width={180} />
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} />
                                            <Bar dataKey="value" name="Count" fill="#f05c6e" radius={[0, 4, 4, 0]} />
                                        </BarChart>
                                    </ResponsiveContainer>
                                </div>

                                {/* Payment method pie */}
                                <div className="chart-container">
                                    <div className="chart-title">💳 Method Distribution</div>
                                    <ResponsiveContainer width="100%" height={240}>
                                        <PieChart>
                                            <Pie
                                                data={Object.entries(stats.methodBreakdown)
                                                    .filter(([m]) => ['USDT-TRC20', 'USDT-ERC20', 'RISE', 'PAYPAL'].includes(m))
                                                    .map(([name, value]) => ({ name, value }))}
                                                cx="50%" cy="50%"
                                                outerRadius={80}
                                                paddingAngle={3}
                                                dataKey="value"
                                                label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                                            >
                                                {['#2dd99e', '#4f8ef7', '#f5c842', '#a78bfa'].map((c, i) => (
                                                    <Cell key={i} fill={c} />
                                                ))}
                                            </Pie>
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} />
                                        </PieChart>
                                    </ResponsiveContainer>
                                </div>

                                {/* Month-over-month approval rate */}
                                <div className="chart-container">
                                    <div className="chart-title">📈 Approval Rate by Month</div>
                                    <ResponsiveContainer width="100%" height={240}>
                                        <BarChart
                                            data={stats.monthlyData.map(m => ({
                                                month: m.month.slice(0, 3),
                                                rate: m.requests > 0 ? Math.round((m.paid / m.requests) * 100) : 0,
                                            }))}
                                            margin={{ top: 4, right: 8, bottom: 4, left: 8 }}
                                        >
                                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                                            <XAxis dataKey="month" tick={{ fill: '#4a566a', fontSize: 10 }} />
                                            <YAxis tick={{ fill: '#4a566a', fontSize: 10 }} unit="%" />
                                            <Tooltip contentStyle={CUSTOM_TOOLTIP_STYLE} formatter={(v: number) => `${v}%`} />
                                            <Bar dataKey="rate" name="Approval %" fill="#f5c842" radius={[4, 4, 0, 0]} />
                                        </BarChart>
                                    </ResponsiveContainer>
                                </div>
                            </div>

                            {/* Top traders */}
                            <div className="section-card">
                                <div className="section-header">
                                    <div className="section-title">
                                        <div className="section-title-icon">👥</div>
                                        Top Traders by Paid Amount
                                    </div>
                                </div>
                                <div className="table-container" style={{ maxHeight: 400 }}>
                                    <table className="data-table">
                                        <thead>
                                            <tr>
                                                <th>#</th>
                                                <th>Name</th>
                                                <th>Email</th>
                                                <th>Requests</th>
                                                <th>Paid</th>
                                                <th>Rejected</th>
                                                <th>Total Paid</th>
                                                <th>Avg Payout</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {useMemo(() => {
                                                const byEmail: Record<string, { name: string; email: string; requests: number; paid: number; rejected: number; totalPaid: number }> = {};
                                                for (const r of records) {
                                                    if (!r.email) continue;
                                                    if (!byEmail[r.email]) byEmail[r.email] = { name: r.name, email: r.email, requests: 0, paid: 0, rejected: 0, totalPaid: 0 };
                                                    byEmail[r.email].requests++;
                                                    if (r.condition === 'Paid') { byEmail[r.email].paid++; byEmail[r.email].totalPaid += (r.total ?? r.submitted); }
                                                    if (r.condition === 'Rejected') byEmail[r.email].rejected++;
                                                }
                                                return Object.values(byEmail).sort((a, b) => b.totalPaid - a.totalPaid).slice(0, 20);
                                            }, [records]).map((t, i) => (
                                                <tr key={t.email}>
                                                    <td style={{ color: i < 3 ? 'var(--accent-gold)' : 'var(--text-muted)', fontWeight: 700 }}>
                                                        {i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : i + 1}
                                                    </td>
                                                    <td style={{ color: 'var(--text-primary)', fontWeight: 600 }}>{t.name}</td>
                                                    <td style={{ color: 'var(--text-muted)', fontSize: 11 }}>{t.email}</td>
                                                    <td>{t.requests}</td>
                                                    <td style={{ color: 'var(--green)' }}>{t.paid}</td>
                                                    <td style={{ color: 'var(--red)' }}>{t.rejected}</td>
                                                    <td style={{ color: 'var(--accent-gold)', fontWeight: 700 }}>{fmtCurrency(t.totalPaid)}</td>
                                                    <td style={{ color: 'var(--text-secondary)' }}>{t.paid > 0 ? fmtCurrency(t.totalPaid / t.paid) : '—'}</td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </>
                    )}
                </div>
            </div>
        </div>
    );
}
