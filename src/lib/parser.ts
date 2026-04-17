import * as XLSX from 'xlsx';
import { PayoutRecord, PayoutStatus, DashboardStats, MonthlyData } from './types';

const STATUS_MAP: Record<string, PayoutStatus> = {
    'paid': 'Paid',
    'pagado': 'Paid',
    'rejected': 'Rejected',
    'rechazado': 'Rejected',
    'pending': 'Pending',
    'pendiente': 'Pending',
    'approved': 'Approved',
    'aprobado': 'Approved',
    'submitted': 'Submitted',
    'rise pending': 'Pending',
};

function normalizeStatus(raw: unknown): PayoutStatus {
    if (!raw) return 'Pending';
    const str = String(raw).trim().toLowerCase();
    return STATUS_MAP[str] ?? 'Pending';
}

function normalizeMethod(raw: unknown): string {
    if (!raw) return 'Unknown';
    const str = String(raw).trim();
    // If it looks like a wallet address, return Unknown
    if (str.startsWith('0x') || str.startsWith('T') && str.length > 20) return 'Unknown';
    if (str.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) return 'Unknown';
    if (str.startsWith('http')) return 'Unknown';
    return str;
}

function parseDate(val: unknown): string {
    if (!val) return '';
    if (val instanceof Date) {
        return val.toISOString().split('T')[0];
    }
    if (typeof val === 'number') {
        // Excel serial date
        const d = XLSX.SSF.parse_date_code(val);
        if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }
    const s = String(val);
    const match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (match) return `${match[1]}-${match[2]}-${match[3]}`;
    return '';
}

function parseNum(val: unknown): number {
    if (!val) return 0;
    const n = Number(val);
    return isNaN(n) ? 0 : n;
}

// Detect if sheet uses new format (has 'Account' column) or old format (Spanish)
function detectFormat(headers: unknown[]): 'new' | 'old' {
    const headerStr = headers.map(h => String(h ?? '').toLowerCase()).join(',');
    if (headerStr.includes('account')) return 'new';
    return 'old';
}

export function parseXlsx(buffer: ArrayBuffer): PayoutRecord[] {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const records: PayoutRecord[] = [];

    const MONTH_ORDER = [
        'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST',
        'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER',
        'JANUARY-2026', 'February-2026'
    ];

    for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null }) as unknown[][];
        if (!rows || rows.length < 2) continue;

        const headers = rows[0] as unknown[];
        const format = detectFormat(headers);
        const dataRows = rows.slice(1);

        for (let i = 0; i < dataRows.length; i++) {
            const row = dataRows[i] as unknown[];
            if (!row || row.every(c => c === null || c === undefined || c === '')) continue;

            let record: PayoutRecord;

            if (format === 'new') {
                // Columns: Date(0), Account(1), Name(2), Email(3), Method(4), Wallet(5),
                //          Submitted(6), Deduction(7), Total(8), Condition(9), BackOffice(10), Comment(11)
                record = {
                    id: `${sheetName}-${i}`,
                    date: parseDate(row[0]),
                    account: row[1] ? String(row[1]).replace('.0', '') : '',
                    name: String(row[2] ?? '').trim(),
                    email: String(row[3] ?? '').trim().toLowerCase(),
                    method: normalizeMethod(row[4]),
                    wallet: String(row[5] ?? '').trim(),
                    submitted: parseNum(row[6]),
                    deduction: parseNum(row[7]),
                    total: row[8] ? parseNum(row[8]) : null,
                    condition: normalizeStatus(row[9]),
                    backOffice: !!row[10],
                    comment: String(row[11] ?? '').trim(),
                    month: sheetName.toUpperCase(),
                    sheet: sheetName,
                };
            } else {
                // Old format (Spanish): Fecha(0), Nombre(1), Correo(2), Metodo(3), Wallet(4),
                //                       Solicitado(5), Descuento(6), Total(7), Estado(8), Fecha estado(9), Comentario(10), BackOffice(11)
                record = {
                    id: `${sheetName}-${i}`,
                    date: parseDate(row[0]),
                    account: '',
                    name: String(row[1] ?? '').trim(),
                    email: String(row[2] ?? '').trim().toLowerCase(),
                    method: normalizeMethod(row[3]),
                    wallet: String(row[4] ?? '').trim(),
                    submitted: parseNum(row[5]),
                    deduction: parseNum(row[6]),
                    total: row[7] ? parseNum(row[7]) : null,
                    condition: normalizeStatus(row[8]),
                    backOffice: !!row[11],
                    comment: String(row[10] ?? '').trim(),
                    month: sheetName.toUpperCase(),
                    sheet: sheetName,
                };
            }

            // Filter out garbage rows (no name, no amount)
            if (!record.name || record.submitted === 0) continue;
            records.push(record);
        }
    }

    return records;
}

export function computeStats(records: PayoutRecord[]): DashboardStats {
    const paid = records.filter(r => r.condition === 'Paid');
    const rejected = records.filter(r => r.condition === 'Rejected');
    const pending = records.filter(r => ['Pending', 'Submitted', 'Approved'].includes(r.condition));

    const totalPaidAmount = paid.reduce((s, r) => s + (r.total ?? r.submitted), 0);
    const totalSubmittedAmount = records.reduce((s, r) => s + r.submitted, 0);
    const totalDeductions = records.reduce((s, r) => s + r.deduction, 0);

    // Method breakdown
    const methodBreakdown: Record<string, number> = {};
    for (const r of records) {
        const m = r.method || 'Unknown';
        methodBreakdown[m] = (methodBreakdown[m] ?? 0) + 1;
    }

    // Rejection reasons
    const rejectionReasons: Record<string, number> = {};
    for (const r of rejected) {
        if (!r.comment) continue;
        const parts = r.comment.split(/\/|,/).map(p => p.trim()).filter(Boolean);
        for (const p of parts) {
            const key = p.replace(/\s+/g, ' ').trim();
            if (key) rejectionReasons[key] = (rejectionReasons[key] ?? 0) + 1;
        }
    }

    // Monthly data
    const monthMap: Record<string, MonthlyData> = {};
    for (const r of records) {
        if (!monthMap[r.sheet]) {
            monthMap[r.sheet] = {
                month: r.sheet,
                requests: 0, paid: 0, rejected: 0,
                totalPaidAmount: 0, totalSubmittedAmount: 0,
            };
        }
        const m = monthMap[r.sheet];
        m.requests++;
        m.totalSubmittedAmount += r.submitted;
        if (r.condition === 'Paid') { m.paid++; m.totalPaidAmount += (r.total ?? r.submitted); }
        if (r.condition === 'Rejected') m.rejected++;
    }

    const SHEET_ORDER = [
        'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST',
        'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER',
        'JANUARY-2026', 'February-2026'
    ];

    const monthlyData = [...Object.values(monthMap)].sort((a, b) => {
        const ai = SHEET_ORDER.findIndex(s => s.toUpperCase() === a.month.toUpperCase());
        const bi = SHEET_ORDER.findIndex(s => s.toUpperCase() === b.month.toUpperCase());
        return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
    });

    // Daily data (last 30 days of paid records)
    const dailyMap: Record<string, { requests: number; paid: number; rejected: number; amount: number }> = {};
    for (const r of records) {
        if (!r.date) continue;
        if (!dailyMap[r.date]) dailyMap[r.date] = { requests: 0, paid: 0, rejected: 0, amount: 0 };
        dailyMap[r.date].requests++;
        if (r.condition === 'Paid') { dailyMap[r.date].paid++; dailyMap[r.date].amount += (r.total ?? r.submitted); }
        if (r.condition === 'Rejected') dailyMap[r.date].rejected++;
    }

    const dailyData = Object.entries(dailyMap)
        .sort(([a], [b]) => a.localeCompare(b))
        .slice(-60)
        .map(([date, v]) => ({ date, ...v }));

    const uniqueEmails = new Set(records.map(r => r.email).filter(Boolean));

    return {
        totalRequests: records.length,
        totalPaid: paid.length,
        totalRejected: rejected.length,
        totalPending: pending.length,
        totalSubmittedAmount,
        totalPaidAmount,
        totalDeductions,
        approvalRate: records.length > 0 ? (paid.length / records.length) * 100 : 0,
        rejectionRate: records.length > 0 ? (rejected.length / records.length) * 100 : 0,
        avgPaidAmount: paid.length > 0 ? totalPaidAmount / paid.length : 0,
        uniqueTraders: uniqueEmails.size,
        methodBreakdown,
        rejectionReasons,
        monthlyData,
        dailyData,
        statusDistribution: [
            { name: 'Paid', value: paid.length, color: '#2dd99e' },
            { name: 'Rejected', value: rejected.length, color: '#f05c6e' },
            { name: 'Pending/Other', value: pending.length, color: '#f5a623' },
        ],
    };
}

export function filterRecords(records: PayoutRecord[], filters: {
    search?: string;
    status?: string;
    method?: string;
    month?: string;
    dateFrom?: string;
    dateTo?: string;
    minAmount?: string;
    maxAmount?: string;
    rejectionReason?: string;
}): PayoutRecord[] {
    return records.filter(r => {
        if (filters.search) {
            const q = filters.search.toLowerCase();
            if (
                !r.name.toLowerCase().includes(q) &&
                !r.email.toLowerCase().includes(q) &&
                !r.account.toLowerCase().includes(q) &&
                !r.comment.toLowerCase().includes(q)
            ) return false;
        }
        if (filters.status && filters.status !== 'all' && r.condition !== filters.status) return false;
        if (filters.method && filters.method !== 'all' && r.method !== filters.method) return false;
        if (filters.month && filters.month !== 'all' && r.sheet !== filters.month) return false;
        if (filters.dateFrom && r.date && r.date < filters.dateFrom) return false;
        if (filters.dateTo && r.date && r.date > filters.dateTo) return false;
        if (filters.minAmount) {
            const min = parseFloat(filters.minAmount);
            if (!isNaN(min) && r.submitted < min) return false;
        }
        if (filters.maxAmount) {
            const max = parseFloat(filters.maxAmount);
            if (!isNaN(max) && r.submitted > max) return false;
        }
        if (filters.rejectionReason && filters.rejectionReason !== 'all') {
            if (!r.comment.toLowerCase().includes(filters.rejectionReason.toLowerCase())) return false;
        }
        return true;
    });
}

export function exportToXlsx(records: PayoutRecord[], filename: string): void {
    const data = records.map(r => ({
        'Date': r.date,
        'Account': r.account,
        'Name': r.name,
        'Email': r.email,
        'Method': r.method,
        'Wallet': r.wallet,
        'Submitted ($)': r.submitted,
        'Deduction ($)': r.deduction || '',
        'Total ($)': r.total ?? '',
        'Status': r.condition,
        'Comment': r.comment,
        'Month': r.sheet,
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();

    // Column widths
    ws['!cols'] = [
        { wch: 12 }, { wch: 14 }, { wch: 26 }, { wch: 32 },
        { wch: 14 }, { wch: 46 }, { wch: 14 }, { wch: 14 },
        { wch: 14 }, { wch: 12 }, { wch: 50 }, { wch: 14 },
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Payouts Export');
    XLSX.writeFile(wb, filename);
}
