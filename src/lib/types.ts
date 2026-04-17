export type PayoutStatus = 'Paid' | 'Rejected' | 'Pending' | 'Approved' | 'Submitted';

export interface PayoutRecord {
    id: string;
    date: string; // ISO date string
    account: string;
    name: string;
    email: string;
    method: string;
    wallet: string;
    submitted: number;
    deduction: number;
    total: number | null;
    condition: PayoutStatus;
    backOffice: boolean;
    comment: string;
    month: string; // e.g. "NOVEMBER"
    sheet: string;
}

export interface DashboardStats {
    totalRequests: number;
    totalPaid: number;
    totalRejected: number;
    totalPending: number;
    totalSubmittedAmount: number;
    totalPaidAmount: number;
    totalDeductions: number;
    approvalRate: number;
    rejectionRate: number;
    avgPaidAmount: number;
    uniqueTraders: number;
    methodBreakdown: Record<string, number>;
    rejectionReasons: Record<string, number>;
    monthlyData: MonthlyData[];
    dailyData: DailyData[];
    statusDistribution: StatusItem[];
}

export interface MonthlyData {
    month: string;
    requests: number;
    paid: number;
    rejected: number;
    totalPaidAmount: number;
    totalSubmittedAmount: number;
}

export interface DailyData {
    date: string;
    requests: number;
    paid: number;
    rejected: number;
    amount: number;
}

export interface StatusItem {
    name: string;
    value: number;
    color: string;
}

export interface FilterState {
    search: string;
    status: string;
    method: string;
    month: string;
    dateFrom: string;
    dateTo: string;
    minAmount: string;
    maxAmount: string;
    rejectionReason: string;
}
