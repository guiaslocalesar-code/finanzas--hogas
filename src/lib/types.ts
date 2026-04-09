export const SHEET_NAME = "Transacciones";

export const DEFAULT_HEADERS = [
  "id",
  "type",
  "amount",
  "category",
  "description",
  "date",
  "createdAt",
  "dueDate"
] as const;

export const CARD_SHEET_NAME = "Tarjetas";
export const CARD_HEADERS = [
  "cardId",
  "issuer",
  "brand",
  "bank",
  "holder",
  "last4",
  "closeDay",
  "dueDay",
  "active",
  "createdAt"
] as const;

export const CARD_SUMMARY_SHEET_NAME = "ResumenesTarjeta";
export const CARD_SUMMARY_HEADERS = [
  "summaryId",
  "cardId",
  "issuer",
  "brand",
  "bank",
  "holder",
  "fileName",
  "sourceType",
  "statementDate",
  "closingDate",
  "dueDate",
  "nextDueDate",
  "totalAmount",
  "minimumPayment",
  "currency",
  "rawText",
  "rawDetectedData",
  "warnings",
  "parseStatus",
  "createdAt"
] as const;

export const INSTALLMENT_PROJECTION_SHEET_NAME = "CuotasProyectadas";
export const INSTALLMENT_PROJECTION_HEADERS = [
  "projectionId",
  "summaryId",
  "cardId",
  "issuer",
  "monthLabel",
  "yearMonth",
  "amount",
  "sourceType",
  "confirmed",
  "createdAt"
] as const;

export const INSTALLMENT_DETAIL_SHEET_NAME = "CuotasDetalle";
export const INSTALLMENT_DETAIL_HEADERS = [
  "installmentId",
  "summaryId",
  "cardId",
  "purchaseDate",
  "merchant",
  "installmentNumber",
  "installmentTotal",
  "amount",
  "dueMonth",
  "dueDate",
  "ownerType",
  "businessPercent",
  "businessAmount",
  "personalAmount",
  "reimbursementStatus",
  "notes",
  "createdAt"
] as const;

export const BUSINESS_REIMBURSEMENT_SHEET_NAME = "ReintegrosNegocio";
export const BUSINESS_REIMBURSEMENT_HEADERS = [
  "reimbursementId",
  "sourceType",
  "sourceId",
  "cardId",
  "concept",
  "totalPaid",
  "businessAmount",
  "personalAmount",
  "reimbursementStatus",
  "reimbursementDueDate",
  "reimbursedAmount",
  "reimbursedDate",
  "notes",
  "createdAt"
] as const;

export type TransactionType = "income" | "expense";
export type OwnerType = "personal" | "business" | "mixed";
export type ReimbursementStatus = "pending" | "partial" | "paid";
export type AppErrorStep =
  | "env"
  | "jwt"
  | "google-token"
  | "form-data"
  | "validate"
  | "sheet-read"
  | "sheet-write"
  | "sheet-delete"
  | "parse-pdf"
  | "detect-issuer"
  | "parse"
  | "setup";

export interface AssetFetcher {
  fetch(input: Request | string | URL, init?: RequestInit): Promise<Response>;
}

export interface Env {
  SPREADSHEET_ID: string;
  SERVICE_ACCOUNT_EMAIL: string;
  PRIVATE_KEY: string;
  ASSETS: AssetFetcher;
}

export interface Transaction {
  id: string;
  type: TransactionType;
  amount: number;
  category: string;
  description: string;
  date: string;
  createdAt: string;
  dueDate: string;
}

export type TransactionRecord = Transaction & Record<string, string | number>;

export interface CreateTransactionInput {
  type: TransactionType;
  amount: number;
  category: string;
  description: string;
  date: string;
  dueDate: string;
}

export interface UpdateTransactionInput {
  id: string;
  type?: TransactionType;
  amount?: number;
  category?: string;
  description?: string;
  date?: string;
  dueDate?: string;
}

export interface CardRecord extends Record<string, string | number | boolean> {
  cardId: string;
  issuer: string;
  brand: string;
  bank: string;
  holder: string;
  last4: string;
  closeDay: number;
  dueDay: number;
  active: boolean;
  createdAt: string;
}

export interface CardSummaryRecord extends Record<string, string | number | boolean> {
  summaryId: string;
  cardId: string;
  issuer: string;
  brand: string;
  bank: string;
  holder: string;
  fileName: string;
  sourceType: string;
  statementDate: string;
  closingDate: string;
  dueDate: string;
  nextDueDate: string;
  totalAmount: number;
  minimumPayment: number;
  currency: string;
  rawText: string;
  rawDetectedData: string;
  warnings: string;
  parseStatus: string;
  createdAt: string;
}

export interface InstallmentProjectionRecord extends Record<string, string | number | boolean> {
  projectionId: string;
  summaryId: string;
  cardId: string;
  issuer: string;
  monthLabel: string;
  yearMonth: string;
  amount: number;
  sourceType: string;
  confirmed: boolean;
  createdAt: string;
}

export interface InstallmentDetailRecord extends Record<string, string | number | boolean> {
  installmentId: string;
  summaryId: string;
  cardId: string;
  purchaseDate: string;
  merchant: string;
  installmentNumber: number;
  installmentTotal: number;
  amount: number;
  dueMonth: string;
  dueDate: string;
  ownerType: OwnerType;
  businessPercent: number;
  businessAmount: number;
  personalAmount: number;
  reimbursementStatus: ReimbursementStatus;
  notes: string;
  createdAt: string;
}

export interface BusinessReimbursementRecord extends Record<string, string | number | boolean> {
  reimbursementId: string;
  sourceType: string;
  sourceId: string;
  cardId: string;
  concept: string;
  totalPaid: number;
  businessAmount: number;
  personalAmount: number;
  reimbursementStatus: ReimbursementStatus;
  reimbursementDueDate: string;
  reimbursedAmount: number;
  reimbursedDate: string;
  notes: string;
  createdAt: string;
}

export interface ModuleSheetSetupStatus {
  sheetName: string;
  created: boolean;
  addedHeaders: string[];
  headers: string[];
}

export interface FinanceModuleSetupResult {
  ok: true;
  sheets: ModuleSheetSetupStatus[];
}

export interface CardSummaryImportResult {
  summary: CardSummaryRecord;
  installments: InstallmentDetailRecord[];
  projections: InstallmentProjectionRecord[];
}

export interface InstallmentOutlook {
  currentMonth: {
    yearMonth: string;
    projections: InstallmentProjectionRecord[];
    details: InstallmentDetailRecord[];
  };
  nextMonth: {
    yearMonth: string;
    projections: InstallmentProjectionRecord[];
    details: InstallmentDetailRecord[];
  };
  followingMonth: {
    yearMonth: string;
    projections: InstallmentProjectionRecord[];
    details: InstallmentDetailRecord[];
  };
}

export interface InstallmentFilters {
  cardId?: string;
  ownerType?: OwnerType;
  reimbursementStatus?: ReimbursementStatus;
}

export interface InstallmentForecastBucket {
  yearMonth: string;
  totalAmount: number;
  items: InstallmentDetailRecord[];
}

export interface InstallmentForecastResponse {
  thisMonth: InstallmentForecastBucket;
  nextMonth: InstallmentForecastBucket;
  thirdMonth: InstallmentForecastBucket;
  totalPending: number;
  businessPending: number;
  personalPending: number;
  filters: InstallmentFilters;
}

export interface CardsDashboardResponse {
  activeCards: number;
  inactiveCards: number;
  statementCount: number;
  statementsTotal: number;
  minimumPaymentsTotal: number;
  pendingInstallmentsThisMonth: number;
  pendingInstallmentsNextMonth: number;
  pendingInstallmentsThirdMonth: number;
  businessPending: number;
  personalPending: number;
  reimbursementsPending: number;
  reimbursementsPartial: number;
  reimbursementsPaid: number;
  reimbursableAmountPending: number;
}

export interface CardStatementProjectionPreview {
  monthLabel: string;
  yearMonth: string;
  amount: number;
}

export interface ParsedCardStatementPreview {
  issuer: string;
  brand: string;
  bank: string;
  holder: string;
  closingDate: string;
  dueDate: string;
  nextDueDate: string;
  totalAmount: number;
  minimumPayment: number;
  projections: CardStatementProjectionPreview[];
  rawDetectedData: Record<string, unknown>;
  warnings: string[];
}

export interface UploadedCardStatementResult {
  ok: true;
  fileName: string;
  preview: ParsedCardStatementPreview;
  summary: CardSummaryRecord | null;
  projections: InstallmentProjectionRecord[];
}

export interface CardStatementUploadDebugResult {
  ok: true;
  fileName: string;
  mimeType: string;
  size: number;
  textExtractedLength: number;
  detectedIssuer: string;
  parseWarnings: string[];
  saveAttempted: boolean;
  preview: ParsedCardStatementPreview | null;
}

export interface DebugCardStatementResult {
  ok: true;
  summary: CardSummaryRecord;
  projections: InstallmentProjectionRecord[];
  installments: InstallmentDetailRecord[];
  parsed: ParsedCardStatementPreview;
  parserDebug: Record<string, unknown>;
}

export interface CardStatementDetailResult extends DebugCardStatementResult {}

export interface GoogleTokenResponse {
  access_token: string;
  expires_in: number;
  token_type: string;
}

export interface GoogleValuesResponse {
  range?: string;
  majorDimension?: string;
  values?: string[][];
}

export interface GoogleSpreadsheetResponse {
  sheets?: Array<{
    properties?: {
      sheetId?: number;
      title?: string;
    };
  }>;
}

export interface SheetSnapshot {
  headers: string[];
  rows: string[][];
}

export interface DebugSheetsResult {
  spreadsheetId: string;
  sheetName: string;
  expectedHeaders: readonly string[];
  detectedHeaders: string[];
  rowCount: number;
  firstRecord: TransactionRecord | null;
}

export class AppError extends Error {
  readonly step: AppErrorStep;
  readonly details?: string;
  readonly status: number;

  constructor(step: AppErrorStep, message: string, details?: string, status = 500) {
    super(message);
    this.name = "AppError";
    this.step = step;
    this.details = details;
    this.status = status;
  }
}
