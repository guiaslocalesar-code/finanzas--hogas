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

export type TransactionType = "income" | "expense";
export type AppErrorStep =
  | "env"
  | "jwt"
  | "google-token"
  | "sheet-read"
  | "sheet-write"
  | "sheet-delete"
  | "parse";

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
