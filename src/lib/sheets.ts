import { getGoogleAccessToken } from "./google-auth";
import {
  DEFAULT_HEADERS,
  SHEET_NAME,
  type CreateTransactionInput,
  type Env,
  type GoogleSpreadsheetResponse,
  type GoogleValuesResponse,
  type SheetSnapshot,
  type TransactionRecord,
  type UpdateTransactionInput
} from "./types";

const GOOGLE_SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
const READ_RANGE = `${SHEET_NAME}!A:Z`;

export async function listTransactions(env: Env): Promise<TransactionRecord[]> {
  const snapshot = await getSheetSnapshot(env);

  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToTransaction(snapshot.headers, row));
}

export async function createTransaction(
  env: Env,
  transaction: CreateTransactionInput & { id: string; createdAt: string }
): Promise<TransactionRecord> {
  const snapshot = await getSheetSnapshot(env);
  const rowValues = buildRowValues(snapshot.headers, transaction);

  await sheetsRequest(
    env,
    `/values/${encodeURIComponent(`${SHEET_NAME}!A1:Z1`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      body: JSON.stringify({
        majorDimension: "ROWS",
        values: [rowValues]
      })
    }
  );

  return rowToTransaction(snapshot.headers, rowValues);
}

export async function updateTransaction(
  env: Env,
  patch: UpdateTransactionInput
): Promise<TransactionRecord | null> {
  const snapshot = await getSheetSnapshot(env);
  const match = findRowById(snapshot, patch.id);

  if (!match) {
    return null;
  }

  const currentRecord = rowToObject(snapshot.headers, match.row);
  const mergedRecord = {
    ...currentRecord,
    ...patch
  };
  const updatedRowValues = buildRowValues(snapshot.headers, mergedRecord);
  const endColumn = columnNumberToLetter(snapshot.headers.length);
  const rowNumber = match.index + 2;

  await sheetsRequest(
    env,
    `/values/${encodeURIComponent(`${SHEET_NAME}!A${rowNumber}:${endColumn}${rowNumber}`)}?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      body: JSON.stringify({
        majorDimension: "ROWS",
        values: [updatedRowValues]
      })
    }
  );

  return rowToTransaction(snapshot.headers, updatedRowValues);
}

export async function deleteTransaction(env: Env, id: string): Promise<boolean> {
  const snapshot = await getSheetSnapshot(env);
  const match = findRowById(snapshot, id);

  if (!match) {
    return false;
  }

  const sheetId = await getSheetId(env);
  const startIndex = match.index + 1;

  await sheetsRequest(env, ":batchUpdate", {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId,
              dimension: "ROWS",
              startIndex,
              endIndex: startIndex + 1
            }
          }
        }
      ]
    })
  });

  return true;
}

async function getSheetSnapshot(env: Env): Promise<SheetSnapshot> {
  const values = await getSheetValues(env);
  const rawHeaders = values[0]?.map((header) => header.trim()).filter(Boolean) ?? [];

  if (rawHeaders.length > 0) {
    return {
      headers: rawHeaders,
      rows: values.slice(1)
    };
  }

  await initializeHeaders(env);

  return {
    headers: [...DEFAULT_HEADERS],
    rows: []
  };
}

async function initializeHeaders(env: Env): Promise<void> {
  const range = `${SHEET_NAME}!A1:${columnNumberToLetter(DEFAULT_HEADERS.length)}1`;

  await sheetsRequest(env, `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, {
    method: "PUT",
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [DEFAULT_HEADERS]
    })
  });
}

async function getSheetValues(env: Env): Promise<string[][]> {
  const response = await sheetsRequest(env, `/values/${encodeURIComponent(READ_RANGE)}?majorDimension=ROWS`);
  const data = (await response.json()) as GoogleValuesResponse;
  return data.values ?? [];
}

async function getSheetId(env: Env): Promise<number> {
  const response = await sheetsRequest(
    env,
    `?fields=${encodeURIComponent("sheets(properties(sheetId,title))")}`
  );
  const data = (await response.json()) as GoogleSpreadsheetResponse;
  const matchingSheet = data.sheets?.find(
    (sheet) => sheet.properties?.title === SHEET_NAME && typeof sheet.properties.sheetId === "number"
  );

  if (typeof matchingSheet?.properties?.sheetId !== "number") {
    throw new Error(`Sheet "${SHEET_NAME}" was not found in spreadsheet ${env.SPREADSHEET_ID}`);
  }

  return matchingSheet.properties.sheetId;
}

async function sheetsRequest(env: Env, path: string, init?: RequestInit): Promise<Response> {
  const accessToken = await getGoogleAccessToken(env);
  const response = await fetch(`${GOOGLE_SHEETS_API}/${env.SPREADSHEET_ID}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Google Sheets API request failed: ${response.status} ${errorText}`);
  }

  return response;
}

function findRowById(snapshot: SheetSnapshot, id: string): { row: string[]; index: number } | null {
  const idColumnIndex = snapshot.headers.indexOf("id");

  if (idColumnIndex === -1) {
    throw new Error('The sheet is missing the "id" header.');
  }

  for (let index = 0; index < snapshot.rows.length; index += 1) {
    const row = snapshot.rows[index] ?? [];
    if (normalizeCell(row[idColumnIndex]) === id) {
      return { row, index };
    }
  }

  return null;
}

function rowToTransaction(headers: string[], row: string[]): TransactionRecord {
  const record = rowToObject(headers, row);

  return {
    ...record,
    id: stringValue(record.id),
    type: stringValue(record.type) === "income" ? "income" : "expense",
    amount: toNumber(record.amount),
    category: stringValue(record.category),
    description: stringValue(record.description),
    date: stringValue(record.date),
    createdAt: stringValue(record.createdAt),
    dueDate: stringValue(record.dueDate)
  };
}

function rowToObject(headers: string[], row: string[]): Record<string, string> {
  return headers.reduce<Record<string, string>>((result, header, index) => {
    result[header] = normalizeCell(row[index]);
    return result;
  }, {});
}

function buildRowValues<T extends object>(headers: string[], record: T): string[] {
  return headers.map((header) => stringifyCellValue(record[header as keyof T]));
}

function stringifyCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value.toString() : "";
  }

  return String(value).trim();
}

function normalizeCell(value: string | undefined): string {
  return typeof value === "string" ? value.trim() : "";
}

function stringValue(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function toNumber(value: unknown): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function columnNumberToLetter(columnNumber: number): string {
  let current = columnNumber;
  let letter = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    current = Math.floor((current - remainder) / 26);
  }

  return letter || "A";
}
