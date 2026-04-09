import { getGoogleAccessToken } from "./google-auth";
import {
  AppError,
  DEFAULT_HEADERS,
  SHEET_NAME,
  type DebugSheetsResult,
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
  const transactions: TransactionRecord[] = [];

  try {
    for (const row of snapshot.rows.filter((candidate) => candidate.some((cell) => normalizeCell(cell) !== ""))) {
      transactions.push(rowToTransaction(snapshot.headers, row));
    }
  } catch (error) {
    console.error("[sheets] Failed to parse transaction rows", error);
    throw new AppError(
      "parse",
      "Unable to parse rows from Google Sheets.",
      error instanceof Error ? error.message : String(error)
    );
  }

  return transactions;
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
    "sheet-write",
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
    "sheet-write",
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

  await sheetsRequest(env, ":batchUpdate", "sheet-delete", {
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

export async function debugSheets(env: Env): Promise<DebugSheetsResult> {
  const snapshot = await getSheetSnapshot(env);
  const transactions = await listTransactions(env);

  return {
    spreadsheetId: getSanitizedSpreadsheetId(env),
    sheetName: SHEET_NAME,
    expectedHeaders: DEFAULT_HEADERS,
    detectedHeaders: snapshot.headers,
    rowCount: snapshot.rows.filter((row) => row.some((cell) => normalizeCell(cell) !== "")).length,
    firstRecord: transactions[0] ?? null
  };
}

async function getSheetSnapshot(env: Env): Promise<SheetSnapshot> {
  const values = await getSheetValues(env);

  try {
    const rawHeaders = values[0]?.map((header) => normalizeCell(header)).filter(Boolean) ?? [];
    console.log("[sheets] Snapshot loaded", {
      spreadsheetId: env.SPREADSHEET_ID,
      sheetName: SHEET_NAME,
      range: READ_RANGE,
      rawRowCount: values.length,
      detectedHeaders: rawHeaders
    });

    if (rawHeaders.length > 0) {
      if (!rawHeaders.includes("dueDate")) {
        console.log("[sheets] dueDate header missing, migrating sheet", {
          spreadsheetId: env.SPREADSHEET_ID,
          sheetName: SHEET_NAME,
          detectedHeaders: rawHeaders
        });

        await appendHeader(env, rawHeaders.length + 1, "dueDate");
        rawHeaders.push("dueDate");
      }

      const missingHeaders = DEFAULT_HEADERS.filter((header) => !rawHeaders.includes(header));

      if (missingHeaders.length > 0) {
        console.error("[sheets] Required headers are missing", {
          spreadsheetId: env.SPREADSHEET_ID,
          sheetName: SHEET_NAME,
          detectedHeaders: rawHeaders,
          missingHeaders
        });
        throw new AppError(
          "parse",
          "The Google Sheet is missing required headers.",
          `Missing headers: ${missingHeaders.join(", ")}`
        );
      }

      return {
        headers: rawHeaders,
        rows: values.slice(1)
      };
    }

    console.log("[sheets] Sheet is empty, initializing headers", {
      spreadsheetId: env.SPREADSHEET_ID,
      sheetName: SHEET_NAME,
      headers: DEFAULT_HEADERS
    });

    await initializeHeaders(env);

    return {
      headers: [...DEFAULT_HEADERS],
      rows: []
    };
  } catch (error) {
    if (error instanceof AppError) {
      throw error;
    }

    console.error("[sheets] Failed to build sheet snapshot", error);
    throw new AppError(
      "parse",
      "Unable to build sheet snapshot from Google Sheets response.",
      error instanceof Error ? error.message : String(error)
    );
  }
}

async function initializeHeaders(env: Env): Promise<void> {
  const range = `${SHEET_NAME}!A1:${columnNumberToLetter(DEFAULT_HEADERS.length)}1`;

  await sheetsRequest(env, `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, "sheet-write", {
    method: "PUT",
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [DEFAULT_HEADERS]
    })
  });
}

async function appendHeader(env: Env, columnNumber: number, header: string): Promise<void> {
  const range = `${SHEET_NAME}!${columnNumberToLetter(columnNumber)}1`;

  await sheetsRequest(env, `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, "sheet-write", {
    method: "PUT",
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [[header]]
    })
  });
}

async function getSheetValues(env: Env): Promise<string[][]> {
  console.log("[sheets] Reading values", {
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName: SHEET_NAME,
    range: READ_RANGE
  });

  const response = await sheetsRequest(env, `/values/${encodeURIComponent(READ_RANGE)}?majorDimension=ROWS`, "sheet-read");

  try {
    const data = (await response.json()) as GoogleValuesResponse;
    return data.values ?? [];
  } catch (error) {
    console.error("[sheets] Failed to parse sheet values JSON", error);
    throw new AppError(
      "parse",
      "Unable to parse Google Sheets values response.",
      error instanceof Error ? error.message : String(error)
    );
  }
}

async function getSheetId(env: Env): Promise<number> {
  const response = await sheetsRequest(
    env,
    `?fields=${encodeURIComponent("sheets(properties(sheetId,title))")}`,
    "sheet-read"
  );
  let data: GoogleSpreadsheetResponse;

  try {
    data = (await response.json()) as GoogleSpreadsheetResponse;
  } catch (error) {
    throw new AppError(
      "parse",
      "Unable to parse Google Sheets metadata response.",
      error instanceof Error ? error.message : String(error)
    );
  }

  const matchingSheet = data.sheets?.find(
    (sheet) => sheet.properties?.title === SHEET_NAME && typeof sheet.properties.sheetId === "number"
  );

  if (typeof matchingSheet?.properties?.sheetId !== "number") {
    throw new AppError(
      "sheet-read",
      `Sheet "${SHEET_NAME}" was not found in the configured spreadsheet.`,
      `spreadsheetId=${env.SPREADSHEET_ID}`
    );
  }

  return matchingSheet.properties.sheetId;
}

async function sheetsRequest(
  env: Env,
  path: string,
  step: "sheet-read" | "sheet-write" | "sheet-delete" = "sheet-read",
  init?: RequestInit
): Promise<Response> {
  const accessToken = await getGoogleAccessToken(env);
  const spreadsheetId = getSanitizedSpreadsheetId(env);
  const url = `${GOOGLE_SHEETS_API}/${spreadsheetId}${path}`;

  console.log("[sheets] Requesting Google Sheets API", {
    step,
    method: init?.method ?? "GET",
    url
  });

  const response = await fetchSheetsWithRetry(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error("[sheets] Google Sheets API request failed", {
      step,
      status: response.status,
      body: errorText
    });
    throw new AppError(step, "Google Sheets API request failed.", `status=${response.status} body=${errorText}`);
  }

  return response;
}

async function fetchSheetsWithRetry(url: string, init: RequestInit): Promise<Response> {
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const response = await fetch(url, init);

    if (response.ok || !shouldRetrySheetsResponse(response.status) || attempt === maxAttempts) {
      return response;
    }

    response.body?.cancel().catch(() => undefined);
    await sleep(250 * attempt);
  }

  return fetch(url, init);
}

function shouldRetrySheetsResponse(status: number): boolean {
  return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
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
  if (!headers.length) {
    throw new Error("No headers detected in the Google Sheet.");
  }

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

function getSanitizedSpreadsheetId(env: Env): string {
  return env.SPREADSHEET_ID.trim().replace(/\/+$/, "");
}
