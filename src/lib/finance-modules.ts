import { getGoogleAccessToken } from "./google-auth";
import { parseNaranjaWithGemini } from "./gemini";
import { detectIssuerFromText, extractPdfText, parseDetectedStatement } from "./pdf-statements";
import {
  AppError,
  BUSINESS_REIMBURSEMENT_HEADERS,
  BUSINESS_REIMBURSEMENT_SHEET_NAME,
  CARD_HEADERS,
  CARD_SHEET_NAME,
  CARD_SUMMARY_HEADERS,
  CARD_SUMMARY_SHEET_NAME,
  INSTALLMENT_DETAIL_HEADERS,
  INSTALLMENT_DETAIL_SHEET_NAME,
  INSTALLMENT_PROJECTION_HEADERS,
  INSTALLMENT_PROJECTION_SHEET_NAME,
  type AppErrorStep,
  type BusinessReimbursementRecord,
  type CardStatementUploadDebugResult,
  type CardStatementDetailResult,
  type CardsDashboardResponse,
  type CardRecord,
  type CardSummaryImportResult,
  type CardSummaryRecord,
  type DebugCardStatementResult,
  type Env,
  type FinanceModuleSetupResult,
  type GoogleSpreadsheetResponse,
  type GoogleValuesResponse,
  type InstallmentDetailRecord,
  type InstallmentFilters,
  type InstallmentForecastResponse,
  type InstallmentOutlook,
  type InstallmentProjectionRecord,
  type ModuleSheetSetupStatus,
  type OwnerType,
  type ParsedCardStatementPreview,
  type ReimbursementStatus,
  type SheetSnapshot,
  type UploadedCardStatementResult
} from "./types";

const GOOGLE_SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";
const DEFAULT_SHEET_RANGE_SUFFIX = "!A:Z";

type ModuleSheetConfig = {
  sheetName: string;
  headers: readonly string[];
};

type ParsedSummaryImport = {
  statementDate: string;
  closingDate: string;
  dueDate: string;
  totalAmount: number;
  minimumPayment: number;
  currency: string;
  parseStatus: string;
  installments: Array<{
    merchant: string;
    installmentNumber: number;
    installmentTotal: number;
    amount: number;
    purchaseDate: string;
  }>;
};

const MODULE_SHEETS: ModuleSheetConfig[] = [
  { sheetName: CARD_SHEET_NAME, headers: CARD_HEADERS },
  { sheetName: CARD_SUMMARY_SHEET_NAME, headers: CARD_SUMMARY_HEADERS },
  { sheetName: INSTALLMENT_PROJECTION_SHEET_NAME, headers: INSTALLMENT_PROJECTION_HEADERS },
  { sheetName: INSTALLMENT_DETAIL_SHEET_NAME, headers: INSTALLMENT_DETAIL_HEADERS },
  { sheetName: BUSINESS_REIMBURSEMENT_SHEET_NAME, headers: BUSINESS_REIMBURSEMENT_HEADERS }
];

export async function initializeFinanceModuleSheets(env: Env): Promise<FinanceModuleSetupResult> {
  const statuses: ModuleSheetSetupStatus[] = [];

  for (const config of MODULE_SHEETS) {
    statuses.push(await ensureModuleSheet(env, config));
  }

  return {
    ok: true,
    sheets: statuses
  };
}

export async function listCards(env: Env): Promise<CardRecord[]> {
  const snapshot = await readModuleSheet(env, CARD_SHEET_NAME, CARD_HEADERS);
  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToCard(snapshot.headers, row))
    .sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function getCard(env: Env, cardId: string): Promise<CardRecord | null> {
  return getCardById(env, cardId);
}

export async function createCard(env: Env, input: Record<string, unknown>): Promise<CardRecord> {
  const snapshot = await readModuleSheet(env, CARD_SHEET_NAME, CARD_HEADERS);
  const record: CardRecord = {
    cardId: generateId("card"),
    issuer: stringValue(input.issuer),
    brand: stringValue(input.brand),
    bank: stringValue(input.bank),
    holder: stringValue(input.holder),
    last4: stringValue(input.last4).slice(-4),
    closeDay: boundedDay(input.closeDay),
    dueDay: boundedDay(input.dueDay),
    active: booleanValue(input.active, true),
    createdAt: new Date().toISOString()
  };

  await appendRow(env, CARD_SHEET_NAME, snapshot.headers, record);
  return record;
}

export async function updateCard(env: Env, input: Record<string, unknown>): Promise<CardRecord | null> {
  const snapshot = await readModuleSheet(env, CARD_SHEET_NAME, CARD_HEADERS);
  const cardId = stringValue(input.cardId);
  const match = findRowByKey(snapshot, "cardId", cardId);

  if (!match) {
    return null;
  }

  const current = rowToCard(snapshot.headers, match.row);
  const updated: CardRecord = {
    ...current,
    issuer: "issuer" in input ? stringValue(input.issuer) : current.issuer,
    brand: "brand" in input ? stringValue(input.brand) : current.brand,
    bank: "bank" in input ? stringValue(input.bank) : current.bank,
    holder: "holder" in input ? stringValue(input.holder) : current.holder,
    last4: "last4" in input ? stringValue(input.last4).slice(-4) : current.last4,
    closeDay: "closeDay" in input ? boundedDay(input.closeDay) : current.closeDay,
    dueDay: "dueDay" in input ? boundedDay(input.dueDay) : current.dueDay,
    active: "active" in input ? booleanValue(input.active, current.active) : current.active
  };

  await updateRow(env, CARD_SHEET_NAME, snapshot.headers, match.index, updated);
  return updated;
}

export async function deleteCard(env: Env, cardId: string): Promise<boolean> {
  const snapshot = await readModuleSheet(env, CARD_SHEET_NAME, CARD_HEADERS);
  const match = findRowByKey(snapshot, "cardId", cardId);

  if (!match) {
    return false;
  }

  await deleteSheetRow(env, CARD_SHEET_NAME, match.index);
  return true;
}

export async function listCardSummaries(env: Env): Promise<CardSummaryRecord[]> {
  const snapshot = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToCardSummary(snapshot.headers, row))
    .sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function getCardStatement(env: Env, summaryId: string): Promise<CardSummaryRecord | null> {
  const snapshot = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  const match = findRowByKey(snapshot, "summaryId", summaryId);
  return match ? rowToCardSummary(snapshot.headers, match.row) : null;
}

export async function createCardSummary(env: Env, input: Record<string, unknown>): Promise<CardSummaryRecord> {
  const snapshot = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  const card = stringValue(input.cardId) ? await getCardById(env, stringValue(input.cardId)) : null;
  const closingDate = normalizeIsoDateValue(stringValue(input.closingDate));
  const dueDate = normalizeIsoDateValue(stringValue(input.dueDate)) || guessDueDateFromCard(card, closingDate);
  const nextDueDate =
    normalizeIsoDateValue(stringValue(input.nextDueDate)) || addMonthsToDateString(dueDate, 1);

  const record: CardSummaryRecord = {
    summaryId: generateId("summary"),
    cardId: stringValue(input.cardId),
    issuer: stringValue(input.issuer) || card?.issuer || "",
    brand: stringValue(input.brand) || card?.brand || "",
    bank: stringValue(input.bank) || card?.bank || "",
    holder: stringValue(input.holder) || card?.holder || "",
    fileName: stringValue(input.fileName),
    sourceType: stringValue(input.sourceType) || "manual",
    statementDate: normalizeIsoDateValue(stringValue(input.statementDate)),
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount: numberValue(input.totalAmount),
    minimumPayment: numberValue(input.minimumPayment),
    currency: stringValue(input.currency) || "ARS",
    rawText: stringValue(input.rawText),
    rawDetectedData: stringifyJsonField(input.rawDetectedData),
    warnings: stringifyJsonField(input.warnings),
    parseStatus: stringValue(input.parseStatus) || "manual",
    createdAt: new Date().toISOString()
  };

  await appendRow(env, CARD_SUMMARY_SHEET_NAME, snapshot.headers, record);
  return record;
}

export async function updateCardStatement(
  env: Env,
  summaryId: string,
  input: Record<string, unknown>
): Promise<CardSummaryRecord | null> {
  const snapshot = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  const match = findRowByKey(snapshot, "summaryId", summaryId);

  if (!match) {
    return null;
  }

  const current = rowToCardSummary(snapshot.headers, match.row);
  const card = current.cardId ? await getCardById(env, current.cardId) : null;
  const closingDate = "closingDate" in input
    ? normalizeIsoDateValue(stringValue(input.closingDate))
    : current.closingDate;
  const dueDate =
    "dueDate" in input
      ? normalizeIsoDateValue(stringValue(input.dueDate))
      : current.dueDate || guessDueDateFromCard(card, closingDate);

  const updated: CardSummaryRecord = {
    ...current,
    cardId: "cardId" in input ? stringValue(input.cardId) : current.cardId,
    issuer: "issuer" in input ? stringValue(input.issuer) : current.issuer,
    brand: "brand" in input ? stringValue(input.brand) : current.brand,
    bank: "bank" in input ? stringValue(input.bank) : current.bank,
    holder: "holder" in input ? stringValue(input.holder) : current.holder,
    fileName: "fileName" in input ? stringValue(input.fileName) : current.fileName,
    sourceType: "sourceType" in input ? stringValue(input.sourceType) : current.sourceType,
    statementDate:
      "statementDate" in input ? normalizeIsoDateValue(stringValue(input.statementDate)) : current.statementDate,
    closingDate,
    dueDate,
    nextDueDate:
      "nextDueDate" in input
        ? normalizeIsoDateValue(stringValue(input.nextDueDate))
        : current.nextDueDate || addMonthsToDateString(dueDate, 1),
    totalAmount: "totalAmount" in input ? numberValue(input.totalAmount) : current.totalAmount,
    minimumPayment: "minimumPayment" in input ? numberValue(input.minimumPayment) : current.minimumPayment,
    currency: "currency" in input ? stringValue(input.currency) : current.currency,
    rawText: "rawText" in input ? stringValue(input.rawText) : current.rawText,
    rawDetectedData: "rawDetectedData" in input ? stringifyJsonField(input.rawDetectedData) : current.rawDetectedData,
    warnings: "warnings" in input ? stringifyJsonField(input.warnings) : current.warnings,
    parseStatus: "parseStatus" in input ? stringValue(input.parseStatus) : current.parseStatus
  };

  await updateRow(env, CARD_SUMMARY_SHEET_NAME, snapshot.headers, match.index, updated);
  return updated;
}

export async function uploadCardStatementPdf(
  env: Env,
  input: {
    fileName: string;
    pdfBytes: ArrayBuffer;
    cardId?: string;
    previewOnly?: boolean;
  }
): Promise<UploadedCardStatementResult> {
  const fileName = input.fileName.trim();
  const cardId = input.cardId?.trim() ?? "";
  logUploadStage("start", {
    fileName,
    cardId,
    previewOnly: Boolean(input.previewOnly),
    bytes: input.pdfBytes.byteLength
  });

  if (!fileName) {
    throw new AppError("validate", 'Field "fileName" is required.', undefined, 400);
  }

  if (input.pdfBytes.byteLength === 0) {
    throw new AppError("validate", "The uploaded PDF is empty.", undefined, 400);
  }

  const setupResult = await initializeFinanceModuleSheets(env);
  logUploadStage("setup-complete", {
    sheets: setupResult.sheets.map((sheet) => ({
      sheetName: sheet.sheetName,
      created: sheet.created,
      headers: sheet.headers.length
    }))
  });

  const analysis = await analyzeUploadedCardStatement(env, fileName, input.pdfBytes);

  if (input.previewOnly) {
    logUploadStage("preview-return", {
      fileName,
      detectedIssuer: analysis.detectedIssuer,
      warnings: analysis.preview.warnings.length,
      textExtractedLength: analysis.textExtractedLength
    });
    return {
      ok: true,
      fileName,
      preview: analysis.preview,
      summary: null,
      projections: [],
      installmentsDetail: []
    };
  }

  const card = cardId ? await getCardById(env, cardId) : null;
  if (cardId && !card) {
    throw new AppError("validate", `Card "${cardId}" was not found.`, undefined, 400);
  }

  const dueDate =
    normalizeIsoDateValue(analysis.preview.dueDate) ||
    guessDueDateFromCard(card, normalizeIsoDateValue(analysis.preview.closingDate));
  const nextDueDate =
    normalizeIsoDateValue(analysis.preview.nextDueDate) || addMonthsToDateString(dueDate, 1);

  if (!analysis.canSave) {
    throw new AppError(
      "parse-pdf",
      "No se pudo extraer suficiente información del PDF para guardarlo.",
      JSON.stringify({
        warnings: analysis.preview.warnings,
        detectedIssuer: analysis.detectedIssuer,
        textExtractedLength: analysis.textExtractedLength
      }),
      400
    );
  }

  logUploadStage("save-summary-start", {
    fileName,
    cardId,
    dueDate,
    nextDueDate,
    detectedIssuer: analysis.detectedIssuer,
    warnings: analysis.preview.warnings.length
  });

  const summarySheet = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  const projectionSheet = await readModuleSheet(
    env,
    INSTALLMENT_PROJECTION_SHEET_NAME,
    INSTALLMENT_PROJECTION_HEADERS
  );
  const installmentSheet = await readModuleSheet(env, INSTALLMENT_DETAIL_SHEET_NAME, INSTALLMENT_DETAIL_HEADERS);

  const summary = await appendCardSummaryFromUpload(env, summarySheet.headers, {
    cardId,
    card,
    fileName,
    rawText: analysis.rawText,
    preview: analysis.preview,
    dueDate,
    nextDueDate
  });

  const projections = buildProjectionRecordsFromPreview(summary, analysis.preview);
  const historicalInstallments = installmentSheet.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToInstallmentDetail(installmentSheet.headers, row));
  const installmentsDetail = inheritInstallmentClassifications(
    buildInstallmentDetailRecordsFromPreview(summary, analysis.preview),
    historicalInstallments
  );

  if (projections.length > 0) {
    logUploadStage("save-projections-start", {
      summaryId: summary.summaryId,
      count: projections.length
    });
    await appendRows(env, INSTALLMENT_PROJECTION_SHEET_NAME, projectionSheet.headers, projections);
  }

  if (installmentsDetail.length > 0) {
    logUploadStage("save-installments-detail-start", {
      summaryId: summary.summaryId,
      count: installmentsDetail.length
    });
    await appendRows(env, INSTALLMENT_DETAIL_SHEET_NAME, installmentSheet.headers, installmentsDetail);
    await Promise.all(installmentsDetail.map((installment) => syncInstallmentReimbursement(env, installment)));
  }

  logUploadStage("save-complete", {
    fileName,
    summaryId: summary.summaryId,
    projections: projections.length,
    installmentsDetail: installmentsDetail.length
  });

  return {
    ok: true,
    fileName,
    preview: analysis.preview,
    summary,
    projections,
    installmentsDetail
  };
}

export async function debugUploadCardStatementPdf(
  env: Env,
  input: {
    fileName: string;
    mimeType: string;
    pdfBytes: ArrayBuffer;
    cardId?: string;
  }
): Promise<CardStatementUploadDebugResult> {
  const setupResult = await initializeFinanceModuleSheets(env);
  logUploadStage("debug-setup-complete", {
    fileName: input.fileName,
    sheetsReady: setupResult.sheets.length
  });

  const analysis = await analyzeUploadedCardStatement(env, input.fileName, input.pdfBytes);

  return {
    ok: true,
    fileName: input.fileName.trim(),
    mimeType: input.mimeType.trim(),
    size: input.pdfBytes.byteLength,
    textExtractedLength: analysis.textExtractedLength,
    detectedIssuer: analysis.detectedIssuer,
    parseWarnings: analysis.preview.warnings,
    saveAttempted: false,
    preview: analysis.preview
  };
}

export async function getCardStatementDetail(
  env: Env,
  summaryId: string
): Promise<CardStatementDetailResult | null> {
  const normalizedSummaryId = summaryId.trim();
  const snapshot = await readModuleSheet(env, CARD_SUMMARY_SHEET_NAME, CARD_SUMMARY_HEADERS);
  const match = findRowByKey(snapshot, "summaryId", normalizedSummaryId);
  console.log(
    "[card-statements/detail]",
    JSON.stringify({
      stage: "lookup-summary",
      summaryId: normalizedSummaryId,
      found: Boolean(match),
      rowIndex: match ? match.index + 2 : null
    })
  );

  if (!match) {
    return null;
  }

  const summary = rowToCardSummary(snapshot.headers, match.row);
  const projections = (await listInstallmentProjections(env)).filter((item) => item.summaryId === summary.summaryId);
  const installments = (await listInstallments(env)).filter((item) => item.summaryId === summary.summaryId);
  const storedParserDebug = parseStoredRecord(summary.rawDetectedData);
  const reparsed =
    summary.rawText && (summary.totalAmount <= 0 || summary.minimumPayment <= 0 || !summary.nextDueDate || projections.length === 0)
      ? parseDetectedStatement(summary.rawText, summary.fileName)
      : null;
  const parserDebug = reparsed
    ? {
        ...storedParserDebug,
        reparsedFromRawText: true,
        reparsedPreview: reparsed.rawDetectedData
      }
    : storedParserDebug;
  const parsedWarnings = Array.from(
    new Set([...parseStoredStringArray(summary.warnings), ...(reparsed?.warnings ?? [])].filter(Boolean))
  );
  const parsedProjections =
    projections.length > 0
      ? projections.map((item) => ({
          monthLabel: item.monthLabel,
          yearMonth: item.yearMonth,
          amount: item.amount
        }))
      : reparsed?.projections ?? [];
  const installmentsDetail =
    installments.length > 0
      ? installments
      : reparsed
        ? buildInstallmentDetailRecordsFromPreview(summary, reparsed)
        : [];
  const parsed: ParsedCardStatementPreview = {
    issuer: reparsed?.issuer || summary.issuer,
    brand: reparsed?.brand || summary.brand,
    bank: reparsed?.bank || summary.bank,
    holder: reparsed?.holder || summary.holder,
    closingDate: summary.closingDate,
    dueDate: summary.dueDate,
    nextDueDate: summary.nextDueDate || reparsed?.nextDueDate || "",
    totalAmount: summary.totalAmount > 0 ? summary.totalAmount : reparsed?.totalAmount ?? 0,
    minimumPayment: summary.minimumPayment > 0 ? summary.minimumPayment : reparsed?.minimumPayment ?? 0,
    projections: parsedProjections,
    installmentsDetail: reparsed?.installmentsDetail ?? [],
    rawDetectedData: parserDebug,
    warnings: parsedWarnings
  };

  console.log(
    "[card-statements/detail]",
    JSON.stringify({
      stage: "detail-built",
      summaryId: summary.summaryId,
      projections: projections.length,
      installmentsDetail: installmentsDetail.length,
      reparsedFromRawText: Boolean(reparsed),
      warnings: parsedWarnings.length
    })
  );

  return {
    ok: true,
    summary,
    projections,
    installmentsDetail,
    warnings: parsedWarnings,
    parseSource: reparsed ? "rawText-reparsed" : summary.parseStatus || "stored-summary",
    confidenceScore:
      reparsed
        ? parsed.totalAmount > 0 || parsed.minimumPayment > 0 || parsed.projections.length > 0
          ? 0.75
          : 0.35
        : summary.parseStatus === "parsed"
          ? 0.9
          : summary.parseStatus === "parsed-with-warnings"
            ? 0.7
            : null,
    parsed,
    parserDebug
  };
}

export async function debugCardStatement(env: Env, summaryId: string): Promise<CardStatementDetailResult | null> {
  return getCardStatementDetail(env, summaryId);
}

async function analyzeUploadedCardStatement(
  env: Env,
  fileName: string,
  pdfBytes: ArrayBuffer
): Promise<{
  rawText: string;
  preview: ParsedCardStatementPreview;
  detectedIssuer: string;
  textExtractedLength: number;
  canSave: boolean;
}> {
  let rawText = "";
  let detectedIssuer = "unknown";

  logUploadStage("extract-text-start", { fileName, bytes: pdfBytes.byteLength });
  try {
    rawText = await extractPdfText(pdfBytes);
    logUploadStage("extract-text-success", { fileName, chars: rawText.length });
  } catch (error) {
    const appError = error instanceof AppError ? error : new AppError("parse-pdf", "No se pudo extraer texto del PDF.", String(error), 400);
    const fallbackIssuer = safeDetectIssuerFromUpload(fileName);
    logUploadStage("extract-text-failed", {
      fileName,
      detectedIssuer: fallbackIssuer,
      error: appError.message,
      details: appError.details ?? ""
    });
    const fallbackPreview = fallbackIssuer === "naranja-x"
      ? await tryParseNaranjaWithGemini(env, pdfBytes, buildPdfFallbackPreview(fileName, appError.message, fallbackIssuer))
      : buildPdfFallbackPreview(fileName, appError.message, fallbackIssuer);
    return buildAnalysisResult("", fallbackPreview, fallbackIssuer, 0);
  }

  try {
    detectedIssuer = detectIssuerFromText(rawText, fileName);
    logUploadStage("detect-issuer-success", { fileName, detectedIssuer });
  } catch (error) {
    const appError = error instanceof AppError ? error : new AppError("detect-issuer", "No se pudo detectar el emisor.", String(error), 400);
    logUploadStage("detect-issuer-warning", {
      fileName,
      error: appError.message
    });
  }

  let preview: ParsedCardStatementPreview;
  try {
    preview = parseDetectedStatement(rawText, fileName);
    logUploadStage("parse-success", {
      fileName,
      detectedIssuer,
      warnings: preview.warnings.length,
      dueDate: preview.dueDate,
      totalAmount: preview.totalAmount
    });
  } catch (error) {
    const appError = error instanceof AppError ? error : new AppError("parse", "No se pudo parsear el PDF.", String(error), 400);
    logUploadStage("parse-failed", {
      fileName,
      error: appError.message,
      details: appError.details ?? ""
    });
    preview = buildPdfFallbackPreview(fileName, appError.message, detectedIssuer);
  }

  if (detectedIssuer === "naranja-x" && shouldUseGeminiForNaranja(preview)) {
    preview = await tryParseNaranjaWithGemini(env, pdfBytes, preview);
  }

  return buildAnalysisResult(rawText, preview, detectedIssuer, rawText.length);
}

async function appendCardSummaryFromUpload(
  env: Env,
  headers: string[],
  input: {
    cardId: string;
    card: CardRecord | null;
    fileName: string;
    rawText: string;
    preview: ParsedCardStatementPreview;
    dueDate: string;
    nextDueDate: string;
  }
): Promise<CardSummaryRecord> {
  const record: CardSummaryRecord = {
    summaryId: generateId("summary"),
    cardId: input.cardId,
    issuer: input.preview.issuer || input.card?.issuer || "",
    brand: input.preview.brand || input.card?.brand || "",
    bank: input.preview.bank || input.card?.bank || "",
    holder: input.preview.holder || input.card?.holder || "",
    fileName: input.fileName,
    sourceType: "pdf-upload",
    statementDate: input.preview.closingDate || input.dueDate,
    closingDate: input.preview.closingDate,
    dueDate: input.dueDate,
    nextDueDate: input.nextDueDate,
    totalAmount: input.preview.totalAmount,
    minimumPayment: input.preview.minimumPayment,
    currency: "ARS",
    rawText: input.rawText,
    rawDetectedData: stringifyJsonField(input.preview.rawDetectedData),
    warnings: stringifyJsonField(input.preview.warnings),
    parseStatus: input.preview.warnings.length > 0 ? "parsed-with-warnings" : "parsed",
    createdAt: new Date().toISOString()
  };

  await appendRow(env, CARD_SUMMARY_SHEET_NAME, headers, record);
  return record;
}

function buildPdfFallbackPreview(fileName: string, message: string, detectedIssuer = "unknown"): ParsedCardStatementPreview {
  const issuerLabel = detectedIssuer === "naranja-x"
    ? "Naranja X"
    : detectedIssuer === "mastercard-bna"
      ? "Mastercard BNA"
      : detectedIssuer === "visa-bna"
        ? "VISA BNA"
        : detectedIssuer === "visa-santander"
          ? "VISA Santander"
          : "Desconocido";
  return {
    issuer: issuerLabel,
    brand: detectedIssuer === "naranja-x" ? "Naranja X" : "",
    bank: detectedIssuer === "naranja-x" ? "Naranja X" : "",
    holder: "",
    closingDate: "",
    dueDate: "",
    nextDueDate: "",
    totalAmount: 0,
    minimumPayment: 0,
    projections: [],
    installmentsDetail: [],
    rawDetectedData: {
      parser: "fallback",
      fileName,
      failure: message,
      detectedIssuer,
      parseSource: "fallback-manual-required",
      textState: detectedIssuer === "naranja-x" ? "not-legible-or-image-based" : "not-legible"
    },
    warnings: [
      "No se pudo parsear automáticamente este PDF.",
      detectedIssuer === "naranja-x"
        ? "El PDF de Naranja X requiere un fallback especial porque no expone texto legible con el extractor actual."
        : "No se detectó una tabla confiable de cuotas futuras.",
      message,
      "Podés continuar con la carga manual."
    ]
  };
}

async function tryParseNaranjaWithGemini(
  env: Env,
  pdfBytes: ArrayBuffer,
  fallbackPreview: ParsedCardStatementPreview
): Promise<ParsedCardStatementPreview> {
  logUploadStage("gemini-naranja-start", {
    fallbackWarnings: fallbackPreview.warnings.length
  });

  try {
    const geminiPreview = await parseNaranjaWithGemini(pdfBytes, env);
    logUploadStage("gemini-naranja-success", {
      dueDate: geminiPreview.dueDate,
      totalAmount: geminiPreview.totalAmount,
      projections: geminiPreview.projections.length,
      installmentsDetail: geminiPreview.installmentsDetail.length,
      warnings: geminiPreview.warnings.length
    });

    return {
      ...geminiPreview,
      rawDetectedData: {
        ...geminiPreview.rawDetectedData,
        parseSource: "gemini",
        fallbackWarnings: fallbackPreview.warnings
      },
      warnings: Array.from(new Set([...geminiPreview.warnings, ...fallbackPreview.warnings])).filter(
        (warning) => !/continuar|carga manual|fallback especial|no expone texto legible/i.test(warning)
      )
    };
  } catch (error) {
    const appError = error instanceof AppError
      ? error
      : new AppError("parse", "Gemini fallback failed.", String(error), 400);
    logUploadStage("gemini-naranja-failed", {
      error: appError.message,
      details: appError.details ?? ""
    });

    return {
      ...fallbackPreview,
      rawDetectedData: {
        ...fallbackPreview.rawDetectedData,
        geminiFallbackAttempted: true,
        geminiFallbackError: appError.message,
        parseSource: "fallback-manual-required"
      },
      warnings: Array.from(
        new Set([
          ...fallbackPreview.warnings,
          "Fallback Gemini para Naranja no disponible. Podés continuar con la carga manual.",
          appError.message
        ])
      )
    };
  }
}

function shouldUseGeminiForNaranja(preview: ParsedCardStatementPreview): boolean {
  const source = String(preview.rawDetectedData.parseSource ?? "");
  const isNaranja = preview.issuer.toLowerCase().includes("naranja") || preview.brand.toLowerCase().includes("naranja");
  const hasUsefulData = Boolean(preview.dueDate || preview.totalAmount > 0 || preview.projections.length > 0);
  const hasExtractorWarning = preview.warnings.some((warning) =>
    /no readable text|no expone texto|fallback especial|carga manual|no se pudo parsear/i.test(warning)
  );

  return isNaranja && (!hasUsefulData || source === "fallback-manual-required" || hasExtractorWarning);
}

function buildAnalysisResult(
  rawText: string,
  preview: ParsedCardStatementPreview,
  detectedIssuer: string,
  textExtractedLength: number
): {
  rawText: string;
  preview: ParsedCardStatementPreview;
  detectedIssuer: string;
  textExtractedLength: number;
  canSave: boolean;
} {
  return {
    rawText,
    preview,
    detectedIssuer,
    textExtractedLength,
    canSave: Boolean(preview.dueDate || preview.totalAmount > 0 || preview.projections.length > 0)
  };
}

function safeDetectIssuerFromUpload(fileName: string): string {
  try {
    return detectIssuerFromText("", fileName);
  } catch {
    return "unknown";
  }
}

function logUploadStage(stage: string, payload: Record<string, unknown>): void {
  console.log("[card-statements/upload]", JSON.stringify({ stage, ...payload }));
}

export async function importCardSummaryFromText(
  env: Env,
  input: Record<string, unknown>
): Promise<CardSummaryImportResult> {
  await initializeFinanceModuleSheets(env);

  const cardId = stringValue(input.cardId);
  const card = cardId ? await getCardById(env, cardId) : null;
  const rawText = stringValue(input.rawText);

  if (!rawText) {
    throw new AppError("parse", 'Field "rawText" is required to import a card statement.', undefined, 400);
  }

  const parsed = parseCardSummaryText(rawText, {
    statementDate: stringValue(input.statementDate),
    closingDate: stringValue(input.closingDate),
    dueDate: stringValue(input.dueDate),
    currency: stringValue(input.currency) || "ARS"
  });

  const summary = await createCardSummary(env, {
    cardId,
    issuer: stringValue(input.issuer) || card?.issuer || "",
    bank: stringValue(input.bank) || card?.bank || "",
    holder: stringValue(input.holder) || card?.holder || "",
    fileName: stringValue(input.fileName),
    statementDate: parsed.statementDate,
    closingDate: parsed.closingDate,
    dueDate: parsed.dueDate || guessDueDateFromCard(card, parsed.closingDate),
    nextDueDate: addMonthsToDateString(parsed.dueDate || guessDueDateFromCard(card, parsed.closingDate), 1),
    totalAmount: parsed.totalAmount,
    minimumPayment: parsed.minimumPayment,
    currency: parsed.currency,
    rawText,
    parseStatus: parsed.parseStatus
  });

  const installmentSnapshot = await readModuleSheet(env, INSTALLMENT_DETAIL_SHEET_NAME, INSTALLMENT_DETAIL_HEADERS);
  const projectionSnapshot = await readModuleSheet(
    env,
    INSTALLMENT_PROJECTION_SHEET_NAME,
    INSTALLMENT_PROJECTION_HEADERS
  );
  const installments = buildInstallmentDetails(summary, parsed.installments);
  const projections = buildInstallmentProjections(summary, installments);

  if (installments.length > 0) {
    await appendRows(env, INSTALLMENT_DETAIL_SHEET_NAME, installmentSnapshot.headers, installments);
  }

  if (projections.length > 0) {
    await appendRows(env, INSTALLMENT_PROJECTION_SHEET_NAME, projectionSnapshot.headers, projections);
  }

  return {
    summary,
    installments,
    projections
  };
}

export async function listInstallmentProjections(env: Env): Promise<InstallmentProjectionRecord[]> {
  const snapshot = await readModuleSheet(env, INSTALLMENT_PROJECTION_SHEET_NAME, INSTALLMENT_PROJECTION_HEADERS);
  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToInstallmentProjection(snapshot.headers, row))
    .sort((left, right) => left.yearMonth.localeCompare(right.yearMonth) || left.issuer.localeCompare(right.issuer));
}

export async function listInstallmentDetails(env: Env): Promise<InstallmentDetailRecord[]> {
  const snapshot = await readModuleSheet(env, INSTALLMENT_DETAIL_SHEET_NAME, INSTALLMENT_DETAIL_HEADERS);
  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToInstallmentDetail(snapshot.headers, row))
    .sort((left, right) => left.dueMonth.localeCompare(right.dueMonth) || left.createdAt.localeCompare(right.createdAt));
}

export async function listInstallments(env: Env, filters: InstallmentFilters = {}): Promise<InstallmentDetailRecord[]> {
  const details = await listInstallmentDetails(env);
  return applyInstallmentFilters(details, filters);
}

export async function createManualInstallment(
  env: Env,
  input: Record<string, unknown>
): Promise<InstallmentDetailRecord> {
  const snapshot = await readModuleSheet(env, INSTALLMENT_DETAIL_SHEET_NAME, INSTALLMENT_DETAIL_HEADERS);
  const ownerType = parseOwnerType(input.ownerType) ?? "personal";
  const amount = numberValue(input.amount);
  const businessPercent = normalizeBusinessPercent(numberValue(input.businessPercent), ownerType);
  const businessAmount =
    "businessAmount" in input ? numberValue(input.businessAmount) : roundCurrency(amount * (businessPercent / 100));
  const personalAmount =
    "personalAmount" in input ? numberValue(input.personalAmount) : roundCurrency(amount - businessAmount);
  const reimbursementStatus = deriveReimbursementStatus(
    parseReimbursementStatus(input.reimbursementStatus) ?? "pending",
    0,
    businessAmount
  );
  const dueDate = normalizeIsoDateValue(stringValue(input.dueDate));
  const dueMonth = normalizeYearMonthValue(stringValue(input.dueMonth) || dueDate.slice(0, 7));

  const record: InstallmentDetailRecord = {
    installmentId: generateId("installment"),
    summaryId: stringValue(input.summaryId),
    cardId: stringValue(input.cardId),
    purchaseDate: stringValue(input.purchaseDate),
    merchant: stringValue(input.merchant),
    installmentNumber: positiveInteger(input.installmentNumber, 1),
    installmentTotal: positiveInteger(input.installmentTotal, 1),
    amount,
    dueMonth,
    dueDate,
    ownerType,
    businessPercent,
    businessAmount,
    personalAmount,
    reimbursementStatus,
    notes: stringValue(input.notes),
    createdAt: new Date().toISOString()
  };

  await appendRow(env, INSTALLMENT_DETAIL_SHEET_NAME, snapshot.headers, record);
  await syncInstallmentReimbursement(env, record);
  return record;
}

export async function updateInstallmentDetail(
  env: Env,
  input: Record<string, unknown>
): Promise<InstallmentDetailRecord | null> {
  const snapshot = await readModuleSheet(env, INSTALLMENT_DETAIL_SHEET_NAME, INSTALLMENT_DETAIL_HEADERS);
  const installmentId = stringValue(input.installmentId);
  const match = findRowByKey(snapshot, "installmentId", installmentId);

  if (!match) {
    return null;
  }

  const current = rowToInstallmentDetail(snapshot.headers, match.row);
  const ownerType = parseOwnerType(input.ownerType) ?? current.ownerType;
  const nextAmount = "amount" in input ? numberValue(input.amount) : current.amount;
  const businessPercent = normalizeBusinessPercent(
    "businessPercent" in input ? numberValue(input.businessPercent) : current.businessPercent,
    ownerType
  );
  const businessAmount =
    "businessAmount" in input ? numberValue(input.businessAmount) : roundCurrency(nextAmount * (businessPercent / 100));
  const personalAmount =
    "personalAmount" in input ? numberValue(input.personalAmount) : roundCurrency(nextAmount - businessAmount);
  const reimbursementStatus =
    parseReimbursementStatus(input.reimbursementStatus) ?? current.reimbursementStatus;

  const updated: InstallmentDetailRecord = {
    ...current,
    summaryId: "summaryId" in input ? stringValue(input.summaryId) : current.summaryId,
    cardId: "cardId" in input ? stringValue(input.cardId) : current.cardId,
    purchaseDate: "purchaseDate" in input ? stringValue(input.purchaseDate) : current.purchaseDate,
    merchant: "merchant" in input ? stringValue(input.merchant) : current.merchant,
    installmentNumber:
      "installmentNumber" in input ? positiveInteger(input.installmentNumber, current.installmentNumber) : current.installmentNumber,
    installmentTotal:
      "installmentTotal" in input ? positiveInteger(input.installmentTotal, current.installmentTotal) : current.installmentTotal,
    amount: nextAmount,
    dueMonth:
      "dueMonth" in input
        ? normalizeYearMonthValue(stringValue(input.dueMonth))
        : ("dueDate" in input
          ? normalizeYearMonthValue(normalizeIsoDateValue(stringValue(input.dueDate)).slice(0, 7))
          : current.dueMonth),
    dueDate: "dueDate" in input ? normalizeIsoDateValue(stringValue(input.dueDate)) : current.dueDate,
    ownerType,
    businessPercent,
    businessAmount,
    personalAmount,
    reimbursementStatus,
    notes: "notes" in input ? stringValue(input.notes) : current.notes
  };

  await updateRow(env, INSTALLMENT_DETAIL_SHEET_NAME, snapshot.headers, match.index, updated);
  await syncInstallmentReimbursement(env, updated);
  return updated;
}

export async function getInstallment(env: Env, installmentId: string): Promise<InstallmentDetailRecord | null> {
  const details = await listInstallmentDetails(env);
  return details.find((item) => item.installmentId === installmentId) ?? null;
}

export async function getInstallmentForecast(
  env: Env,
  filters: InstallmentFilters = {}
): Promise<InstallmentForecastResponse> {
  const items = (await listInstallments(env, filters)).flatMap((item) => expandInstallmentDetail(item));
  const baseYearMonth = currentYearMonthString();
  const nextYearMonth = addMonthsToYearMonth(baseYearMonth, 1);
  const thirdYearMonth = addMonthsToYearMonth(baseYearMonth, 2);
  const thisMonthItems = items.filter((item) => item.dueMonth === baseYearMonth);
  const nextMonthItems = items.filter((item) => item.dueMonth === nextYearMonth);
  const thirdMonthItems = items.filter((item) => item.dueMonth === thirdYearMonth);
  const totalPending = roundCurrency(items.reduce((sum, item) => sum + item.amount, 0));
  const businessPending = roundCurrency(items.reduce((sum, item) => sum + item.businessAmount, 0));
  const personalPending = roundCurrency(items.reduce((sum, item) => sum + item.personalAmount, 0));

  return {
    thisMonth: buildForecastBucket(baseYearMonth, thisMonthItems),
    nextMonth: buildForecastBucket(nextYearMonth, nextMonthItems),
    thirdMonth: buildForecastBucket(thirdYearMonth, thirdMonthItems),
    totalPending,
    businessPending,
    personalPending,
    filters
  };
}

export async function getInstallmentOutlook(env: Env, baseYearMonth?: string): Promise<InstallmentOutlook> {
  const projections = await listInstallmentProjections(env);
  const details = await listInstallmentDetails(env);
  const currentYearMonth = baseYearMonth || currentYearMonthString();
  const nextYearMonth = addMonthsToYearMonth(currentYearMonth, 1);
  const followingYearMonth = addMonthsToYearMonth(currentYearMonth, 2);

  return {
    currentMonth: {
      yearMonth: currentYearMonth,
      projections: projections.filter((item) => item.yearMonth === currentYearMonth),
      details: details.filter((item) => item.dueMonth === currentYearMonth)
    },
    nextMonth: {
      yearMonth: nextYearMonth,
      projections: projections.filter((item) => item.yearMonth === nextYearMonth),
      details: details.filter((item) => item.dueMonth === nextYearMonth)
    },
    followingMonth: {
      yearMonth: followingYearMonth,
      projections: projections.filter((item) => item.yearMonth === followingYearMonth),
      details: details.filter((item) => item.dueMonth === followingYearMonth)
    }
  };
}

export async function listBusinessReimbursements(env: Env): Promise<BusinessReimbursementRecord[]> {
  const snapshot = await readModuleSheet(env, BUSINESS_REIMBURSEMENT_SHEET_NAME, BUSINESS_REIMBURSEMENT_HEADERS);
  return snapshot.rows
    .filter((row) => row.some((cell) => normalizeCell(cell) !== ""))
    .map((row) => rowToBusinessReimbursement(snapshot.headers, row))
    .sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function getBusinessReimbursement(
  env: Env,
  reimbursementId: string
): Promise<BusinessReimbursementRecord | null> {
  const reimbursements = await listBusinessReimbursements(env);
  return reimbursements.find((item) => item.reimbursementId === reimbursementId) ?? null;
}

export async function createBusinessReimbursement(
  env: Env,
  input: Record<string, unknown>
): Promise<BusinessReimbursementRecord> {
  const snapshot = await readModuleSheet(env, BUSINESS_REIMBURSEMENT_SHEET_NAME, BUSINESS_REIMBURSEMENT_HEADERS);
  const totalPaid = numberValue(input.totalPaid);
  const businessAmount = numberValue(input.businessAmount);
  const reimbursedAmount = numberValue(input.reimbursedAmount);
  const reimbursementStatus = deriveReimbursementStatus(
    parseReimbursementStatus(input.reimbursementStatus) ?? "pending",
    reimbursedAmount,
    businessAmount
  );

  const record: BusinessReimbursementRecord = {
    reimbursementId: generateId("reimbursement"),
    sourceType: stringValue(input.sourceType),
    sourceId: stringValue(input.sourceId),
    cardId: stringValue(input.cardId),
    concept: stringValue(input.concept),
    totalPaid,
    businessAmount,
    personalAmount: "personalAmount" in input ? numberValue(input.personalAmount) : roundCurrency(totalPaid - businessAmount),
    reimbursementStatus,
    reimbursementDueDate: normalizeIsoDateValue(stringValue(input.reimbursementDueDate)),
    reimbursedAmount,
    reimbursedDate: normalizeIsoDateValue(stringValue(input.reimbursedDate)),
    notes: stringValue(input.notes),
    createdAt: new Date().toISOString()
  };

  await appendRow(env, BUSINESS_REIMBURSEMENT_SHEET_NAME, snapshot.headers, record);
  return record;
}

export async function updateBusinessReimbursement(
  env: Env,
  input: Record<string, unknown>
): Promise<BusinessReimbursementRecord | null> {
  const snapshot = await readModuleSheet(env, BUSINESS_REIMBURSEMENT_SHEET_NAME, BUSINESS_REIMBURSEMENT_HEADERS);
  const reimbursementId = stringValue(input.reimbursementId);
  const match = findRowByKey(snapshot, "reimbursementId", reimbursementId);

  if (!match) {
    return null;
  }

  const current = rowToBusinessReimbursement(snapshot.headers, match.row);
  const reimbursedAmount = "reimbursedAmount" in input ? numberValue(input.reimbursedAmount) : current.reimbursedAmount;
  const businessAmount = "businessAmount" in input ? numberValue(input.businessAmount) : current.businessAmount;
  const reimbursementStatus = deriveReimbursementStatus(
    parseReimbursementStatus(input.reimbursementStatus) ?? current.reimbursementStatus,
    reimbursedAmount,
    businessAmount
  );

  const updated: BusinessReimbursementRecord = {
    ...current,
    sourceType: "sourceType" in input ? stringValue(input.sourceType) : current.sourceType,
    sourceId: "sourceId" in input ? stringValue(input.sourceId) : current.sourceId,
    cardId: "cardId" in input ? stringValue(input.cardId) : current.cardId,
    concept: "concept" in input ? stringValue(input.concept) : current.concept,
    totalPaid: "totalPaid" in input ? numberValue(input.totalPaid) : current.totalPaid,
    businessAmount,
    personalAmount:
      "personalAmount" in input
        ? numberValue(input.personalAmount)
        : roundCurrency(("totalPaid" in input ? numberValue(input.totalPaid) : current.totalPaid) - businessAmount),
    reimbursementStatus,
    reimbursementDueDate:
      "reimbursementDueDate" in input
        ? normalizeIsoDateValue(stringValue(input.reimbursementDueDate))
        : current.reimbursementDueDate,
    reimbursedAmount,
    reimbursedDate:
      "reimbursedDate" in input ? normalizeIsoDateValue(stringValue(input.reimbursedDate)) : current.reimbursedDate,
    notes: "notes" in input ? stringValue(input.notes) : current.notes
  };

  await updateRow(env, BUSINESS_REIMBURSEMENT_SHEET_NAME, snapshot.headers, match.index, updated);
  return updated;
}

export async function registerReimbursementPayment(
  env: Env,
  input: Record<string, unknown>
): Promise<BusinessReimbursementRecord | null> {
  const reimbursementId = stringValue(input.reimbursementId);
  const reimbursement = await getBusinessReimbursement(env, reimbursementId);

  if (!reimbursement) {
    return null;
  }

  const paymentAmount = numberValue(input.paymentAmount);
  const updatedReimbursedAmount = roundCurrency(reimbursement.reimbursedAmount + paymentAmount);

  return updateBusinessReimbursement(env, {
    reimbursementId,
    reimbursedAmount: updatedReimbursedAmount,
    reimbursedDate: stringValue(input.reimbursedDate) || new Date().toISOString().slice(0, 10)
  });
}

export async function getCardsDashboard(env: Env): Promise<CardsDashboardResponse> {
  const cards = await listCards(env);
  const statements = await listCardSummaries(env);
  const forecast = await getInstallmentForecast(env);
  const reimbursements = await listBusinessReimbursements(env);

  return {
    activeCards: cards.filter((card) => card.active).length,
    inactiveCards: cards.filter((card) => !card.active).length,
    statementCount: statements.length,
    statementsTotal: roundCurrency(statements.reduce((sum, item) => sum + item.totalAmount, 0)),
    minimumPaymentsTotal: roundCurrency(statements.reduce((sum, item) => sum + item.minimumPayment, 0)),
    pendingInstallmentsThisMonth: forecast.thisMonth.items.length,
    pendingInstallmentsNextMonth: forecast.nextMonth.items.length,
    pendingInstallmentsThirdMonth: forecast.thirdMonth.items.length,
    businessPending: forecast.businessPending,
    personalPending: forecast.personalPending,
    reimbursementsPending: reimbursements.filter((item) => item.reimbursementStatus === "pending").length,
    reimbursementsPartial: reimbursements.filter((item) => item.reimbursementStatus === "partial").length,
    reimbursementsPaid: reimbursements.filter((item) => item.reimbursementStatus === "paid").length,
    reimbursableAmountPending: roundCurrency(
      reimbursements
        .filter((item) => item.reimbursementStatus !== "paid")
        .reduce((sum, item) => sum + Math.max(0, item.businessAmount - item.reimbursedAmount), 0)
    )
  };
}

function applyInstallmentFilters(
  installments: InstallmentDetailRecord[],
  filters: InstallmentFilters
): InstallmentDetailRecord[] {
  return installments.filter((item) => {
    if (filters.cardId && item.cardId !== filters.cardId) {
      return false;
    }

    if (filters.ownerType && item.ownerType !== filters.ownerType) {
      return false;
    }

    if (filters.reimbursementStatus && item.reimbursementStatus !== filters.reimbursementStatus) {
      return false;
    }

    return true;
  });
}

function buildForecastBucket(yearMonth: string, items: InstallmentDetailRecord[]) {
  return {
    yearMonth,
    totalAmount: roundCurrency(items.reduce((sum, item) => sum + item.amount, 0)),
    items
  };
}

function expandInstallmentDetail(installment: InstallmentDetailRecord): InstallmentDetailRecord[] {
  const remainingCount = Math.max(1, installment.installmentTotal - installment.installmentNumber + 1);
  const expanded: InstallmentDetailRecord[] = [];

  for (let offset = 0; offset < remainingCount; offset += 1) {
    const dueDate = addMonthsToDateString(installment.dueDate, offset);

    expanded.push({
      ...installment,
      dueDate,
      dueMonth: dueDate.slice(0, 7)
    });
  }

  return expanded;
}

async function syncInstallmentReimbursement(env: Env, installment: InstallmentDetailRecord): Promise<void> {
  const existing = await getReimbursementBySource(env, "installment", installment.installmentId);

  if (installment.businessAmount <= 0 || installment.ownerType === "personal") {
    if (existing) {
      await updateBusinessReimbursement(env, {
        reimbursementId: existing.reimbursementId,
        businessAmount: 0,
        personalAmount: installment.amount,
        reimbursementStatus: "paid",
        notes: installment.notes
      });
    }

    return;
  }

  const reimbursementPayload = {
    sourceType: "installment",
    sourceId: installment.installmentId,
    cardId: installment.cardId,
    concept: installment.merchant || `Cuota ${installment.installmentNumber}/${installment.installmentTotal}`,
    totalPaid: installment.amount,
    businessAmount: installment.businessAmount,
    personalAmount: installment.personalAmount,
    reimbursementStatus: installment.reimbursementStatus,
    reimbursementDueDate: installment.dueDate,
    notes: installment.notes
  };

  if (existing) {
    await updateBusinessReimbursement(env, {
      reimbursementId: existing.reimbursementId,
      ...reimbursementPayload
    });
    return;
  }

  await createBusinessReimbursement(env, reimbursementPayload);
}

async function getReimbursementBySource(
  env: Env,
  sourceType: string,
  sourceId: string
): Promise<BusinessReimbursementRecord | null> {
  const reimbursements = await listBusinessReimbursements(env);
  return (
    reimbursements.find((item) => item.sourceType === sourceType && item.sourceId === sourceId) ?? null
  );
}

async function getCardById(env: Env, cardId: string): Promise<CardRecord | null> {
  const cards = await listCards(env);
  return cards.find((card) => card.cardId === cardId) ?? null;
}

async function ensureModuleSheet(env: Env, config: ModuleSheetConfig): Promise<ModuleSheetSetupStatus> {
  const existingSheets = await getSpreadsheetSheets(env);
  const existingSheet = existingSheets.find((sheet) => sheet.title === config.sheetName);
  let created = false;
  let addedHeaders: string[] = [];

  if (!existingSheet) {
    await addSheet(env, config.sheetName);
    created = true;
  }

  const currentHeaders = await readHeaders(env, config.sheetName);

  if (currentHeaders.length === 0) {
    await writeHeaders(env, config.sheetName, config.headers);
    return {
      sheetName: config.sheetName,
      created,
      addedHeaders: [...config.headers],
      headers: [...config.headers]
    };
  }

  const missingHeaders = config.headers.filter((header) => !currentHeaders.includes(header));
  if (missingHeaders.length > 0) {
    await appendHeaders(env, config.sheetName, currentHeaders.length + 1, missingHeaders);
    addedHeaders = [...missingHeaders];
  }

  return {
    sheetName: config.sheetName,
    created,
    addedHeaders,
    headers: [...currentHeaders, ...addedHeaders]
  };
}

async function readModuleSheet(
  env: Env,
  sheetName: string,
  headers: readonly string[]
): Promise<SheetSnapshot> {
  await ensureModuleSheet(env, { sheetName, headers });
  const response = await sheetsRequest(
    env,
    `/values/${encodeURIComponent(`${sheetName}${DEFAULT_SHEET_RANGE_SUFFIX}`)}?majorDimension=ROWS`,
    "sheet-read"
  );
  const data = (await response.json()) as GoogleValuesResponse;
  const values = data.values ?? [];
  const detectedHeaders = values[0]?.map((header) => normalizeCell(header)).filter(Boolean) ?? [...headers];

  return {
    headers: detectedHeaders,
    rows: values.slice(1)
  };
}

async function appendRow(
  env: Env,
  sheetName: string,
  headers: string[],
  record: Record<string, string | number | boolean>
): Promise<void> {
  await appendRows(env, sheetName, headers, [record]);
}

async function appendRows(
  env: Env,
  sheetName: string,
  headers: string[],
  records: Array<Record<string, string | number | boolean>>
): Promise<void> {
  if (records.length === 0) {
    return;
  }

  await sheetsRequest(
    env,
    `/values/${encodeURIComponent(`${sheetName}!A1:Z1`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    "sheet-write",
    {
      method: "POST",
      body: JSON.stringify({
        majorDimension: "ROWS",
        values: records.map((record) => buildRowValues(headers, record))
      })
    }
  );
}

async function updateRow(
  env: Env,
  sheetName: string,
  headers: string[],
  rowIndex: number,
  record: Record<string, string | number | boolean>
): Promise<void> {
  const rowNumber = rowIndex + 2;
  const endColumn = columnNumberToLetter(headers.length);

  await sheetsRequest(
    env,
    `/values/${encodeURIComponent(`${sheetName}!A${rowNumber}:${endColumn}${rowNumber}`)}?valueInputOption=USER_ENTERED`,
    "sheet-write",
    {
      method: "PUT",
      body: JSON.stringify({
        majorDimension: "ROWS",
        values: [buildRowValues(headers, record)]
      })
    }
  );
}

async function deleteSheetRow(env: Env, sheetName: string, rowIndex: number): Promise<void> {
  const sheetId = await getSheetIdByName(env, sheetName);
  const startIndex = rowIndex + 1;

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
}

async function getSpreadsheetSheets(env: Env): Promise<Array<{ sheetId: number; title: string }>> {
  const response = await sheetsRequest(
    env,
    `?fields=${encodeURIComponent("sheets(properties(sheetId,title))")}`,
    "sheet-read"
  );
  const data = (await response.json()) as GoogleSpreadsheetResponse;

  return (data.sheets ?? [])
    .map((sheet) => ({
      sheetId: sheet.properties?.sheetId ?? -1,
      title: sheet.properties?.title ?? ""
    }))
    .filter((sheet) => sheet.sheetId >= 0 && sheet.title);
}

async function getSheetIdByName(env: Env, sheetName: string): Promise<number> {
  const sheets = await getSpreadsheetSheets(env);
  const match = sheets.find((sheet) => sheet.title === sheetName);

  if (!match) {
    throw new AppError("sheet-read", `Sheet "${sheetName}" was not found in the configured spreadsheet.`);
  }

  return match.sheetId;
}

async function addSheet(env: Env, sheetName: string): Promise<void> {
  await sheetsRequest(env, ":batchUpdate", "setup", {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
              gridProperties: {
                frozenRowCount: 1
              }
            }
          }
        }
      ]
    })
  });
}

async function readHeaders(env: Env, sheetName: string): Promise<string[]> {
  const response = await sheetsRequest(env, `/values/${encodeURIComponent(`${sheetName}!1:1`)}?majorDimension=ROWS`, "sheet-read");
  const data = (await response.json()) as GoogleValuesResponse;
  return data.values?.[0]?.map((header) => normalizeCell(header)).filter(Boolean) ?? [];
}

async function writeHeaders(env: Env, sheetName: string, headers: readonly string[]): Promise<void> {
  const range = `${sheetName}!A1:${columnNumberToLetter(headers.length)}1`;

  await sheetsRequest(env, `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, "sheet-write", {
    method: "PUT",
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [headers]
    })
  });
}

async function appendHeaders(
  env: Env,
  sheetName: string,
  startColumnNumber: number,
  headers: readonly string[]
): Promise<void> {
  const startColumn = columnNumberToLetter(startColumnNumber);
  const endColumn = columnNumberToLetter(startColumnNumber + headers.length - 1);
  const range = `${sheetName}!${startColumn}1:${endColumn}1`;

  await sheetsRequest(env, `/values/${encodeURIComponent(range)}?valueInputOption=RAW`, "sheet-write", {
    method: "PUT",
    body: JSON.stringify({
      majorDimension: "ROWS",
      values: [headers]
    })
  });
}

async function sheetsRequest(
  env: Env,
  path: string,
  step: AppErrorStep,
  init?: RequestInit
): Promise<Response> {
  const accessToken = await getGoogleAccessToken(env);
  const spreadsheetId = getSanitizedSpreadsheetId(env);
  const url = `${GOOGLE_SHEETS_API}/${spreadsheetId}${path}`;

  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {})
    }
  });

  if (!response.ok) {
    const body = await response.text();
    throw new AppError(step, "Google Sheets module request failed.", `status=${response.status} body=${body}`);
  }

  return response;
}

function parseCardSummaryText(
  rawText: string,
  fallback: { statementDate: string; closingDate: string; dueDate: string; currency: string }
): ParsedSummaryImport {
  const normalizedText = rawText.replace(/\r/g, "");
  const lines = normalizedText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const referenceYear = currentYear();
  const closingDate =
    detectLabeledDate(normalizedText, [/fecha de cierre/i, /\bcierre\b/i], referenceYear) || fallback.closingDate;
  const dueDate =
    detectLabeledDate(normalizedText, [/fecha de vencimiento/i, /\bvencimiento\b/i, /pago hasta/i], referenceYear) ||
    fallback.dueDate;
  const statementDate =
    detectLabeledDate(normalizedText, [/fecha del resumen/i, /\bresumen\b/i, /periodo/i], referenceYear) ||
    fallback.statementDate ||
    closingDate;
  const totalAmount =
    detectLabeledAmount(normalizedText, [/total a pagar/i, /saldo total/i, /\btotal\b/i]) ?? 0;
  const minimumPayment =
    detectLabeledAmount(normalizedText, [/pago minimo/i, /minimo a pagar/i, /\bminimo\b/i]) ?? 0;
  const currency = /u\$s|usd/i.test(normalizedText) ? "USD" : fallback.currency || "ARS";
  const installments = parseInstallmentLines(lines, statementDate || closingDate, dueDate);

  return {
    statementDate,
    closingDate,
    dueDate,
    totalAmount,
    minimumPayment,
    currency,
    parseStatus: installments.length > 0 || totalAmount > 0 || dueDate ? "parsed" : "partial",
    installments
  };
}

function parseInstallmentLines(
  lines: string[],
  purchaseDate: string,
  dueDate: string
): ParsedSummaryImport["installments"] {
  const parsed: ParsedSummaryImport["installments"] = [];

  for (const line of lines) {
    const installmentMatch = line.match(/(.+?)(?:\s+|-)(?:cuota\s*)?(\d{1,2})\/(\d{1,2})/i);
    if (!installmentMatch) {
      continue;
    }

    const amount = extractLastAmount(line);
    if (amount === null || amount <= 0) {
      continue;
    }

    parsed.push({
      merchant: cleanupMerchant(installmentMatch[1]),
      installmentNumber: Number(installmentMatch[2]),
      installmentTotal: Number(installmentMatch[3]),
      amount,
      purchaseDate: purchaseDate || dueDate || ""
    });
  }

  return parsed;
}

function buildInstallmentDetails(
  summary: CardSummaryRecord,
  parsedInstallments: ParsedSummaryImport["installments"]
): InstallmentDetailRecord[] {
  return parsedInstallments.map((item) => ({
    installmentId: generateId("installment"),
    summaryId: summary.summaryId,
    cardId: summary.cardId,
    purchaseDate: item.purchaseDate,
    merchant: item.merchant,
    installmentNumber: item.installmentNumber,
    installmentTotal: item.installmentTotal,
    amount: item.amount,
    dueMonth: summary.dueDate.slice(0, 7),
    dueDate: summary.dueDate,
    ownerType: "personal",
    businessPercent: 0,
    businessAmount: 0,
    personalAmount: item.amount,
    reimbursementStatus: "pending",
    notes: "",
    createdAt: new Date().toISOString()
  }));
}

function buildInstallmentProjections(
  summary: CardSummaryRecord,
  installments: InstallmentDetailRecord[]
): InstallmentProjectionRecord[] {
  const projections: InstallmentProjectionRecord[] = [];
  const createdAt = new Date().toISOString();

  if (installments.length === 0 && summary.totalAmount > 0 && summary.dueDate) {
    projections.push({
      projectionId: generateId("projection"),
      summaryId: summary.summaryId,
      cardId: summary.cardId,
      issuer: summary.issuer,
      monthLabel: formatMonthLabel(summary.dueDate.slice(0, 7)),
      yearMonth: summary.dueDate.slice(0, 7),
      amount: summary.totalAmount,
      sourceType: "statement-total",
      confirmed: true,
      createdAt
    });

    return projections;
  }

  for (const installment of installments) {
    for (let offset = 0; offset <= installment.installmentTotal - installment.installmentNumber; offset += 1) {
      const projectedDueDate = addMonthsToDateString(installment.dueDate, offset);
      const yearMonth = projectedDueDate.slice(0, 7);

      projections.push({
        projectionId: generateId("projection"),
        summaryId: summary.summaryId,
        cardId: summary.cardId,
        issuer: summary.issuer,
        monthLabel: formatMonthLabel(yearMonth),
        yearMonth,
        amount: installment.amount,
        sourceType: "installment",
        confirmed: offset === 0,
        createdAt
      });
    }
  }

  return projections;
}

function buildProjectionRecordsFromPreview(
  summary: CardSummaryRecord,
  preview: ParsedCardStatementPreview
): InstallmentProjectionRecord[] {
  const createdAt = new Date().toISOString();
  const previewProjections = preview.projections
    .filter((item) => item.yearMonth && item.amount > 0)
    .map((item) => ({
      projectionId: generateId("projection"),
      summaryId: summary.summaryId,
      cardId: summary.cardId,
      issuer: summary.issuer,
      monthLabel: item.monthLabel || formatMonthLabel(normalizeYearMonthValue(item.yearMonth)),
      yearMonth: normalizeYearMonthValue(item.yearMonth),
      amount: roundCurrency(item.amount),
      sourceType: "pdf-upload",
      confirmed: true,
      createdAt
    }))
    .filter((item) => item.yearMonth !== "");

  if (previewProjections.length > 0) {
    return previewProjections;
  }

  if (summary.totalAmount > 0 && summary.dueDate) {
    return [
      {
        projectionId: generateId("projection"),
        summaryId: summary.summaryId,
        cardId: summary.cardId,
        issuer: summary.issuer,
        monthLabel: formatMonthLabel(summary.dueDate.slice(0, 7)),
        yearMonth: summary.dueDate.slice(0, 7),
        amount: summary.totalAmount,
        sourceType: "statement-total",
        confirmed: true,
        createdAt
      }
    ];
  }

  return [];
}

function buildInstallmentDetailRecordsFromPreview(
  summary: CardSummaryRecord,
  preview: ParsedCardStatementPreview
): InstallmentDetailRecord[] {
  const createdAt = new Date().toISOString();
  const dueDate = summary.dueDate || preview.dueDate;
  const dueMonth = dueDate ? dueDate.slice(0, 7) : "";

  return preview.installmentsDetail
    .filter((item) => item.merchant && item.amount > 0 && item.installmentTotal > 1)
    .map((item) => ({
      installmentId: generateId("installment"),
      summaryId: summary.summaryId,
      cardId: summary.cardId,
      purchaseDate: "",
      merchant: item.merchant,
      installmentNumber: item.installmentNumber,
      installmentTotal: item.installmentTotal,
      amount: roundCurrency(item.amount),
      dueMonth,
      dueDate,
      ownerType: "personal" as const,
      businessPercent: 0,
      businessAmount: 0,
      personalAmount: roundCurrency(item.amount),
      reimbursementStatus: "pending" as const,
      notes: `Cuota ${item.installmentNumber}/${item.installmentTotal}. Restantes: ${item.remainingInstallments}`,
      createdAt
    }));
}

function inheritInstallmentClassifications(
  newInstallments: InstallmentDetailRecord[],
  historicalInstallments: InstallmentDetailRecord[]
): InstallmentDetailRecord[] {
  return newInstallments.map((installment) => {
    const previous = findPreviousInstallmentMatch(installment, historicalInstallments);

    if (!previous) {
      return installment;
    }

    return {
      ...installment,
      ownerType: previous.ownerType,
      businessPercent: previous.businessPercent,
      businessAmount: previous.businessAmount,
      personalAmount: previous.personalAmount,
      reimbursementStatus: previous.reimbursementStatus === "paid" ? "pending" : previous.reimbursementStatus,
      notes: previous.notes
    };
  });
}

export function findPreviousInstallmentMatch(
  newInstallment: InstallmentDetailRecord,
  historicalInstallments: InstallmentDetailRecord[]
): InstallmentDetailRecord | null {
  const normalizedMerchant = normalizeInstallmentMerchantForMatch(newInstallment.merchant);

  if (!normalizedMerchant || newInstallment.amount <= 0 || newInstallment.installmentTotal <= 1) {
    return null;
  }

  return historicalInstallments
    .filter((candidate) => {
      const candidateMerchant = normalizeInstallmentMerchantForMatch(candidate.merchant);
      const merchantMatches =
        candidateMerchant === normalizedMerchant ||
        (candidateMerchant.length >= 5 &&
          normalizedMerchant.length >= 5 &&
          (candidateMerchant.includes(normalizedMerchant) || normalizedMerchant.includes(candidateMerchant)));

      return (
        candidate.installmentId !== newInstallment.installmentId &&
        candidate.installmentTotal === newInstallment.installmentTotal &&
        candidate.installmentNumber < newInstallment.installmentNumber &&
        roundCurrency(candidate.amount) === roundCurrency(newInstallment.amount) &&
        merchantMatches
      );
    })
    .sort(
      (left, right) =>
        right.installmentNumber - left.installmentNumber || right.createdAt.localeCompare(left.createdAt)
    )[0] ?? null;
}

function normalizeInstallmentMerchantForMatch(value: string): string {
  return stringValue(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
}

function rowToCard(headers: string[], row: string[]): CardRecord {
  const record = rowToObject(headers, row);
  return {
    ...record,
    cardId: stringValue(record.cardId),
    issuer: stringValue(record.issuer),
    brand: stringValue(record.brand),
    bank: stringValue(record.bank),
    holder: stringValue(record.holder),
    last4: stringValue(record.last4),
    closeDay: numberValue(record.closeDay),
    dueDay: numberValue(record.dueDay),
    active: booleanValue(record.active, true),
    createdAt: stringValue(record.createdAt)
  };
}

function rowToCardSummary(headers: string[], row: string[]): CardSummaryRecord {
  const record = rowToObject(headers, row);
  return {
    ...record,
    summaryId: stringValue(record.summaryId),
    cardId: stringValue(record.cardId),
    issuer: stringValue(record.issuer),
    brand: stringValue(record.brand),
    bank: stringValue(record.bank),
    holder: stringValue(record.holder),
    fileName: stringValue(record.fileName),
    sourceType: stringValue(record.sourceType),
    statementDate: stringValue(record.statementDate),
    closingDate: stringValue(record.closingDate),
    dueDate: stringValue(record.dueDate),
    nextDueDate: stringValue(record.nextDueDate),
    totalAmount: numberValue(record.totalAmount),
    minimumPayment: numberValue(record.minimumPayment),
    currency: stringValue(record.currency),
    rawText: stringValue(record.rawText),
    rawDetectedData: stringValue(record.rawDetectedData),
    warnings: stringValue(record.warnings),
    parseStatus: stringValue(record.parseStatus),
    createdAt: stringValue(record.createdAt)
  };
}

function rowToInstallmentProjection(headers: string[], row: string[]): InstallmentProjectionRecord {
  const record = rowToObject(headers, row);
  return {
    ...record,
    projectionId: stringValue(record.projectionId),
    summaryId: stringValue(record.summaryId),
    cardId: stringValue(record.cardId),
    issuer: stringValue(record.issuer),
    monthLabel: stringValue(record.monthLabel),
    yearMonth: stringValue(record.yearMonth),
    amount: numberValue(record.amount),
    sourceType: stringValue(record.sourceType),
    confirmed: booleanValue(record.confirmed, false),
    createdAt: stringValue(record.createdAt)
  };
}

function rowToInstallmentDetail(headers: string[], row: string[]): InstallmentDetailRecord {
  const record = rowToObject(headers, row);
  return {
    ...record,
    installmentId: stringValue(record.installmentId),
    summaryId: stringValue(record.summaryId),
    cardId: stringValue(record.cardId),
    purchaseDate: stringValue(record.purchaseDate),
    merchant: stringValue(record.merchant),
    installmentNumber: numberValue(record.installmentNumber),
    installmentTotal: numberValue(record.installmentTotal),
    amount: numberValue(record.amount),
    dueMonth: stringValue(record.dueMonth),
    dueDate: stringValue(record.dueDate),
    ownerType: parseOwnerType(record.ownerType) ?? "personal",
    businessPercent: numberValue(record.businessPercent),
    businessAmount: numberValue(record.businessAmount),
    personalAmount: numberValue(record.personalAmount),
    reimbursementStatus: parseReimbursementStatus(record.reimbursementStatus) ?? "pending",
    notes: stringValue(record.notes),
    createdAt: stringValue(record.createdAt)
  };
}

function rowToBusinessReimbursement(headers: string[], row: string[]): BusinessReimbursementRecord {
  const record = rowToObject(headers, row);
  return {
    ...record,
    reimbursementId: stringValue(record.reimbursementId),
    sourceType: stringValue(record.sourceType),
    sourceId: stringValue(record.sourceId),
    cardId: stringValue(record.cardId),
    concept: stringValue(record.concept),
    totalPaid: numberValue(record.totalPaid),
    businessAmount: numberValue(record.businessAmount),
    personalAmount: numberValue(record.personalAmount),
    reimbursementStatus: parseReimbursementStatus(record.reimbursementStatus) ?? "pending",
    reimbursementDueDate: stringValue(record.reimbursementDueDate),
    reimbursedAmount: numberValue(record.reimbursedAmount),
    reimbursedDate: stringValue(record.reimbursedDate),
    notes: stringValue(record.notes),
    createdAt: stringValue(record.createdAt)
  };
}

function rowToObject(headers: string[], row: string[]): Record<string, string> {
  return headers.reduce<Record<string, string>>((result, header, index) => {
    result[header] = normalizeCell(row[index]);
    return result;
  }, {});
}

function findRowByKey(snapshot: SheetSnapshot, key: string, value: string): { row: string[]; index: number } | null {
  const keyIndex = snapshot.headers.indexOf(key);

  if (keyIndex === -1) {
    throw new AppError("parse", `Header "${key}" was not found in sheet snapshot.`);
  }

  for (let index = 0; index < snapshot.rows.length; index += 1) {
    if (normalizeCell(snapshot.rows[index]?.[keyIndex]) === value) {
      return { row: snapshot.rows[index] ?? [], index };
    }
  }

  return null;
}

function buildRowValues(headers: string[], record: Record<string, string | number | boolean>): string[] {
  return headers.map((header) => stringifyCellValue(record[header]));
}

function stringifyCellValue(value: string | number | boolean | undefined): string {
  if (value === undefined || value === null) {
    return "";
  }

  if (typeof value === "boolean") {
    return value ? "true" : "false";
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? String(value) : "";
  }

  return String(value).trim();
}

function stringifyJsonField(value: unknown): string {
  if (typeof value === "string") {
    return value.trim();
  }

  if (Array.isArray(value) || (typeof value === "object" && value !== null)) {
    try {
      return JSON.stringify(value);
    } catch {
      return "";
    }
  }

  return "";
}

function parseStoredRecord(value: string): Record<string, string | number | boolean | string[]> {
  if (!value) {
    return {};
  }

  try {
    const parsed = JSON.parse(value);
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed as Record<string, string | number | boolean | string[]>;
    }
  } catch {
    return { raw: value };
  }

  return {};
}

function parseStoredStringArray(value: string): string[] {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) {
      return parsed.map(stringifyUnknownArrayItem).filter(Boolean);
    }
    if (parsed && typeof parsed === "object") {
      return Object.entries(parsed)
        .map(([key, item]) => `${key}: ${stringifyUnknownArrayItem(item)}`.trim())
        .filter(Boolean);
    }
    if (typeof parsed === "string") {
      return [parsed.trim()].filter(Boolean);
    }
  } catch {
    return [value.trim()].filter(Boolean);
  }

  return [];
}

function stringifyUnknownArrayItem(value: unknown): string {
  if (typeof value === "string") {
    return value.trim();
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (value && typeof value === "object") {
    return Object.entries(value as Record<string, unknown>)
      .map(([key, item]) => `${key}: ${stringifyUnknownArrayItem(item)}`.trim())
      .filter(Boolean)
      .join(", ");
  }

  return "";
}

function normalizeIsoDateValue(value: string): string {
  if (!value) {
    return "";
  }

  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    return "";
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day) || month < 1 || month > 12 || day < 1 || day > 31) {
    return "";
  }

  return `${match[1]}-${match[2]}-${match[3]}`;
}

function normalizeYearMonthValue(value: string): string {
  if (!value) {
    return "";
  }

  const match = value.match(/^(\d{4})-(\d{2})$/);
  if (!match) {
    return "";
  }

  const month = Number(match[2]);
  if (!Number.isFinite(month) || month < 1 || month > 12) {
    return "";
  }

  return `${match[1]}-${match[2]}`;
}

function detectLabeledDate(text: string, labels: RegExp[], fallbackYear: number): string {
  for (const label of labels) {
    const index = text.search(label);

    if (index === -1) {
      continue;
    }

    const window = text.slice(index, index + 120);
    const isoMatch = window.match(/\b(\d{4})[-/](\d{2})[-/](\d{2})\b/);

    if (isoMatch) {
      return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    }

    const shortMatch = window.match(/\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b/);

    if (shortMatch) {
      return normalizeDate(shortMatch[0], fallbackYear);
    }
  }

  return "";
}

function detectLabeledAmount(text: string, labels: RegExp[]): number | null {
  for (const label of labels) {
    const index = text.search(label);

    if (index === -1) {
      continue;
    }

    const window = text.slice(index, index + 160);
    const amount = extractLastAmount(window);

    if (amount !== null) {
      return amount;
    }
  }

  return null;
}

function extractLastAmount(text: string): number | null {
  const matches = Array.from(text.matchAll(/-?\$?\s*\d[\d.\s,]*/g));

  for (let index = matches.length - 1; index >= 0; index -= 1) {
    const amount = parseLocalizedNumber(matches[index]?.[0] ?? "");

    if (amount !== null) {
      return amount;
    }
  }

  return null;
}

function parseLocalizedNumber(value: string): number | null {
  const cleaned = value.replace(/[^\d,.-]/g, "").trim();

  if (!cleaned) {
    return null;
  }

  const lastComma = cleaned.lastIndexOf(",");
  const lastDot = cleaned.lastIndexOf(".");

  if (lastComma > lastDot) {
    const normalized = cleaned.replace(/\./g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? roundCurrency(parsed) : null;
  }

  const normalized = cleaned.replace(/,/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? roundCurrency(parsed) : null;
}

function normalizeDate(value: string, fallbackYear: number): string {
  const isoMatch = value.match(/\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b/);

  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);

    if (year > 0 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day
        .toString()
        .padStart(2, "0")}`;
    }
  }

  const shortMatch = value.match(/\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b/);

  if (!shortMatch) {
    return "";
  }

  const day = Number(shortMatch[1]);
  const month = Number(shortMatch[2]);
  const rawYear = shortMatch[3];
  const year = rawYear ? (rawYear.length === 2 ? 2000 + Number(rawYear) : Number(rawYear)) : fallbackYear;

  if (month < 1 || month > 12 || day < 1 || day > 31 || !Number.isFinite(year)) {
    return "";
  }

  return `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day
    .toString()
    .padStart(2, "0")}`;
}

function guessDueDateFromCard(card: CardRecord | null, closingDate: string): string {
  if (!closingDate) {
    return "";
  }

  if (!card || card.dueDay <= 0) {
    return closingDate;
  }

  const [yearValue, monthValue] = closingDate.slice(0, 7).split("-");
  const year = Number(yearValue);
  const month = Number(monthValue);

  if (!Number.isFinite(year) || !Number.isFinite(month)) {
    return closingDate;
  }

  const closingDay = Number(closingDate.slice(8, 10));
  const targetMonth = closingDay > card.dueDay ? month + 1 : month;
  const targetDate = new Date(Date.UTC(year, targetMonth - 1, 1));
  const safeDay = Math.min(card.dueDay, daysInMonth(targetDate.getUTCFullYear(), targetDate.getUTCMonth() + 1));

  return `${targetDate.getUTCFullYear().toString().padStart(4, "0")}-${(targetDate.getUTCMonth() + 1)
    .toString()
    .padStart(2, "0")}-${safeDay.toString().padStart(2, "0")}`;
}

function addMonthsToDateString(dateValue: string, monthsToAdd: number): string {
  if (!dateValue) {
    return "";
  }

  const [yearValue, monthValue, dayValue] = dateValue.split("-");
  const year = Number(yearValue);
  const month = Number(monthValue);
  const day = Number(dayValue);

  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return dateValue;
  }

  const target = new Date(Date.UTC(year, month - 1 + monthsToAdd, 1));
  const safeDay = Math.min(day, daysInMonth(target.getUTCFullYear(), target.getUTCMonth() + 1));

  return `${target.getUTCFullYear().toString().padStart(4, "0")}-${(target.getUTCMonth() + 1)
    .toString()
    .padStart(2, "0")}-${safeDay.toString().padStart(2, "0")}`;
}

function addMonthsToYearMonth(yearMonth: string, monthsToAdd: number): string {
  const [yearValue, monthValue] = yearMonth.split("-");
  const year = Number(yearValue);
  const month = Number(monthValue);

  if (!Number.isFinite(year) || !Number.isFinite(month)) {
    return yearMonth;
  }

  const target = new Date(Date.UTC(year, month - 1 + monthsToAdd, 1));
  return `${target.getUTCFullYear().toString().padStart(4, "0")}-${(target.getUTCMonth() + 1)
    .toString()
    .padStart(2, "0")}`;
}

function formatMonthLabel(yearMonth: string): string {
  const [yearValue, monthValue] = yearMonth.split("-");
  const year = Number(yearValue);
  const month = Number(monthValue);

  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) {
    return yearMonth;
  }

  return new Intl.DateTimeFormat("es-AR", { month: "long", year: "numeric", timeZone: "UTC" }).format(
    new Date(Date.UTC(year, month - 1, 1))
  );
}

function currentYearMonthString(): string {
  const now = new Date();
  return `${now.getUTCFullYear().toString().padStart(4, "0")}-${(now.getUTCMonth() + 1)
    .toString()
    .padStart(2, "0")}`;
}

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function currentYear(): number {
  return new Date().getUTCFullYear();
}

function cleanupMerchant(value: string): string {
  return value
    .replace(/\s{2,}/g, " ")
    .replace(/\bcuota\s+\d+\s*\/\s*\d+\b/i, "")
    .trim();
}

function deriveReimbursementStatus(
  preferredStatus: ReimbursementStatus,
  reimbursedAmount: number,
  businessAmount: number
): ReimbursementStatus {
  if (businessAmount <= 0) {
    return "paid";
  }

  if (reimbursedAmount >= businessAmount) {
    return "paid";
  }

  if (reimbursedAmount > 0) {
    return "partial";
  }

  return preferredStatus === "paid" ? "pending" : preferredStatus;
}

function normalizeBusinessPercent(value: number, ownerType: OwnerType): number {
  if (ownerType === "personal") {
    return 0;
  }

  if (ownerType === "business") {
    return 100;
  }

  if (!Number.isFinite(value)) {
    return 50;
  }

  return Math.max(0, Math.min(100, roundCurrency(value)));
}

function parseOwnerType(value: unknown): OwnerType | null {
  return value === "personal" || value === "business" || value === "mixed" ? value : null;
}

function parseReimbursementStatus(value: unknown): ReimbursementStatus | null {
  return value === "pending" || value === "partial" || value === "paid" ? value : null;
}

function boundedDay(value: unknown): number {
  const parsed = Math.floor(numberValue(value));
  return Math.max(1, Math.min(31, parsed || 1));
}

function positiveInteger(value: unknown, fallback: number): number {
  const parsed = Math.floor(numberValue(value));
  return parsed > 0 ? parsed : fallback;
}

function booleanValue(value: unknown, fallback: boolean): boolean {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();

    if (normalized === "true") {
      return true;
    }

    if (normalized === "false") {
      return false;
    }
  }

  return fallback;
}

function numberValue(value: unknown): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const parsed = parseLocalizedNumber(stringValue(value));
  return parsed ?? 0;
}

function stringValue(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeCell(value: string | undefined): string {
  return typeof value === "string" ? value.trim() : "";
}

function roundCurrency(value: number): number {
  return Math.round(value * 100) / 100;
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

function generateId(prefix: string): string {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function getSanitizedSpreadsheetId(env: Env): string {
  return env.SPREADSHEET_ID.trim().replace(/\/+$/, "");
}
