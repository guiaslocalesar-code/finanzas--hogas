import { Hono } from "hono";
import { runParserSmokeTests } from "./lib/pdf-statements";
import {
  createBusinessReimbursement,
  createCard,
  createCardSummary,
  createManualInstallment,
  debugCardStatement,
  debugUploadCardStatementPdf,
  deleteCard,
  getCardsDashboard,
  getCardStatement,
  getInstallmentForecast,
  initializeFinanceModuleSheets,
  listBusinessReimbursements,
  listCardSummaries,
  listCards,
  listInstallmentProjections,
  listInstallments,
  registerReimbursementPayment,
  uploadCardStatementPdf,
  updateBusinessReimbursement,
  updateCard,
  updateCardStatement,
  updateInstallmentDetail
} from "./lib/finance-modules";
import {
  createTransaction,
  debugSheets,
  deleteTransaction,
  listTransactions,
  updateTransaction
} from "./lib/sheets";
import type {
  AppErrorStep,
  InstallmentFilters,
  OwnerType,
  ReimbursementStatus,
  CreateTransactionInput,
  Env,
  TransactionType,
  UpdateTransactionInput
} from "./lib/types";
import { AppError, DEFAULT_HEADERS, SHEET_NAME } from "./lib/types";

const app = new Hono<{ Bindings: Env }>();

app.get("/", async (c) => {
  return c.env.ASSETS.fetch(new URL("/index.html", c.req.url));
});

app.get("/health", (c) => {
  return c.json({ ok: true });
});

app.get("/api/transactions", async (c) => {
  try {
    const transactions = await listTransactions(c.env);
    return c.json(transactions);
  } catch (error) {
    const appError = toAppError(error, "sheet-read");
    return c.json(
      {
        ok: false,
        step: appError.step,
        error: appError.message,
        details: appError.details ?? null
      },
      { status: appError.status as 500 }
    );
  }
});

app.post("/api/transactions", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const input = validateCreateInput(payload);

  if (!input.ok) {
    return c.json({ error: input.error }, 400);
  }

  const transaction = await createTransaction(c.env, {
    ...input.value,
    id: Date.now().toString(),
    createdAt: new Date().toISOString()
  });

  return c.json(transaction, 201);
});

app.patch("/api/transactions", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const input = validateUpdateInput(payload);

  if (!input.ok) {
    return c.json({ error: input.error }, 400);
  }

  const transaction = await updateTransaction(c.env, input.value);

  if (!transaction) {
    return c.json({ error: "Transaction not found." }, 404);
  }

  return c.json(transaction);
});

app.delete("/api/transactions/:id", async (c) => {
  const id = c.req.param("id").trim();

  if (!id) {
    return c.json({ error: "Transaction id is required." }, 400);
  }

  const deleted = await deleteTransaction(c.env, id);

  if (!deleted) {
    return c.json({ error: "Transaction not found." }, 404);
  }

  return c.json({ ok: true, id });
});

app.get("/debug/sheets", async (c) => {
  try {
    const result = await debugSheets(c.env);
    return c.json({
      ok: true,
      ...result
    });
  } catch (error) {
    const appError = toAppError(error, "sheet-read");
    return c.json(
      {
        ok: false,
        spreadsheetId: c.env.SPREADSHEET_ID ?? "",
        sheetName: SHEET_NAME,
        expectedHeaders: DEFAULT_HEADERS,
        detectedHeaders: [],
        rowCount: 0,
        firstRecord: null,
        step: appError.step,
        error: appError.message,
        details: appError.details ?? null
      },
      { status: appError.status as 500 }
    );
  }
});

app.post("/api/setup/modules", async (c) => {
  const result = await initializeFinanceModuleSheets(c.env);
  return c.json(result);
});

app.get("/api/cards", async (c) => {
  const cards = await listCards(c.env);
  return c.json(cards);
});

app.post("/api/cards", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const cardInput = validateCardPayload(payload, false);
  if (!cardInput.ok) {
    return c.json({ error: cardInput.error }, 400);
  }

  const card = await createCard(c.env, cardInput.value);
  return c.json(card, 201);
});

app.patch("/api/cards/:id", async (c) => {
  const payload = await safeJson(c);
  const cardId = c.req.param("id").trim();

  if (!payload || !cardId) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const cardInput = validateCardPayload(payload, true);
  if (!cardInput.ok) {
    return c.json({ error: cardInput.error }, 400);
  }

  const card = await updateCard(c.env, { ...cardInput.value, cardId });

  if (!card) {
    return c.json({ error: "Card not found." }, 404);
  }

  return c.json(card);
});

app.delete("/api/cards/:id", async (c) => {
  const cardId = c.req.param("id").trim();

  if (!cardId) {
    return c.json({ error: "Card id is required." }, 400);
  }

  const deleted = await deleteCard(c.env, cardId);

  if (!deleted) {
    return c.json({ error: "Card not found." }, 404);
  }

  return c.json({ ok: true, cardId });
});

app.get("/api/card-statements", async (c) => {
  const summaries = await listCardSummaries(c.env);
  return c.json(summaries);
});

app.get("/api/card-statements/:id", async (c) => {
  const summaryId = c.req.param("id").trim();
  const summary = await getCardStatement(c.env, summaryId);

  if (!summary) {
    return c.json({ error: "Statement not found." }, 404);
  }

  return c.json(summary);
});

app.post("/api/card-statements/manual", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const statementInput = validateCardStatementPayload(payload, false);
  if (!statementInput.ok) {
    return c.json({ error: statementInput.error }, 400);
  }

  const summary = await createCardSummary(c.env, statementInput.value);
  return c.json(summary, 201);
});

app.post("/api/card-statements/upload", async (c) => {
  console.log("[card-statements/upload]", JSON.stringify({ stage: "endpoint-enter" }));

  try {
    const uploadInput = await parseCardStatementUploadRequest(c);
    console.log(
      "[card-statements/upload]",
      JSON.stringify({
        stage: "request-validated",
        fileName: uploadInput.fileName,
        mimeType: uploadInput.mimeType,
        size: uploadInput.pdfBytes.byteLength,
        cardId: uploadInput.cardId,
        previewOnly: uploadInput.previewOnly
      })
    );

    const result = await uploadCardStatementPdf(c.env, uploadInput);
    return c.json(result, result.summary ? 201 : 200);
  } catch (error) {
    return c.json(buildUploadErrorResponse(error, "parse-pdf"), getErrorStatus(error));
  }
});

app.post("/debug/card-statements/upload-test", async (c) => {
  console.log("[card-statements/upload-test]", JSON.stringify({ stage: "endpoint-enter" }));

  try {
    const uploadInput = await parseCardStatementUploadRequest(c);
    const result = await debugUploadCardStatementPdf(c.env, uploadInput);
    return c.json(result);
  } catch (error) {
    return c.json(buildUploadErrorResponse(error, "parse-pdf"), getErrorStatus(error));
  }
});

app.patch("/api/card-statements/:id", async (c) => {
  const payload = await safeJson(c);
  const summaryId = c.req.param("id").trim();

  if (!payload || !summaryId) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const statementInput = validateCardStatementPayload(payload, true);
  if (!statementInput.ok) {
    return c.json({ error: statementInput.error }, 400);
  }

  const summary = await updateCardStatement(c.env, summaryId, statementInput.value);

  if (!summary) {
    return c.json({ error: "Statement not found." }, 404);
  }

  return c.json(summary);
});

app.get("/debug/card-statement/:id", async (c) => {
  const summaryId = c.req.param("id").trim();

  if (!summaryId) {
    return c.json({ error: "Statement id is required." }, 400);
  }

  const result = await debugCardStatement(c.env, summaryId);

  if (!result) {
    return c.json({ error: "Statement not found." }, 404);
  }

  return c.json(result);
});

app.get("/debug/cards", async (c) => {
  const cards = await listCards(c.env);
  return c.json({
    ok: true,
    count: cards.length,
    active: cards.filter((card) => card.active).length,
    cards
  });
});

app.get("/debug/statements", async (c) => {
  const statements = await listCardSummaries(c.env);
  const parserSmoke = runParserSmokeTests({
    visaSantander:
      "SANTANDER RIO VISA RESUMEN DE CUENTA TITULAR DE CUENTA: JUAN PEREZ CIERRE ACTUAL: 12 mar. 26 VENCIMIENTO ACTUAL: 25 mar. 26 SALDO ACTUAL: $ 515.721,14 PAGO MINIMO: $ 60.096,00 Cuotas a vencer Abril/26 123.456,78 Mayo/26 22.000,00",
    visaBna:
      "VISA PLATINUM BANCO NACION TITULAR DE CUENTA: MERCADO LEANDRO EZEQUIEL CIERRE ACTUAL: 12 mar. 26 VENCIMIENTO ACTUAL: 25 mar. 26 SALDO ACTUAL: $ 515.721,14 PAGO MINIMO: $ 60.096,00 Cuotas a vencer Abril/26 54.000,00 Mayo/26 43.210,00",
    mastercardBna:
      "MASTERCARD PLATINUM BANCO NACION Estado de cuenta al: 12-Mar-26 Vencimiento actual: 25-Mar-26 Saldo actual: $ 975.446,16 Pago Minimo: $ 112.510,00 Cuotas a vencer Abril-26 33.000,00 Mayo-26 22.000,00",
    naranjaX:
      "NARANJA X JOHANA LUCIANA ZAMUDIO Tu total a pagar es $608.101,95 y vence el 10/04/26. LA MENOR ENTREGA $351.900,00 El resumen actual cerro el 27/03. El proximo resumen cierra el 27/04, y vence el 10/05. Cuotas futuras Mayo/26 $187.257,30 Junio/26 $56.670,00"
  });
  return c.json({
    ok: true,
    count: statements.length,
    parsedWithWarnings: statements.filter((statement) => statement.parseStatus === "parsed-with-warnings").length,
    parserSmoke,
    statements
  });
});

app.get("/debug/installments", async (c) => {
  const filters = getInstallmentFilters(c);
  const installments = await listInstallments(c.env, filters);
  const projections = await listInstallmentProjections(c.env);
  const forecast = await getInstallmentForecast(c.env, filters);
  return c.json({
    ok: true,
    filters,
    installmentCount: installments.length,
    projectionCount: projections.length,
    forecast,
    installments,
    projections
  });
});

app.get("/debug/reimbursements", async (c) => {
  const reimbursements = await listBusinessReimbursements(c.env);
  return c.json({
    ok: true,
    count: reimbursements.length,
    pending: reimbursements.filter((item) => item.reimbursementStatus === "pending").length,
    partial: reimbursements.filter((item) => item.reimbursementStatus === "partial").length,
    paid: reimbursements.filter((item) => item.reimbursementStatus === "paid").length,
    reimbursements
  });
});

app.get("/api/installments/forecast", async (c) => {
  const forecast = await getInstallmentForecast(c.env, getInstallmentFilters(c));
  return c.json(forecast);
});

app.get("/api/installments", async (c) => {
  const details = await listInstallments(c.env, getInstallmentFilters(c));
  return c.json(details);
});

app.post("/api/installments/manual", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const installmentInput = validateInstallmentPayload(payload, false);
  if (!installmentInput.ok) {
    return c.json({ error: installmentInput.error }, 400);
  }

  const detail = await createManualInstallment(c.env, installmentInput.value);
  return c.json(detail, 201);
});

app.patch("/api/installments/:id", async (c) => {
  const payload = await safeJson(c);
  const installmentId = c.req.param("id").trim();

  if (!payload || !installmentId) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const installmentInput = validateInstallmentPayload(payload, true);
  if (!installmentInput.ok) {
    return c.json({ error: installmentInput.error }, 400);
  }

  const detail = await updateInstallmentDetail(c.env, { ...installmentInput.value, installmentId });

  if (!detail) {
    return c.json({ error: "Installment not found." }, 404);
  }

  return c.json(detail);
});

app.get("/api/reimbursements", async (c) => {
  const reimbursements = await listBusinessReimbursements(c.env);
  return c.json(reimbursements);
});

app.post("/api/reimbursements", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const reimbursementInput = validateReimbursementPayload(payload, false);
  if (!reimbursementInput.ok) {
    return c.json({ error: reimbursementInput.error }, 400);
  }

  const reimbursement = await createBusinessReimbursement(c.env, reimbursementInput.value);
  return c.json(reimbursement, 201);
});

app.patch("/api/reimbursements/:id", async (c) => {
  const payload = await safeJson(c);
  const reimbursementId = c.req.param("id").trim();

  if (!payload || !reimbursementId) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const reimbursementInput = validateReimbursementPayload(payload, true);
  if (!reimbursementInput.ok) {
    return c.json({ error: reimbursementInput.error }, 400);
  }

  const reimbursement = await updateBusinessReimbursement(c.env, {
    ...reimbursementInput.value,
    reimbursementId
  });

  if (!reimbursement) {
    return c.json({ error: "Reimbursement not found." }, 404);
  }

  return c.json(reimbursement);
});

app.post("/api/reimbursements/register-payment", async (c) => {
  const payload = await safeJson(c);

  if (!payload) {
    return c.json({ error: "Invalid JSON body." }, 400);
  }

  const reimbursementId = stringField(payload.reimbursementId);
  const paymentAmount = parseAmount(payload.paymentAmount);

  if (!reimbursementId) {
    return c.json({ error: 'Field "reimbursementId" is required.' }, 400);
  }

  if (paymentAmount === null || paymentAmount <= 0) {
    return c.json({ error: 'Field "paymentAmount" must be a valid number greater than zero.' }, 400);
  }

  const reimbursement = await registerReimbursementPayment(c.env, {
    reimbursementId,
    paymentAmount,
    reimbursedDate: dateField(payload.reimbursedDate)
  });

  if (!reimbursement) {
    return c.json({ error: "Reimbursement not found." }, 404);
  }

  return c.json(reimbursement);
});

app.get("/api/cards/dashboard", async (c) => {
  const dashboard = await getCardsDashboard(c.env);
  return c.json(dashboard);
});

app.notFound(async (c) => {
  if (c.req.method === "GET" && !c.req.path.startsWith("/api/")) {
    return c.env.ASSETS.fetch(new URL("/index.html", c.req.url));
  }

  return c.json({ error: "Not found." }, 404);
});

app.onError((error, c) => {
  console.error(error);
  const appError = toAppError(error, "parse");
  return c.json(
    {
      ok: false,
      step: appError.step,
      error: appError.message,
      details: appError.details ?? null
    },
    { status: appError.status as 500 }
  );
});

export default app;

async function safeJson(c: { req: { json: () => Promise<unknown> } }): Promise<Record<string, unknown> | null> {
  try {
    const payload = await c.req.json();
    return typeof payload === "object" && payload !== null ? (payload as Record<string, unknown>) : null;
  } catch {
    return null;
  }
}

function validateCreateInput(
  payload: Record<string, unknown>
): { ok: true; value: CreateTransactionInput } | { ok: false; error: string } {
  const type = parseTransactionType(payload.type);
  const amount = parseAmount(payload.amount);

  if (!type) {
    return { ok: false, error: 'Field "type" must be "income" or "expense".' };
  }

  if (amount === null) {
    return { ok: false, error: 'Field "amount" must be a valid number.' };
  }

  return {
    ok: true,
    value: {
      type,
      amount,
      category: stringField(payload.category),
      description: stringField(payload.description),
      date: dateField(payload.date),
      dueDate: dateField(payload.dueDate)
    }
  };
}

function validateUpdateInput(
  payload: Record<string, unknown>
): { ok: true; value: UpdateTransactionInput } | { ok: false; error: string } {
  const id = stringField(payload.id);

  if (!id) {
    return { ok: false, error: 'Field "id" is required.' };
  }

  const value: UpdateTransactionInput = { id };

  if ("type" in payload) {
    const type = parseTransactionType(payload.type);
    if (!type) {
      return { ok: false, error: 'Field "type" must be "income" or "expense".' };
    }
    value.type = type;
  }

  if ("amount" in payload) {
    const amount = parseAmount(payload.amount);
    if (amount === null) {
      return { ok: false, error: 'Field "amount" must be a valid number.' };
    }
    value.amount = amount;
  }

  if ("category" in payload) {
    value.category = stringField(payload.category);
  }

  if ("description" in payload) {
    value.description = stringField(payload.description);
  }

  if ("date" in payload) {
    value.date = dateField(payload.date);
  }

  if ("dueDate" in payload) {
    value.dueDate = dateField(payload.dueDate);
  }

  return { ok: true, value };
}

function validateCardPayload(
  payload: Record<string, unknown>,
  partial: boolean
): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  if (!partial && !stringField(payload.issuer)) {
    return { ok: false, error: 'Field "issuer" is required.' };
  }

  if (!partial && !stringField(payload.holder)) {
    return { ok: false, error: 'Field "holder" is required.' };
  }

  const value: Record<string, unknown> = {};
  setIfPresent(value, payload, "issuer", stringField);
  setIfPresent(value, payload, "brand", stringField);
  setIfPresent(value, payload, "bank", stringField);
  setIfPresent(value, payload, "holder", stringField);
  setIfPresent(value, payload, "last4", stringField);
  setIfPresent(value, payload, "closeDay", numericField);
  setIfPresent(value, payload, "dueDay", numericField);
  setIfPresent(value, payload, "active", booleanField);

  return { ok: true, value };
}

function validateCardStatementPayload(
  payload: Record<string, unknown>,
  partial: boolean
): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  if (!partial && !stringField(payload.cardId)) {
    return { ok: false, error: 'Field "cardId" is required.' };
  }

  if (!partial && !dateField(payload.dueDate) && !dateField(payload.closingDate)) {
    return { ok: false, error: 'Field "dueDate" or "closingDate" is required.' };
  }

  const value: Record<string, unknown> = {};
  setIfPresent(value, payload, "cardId", stringField);
  setIfPresent(value, payload, "issuer", stringField);
  setIfPresent(value, payload, "brand", stringField);
  setIfPresent(value, payload, "bank", stringField);
  setIfPresent(value, payload, "holder", stringField);
  setIfPresent(value, payload, "fileName", stringField);
  setIfPresent(value, payload, "sourceType", stringField);
  setIfPresent(value, payload, "statementDate", dateField);
  setIfPresent(value, payload, "closingDate", dateField);
  setIfPresent(value, payload, "dueDate", dateField);
  setIfPresent(value, payload, "nextDueDate", dateField);
  setIfPresent(value, payload, "totalAmount", numericField);
  setIfPresent(value, payload, "minimumPayment", numericField);
  setIfPresent(value, payload, "currency", stringField);
  setIfPresent(value, payload, "rawText", stringField);
  setIfPresent(value, payload, "rawDetectedData", rawJsonField);
  setIfPresent(value, payload, "warnings", rawJsonField);
  if ("parseStatus" in payload) {
    value.parseStatus = stringField(payload.parseStatus) || "manual";
  } else if (!partial) {
    value.parseStatus = "manual";
  }

  return { ok: true, value };
}

function validateInstallmentPayload(
  payload: Record<string, unknown>,
  partial: boolean
): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  const amount = numericField(payload.amount);
  const ownerType = parseOwnerType(payload.ownerType);
  const reimbursementStatus = parseReimbursementStatus(payload.reimbursementStatus);

  if (!partial && !stringField(payload.merchant)) {
    return { ok: false, error: 'Field "merchant" is required.' };
  }

  if (!partial && (!amount || amount <= 0)) {
    return { ok: false, error: 'Field "amount" must be a valid number greater than zero.' };
  }

  if (!partial && !dateField(payload.dueDate)) {
    return { ok: false, error: 'Field "dueDate" is required.' };
  }

  if ("ownerType" in payload && !ownerType) {
    return { ok: false, error: 'Field "ownerType" must be "personal", "business" or "mixed".' };
  }

  if ("reimbursementStatus" in payload && !reimbursementStatus) {
    return { ok: false, error: 'Field "reimbursementStatus" must be "pending", "partial" or "paid".' };
  }

  const value: Record<string, unknown> = {};
  setIfPresent(value, payload, "summaryId", stringField);
  setIfPresent(value, payload, "cardId", stringField);
  setIfPresent(value, payload, "purchaseDate", dateField);
  setIfPresent(value, payload, "merchant", stringField);
  setIfPresent(value, payload, "installmentNumber", numericField);
  setIfPresent(value, payload, "installmentTotal", numericField);
  if ("amount" in payload) {
    value.amount = amount;
  }
  setIfPresent(value, payload, "dueMonth", stringField);
  setIfPresent(value, payload, "dueDate", dateField);
  if ("ownerType" in payload) {
    value.ownerType = ownerType ?? undefined;
  }
  setIfPresent(value, payload, "businessPercent", numericField);
  setIfPresent(value, payload, "businessAmount", numericField);
  setIfPresent(value, payload, "personalAmount", numericField);
  if ("reimbursementStatus" in payload) {
    value.reimbursementStatus = reimbursementStatus ?? undefined;
  }
  setIfPresent(value, payload, "notes", stringField);

  return { ok: true, value };
}

function validateReimbursementPayload(
  payload: Record<string, unknown>,
  partial: boolean
): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  const reimbursementStatus = parseReimbursementStatus(payload.reimbursementStatus);

  if (!partial && !stringField(payload.concept)) {
    return { ok: false, error: 'Field "concept" is required.' };
  }

  if ("reimbursementStatus" in payload && !reimbursementStatus) {
    return { ok: false, error: 'Field "reimbursementStatus" must be "pending", "partial" or "paid".' };
  }

  const value: Record<string, unknown> = {};
  setIfPresent(value, payload, "sourceType", stringField);
  setIfPresent(value, payload, "sourceId", stringField);
  setIfPresent(value, payload, "cardId", stringField);
  setIfPresent(value, payload, "concept", stringField);
  setIfPresent(value, payload, "totalPaid", numericField);
  setIfPresent(value, payload, "businessAmount", numericField);
  setIfPresent(value, payload, "personalAmount", numericField);
  if ("reimbursementStatus" in payload) {
    value.reimbursementStatus = reimbursementStatus ?? undefined;
  }
  setIfPresent(value, payload, "reimbursementDueDate", dateField);
  setIfPresent(value, payload, "reimbursedAmount", numericField);
  setIfPresent(value, payload, "reimbursedDate", dateField);
  setIfPresent(value, payload, "notes", stringField);

  return { ok: true, value };
}

function getInstallmentFilters(c: {
  req: { query: (name: string) => string | undefined };
}): InstallmentFilters {
  const ownerType = parseOwnerType(c.req.query("ownerType"));
  const reimbursementStatus = parseReimbursementStatus(c.req.query("reimbursementStatus"));

  return {
    cardId: stringField(c.req.query("cardId")),
    ownerType: ownerType ?? undefined,
    reimbursementStatus: reimbursementStatus ?? undefined
  };
}

function parseTransactionType(value: unknown): TransactionType | null {
  return value === "income" || value === "expense" ? value : null;
}

function parseAmount(value: unknown): number | null {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === "string" && value.trim() === "") {
    return null;
  }

  const amount = typeof value === "number" ? value : Number(String(value ?? "").trim());
  return Number.isFinite(amount) ? amount : null;
}

function stringField(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function dateField(value: unknown): string {
  if (typeof value !== "string") {
    return "";
  }

  const trimmed = value.trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(trimmed) ? trimmed : "";
}

function rawJsonField(value: unknown): unknown {
  return value;
}

function numericField(value: unknown): number | undefined {
  const parsed = parseAmount(value);
  return parsed === null ? undefined : parsed;
}

function setIfPresent<T>(
  target: Record<string, unknown>,
  source: Record<string, unknown>,
  key: string,
  parser: (value: unknown) => T
): void {
  if (key in source) {
    target[key] = parser(source[key]);
  }
}

function booleanField(value: unknown): boolean | undefined {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    if (value.trim().toLowerCase() === "true") {
      return true;
    }

    if (value.trim().toLowerCase() === "false") {
      return false;
    }
  }

  return undefined;
}

function formDataString(value: FormDataEntryValue | null): string {
  return typeof value === "string" ? value.trim() : "";
}

function parseBooleanFlag(value: FormDataEntryValue | null): boolean {
  if (typeof value !== "string") {
    return false;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
}

async function parseCardStatementUploadRequest(c: {
  req: { formData: () => Promise<FormData> };
}): Promise<{
  fileName: string;
  mimeType: string;
  pdfBytes: ArrayBuffer;
  cardId: string;
  previewOnly: boolean;
}> {
  let formData: FormData;

  try {
    formData = await c.req.formData();
    console.log("[card-statements/upload]", JSON.stringify({ stage: "form-data-read" }));
  } catch (error) {
    throw new AppError("form-data", "No se pudo leer multipart/form-data.", String(error), 400);
  }

  const fileEntry = formData.get("file") ?? formData.get("pdf");
  const cardId = formDataString(formData.get("cardId"));
  const previewOnly = parseBooleanFlag(formData.get("previewOnly"));

  console.log(
    "[card-statements/upload]",
    JSON.stringify({
      stage: "form-data-fields",
      fileFieldFound: Boolean(fileEntry),
      cardId,
      previewOnly
    })
  );

  if (!(fileEntry instanceof File)) {
    throw new AppError("validate", 'Field "file" is required and must be a PDF file.', undefined, 400);
  }

  const fileName = fileEntry.name.trim();
  const mimeType = fileEntry.type.trim().toLowerCase();

  if (!fileName) {
    throw new AppError("validate", "The uploaded file must include a file name.", undefined, 400);
  }

  if (!fileName.toLowerCase().endsWith(".pdf") && mimeType !== "application/pdf") {
    throw new AppError("validate", "Only PDF uploads are supported.", JSON.stringify({ fileName, mimeType }), 400);
  }

  if (fileEntry.size <= 0) {
    throw new AppError("validate", "The uploaded PDF is empty.", JSON.stringify({ fileName, mimeType }), 400);
  }

  let pdfBytes: ArrayBuffer;
  try {
    pdfBytes = await fileEntry.arrayBuffer();
    console.log(
      "[card-statements/upload]",
      JSON.stringify({
        stage: "file-bytes-read",
        fileName,
        mimeType,
        size: pdfBytes.byteLength
      })
    );
  } catch (error) {
    throw new AppError("parse-pdf", "No se pudieron leer los bytes del PDF.", String(error), 400);
  }

  if (pdfBytes.byteLength === 0) {
    throw new AppError("validate", "The uploaded PDF is empty.", JSON.stringify({ fileName, mimeType }), 400);
  }

  return {
    fileName,
    mimeType,
    pdfBytes,
    cardId,
    previewOnly
  };
}

function buildUploadErrorResponse(error: unknown, fallbackStage: AppErrorStep) {
  const appError = toAppError(error, fallbackStage);
  let details: unknown = appError.details ?? null;

  if (typeof details === "string") {
    try {
      details = JSON.parse(details);
    } catch {
      details = details || null;
    }
  }

  return {
    ok: false,
    stage: appError.step,
    error: errorCodeForStage(appError.step),
    message: appError.message,
    details
  };
}

function getErrorStatus(error: unknown): 400 | 404 | 500 {
  if (error instanceof AppError) {
    return error.status === 404 ? 404 : error.status === 400 ? 400 : 500;
  }

  return 500;
}

function errorCodeForStage(stage: string): string {
  switch (stage) {
    case "form-data":
      return "FORM_DATA_INVALID";
    case "validate":
      return "UPLOAD_VALIDATION_FAILED";
    case "parse-pdf":
      return "PDF_PARSE_FAILED";
    case "detect-issuer":
      return "PDF_ISSUER_NOT_DETECTED";
    case "sheet-write":
      return "SHEET_SAVE_FAILED";
    case "setup":
      return "SHEET_SETUP_FAILED";
    default:
      return "UPLOAD_FAILED";
  }
}

function parseOwnerType(value: unknown): OwnerType | null {
  return value === "personal" || value === "business" || value === "mixed" ? value : null;
}

function parseReimbursementStatus(value: unknown): ReimbursementStatus | null {
  return value === "pending" || value === "partial" || value === "paid" ? value : null;
}

function toAppError(error: unknown, fallbackStep: AppErrorStep): AppError {
  if (error instanceof AppError) {
    return error;
  }

  return new AppError(
    fallbackStep,
    error instanceof Error ? error.message : "Internal server error.",
    error instanceof Error ? error.stack : String(error)
  );
}
