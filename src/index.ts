import { Hono } from "hono";
import {
  createTransaction,
  debugSheets,
  deleteTransaction,
  listTransactions,
  updateTransaction
} from "./lib/sheets";
import type {
  AppErrorStep,
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

function parseTransactionType(value: unknown): TransactionType | null {
  return value === "income" || value === "expense" ? value : null;
}

function parseAmount(value: unknown): number | null {
  const amount = typeof value === "number" ? value : Number(String(value ?? "").trim());
  return Number.isFinite(amount) ? amount : null;
}

function stringField(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function dateField(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
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
