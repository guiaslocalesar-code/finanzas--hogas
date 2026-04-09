import { AppError, type Env, type ParsedCardStatementPreview } from "./types";

const GEMINI_API_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/models";
const GEMINI_MODEL_CANDIDATES = ["gemini-2.5-flash"] as const;

const NARANJA_PROMPT = `Sos un analista financiero experto en extraer datos de resúmenes de tarjetas de crédito argentinas, específicamente de Tarjeta Naranja.
Analizá el documento y devolvé ÚNICAMENTE un JSON válido que coincida exactamente con la estructura requerida. No uses markdown.

Reglas de extracción de compras (MUY IMPORTANTE):
1. Buscá el detalle de todas las compras del mes.
2. Si el consumo indica cuotas (ej. 'Plan Z', 'Z', '02/03', 'Cuota 1 de 3', etc.), extraé el número de la cuota actual en 'installmentNumber' y el total de cuotas en 'installmentTotal'.
3. Si el consumo NO indica cuotas y es un pago único (ej. supermercado, débito automático, suscripciones), poné 'installmentNumber': 1 e 'installmentTotal': 1.
4. Ignorá líneas que sean pagos de la tarjeta, intereses, mantenimiento o impuestos (IVA, Sellos). Solo me interesan los consumos reales.

Estructura JSON obligatoria a devolver:
{
  "summary": {
    "issuer": "Naranja X",
    "holder": string | null,
    "closingDate": string | null (YYYY-MM-DD),
    "dueDate": string | null (YYYY-MM-DD),
    "nextDueDate": string | null (YYYY-MM-DD),
    "totalAmount": number | null,
    "minimumPayment": number | null
  },
  "projections": [
    {
      "monthLabel": string (ej. 'Abril-26'),
      "yearMonth": string (ej. '2026-04'),
      "amount": number
    }
  ],
  "installmentsDetail": [
    {
      "merchant": string (Limpiá códigos raros al inicio),
      "purchaseDate": string | null (YYYY-MM-DD),
      "installmentNumber": number,
      "installmentTotal": number,
      "amount": number (solo el monto de esta cuota, sin $),
      "remainingInstallments": number | null
    }
  ]
}`;

type GeminiProjection = {
  monthLabel?: string | null;
  yearMonth?: string | null;
  amount?: number | string | null;
};

type GeminiInstallment = {
  purchaseDate?: string | null;
  merchant?: string | null;
  installmentNumber?: number | string | null;
  installmentTotal?: number | string | null;
  amount?: number | string | null;
  remainingInstallments?: number | string | null;
};

type GeminiStatementSummary = {
  issuer?: string | null;
  holder?: string | null;
  closingDate?: string | null;
  dueDate?: string | null;
  nextDueDate?: string | null;
  totalAmount?: number | string | null;
  minimumPayment?: number | string | null;
};

type GeminiStatementPayload = GeminiStatementSummary & {
  summary?: GeminiStatementSummary | null;
  projections?: GeminiProjection[] | null;
  installmentsDetail?: GeminiInstallment[] | null;
};

type GeminiResponse = {
  modelVersion?: string;
  candidates?: Array<{
    content?: {
      parts?: Array<{
        text?: string;
      }>;
    };
  }>;
};

export async function parseNaranjaWithGemini(
  pdfBufferOrBase64: ArrayBuffer | string,
  env: Env
): Promise<ParsedCardStatementPreview> {
  const apiKey = env.GEMINI_API_KEY?.trim();

  if (!apiKey) {
    throw new AppError(
      "parse",
      "GEMINI_API_KEY is not configured.",
      "Configure GEMINI_API_KEY as a Cloudflare Worker secret or in .dev.vars for local development.",
      400
    );
  }

  const pdfBase64 =
    typeof pdfBufferOrBase64 === "string" ? pdfBufferOrBase64 : arrayBufferToBase64(pdfBufferOrBase64);
  const geminiResult = await generateGeminiJson(apiKey, pdfBase64);
  const parsed = parseJson<GeminiStatementPayload>(geminiResult.jsonText, "Gemini returned invalid structured JSON.");
  const preview = normalizeGeminiStatement(parsed);
  preview.rawDetectedData.geminiModel = geminiResult.model;
  preview.rawDetectedData.geminiModelVersion = geminiResult.modelVersion ?? "";

  return preview;
}

async function generateGeminiJson(
  apiKey: string,
  pdfBase64: string
): Promise<{ jsonText: string; model: string; modelVersion?: string }> {
  let lastError: AppError | null = null;

  for (const model of GEMINI_MODEL_CANDIDATES) {
    try {
      return await requestGeminiJson(apiKey, pdfBase64, model);
    } catch (error) {
      const appError = error instanceof AppError
        ? error
        : new AppError("parse", "Gemini could not parse the Naranja PDF.", String(error), 400);
      lastError = appError;

      if (!isUnavailableGeminiModelError(appError)) {
        throw appError;
      }
    }
  }

  throw lastError ?? new AppError("parse", "Gemini could not parse the Naranja PDF.", undefined, 400);
}

async function requestGeminiJson(
  apiKey: string,
  pdfBase64: string,
  model: string
): Promise<{ jsonText: string; model: string; modelVersion?: string }> {
  const response = await fetch(`${GEMINI_API_BASE_URL}/${model}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      system_instruction: {
        parts: [{ text: NARANJA_PROMPT }]
      },
      contents: [
        {
          role: "user",
          parts: [
            { text: "Analizá este PDF de Tarjeta Naranja y devolvé el JSON requerido." },
            {
              inline_data: {
                mime_type: "application/pdf",
                data: pdfBase64
              }
            }
          ]
        }
      ],
      generationConfig: {
        responseMimeType: "application/json",
        temperature: 0,
        maxOutputTokens: 8192
      }
    })
  });

  const responseText = await response.text();

  if (!response.ok) {
    throw new AppError("parse", "Gemini could not parse the Naranja PDF.", responseText.slice(0, 1200), 400);
  }

  const gemini = parseJson<GeminiResponse>(responseText, "Gemini returned a non-JSON API response.");
  const jsonText = gemini.candidates?.[0]?.content?.parts?.find((part) => part.text)?.text?.trim();

  if (!jsonText) {
    throw new AppError("parse", "Gemini response did not include JSON text.", responseText.slice(0, 1200), 400);
  }

  return {
    jsonText,
    model,
    modelVersion: gemini.modelVersion
  };
}

function isUnavailableGeminiModelError(error: AppError): boolean {
  return error.details?.includes("no longer available") === true ||
    error.details?.includes('"status": "NOT_FOUND"') === true ||
    error.details?.includes('"status":"NOT_FOUND"') === true;
}

function normalizeGeminiStatement(parsed: GeminiStatementPayload): ParsedCardStatementPreview {
  const summary = parsed.summary ?? parsed;
  const warnings: string[] = [];
  const projections = Array.isArray(parsed.projections)
    ? parsed.projections
        .map((item) => ({
          monthLabel: stringOrEmpty(item.monthLabel),
          yearMonth: stringOrEmpty(item.yearMonth),
          amount: currencyNumber(item.amount)
        }))
        .filter((item) => item.yearMonth && item.amount > 0)
    : [];
  const installmentsDetail = Array.isArray(parsed.installmentsDetail)
    ? parsed.installmentsDetail
        .map((item) => {
          const installmentNumber = integerNumber(item.installmentNumber) || 1;
          const installmentTotal = integerNumber(item.installmentTotal) || 1;
          return {
            purchaseDate: isoDateOrEmpty(item.purchaseDate),
            merchant: stringOrEmpty(item.merchant),
            installmentNumber,
            installmentTotal,
            amount: currencyNumber(item.amount),
            remainingInstallments: Math.max(
              integerNumber(item.remainingInstallments),
              installmentTotal - installmentNumber,
              0
            )
          };
        })
        .filter((item) => item.merchant && item.amount > 0 && item.installmentTotal >= item.installmentNumber)
    : [];

  const preview: ParsedCardStatementPreview = {
    issuer: "Naranja X",
    brand: "Naranja X",
    bank: "Naranja X",
    holder: stringOrEmpty(summary.holder),
    closingDate: isoDateOrEmpty(summary.closingDate),
    dueDate: isoDateOrEmpty(summary.dueDate),
    nextDueDate: isoDateOrEmpty(summary.nextDueDate),
    totalAmount: currencyNumber(summary.totalAmount),
    minimumPayment: currencyNumber(summary.minimumPayment),
    projections,
    installmentsDetail,
    rawDetectedData: {
      parser: "parseNaranjaWithGemini",
      parseSource: "gemini",
      fieldsFound: [],
      projectionsFound: projections.length,
      installmentsDetailFound: installmentsDetail.length
    },
    warnings
  };

  addMissingWarnings(preview);
  preview.rawDetectedData.fieldsFound = Object.entries({
    holder: preview.holder,
    closingDate: preview.closingDate,
    dueDate: preview.dueDate,
    nextDueDate: preview.nextDueDate,
    totalAmount: preview.totalAmount,
    minimumPayment: preview.minimumPayment
  })
    .filter(([, value]) => (typeof value === "number" ? value > 0 : value !== ""))
    .map(([field]) => field);

  return preview;
}

function addMissingWarnings(preview: ParsedCardStatementPreview): void {
  if (!preview.dueDate) {
    preview.warnings.push('Gemini fallback: field "dueDate" could not be detected.');
  }

  if (preview.totalAmount <= 0) {
    preview.warnings.push('Gemini fallback: field "totalAmount" could not be detected.');
  }

  if (preview.projections.length === 0) {
    preview.warnings.push("Gemini fallback: no se detectó tabla de pagos futuros.");
  }
}

function parseJson<T>(value: string, message: string): T {
  try {
    return JSON.parse(value) as T;
  } catch (error) {
    throw new AppError("parse", message, error instanceof Error ? error.message : String(error), 400);
  }
}

function arrayBufferToBase64(value: ArrayBuffer): string {
  const bytes = new Uint8Array(value);
  let binary = "";

  for (let index = 0; index < bytes.length; index += 0x8000) {
    const chunk = bytes.subarray(index, index + 0x8000);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

function stringOrEmpty(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function isoDateOrEmpty(value: unknown): string {
  const candidate = stringOrEmpty(value);
  return /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(candidate) ? candidate : "";
}

function integerNumber(value: unknown): number {
  const parsed = Math.trunc(currencyNumber(value));
  return Number.isFinite(parsed) ? parsed : 0;
}

function currencyNumber(value: unknown): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? Math.round(value * 100) / 100 : 0;
  }

  if (typeof value === "string") {
    const normalized = value.replace(/\./g, "").replace(",", ".").replace(/[^\d.-]/g, "");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? Math.round(parsed * 100) / 100 : 0;
  }

  return 0;
}
