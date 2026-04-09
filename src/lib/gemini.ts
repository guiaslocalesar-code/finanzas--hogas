import { AppError, type Env, type ParsedCardStatementPreview } from "./types";

const GEMINI_MODEL_URL =
  "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent";

const NARANJA_PROMPT =
  "Sos un extractor de datos de resúmenes de tarjetas argentinas. Analizá el PDF de Tarjeta Naranja. Devolvé ÚNICAMENTE un JSON válido. Extraé: issuer (siempre 'Naranja X'), holder, closingDate (YYYY-MM-DD), dueDate (YYYY-MM-DD), nextDueDate (YYYY-MM-DD), totalAmount (number sin $), minimumPayment (number sin $). Además, extraé la tabla de pagos futuros en el array 'projections' (monthLabel, yearMonth, amount) y el detalle de cuotas en el array 'installmentsDetail' (merchant, installmentNumber, installmentTotal, amount). Si falta algo, poné null.";

type GeminiProjection = {
  monthLabel?: string | null;
  yearMonth?: string | null;
  amount?: number | string | null;
};

type GeminiInstallment = {
  merchant?: string | null;
  installmentNumber?: number | string | null;
  installmentTotal?: number | string | null;
  amount?: number | string | null;
};

type GeminiStatementPayload = {
  issuer?: string | null;
  holder?: string | null;
  closingDate?: string | null;
  dueDate?: string | null;
  nextDueDate?: string | null;
  totalAmount?: number | string | null;
  minimumPayment?: number | string | null;
  projections?: GeminiProjection[] | null;
  installmentsDetail?: GeminiInstallment[] | null;
};

type GeminiResponse = {
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
  const response = await fetch(`${GEMINI_MODEL_URL}?key=${encodeURIComponent(apiKey)}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      contents: [
        {
          role: "user",
          parts: [
            { text: NARANJA_PROMPT },
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
        responseSchema: {
          type: "OBJECT",
          properties: {
            issuer: { type: "STRING", nullable: true },
            holder: { type: "STRING", nullable: true },
            closingDate: { type: "STRING", nullable: true },
            dueDate: { type: "STRING", nullable: true },
            nextDueDate: { type: "STRING", nullable: true },
            totalAmount: { type: "NUMBER", nullable: true },
            minimumPayment: { type: "NUMBER", nullable: true },
            projections: {
              type: "ARRAY",
              items: {
                type: "OBJECT",
                properties: {
                  monthLabel: { type: "STRING", nullable: true },
                  yearMonth: { type: "STRING", nullable: true },
                  amount: { type: "NUMBER", nullable: true }
                }
              }
            },
            installmentsDetail: {
              type: "ARRAY",
              items: {
                type: "OBJECT",
                properties: {
                  merchant: { type: "STRING", nullable: true },
                  installmentNumber: { type: "NUMBER", nullable: true },
                  installmentTotal: { type: "NUMBER", nullable: true },
                  amount: { type: "NUMBER", nullable: true }
                }
              }
            }
          }
        }
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

  const parsed = parseJson<GeminiStatementPayload>(jsonText, "Gemini returned invalid structured JSON.");
  return normalizeGeminiStatement(parsed);
}

function normalizeGeminiStatement(parsed: GeminiStatementPayload): ParsedCardStatementPreview {
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
          const installmentNumber = integerNumber(item.installmentNumber);
          const installmentTotal = integerNumber(item.installmentTotal);
          return {
            merchant: stringOrEmpty(item.merchant),
            installmentNumber,
            installmentTotal,
            amount: currencyNumber(item.amount),
            remainingInstallments: Math.max(installmentTotal - installmentNumber, 0)
          };
        })
        .filter((item) => item.merchant && item.amount > 0 && item.installmentTotal > 1)
    : [];

  const preview: ParsedCardStatementPreview = {
    issuer: "Naranja X",
    brand: "Naranja X",
    bank: "Naranja X",
    holder: stringOrEmpty(parsed.holder),
    closingDate: isoDateOrEmpty(parsed.closingDate),
    dueDate: isoDateOrEmpty(parsed.dueDate),
    nextDueDate: isoDateOrEmpty(parsed.nextDueDate),
    totalAmount: currencyNumber(parsed.totalAmount),
    minimumPayment: currencyNumber(parsed.minimumPayment),
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
