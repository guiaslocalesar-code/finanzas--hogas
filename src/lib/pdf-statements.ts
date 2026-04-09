import { AppError, type ParsedCardStatementPreview } from "./types";

type PdfObject = {
  objectNumber: number;
  body: string;
  streamLength: number | null;
  streamBytes: Uint8Array | null;
};

type PdfEncryption = {
  objectNumber: number;
  revision: number;
  version: number;
  keyLengthBytes: number;
  ownerKey: Uint8Array;
  permissions: number;
  fileId: Uint8Array;
};

const PDF_PASSWORD_PADDING = new Uint8Array([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08, 0x2e, 0x2e,
  0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a
]);

const WINDOWS_1252_DECODER = new TextDecoder("windows-1252");

export async function extractPdfText(pdfBytes: ArrayBuffer): Promise<string> {
  const bytes = new Uint8Array(pdfBytes);
  const source = latin1Decode(bytes);
  const lengths = parseObjectLengths(source);
  const encryption = parseEncryption(source);
  const objects = parsePdfObjects(source, bytes, lengths);
  const chunks: string[] = [];

  for (const object of objects) {
    if (!object.streamBytes || object.streamLength === null || !/\/FlateDecode/.test(object.body)) {
      continue;
    }

    let streamBytes = object.streamBytes;

    if (encryption) {
      streamBytes = decryptPdfObjectStream(streamBytes, object.objectNumber, encryption);
    }

    let inflated: Uint8Array;
    try {
      inflated = await inflateBytes(streamBytes);
    } catch {
      continue;
    }

    const streamText = latin1Decode(inflated);
    const extracted = extractTextOperators(streamText);
    const fallbackExtracted = extracted.trim() ? extracted : extractPrintableFragments(streamText);

    if (fallbackExtracted.trim()) {
      chunks.push(fallbackExtracted);
    }
  }

  const result = normalizeExtractedText(chunks.join("\n"));

  if (!result.trim()) {
    throw new AppError("parse-pdf", "No readable text could be extracted from the uploaded PDF.", undefined, 400);
  }

  return result;
}

export function detectIssuerFromText(text: string, fileName = ""):
  | "visa-santander"
  | "visa-bna"
  | "mastercard-bna"
  | "naranja-x" {
  const upperText = compactForDetection(text);
  const upperFileName = fileName.toUpperCase();

  if (upperText.includes("SANTANDER") && upperText.includes("VISA")) {
    return "visa-santander";
  }

  if (upperText.includes("MASTERCARD") && (upperText.includes("BANCONACION") || upperText.includes("BNA"))) {
    return "mastercard-bna";
  }

  if (upperText.includes("VISA") && (upperText.includes("BANCONACION") || upperText.includes("BNA"))) {
    return "visa-bna";
  }

  if (upperText.includes("NARANJAX") || upperText.includes("NARANJA") || upperFileName.includes("NARANJA")) {
    return "naranja-x";
  }

    throw new AppError("detect-issuer", "The uploaded PDF issuer could not be detected.", undefined, 400);
}

export function parseVisaSantanderStatement(text: string): ParsedCardStatementPreview {
  const warnings: string[] = [];
  const holder = detectHolder(
    text,
    [/TITULAR DE CUENTA:\s*([^\n]+)/i, /\b([A-Z][A-Z\s]{8,})\nCAMPILLO/i, /\n([A-Z][A-Z\s]{8,})\s*\nJUAN DEL CAMPILLO/i],
    warnings,
    "holder"
  );
  const closingDate = detectDate(text, [/CIERRE ACTUAL:\s*([^\n]+)/i, /CIERRE\s+(\d{1,2}\s+[A-Za-z]{3}\.?\s+\d{2,4})/i], warnings, "closingDate");
  const dueDate = detectDate(text, [/VENCIMIENTO ACTUAL:\s*([^\n]+)/i, /VENCIMIENTO\s+(\d{1,2}\s+[A-Za-z]{3}\.?\s+\d{2,4})/i], warnings, "dueDate");
  const nextDueDate = detectVisaSantanderNextDueDate(text, warnings);
  const amountAnalysis = detectVisaSantanderAmounts(text);
  const totalAmount = amountAnalysis.totalAmount.value;
  const minimumPayment = amountAnalysis.minimumPayment.value;
  const projectionAnalysis = parseBnaStyleProjections(text);
  const projections = projectionAnalysis.projections;
  const installmentAnalysis = parseVisaInstallmentDetail(text);

  if (totalAmount <= 0) {
    warnings.push('Field "totalAmount" could not be detected.');
  }

  if (minimumPayment <= 0) {
    warnings.push('Field "minimumPayment" could not be detected.');
  }

  if (projections.length === 0) {
    warnings.push("No se detectó una tabla confiable de cuotas futuras en el resumen Visa Santander.");
  }

  return {
    issuer: "VISA Santander",
    brand: "VISA",
    bank: "Santander Rio",
    holder,
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount,
    minimumPayment,
    projections,
    installmentsDetail: installmentAnalysis.installments,
    rawDetectedData: buildParserDebug(
      "parseVisaSantanderStatement",
      text,
      warnings,
      { holder, closingDate, dueDate, nextDueDate, totalAmount, minimumPayment },
      projections,
      projectionAnalysis.candidateBlocks,
      {
        amountDetection: {
          totalAmount: amountAnalysis.totalAmount,
          minimumPayment: amountAnalysis.minimumPayment
        },
        installmentsDetailFound: installmentAnalysis.installments.length,
        installmentCandidateLines: installmentAnalysis.candidateLines
      }
    ),
    warnings
  };
}

export function parseVisaBnaStatement(text: string): ParsedCardStatementPreview {
  const warnings: string[] = [];
  const holder = detectHolder(text, [/TITULAR DE CUENTA:\s*([^\n]+)/i], warnings, "holder");
  const closingDateDetection = detectBnaLabeledDate(text, "CIERRE ACTUAL");
  const dueDateDetection = detectBnaLabeledDate(text, "VENCIMIENTO ACTUAL");
  const closingDate = closingDateDetection.date;
  const dueDate = dueDateDetection.date;
  warnMissingDate(warnings, "closingDate", closingDate);
  warnMissingDate(warnings, "dueDate", dueDate);
  const nextDueDate = detectDate(text, [/PROXIMO VTO\.?\s*([^\n]+)/i], warnings, "nextDueDate");
  const totalAmount = detectAmount(text, [/SALDO ACTUAL:\s*\$?\s*([\d\.\,]+)/i], warnings, "totalAmount");
  const minimumPayment = detectAmount(text, [/PAGO M.{0,4}NIMO:\s*\$?\s*([\d\.\,]+)/i], warnings, "minimumPayment");
  const projectionAnalysis = parseBnaStyleProjections(text);
  const projections = projectionAnalysis.projections;
  const installmentAnalysis = parseVisaBnaInstallmentDetail(text);

  if (projections.length === 0) {
    warnings.push("No se detectó una tabla confiable de cuotas futuras en el resumen Visa BNA.");
  }

  return {
    issuer: "VISA BNA",
    brand: "VISA",
    bank: "Banco Nacion Argentina",
    holder,
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount,
    minimumPayment,
    projections,
    installmentsDetail: installmentAnalysis.installments,
    rawDetectedData: buildParserDebug(
      "parseVisaBnaStatement",
      text,
      warnings,
      { holder, closingDate, dueDate, nextDueDate, totalAmount, minimumPayment },
      projections,
      projectionAnalysis.candidateBlocks,
      {
        dateDetection: {
          closingDate: closingDateDetection,
          dueDate: dueDateDetection
        },
        installmentsDetailFound: installmentAnalysis.installments.length,
        installmentCandidateLines: installmentAnalysis.candidateLines
      }
    ),
    warnings
  };
}

export function parseMastercardBnaStatement(text: string): ParsedCardStatementPreview {
  const warnings: string[] = [];
  const holder = detectHolder(
    text,
    [/TOTAL TITULAR\s+([A-Z][A-Z\s]{6,})/i, /CUIT Entidad:[\s\S]{0,120}\n([A-Z][A-Z\s]{8,})\nCAMPILLO/i, /\n([A-Z][A-Z\s]{8,})\nCAMPILLO/i],
    warnings,
    "holder"
  );
  const closingDate = detectDate(text, [/Estado de cuenta al:\s*([^\n]+)/i], warnings, "closingDate");
  const dueDate = detectDate(text, [/Vencimiento actual:\s*([^\n]+)/i], warnings, "dueDate");
  const nextDueDate = detectDate(text, [/Pr[\s\S]{0,12}?ximo Vencimiento:\s*([^\n]+)/i], warnings, "nextDueDate");
  const totalAmount = detectAmount(text, [/Saldo actual:\s*\$?\s*([\d\.\,]+)/i], warnings, "totalAmount");
  const minimumPayment = detectAmount(text, [/Pago[\s\S]{0,16}?nimo:\s*\$?\s*([\d\.\,]+)/i], warnings, "minimumPayment");
  const projectionAnalysis = parseMastercardBnaProjections(text, dueDate);
  const projections = projectionAnalysis.projections;
  const installmentAnalysis = parseMastercardInstallmentDetail(text);

  if (projections.length === 0) {
    warnings.push("No se detectó una tabla confiable de cuotas futuras en el resumen Mastercard BNA.");
  }

  return {
    issuer: "Mastercard BNA",
    brand: "Mastercard",
    bank: "Banco Nacion Argentina",
    holder,
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount,
    minimumPayment,
    projections,
    installmentsDetail: installmentAnalysis.installments,
    rawDetectedData: buildParserDebug(
      "parseMastercardBnaStatement",
      text,
      warnings,
      { holder, closingDate, dueDate, nextDueDate, totalAmount, minimumPayment },
      projections,
      projectionAnalysis.candidateBlocks,
      {
        installmentsDetailFound: installmentAnalysis.installments.length,
        installmentCandidateLines: installmentAnalysis.candidateLines
      }
    ),
    warnings
  };
}

export function parseNaranjaXStatement(text: string): ParsedCardStatementPreview {
  const warnings: string[] = [];
  const holder = detectHolder(text, [/NARANJA(?:\s+X)?\s+([A-Z][A-Z\s]{6,})/i, /^([A-Z][A-Z\s]{8,})$/m], warnings, "holder");
  const closingDate = detectDate(text, [/El resumen actual cerr[oó] el\s*([0-9]{2}\/[0-9]{2})/i], warnings, "closingDate", {
    inferYearFrom: detectDate(text, [/y vence el\s*([0-9]{2}\/[0-9]{2}\/[0-9]{2,4})/i], [], "dueDate")
  });
  const dueDate = detectDate(text, [/y vence el\s*([0-9]{2}\/[0-9]{2}\/[0-9]{2,4})/i], warnings, "dueDate");
  const nextDueDate = detectDate(text, [/y vence el\s*([0-9]{2}\/[0-9]{2})\./i], warnings, "nextDueDate", {
    fromAfterKeyword: "El próximo resumen cierra",
    inferYearFrom: closingDate
  });
  const totalAmount = detectAmount(text, [/Tu total a pagar es \$\s*([\d\.\,]+)/i, /^Total\s*\$?\s*([\d\.\,]+)/im], warnings, "totalAmount");
  const minimumPayment = detectAmount(
    text,
    [/PAGO MINIMO:\s*\$?\s*([\d\.\,]+)/i, /LA MENOR ENTREGA\s*\$?\s*([\d\.\,]+)/i],
    warnings,
    "minimumPayment"
  );
  const projectionAnalysis = parseNaranjaProjections(text, dueDate);
  const projections = projectionAnalysis.projections;

  if (projections.length === 0) {
    warnings.push("No monthly projections were detected in the Naranja X statement.");
  }

  return {
    issuer: "Naranja X",
    brand: "Naranja X",
    bank: "Naranja X",
    holder,
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount,
    minimumPayment,
    projections,
    installmentsDetail: [],
    rawDetectedData: buildParserDebug("parseNaranjaXStatement", text, warnings, { holder, closingDate, dueDate, nextDueDate, totalAmount, minimumPayment }, projections, projectionAnalysis.candidateBlocks),
    warnings
  };
}

function parseNaranjaXStatementV2(text: string): ParsedCardStatementPreview {
  const warnings: string[] = [];
  const holder = detectHolder(text, [/NARANJA(?:\s+X)?\s+([A-Z][A-Z\s]{6,})/i, /^([A-Z][A-Z\s]{8,})$/m], warnings, "holder");
  const closingDate = detectDate(
    text,
    [/El resumen actual cerr[oÃ³]\s+el\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i, /cierre(?: del resumen)?\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i],
    warnings,
    "closingDate",
    {
      inferYearFrom: detectDate(text, [/y vence el\s*([0-9]{2}\/[0-9]{2}\/[0-9]{2,4})/i], [], "dueDate")
    }
  );
  const dueDate = detectDate(
    text,
    [/Tu total a pagar es[\s\S]{0,80}?vence(?: el)?\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i, /vencimiento(?: del resumen)?\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i, /vence(?: el)?\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i],
    warnings,
    "dueDate"
  );
  const nextDueDate = detectDate(
    text,
    [/El pr.{0,4}ximo resumen cierra[\s\S]{0,80}?vence(?: el)?\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i, /pr.{0,4}ximo vencimiento\s*([0-9]{2}\/[0-9]{2}(?:\/[0-9]{2,4})?)/i],
    warnings,
    "nextDueDate",
    {
      fromAfterKeyword: "El pr",
      inferYearFrom: closingDate
    }
  );
  const totalAmount = detectAmount(
    text,
    [/Tu total a pagar es \$\s*([\d\.\,]+)/i, /total a pagar\s*\$?\s*([\d\.\,]+)/i, /^Total\s*\$?\s*([\d\.\,]+)/im, /saldo actual\s*\$?\s*([\d\.\,]+)/i],
    warnings,
    "totalAmount"
  );
  const minimumPayment = detectAmount(
    text,
    [/PAGO M.{0,10}NIMO:\s*\$?\s*([\d\.\,]+)/i, /LA MENOR ENTREGA\s*\$?\s*([\d\.\,]+)/i, /pago para no generar intereses\s*\$?\s*([\d\.\,]+)/i],
    warnings,
    "minimumPayment"
  );
  const projectionAnalysis = parseNaranjaProjections(text, dueDate);
  const projections = projectionAnalysis.projections;

  if (projections.length === 0) {
    warnings.push("No se detectó una tabla confiable de cuotas futuras en el resumen Naranja X.");
  }

  return {
    issuer: "Naranja X",
    brand: "Naranja X",
    bank: "Naranja X",
    holder,
    closingDate,
    dueDate,
    nextDueDate,
    totalAmount,
    minimumPayment,
    projections,
    installmentsDetail: [],
    rawDetectedData: buildParserDebug("parseNaranjaXStatementV2", text, warnings, { holder, closingDate, dueDate, nextDueDate, totalAmount, minimumPayment }, projections, projectionAnalysis.candidateBlocks),
    warnings
  };
}

export function parseDetectedStatement(text: string, fileName = ""): ParsedCardStatementPreview {
  try {
    switch (detectIssuerFromText(text, fileName)) {
      case "visa-santander":
        return normalizeParsedPreview(parseVisaSantanderStatement(text));
      case "visa-bna":
        return normalizeParsedPreview(parseVisaBnaStatement(text));
      case "mastercard-bna":
        return normalizeParsedPreview(parseMastercardBnaStatement(text));
      case "naranja-x":
        return normalizeParsedPreview(parseNaranjaXStatementV2(text));
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown parser error";
    return normalizeParsedPreview({
      issuer: "Desconocido",
      brand: "",
      bank: "",
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
        compactTextSample: compactForDetection(text).slice(0, 200)
      },
      warnings: [
        "No se pudo detectar el emisor automaticamente.",
        message,
        "Podes confirmar y completar los datos manualmente."
      ]
    });
  }
}

export type ParserSmokeResult = {
  parser: string;
  ok: boolean;
  warnings: string[];
  checks: Record<string, boolean>;
};

export function runParserSmokeTests(samples: {
  visaSantander?: string;
  visaBna?: string;
  mastercardBna?: string;
  naranjaX?: string;
}): ParserSmokeResult[] {
  const results: ParserSmokeResult[] = [];

  const run = (
    parser: string,
    sample: string | undefined,
    fn: (text: string) => ParsedCardStatementPreview
  ): void => {
    if (!sample?.trim()) {
      results.push({
        parser,
        ok: false,
        warnings: ["Missing sample input."],
        checks: {}
      });
      return;
    }

    const parsed = normalizeParsedPreview(fn(sample));
    const checks: Record<string, boolean> = {
      issuer: parsed.issuer.trim() !== "",
      dueDate: parsed.dueDate === "" || isIsoDate(parsed.dueDate),
      totalAmount: parsed.totalAmount >= 0,
      minimumPayment: parsed.minimumPayment >= 0,
      projections: parsed.projections.every((projection) => isYearMonth(projection.yearMonth) && projection.amount >= 0)
    };

    results.push({
      parser,
      ok: Object.values(checks).every(Boolean),
      warnings: parsed.warnings,
      checks
    });
  };

  run("visa-santander", samples.visaSantander, parseVisaSantanderStatement);
  run("visa-bna", samples.visaBna, parseVisaBnaStatement);
  run("mastercard-bna", samples.mastercardBna, parseMastercardBnaStatement);
  run("naranja-x", samples.naranjaX, parseNaranjaXStatement);

  return results;
}

function parsePdfObjects(source: string, bytes: Uint8Array, lengths: Map<number, number>): PdfObject[] {
  const objects: PdfObject[] = [];
  const regex = /(\d+)\s+0\s+obj\b([\s\S]*?)endobj/g;

  for (const match of source.matchAll(regex)) {
    const objectNumber = Number(match[1]);
    const body = match[2] ?? "";
    const hasStream = body.includes("stream");
    let streamLength: number | null = null;
    let streamBytes: Uint8Array | null = null;

    if (hasStream) {
      streamLength = detectStreamLength(body, lengths);

      if (typeof streamLength === "number") {
        const absoluteStreamIndex = source.indexOf("stream", match.index ?? 0);
        let start = absoluteStreamIndex + "stream".length;

        if (bytes[start] === 13 && bytes[start + 1] === 10) {
          start += 2;
        } else if (bytes[start] === 10 || bytes[start] === 13) {
          start += 1;
        }

        streamBytes = bytes.slice(start, start + streamLength);
      }
    }

    objects.push({ objectNumber, body, streamLength, streamBytes });
  }

  return objects;
}

function parseObjectLengths(source: string): Map<number, number> {
  const lengths = new Map<number, number>();
  const regex = /(\d+)\s+0\s+obj\b\s*(\d+)\s*endobj/g;

  for (const match of source.matchAll(regex)) {
    lengths.set(Number(match[1]), Number(match[2]));
  }

  return lengths;
}

function detectStreamLength(body: string, lengths: Map<number, number>): number | null {
  const direct = body.match(/\/Length\s+(\d+)/);
  if (direct) {
    return Number(direct[1]);
  }

  const ref = body.match(/\/Length\s+(\d+)\s+0\s+R/);
  if (ref) {
    return lengths.get(Number(ref[1])) ?? null;
  }

  return null;
}

function parseEncryption(source: string): PdfEncryption | null {
  const trailerEncrypt = source.match(/\/Encrypt\s+(\d+)\s+0\s+R/);
  if (!trailerEncrypt) {
    return null;
  }

  const objectNumber = Number(trailerEncrypt[1]);
  const objectRegex = new RegExp(`${objectNumber}\\s+0\\s+obj\\b([\\s\\S]*?)endobj`);
  const objectMatch = source.match(objectRegex);

  if (!objectMatch) {
    return null;
  }

  const body = objectMatch[1];
  const ownerMatch = body.match(/\/O\s+\(([^)]*)\)/s);
  const permissionsMatch = body.match(/\/P\s+(-?\d+)/);
  const lengthMatch = body.match(/\/Length\s+(\d+)/);
  const revisionMatch = body.match(/\/R\s+(\d+)/);
  const versionMatch = body.match(/\/V\s+(\d+)/);
  const idMatch = source.match(/\/ID\s*\[\s*<([0-9A-Fa-f]+)>/);

  if (!ownerMatch || !permissionsMatch || !lengthMatch || !revisionMatch || !versionMatch || !idMatch) {
    return null;
  }

  return {
    objectNumber,
    revision: Number(revisionMatch[1]),
    version: Number(versionMatch[1]),
    keyLengthBytes: Number(lengthMatch[1]) / 8,
    ownerKey: parsePdfLiteralBytes(ownerMatch[1]),
    permissions: Number(permissionsMatch[1]),
    fileId: hexToBytes(idMatch[1])
  };
}

function decryptPdfObjectStream(streamBytes: Uint8Array, objectNumber: number, encryption: PdfEncryption): Uint8Array {
  if (encryption.version !== 2 || encryption.revision !== 3) {
    throw new AppError("parse-pdf", "The uploaded encrypted PDF uses an unsupported security version.", undefined, 400);
  }

  const fileKey = computePdfFileEncryptionKey(encryption);
  const objectKeySource = new Uint8Array(fileKey.length + 5);
  objectKeySource.set(fileKey, 0);
  objectKeySource.set([objectNumber & 0xff, (objectNumber >> 8) & 0xff, (objectNumber >> 16) & 0xff, 0x00, 0x00], fileKey.length);
  const objectKey = md5Bytes(objectKeySource).slice(0, Math.min(fileKey.length + 5, 16));

  return rc4Bytes(objectKey, streamBytes);
}

function computePdfFileEncryptionKey(encryption: PdfEncryption): Uint8Array {
  const password = PDF_PASSWORD_PADDING.slice(0, 32);
  const permissions = new Uint8Array([
    encryption.permissions & 0xff,
    (encryption.permissions >> 8) & 0xff,
    (encryption.permissions >> 16) & 0xff,
    (encryption.permissions >> 24) & 0xff
  ]);
  let digest = md5Bytes(concatBytes(password, encryption.ownerKey, permissions, encryption.fileId)).slice(
    0,
    encryption.keyLengthBytes
  );

  for (let index = 0; index < 50; index += 1) {
    digest = md5Bytes(digest).slice(0, encryption.keyLengthBytes);
  }

  return digest;
}

async function inflateBytes(input: Uint8Array): Promise<Uint8Array> {
  const sourceBuffer = new ArrayBuffer(input.byteLength);
  new Uint8Array(sourceBuffer).set(input);
  const stream = new Blob([sourceBuffer]).stream().pipeThrough(new DecompressionStream("deflate"));
  const inflatedBuffer = await new Response(stream).arrayBuffer();
  return new Uint8Array(inflatedBuffer);
}

function extractTextOperators(streamText: string): string {
  const chunks: string[] = [];

  for (const match of streamText.matchAll(/(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]+>)\s*Tj/g)) {
    chunks.push(decodePdfTextToken(match[1]));
  }

  for (const match of streamText.matchAll(/\[((?:.|\n)*?)\]\s*TJ/g)) {
    for (const inner of match[1].matchAll(/(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]+>)/g)) {
      chunks.push(decodePdfTextToken(inner[1]));
    }
  }

  for (const match of streamText.matchAll(/(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]+>)\s*['"]/g)) {
    chunks.push(decodePdfTextToken(match[1]));
  }

  return chunks.join("\n");
}

function decodePdfTextToken(token: string): string {
  if (token.startsWith("<")) {
    return decodePdfHexStringToken(token);
  }

  return decodePdfStringToken(token);
}

function decodePdfStringToken(token: string): string {
  const inner = token.replace(/\)\s*Tj$|\)\s*$/, "").replace(/^\(/, "");
  return WINDOWS_1252_DECODER.decode(parsePdfLiteralBytes(inner));
}

function decodePdfHexStringToken(token: string): string {
  const hex = token.replace(/[<>\s]/g, "");

  if (!hex || hex.length < 2) {
    return "";
  }

  const evenHex = hex.length % 2 === 0 ? hex : `${hex}0`;
  const bytes = hexToBytes(evenHex);

  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be").decode(bytes.slice(2));
  }

  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes.slice(2));
  }

  const zeroPrefixedChars = bytes.reduce((count, _, index) => count + (index % 2 === 0 && bytes[index] === 0 ? 1 : 0), 0);
  if (bytes.length >= 4 && zeroPrefixedChars >= Math.floor(bytes.length / 4)) {
    try {
      return new TextDecoder("utf-16be").decode(bytes);
    } catch {
      return WINDOWS_1252_DECODER.decode(bytes);
    }
  }

  return WINDOWS_1252_DECODER.decode(bytes);
}

function parsePdfLiteralBytes(value: string): Uint8Array {
  const bytes: number[] = [];

  for (let index = 0; index < value.length; index += 1) {
    const char = value[index];

    if (char !== "\\") {
      bytes.push(char.charCodeAt(0));
      continue;
    }

    const next = value[index + 1] ?? "";
    index += 1;

    if (/[0-7]/.test(next)) {
      let octal = next;
      for (let extra = 0; extra < 2; extra += 1) {
        const candidate = value[index + 1] ?? "";
        if (!/[0-7]/.test(candidate)) {
          break;
        }
        index += 1;
        octal += candidate;
      }
      bytes.push(parseInt(octal, 8));
      continue;
    }

    const escapes: Record<string, number> = {
      n: 10,
      r: 13,
      t: 9,
      b: 8,
      f: 12,
      "\\": 92,
      "(": 40,
      ")": 41
    };

    bytes.push(escapes[next] ?? next.charCodeAt(0));
  }

  return new Uint8Array(bytes);
}

function normalizeExtractedText(text: string): string {
  return text
    .replace(/\u0000/g, "")
    .replace(/\r/g, "\n")
    .replace(/[^\S\n]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function extractPrintableFragments(streamText: string): string {
  return Array.from(streamText.matchAll(/[A-Za-zÁÉÍÓÚÑáéíóú0-9$\/.,:%\- ]{6,}/g))
    .map((match) => collapseSpaces(match[0] ?? ""))
    .filter((fragment) => /[A-Za-zÁÉÍÓÚÑáéíóú]{3,}/.test(fragment))
    .join("\n");
}

function detectHolder(text: string, patterns: RegExp[], warnings: string[], field: string): string {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) {
      return collapseSpaces(match[1]);
    }
  }

  warnings.push(`Field "${field}" could not be detected.`);
  return "";
}

function detectDate(
  text: string,
  patterns: RegExp[],
  warnings: string[],
  field: string,
  options?: { inferYearFrom?: string; fromAfterKeyword?: string }
): string {
  const baseText = text.replace(/[\u0000-\u001f]+/g, " ");
  const candidateText =
    options?.fromAfterKeyword && text.includes(options.fromAfterKeyword)
      ? baseText.slice(baseText.indexOf(options.fromAfterKeyword))
      : baseText;

  for (const pattern of patterns) {
    const match = candidateText.match(pattern);
    if (match?.[1]) {
      const parsed = parseLooseDate(match[1], options?.inferYearFrom);
      if (parsed) {
        return parsed;
      }
    }
  }

  warnings.push(`Field "${field}" could not be detected.`);
  return "";
}

type LabeledDateDetection = {
  date: string;
  label: string;
  rawCandidate: string;
  matchedRule: string;
};

function detectBnaLabeledDate(text: string, label: "VENCIMIENTO ACTUAL" | "CIERRE ACTUAL"): LabeledDateDetection {
  const normalizedText = text.replace(/\r/g, "\n");
  const labelPattern = label.replace(/\s+/g, "\\s+");
  const spanishDatePattern = "(\\d{1,2}\\s+[A-Za-zÃ±Ã‘Ã¡Ã©Ã­Ã³ÃºÃÃ‰ÃÃ“Ãš]{3,12}\\.?\\s+\\d{2,4})";
  const patterns = [
    {
      rule: "label-then-spanish-date-with-newlines",
      pattern: new RegExp(`${labelPattern}:?[\\s\\S]{0,120}?${spanishDatePattern}`, "i")
    },
    {
      rule: "label-then-spanish-date-same-line",
      pattern: new RegExp(`${labelPattern}:?\\s*${spanishDatePattern}`, "i")
    }
  ];

  for (const { rule, pattern } of patterns) {
    const match = normalizedText.match(pattern);
    const rawCandidate = match?.[1] ?? "";
    if (!rawCandidate) {
      continue;
    }

    const date = parseLooseDate(rawCandidate);
    if (date) {
      return {
        date,
        label,
        rawCandidate: collapseSpaces(rawCandidate),
        matchedRule: rule
      };
    }
  }

  return {
    date: "",
    label,
    rawCandidate: "",
    matchedRule: "not-found"
  };
}

function warnMissingDate(warnings: string[], field: string, value: string): void {
  if (!value) {
    warnings.push(`Field "${field}" could not be detected.`);
  }
}

function detectAmount(text: string, patterns: RegExp[], warnings: string[], field: string): number {
  const candidateText = text.replace(/[\u0000-\u001f]+/g, " ");
  for (const pattern of patterns) {
    const match = candidateText.match(pattern);
    if (match?.[1]) {
      return parseAmount(match[1]);
    }
  }

  warnings.push(`Field "${field}" could not be detected.`);
  return 0;
}

type AmountDetectionCandidate = {
  rule: string;
  value: number;
  snippet: string;
};

type AmountDetectionResult = {
  value: number;
  matchedRule: string;
  candidates: AmountDetectionCandidate[];
};

function detectVisaSantanderAmounts(text: string): {
  totalAmount: AmountDetectionResult;
  minimumPayment: AmountDetectionResult;
} {
  const normalizedText = text.replace(/[\u0000-\u001f]+/g, " ");
  const totalCandidates = dedupeAmountCandidates([
    ...collectRegexAmountCandidates(normalizedText, [
      {
        rule: "debitaremos-cuenta-corriente",
        pattern: /DEBITAREMOS[\s\S]{0,120}?LA SUMA DE\s*\$\s*([\d\.\,]+)/i
      },
      {
        rule: "saldo-actual-inline",
        pattern: /SALDO ACTUAL[\s\S]{0,40}?\$\s*([\d\.\,]+)/i
      },
      {
        rule: "saldo-actual-columna",
        pattern: /Saldo actual:\s*\$\s*([\d\.\,]+)/i
      }
    ]),
    ...collectNearbyLabelAmountCandidates(normalizedText, /SALDO ACTUAL/i, "saldo-actual-window", "max-before")
  ]);
  const minimumCandidates = dedupeAmountCandidates([
    ...collectRegexAmountCandidates(normalizedText, [
      {
        rule: "plan-v-pago-minimo",
        pattern: /Plan V:[\s\S]{0,120}?pago m.{0,4}nimo de\s*\$\s*([\d\.\,]+)/i
      },
      {
        rule: "pago-minimo-inline",
        pattern: /PAGO M.{0,4}NIMO[\s\S]{0,40}?\$\s*([\d\.\,]+)/i
      },
      {
        rule: "pago-minimo-plan-texto",
        pattern: /pago m.{0,4}nimo de\s*\$\s*([\d\.\,]+)/i
      }
    ]),
    ...collectNearbyLabelAmountCandidates(normalizedText, /PAGO M.{0,4}NIMO/i, "pago-minimo-window", "closest-before")
  ]);

  return {
    totalAmount: selectAmountDetection(totalCandidates),
    minimumPayment: selectAmountDetection(minimumCandidates)
  };
}

function detectVisaSantanderNextDueDate(text: string, warnings: string[]): string {
  const normalizedText = text.replace(/[\u0000-\u001f]+/g, " ");
  const match =
    normalizedText.match(/Prox\.?\s*Vto\.?:\s*([0-9]{1,2}\s+[A-Za-z]{3}\s+[0-9]{2,4})/i) ??
    normalizedText.match(/Pr[oó]ximo\s+Vto\.?:\s*([0-9]{1,2}\s+[A-Za-z]{3}\s+[0-9]{2,4})/i);
  if (match?.[1]) {
    const parsed = parseLooseDate(match[1]);
    if (parsed) {
      return parsed;
    }
  }

  warnings.push('Field "nextDueDate" could not be detected.');
  return "";
}

function collectRegexAmountCandidates(
  text: string,
  rules: Array<{ rule: string; pattern: RegExp }>
): AmountDetectionCandidate[] {
  const candidates: AmountDetectionCandidate[] = [];

  for (const rule of rules) {
    const match = text.match(rule.pattern);
    if (!match?.[1]) {
      continue;
    }

    const value = parseAmount(match[1]);
    if (value <= 0) {
      continue;
    }

    candidates.push({
      rule: rule.rule,
      value,
      snippet: collapseSpaces(match[0]).slice(0, 180)
    });
  }

  return candidates;
}

function collectNearbyLabelAmountCandidates(
  text: string,
  labelPattern: RegExp,
  rule: string,
  strategy: "max-before" | "closest-before"
): AmountDetectionCandidate[] {
  const match = text.match(labelPattern);
  if (!match || match.index == null) {
    return [];
  }

  const labelIndex = match.index;
  const start = Math.max(0, labelIndex - 220);
  const end = Math.min(text.length, labelIndex + 220);
  const windowText = text.slice(start, end);
  const labelOffset = labelIndex - start;
  const parsed: Array<{ value: number; snippet: string; distance: number; isBefore: boolean }> = [];

  for (const amountMatch of windowText.matchAll(AMOUNT_REGEX)) {
    const rawValue = amountMatch[1];
    if (!rawValue) {
      continue;
    }

    const value = parseAmount(rawValue);
    if (value <= 0) {
      continue;
    }

    const amountIndex = amountMatch.index ?? 0;
    parsed.push({
      value,
      snippet: collapseSpaces(windowText.slice(Math.max(0, amountIndex - 36), Math.min(windowText.length, amountIndex + 56))),
      distance: Math.abs(labelOffset - amountIndex),
      isBefore: amountIndex <= labelOffset
    });
  }

  const before = parsed.filter((item) => item.isBefore);
  if (before.length === 0) {
    return [];
  }

  const selected =
    strategy === "max-before"
      ? before.reduce((best, item) => (item.value > best.value ? item : best), before[0])
      : before.reduce((best, item) => (item.distance < best.distance ? item : best), before[0]);

  return [
    {
      rule,
      value: selected.value,
      snippet: selected.snippet
    }
  ];
}

function dedupeAmountCandidates(candidates: AmountDetectionCandidate[]): AmountDetectionCandidate[] {
  const seen = new Set<string>();
  const deduped: AmountDetectionCandidate[] = [];

  for (const candidate of candidates) {
    const key = `${candidate.rule}:${candidate.value}:${candidate.snippet}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduped.push(candidate);
  }

  return deduped;
}

function selectAmountDetection(candidates: AmountDetectionCandidate[]): AmountDetectionResult {
  const selected = candidates[0];
  return {
    value: selected?.value ?? 0,
    matchedRule: selected?.rule ?? "",
    candidates
  };
}

function parseBnaStyleProjections(text: string): {
  projections: ParsedCardStatementPreview["projections"];
  candidateBlocks: string[];
} {
  return parseProjectionBlocks(text, [/CUOTAS?\s+A\s+VENCER/i], [/RESUMEN CONSOLIDADO/i, /DETALLE DEL/i, /COMPRAS DEL MES/i]);
}

type InstallmentDetailAnalysis = {
  installments: ParsedCardStatementPreview["installmentsDetail"];
  candidateLines: string[];
};

function parseVisaInstallmentDetail(text: string): InstallmentDetailAnalysis {
  const installments: ParsedCardStatementPreview["installmentsDetail"] = [];
  const candidateLines: string[] = [];

  for (const rawLine of text.split("\n")) {
    const line = collapseSpaces(rawLine);
    const match = line.match(/^(.*?)\bC\.(\d{1,2})\/(\d{1,2})\b\s+([\d\.\,]+)\s*$/i);
    if (!match) {
      continue;
    }

    const installmentNumber = Number(match[2]);
    const installmentTotal = Number(match[3]);
    const amount = parseAmount(match[4] ?? "");
    const merchant = cleanVisaInstallmentMerchant(match[1] ?? "");

    if (!merchant || !isValidInstallment(installmentNumber, installmentTotal, amount)) {
      continue;
    }

    candidateLines.push(line);
    installments.push({
      merchant,
      installmentNumber,
      installmentTotal,
      amount,
      remainingInstallments: Math.max(installmentTotal - installmentNumber, 0)
    });
  }

  for (const rawLine of text.split("\n")) {
    const line = collapseSpaces(rawLine);
    if (!line || /\bC\.\d{1,2}\/\d{1,2}\b/i.test(line) || shouldIgnoreCardConsumption(line)) {
      continue;
    }

    const onePayment = parseVisaOnePaymentLine(line);
    if (!onePayment) {
      continue;
    }

    candidateLines.push(line);
    installments.push(onePayment);
  }

  return {
    installments,
    candidateLines: candidateLines.slice(0, 20)
  };
}

function parseVisaBnaInstallmentDetail(text: string): InstallmentDetailAnalysis {
  const installments: ParsedCardStatementPreview["installmentsDetail"] = [];
  const candidateLines: string[] = [];
  const cleanedText = buildVisaBnaInstallmentSearchText(text);
  const pattern =
    /([A-Z0-9][A-Z0-9\s\.\-&\/]{2,80}?)\s+C\.(\d{1,2})\/(\d{1,2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})/gi;

  for (const match of cleanedText.matchAll(pattern)) {
    const installmentNumber = Number(match[2]);
    const installmentTotal = Number(match[3]);
    const amount = parseAmount(match[4] ?? "");
    const merchant = cleanVisaBnaInstallmentMerchant(match[1] ?? "");

    if (!merchant || !isValidInstallment(installmentNumber, installmentTotal, amount)) {
      continue;
    }

    const candidateLine = collapseSpaces(match[0] ?? "");
    candidateLines.push(candidateLine);
    installments.push({
      merchant,
      installmentNumber,
      installmentTotal,
      amount,
      remainingInstallments: Math.max(installmentTotal - installmentNumber, 0)
    });
  }

  return {
    installments,
    candidateLines: candidateLines.slice(0, 20)
  };
}

function parseMastercardInstallmentDetail(text: string): InstallmentDetailAnalysis {
  const installments: ParsedCardStatementPreview["installmentsDetail"] = [];
  const candidateLines: string[] = [];
  const multilinePattern =
    /(?:^|\n)\s*\d{1,2}[-\/][A-Za-zÁÉÍÓÚáéíóúñÑ]{3,4}\.?(?:[-\/]\d{2,4})?\s*\n\s*(.+?)\s+(\d{1,2})\/(\d{1,2})\s*\n\s*\d+\s*\n\s*([\d\.\,]+)\s*(?=\n|$)/g;

  for (const match of text.matchAll(multilinePattern)) {
    const installmentNumber = Number(match[2]);
    const installmentTotal = Number(match[3]);
    const amount = parseAmount(match[4] ?? "");
    const merchant = cleanMastercardInstallmentMerchant(match[1] ?? "");

    if (!merchant || !isValidInstallment(installmentNumber, installmentTotal, amount)) {
      continue;
    }

    const candidateLine = `${merchant} ${installmentNumber}/${installmentTotal} ${match[4]}`;
    candidateLines.push(candidateLine);
    installments.push({
      merchant,
      installmentNumber,
      installmentTotal,
      amount,
      remainingInstallments: Math.max(installmentTotal - installmentNumber, 0)
    });
  }

  for (const rawLine of text.split("\n")) {
    const line = collapseSpaces(rawLine);
    if (/\bC\.\d{1,2}\/\d{1,2}\b/i.test(line)) {
      continue;
    }

    const match = line.match(/^(.*?)\s+(\d{1,2})\/(\d{1,2})\s+([\d\.\,]+)\s*$/);
    if (!match) {
      continue;
    }

    const installmentNumber = Number(match[2]);
    const installmentTotal = Number(match[3]);
    const amount = parseAmount(match[4] ?? "");
    const merchant = cleanMastercardInstallmentMerchant(match[1] ?? "");

    if (!merchant || !isValidInstallment(installmentNumber, installmentTotal, amount)) {
      continue;
    }

    candidateLines.push(line);
    installments.push({
      merchant,
      installmentNumber,
      installmentTotal,
      amount,
      remainingInstallments: Math.max(installmentTotal - installmentNumber, 0)
    });
  }

  const onePaymentPattern =
    /(?:^|\n)\s*(\d{1,2}[-\/][A-Za-zÁÉÍÓÚáéíóúñÑ]{3,4}\.?(?:[-\/]\d{2,4})?)\s*\n\s*((?!.*\b\d{1,2}\/\d{1,2}\b).+?)\s*\n\s*\d+\s*\n\s*([\d\.\,]+)\s*(?=\n|$)/g;

  for (const match of text.matchAll(onePaymentPattern)) {
    const rawMerchant = match[2] ?? "";
    const merchant = cleanMastercardInstallmentMerchant(rawMerchant);
    const amount = parseAmount(match[3] ?? "");

    if (!merchant || amount <= 0 || shouldIgnoreCardConsumption(merchant)) {
      continue;
    }

    const candidateLine = `${match[1]} ${merchant} 1/1 ${match[3]}`;
    candidateLines.push(candidateLine);
    installments.push({
      purchaseDate: parseCardPurchaseDate(match[1] ?? ""),
      merchant,
      installmentNumber: 1,
      installmentTotal: 1,
      amount,
      remainingInstallments: 0
    });
  }

  return {
    installments,
    candidateLines: candidateLines.slice(0, 20)
  };
}

function buildVisaBnaInstallmentSearchText(text: string): string {
  const lines = text
    .split("\n")
    .map((line) => collapseSpaces(line))
    .filter(Boolean);
  const tokens: string[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const current = lines[index] ?? "";
    const previous = tokens[tokens.length - 1] ?? "";

    if (/^\d{1,2}$/.test(current) && /,\d$/.test(previous)) {
      tokens[tokens.length - 1] = `${previous}${current}`;
      continue;
    }

    if (/^\d$/.test(current) && /,\d$/.test(previous)) {
      tokens[tokens.length - 1] = `${previous}${current}`;
      continue;
    }

    tokens.push(current);
  }

  return collapseSpaces(tokens.join(" "));
}

function cleanVisaInstallmentMerchant(value: string): string {
  return collapseSpaces(
    value
      .replace(/^\d{1,2}\s+/, "")
      .replace(/^\d{3,}\s+/, "")
      .replace(/^\*\s*/, "")
      .replace(/\s+\*+\s*/g, " ")
      .replace(/\b\d{3,}\b/g, "")
  );
}

function cleanVisaBnaInstallmentMerchant(value: string): string {
  const tokens = collapseSpaces(value)
    .replace(/\b\d{1,2}\s+(?:ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)\.?\s+\d{2,4}\b/gi, " ")
    .replace(/\b\d{2,}\b/g, " ")
    .replace(/\s+\*+\s*/g, " ")
    .split(" ")
    .filter(Boolean);
  const tail = tokens.slice(-4).join(" ");

  return collapseSpaces(tail);
}

function cleanMastercardInstallmentMerchant(value: string): string {
  return collapseSpaces(
    value
      .replace(/^\d{1,2}[-\/][A-Za-z]{3}[-\/]\d{2,4}\s+/, "")
      .replace(/^\*\s*/, "")
      .replace(/\s+\*+\s*/g, " ")
  );
}

function parseVisaOnePaymentLine(line: string): ParsedCardStatementPreview["installmentsDetail"][number] | null {
  const match = line.match(
    /^\d{1,2}(?:\s+[A-Za-zÁÉÍÓÚáéíóúñÑ]+\.?\s+\d{2,4})?\s+(?:\d{3,}\s+)?(?:[A-Z]\s+|\*\s+)?(.+?)\s+(USD\s+)?([\d\.\,]+)(?:\s+([\d\.\,]+))?\s*$/i
  );

  if (!match) {
    return null;
  }

  const merchant = cleanOnePaymentMerchant(match[1] ?? "");
  const currency = match[2] ? "USD" : "ARS";
  const amount = parseAmount(match[4] ?? match[3] ?? "");

  if (!merchant || amount <= 0 || shouldIgnoreCardConsumption(merchant)) {
    return null;
  }

  return {
    purchaseDate: parseCardPurchaseDate(line),
    merchant,
    installmentNumber: 1,
    installmentTotal: 1,
    amount,
    currency,
    remainingInstallments: 0
  };
}

function cleanOnePaymentMerchant(value: string): string {
  return collapseSpaces(
    value
      .replace(/^\d{3,}\s+/, "")
      .replace(/^[A-Z]\s+/, "")
      .replace(/^\*\s*/, "")
      .replace(/\s+\*+\s*/g, " ")
      .replace(/\s+USD$/i, "")
  );
}

function shouldIgnoreCardConsumption(value: string): boolean {
  return /SU PAGO|PAGO EN|PAGO MINIMO|TRANSFERENCIA|DEUDA|SALDO|IMPUESTO|SELLOS?|DB IVA|DB\.?\s*RG|IVA\b|IIBB|PERCEP|INTER[EÉ]S|COMISI[OÓ]N|MANTENIMIENTO|CARGO POR|AJUSTE/i.test(
    value
  );
}

function parseCardPurchaseDate(value: string): string {
  const spanishLong = value.match(/\b(\d{1,2})\s+([A-Za-zÁÉÍÓÚáéíóúñÑ]+)\.?\s+(\d{2,4})\b/);
  if (spanishLong) {
    return formatSpanishMonthDate(spanishLong[1] ?? "", spanishLong[2] ?? "", spanishLong[3] ?? "");
  }

  const mastercardShort = value.match(/\b(\d{1,2})[-\/]([A-Za-zÁÉÍÓÚáéíóúñÑ]{3,4})\.?[-\/](\d{2,4})\b/);
  if (mastercardShort) {
    return formatSpanishMonthDate(mastercardShort[1] ?? "", mastercardShort[2] ?? "", mastercardShort[3] ?? "");
  }

  return "";
}

function isValidInstallment(installmentNumber: number, installmentTotal: number, amount: number): boolean {
  return (
    Number.isInteger(installmentNumber) &&
    Number.isInteger(installmentTotal) &&
    installmentNumber >= 1 &&
    installmentTotal >= installmentNumber &&
    installmentTotal >= 1 &&
    amount > 0
  );
}

function parseMastercardBnaProjections(
  text: string,
  dueDate: string
): {
  projections: ParsedCardStatementPreview["projections"];
  candidateBlocks: string[];
} {
  return parseProjectionBlocks(
    text,
    [/CUOTAS?\s+A\s+VENCER/i],
    [/RESUMEN CONSOLIDADO/i, /DETALLE DEL/i, /COMPRAS DEL MES/i],
    dueDate
  );
}

function parseNaranjaProjections(
  text: string,
  dueDate: string
): {
  projections: ParsedCardStatementPreview["projections"];
  candidateBlocks: string[];
} {
  const byBlocks = parseProjectionBlocks(
    text,
    [/CUOTAS?\s+FUTURAS/i, /PLAN DE PAGOS/i, /PROXIMAS?\s+CUOTAS/i, /CUOTAS?\s+A\s+VENCER/i],
    [/MOVIMIENTOS/i, /DETALLE/i, /TOTAL A PAGAR/i],
    dueDate
  );

  if (byBlocks.projections.length > 0) {
    return byBlocks;
  }

  const inferYear = dueDate ? dueDate.slice(0, 4) : "";
  const fallbackMatches = Array.from(
    text.matchAll(
      /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Setiembre|Octubre|Noviembre|Diciembre)(?:\/|-|\s)(\d{2,4})?\s*\$?\s*([\d\.\,]+)/gi
    )
  );
  const projections = dedupeProjections(
    fallbackMatches
      .map((match) => {
        const yearToken = match[2] || inferYear;
        const yearMonth = toYearMonth(match[1], yearToken);
        return {
          monthLabel: `${titleCaseMonth(match[1])}/${normalizeYearToken(yearToken)}`,
          yearMonth,
          amount: parseAmount(match[3])
        };
      })
      .filter((item) => isYearMonth(item.yearMonth) && item.amount > 0)
  );

  return {
    projections,
    candidateBlocks: byBlocks.candidateBlocks
  };
}

function parseProjectionBlocks(
  text: string,
  startPatterns: RegExp[],
  stopPatterns: RegExp[],
  inferYearFromDate = ""
): {
  projections: ParsedCardStatementPreview["projections"];
  candidateBlocks: string[];
} {
  const candidateBlocks = collectProjectionCandidateBlocks(text, startPatterns, stopPatterns);
  const minimumYearMonth = inferYearFromDate ? inferYearFromDate.slice(0, 7) : "";
  const projections = dedupeProjections(
    candidateBlocks.flatMap((block) => parseProjectionPairsFromBlock(block, inferYearFromDate))
  ).filter((item) => item.amount > 0 && (!minimumYearMonth || item.yearMonth >= minimumYearMonth));

  return {
    projections,
    candidateBlocks
  };
}

function collectProjectionCandidateBlocks(text: string, startPatterns: RegExp[], stopPatterns: RegExp[]): string[] {
  const lines = text
    .split(/\n+/)
    .map((line) => collapseSpaces(line))
    .filter(Boolean);
  const blocks: string[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    if (!startPatterns.some((pattern) => pattern.test(lines[index]))) {
      continue;
    }

    const blockLines: string[] = [lines[index]];
    for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
      if (stopPatterns.some((pattern) => pattern.test(lines[cursor]))) {
        break;
      }
      if (blockLines.length >= 14) {
        break;
      }
      blockLines.push(lines[cursor]);
    }

    blocks.push(blockLines.join("\n"));
  }

  return blocks.slice(0, 4);
}

function parseProjectionPairsFromBlock(
  block: string,
  inferYearFromDate: string
): ParsedCardStatementPreview["projections"] {
  const lines = block
    .split(/\n+/)
    .map((line) => collapseSpaces(line))
    .filter(Boolean);
  const projections: ParsedCardStatementPreview["projections"] = [];
  const pendingMonths: Array<{ monthLabel: string; yearMonth: string }> = [];
  const fallbackYear = inferYearFromDate ? inferYearFromDate.slice(0, 4) : String(new Date().getUTCFullYear());

  for (const line of lines) {
    if (/^(pesos|dolares|fecha|concepto|compras|resumen consolidado|detalle del)/i.test(line)) {
      continue;
    }

    const months = parseMonthReferences(line, fallbackYear);
    const amounts = Array.from(line.matchAll(AMOUNT_REGEX))
      .map((match) => parseAmount(match[1]))
      .filter((amount) => amount > 0);

    if (months.length > 0 && amounts.length > 0) {
      for (let index = 0; index < Math.min(months.length, amounts.length); index += 1) {
        projections.push({
          monthLabel: months[index].monthLabel,
          yearMonth: months[index].yearMonth,
          amount: amounts[index]
        });
      }
      if (months.length > amounts.length) {
        pendingMonths.push(...months.slice(amounts.length));
      }
      continue;
    }

    if (months.length > 0) {
      pendingMonths.push(...months);
      continue;
    }

    if (amounts.length > 0 && pendingMonths.length > 0) {
      for (let index = 0; index < Math.min(pendingMonths.length, amounts.length); index += 1) {
        projections.push({
          monthLabel: pendingMonths[index].monthLabel,
          yearMonth: pendingMonths[index].yearMonth,
          amount: amounts[index]
        });
      }
      pendingMonths.splice(0, amounts.length);
    }
  }

  return projections.filter((item) => isYearMonth(item.yearMonth) && item.amount > 0);
}

function parseMonthReferences(line: string, fallbackYear: string): Array<{ monthLabel: string; yearMonth: string }> {
  const results: Array<{ monthLabel: string; yearMonth: string }> = [];

  for (const match of line.matchAll(/(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|SETIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE|ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)(?:\/|-|\s)(\d{2,4})?/gi)) {
    const yearToken = match[2] || fallbackYear;
    const yearMonth = toYearMonth(match[1], yearToken);
    if (!isYearMonth(yearMonth)) {
      continue;
    }
    results.push({
      monthLabel: `${titleCaseMonth(match[1])}/${normalizeYearToken(yearToken)}`,
      yearMonth
    });
  }

  return results;
}

function buildParserDebug(
  parser: string,
  text: string,
  warnings: string[],
  fields: Record<string, string | number>,
  projections: ParsedCardStatementPreview["projections"],
  candidateBlocks: string[],
  extra: Record<string, unknown> = {}
): Record<string, unknown> {
  const fieldsFound = Object.entries(fields)
    .filter(([, value]) => (typeof value === "number" ? value > 0 : String(value).trim() !== ""))
    .map(([key]) => key);
  const missingFields = Object.keys(fields).filter((key) => !fieldsFound.includes(key));

  return {
    parser,
    detectedIssuer: parser.replace(/^parse/i, "").replace(/Statement$/, ""),
    fieldsFound,
    missingFields,
    candidateBlocks: candidateBlocks.map((block) => block.slice(0, 280)),
    projectionsFound: projections.length,
    warningsCount: warnings.length,
    compactTextSample: compactForDetection(text).slice(0, 220),
    ...extra
  };
}

function parseLooseDate(input: string, inferYearFrom?: string): string {
  const cleaned = collapseSpaces(input)
    .replace(/\./g, "")
    .replace(/,/g, "")
    .replace(/ACTUAL/gi, "")
    .trim();

  const slashMatch = cleaned.match(/(\d{2})\/(\d{2})(?:\/(\d{2,4}))?/);
  if (slashMatch) {
    const year = slashMatch[3]
      ? normalizeYearToken(slashMatch[3])
      : inferYearFrom
        ? inferYearFrom.slice(0, 4)
        : String(new Date().getUTCFullYear());
    return `${year}-${slashMatch[2].padStart(2, "0")}-${slashMatch[1].padStart(2, "0")}`;
  }

  const monthMatch = cleaned.match(/(\d{1,2})[-\s]([A-Za-zñÑáéíóúÁÉÍÓÚ]+)[-\s](\d{2,4})/);
  if (monthMatch) {
    const month = monthTokenToNumber(monthMatch[2]);
    if (!month) {
      return "";
    }
    return `${normalizeYearToken(monthMatch[3])}-${month}-${monthMatch[1].padStart(2, "0")}`;
  }

  return "";
}

function parseAmount(value: string): number {
  const normalized = value.replace(/\./g, "").replace(",", ".").replace(/[^\d.-]/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? Math.round(parsed * 100) / 100 : 0;
}

function compactForDetection(text: string): string {
  return text.replace(/[\u0000-\u001f]/g, "").replace(/\s+/g, "").toUpperCase();
}

function collapseSpaces(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeYearToken(value: string): string {
  return value.length === 2 ? `20${value}` : value;
}

function toYearMonth(monthToken: string, yearToken: string): string {
  const month = monthTokenToNumber(monthToken);
  const year = normalizeYearToken(yearToken);
  if (!month || !/^\d{4}$/.test(year)) {
    return "";
  }
  return `${year}-${month}`;
}

function normalizeParsedPreview(parsed: ParsedCardStatementPreview): ParsedCardStatementPreview {
  const projections = dedupeProjections(
    parsed.projections
      .map((item) => ({
        monthLabel: collapseSpaces(item.monthLabel),
        yearMonth: item.yearMonth.trim(),
        amount: Math.max(0, roundCurrency(item.amount))
      }))
      .filter((item) => isYearMonth(item.yearMonth))
  );
  const installmentsDetail = (parsed.installmentsDetail ?? [])
    .map((item) => ({
      purchaseDate: normalizeIsoDate(item.purchaseDate ?? ""),
      merchant: collapseSpaces(item.merchant),
      installmentNumber: Math.max(0, Math.trunc(item.installmentNumber)),
      installmentTotal: Math.max(0, Math.trunc(item.installmentTotal)),
      amount: Math.max(0, roundCurrency(item.amount)),
      remainingInstallments: Math.max(0, Math.trunc(item.remainingInstallments))
    }))
    .filter((item) => item.merchant && isValidInstallment(item.installmentNumber, item.installmentTotal, item.amount));

  return {
    ...parsed,
    closingDate: normalizeIsoDate(parsed.closingDate),
    dueDate: normalizeIsoDate(parsed.dueDate),
    nextDueDate: normalizeIsoDate(parsed.nextDueDate),
    totalAmount: Math.max(0, roundCurrency(parsed.totalAmount)),
    minimumPayment: Math.max(0, roundCurrency(parsed.minimumPayment)),
    projections,
    installmentsDetail,
    warnings: Array.from(new Set(parsed.warnings.map((warning) => collapseSpaces(warning)).filter(Boolean)))
  };
}

function dedupeProjections(items: ParsedCardStatementPreview["projections"]): ParsedCardStatementPreview["projections"] {
  const grouped = new Map<string, ParsedCardStatementPreview["projections"][number]>();

  for (const item of items) {
    const existing = grouped.get(item.yearMonth);
    if (!existing) {
      grouped.set(item.yearMonth, item);
      continue;
    }

    grouped.set(item.yearMonth, {
      ...existing,
      amount: roundCurrency(existing.amount + item.amount),
      monthLabel: existing.monthLabel || item.monthLabel
    });
  }

  return Array.from(grouped.values()).sort((left, right) => left.yearMonth.localeCompare(right.yearMonth));
}

function normalizeIsoDate(value: string): string {
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

function formatSpanishMonthDate(dayToken: string, monthToken: string, yearToken: string): string {
  const normalizedYear = Number(normalizeYearToken(yearToken));
  if (!Number.isFinite(normalizedYear) || normalizedYear < 2020) {
    return "";
  }

  return parseLooseDate(`${dayToken} ${monthToken} ${yearToken}`);
}

function isIsoDate(value: string): boolean {
  return normalizeIsoDate(value) !== "";
}

function isYearMonth(value: string): boolean {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(value);
}

function roundCurrency(value: number): number {
  return Math.round(value * 100) / 100;
}

function titleCaseMonth(monthToken: string): string {
  const normalized = monthToken.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  return normalized.charAt(0).toUpperCase() + normalized.slice(1);
}

function monthTokenToNumber(token: string): string {
  const normalized = token.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const map: Record<string, string> = {
    enero: "01",
    ene: "01",
    febrero: "02",
    feb: "02",
    marzo: "03",
    mar: "03",
    abril: "04",
    abr: "04",
    mayo: "05",
    may: "05",
    junio: "06",
    jun: "06",
    julio: "07",
    jul: "07",
    agosto: "08",
    ago: "08",
    septiembre: "09",
    setiembre: "09",
    sep: "09",
    set: "09",
    octubre: "10",
    oct: "10",
    noviembre: "11",
    nov: "11",
    diciembre: "12",
    dic: "12"
  };

  return map[normalized] ?? "";
}

function latin1Decode(bytes: Uint8Array): string {
  let output = "";

  for (let index = 0; index < bytes.length; index += 1) {
    output += String.fromCharCode(bytes[index] ?? 0);
  }

  return output;
}

function hexToBytes(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length / 2);

  for (let index = 0; index < bytes.length; index += 1) {
    bytes[index] = parseInt(value.slice(index * 2, index * 2 + 2), 16);
  }

  return bytes;
}

function concatBytes(...chunks: Uint8Array[]): Uint8Array {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const output = new Uint8Array(totalLength);
  let offset = 0;

  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }

  return output;
}

function rc4Bytes(key: Uint8Array, data: Uint8Array): Uint8Array {
  const state = new Uint8Array(256);
  for (let index = 0; index < 256; index += 1) {
    state[index] = index;
  }

  let j = 0;
  for (let index = 0; index < 256; index += 1) {
    j = (j + state[index] + key[index % key.length]) & 255;
    [state[index], state[j]] = [state[j], state[index]];
  }

  const output = new Uint8Array(data.length);
  let i = 0;
  j = 0;

  for (let index = 0; index < data.length; index += 1) {
    i = (i + 1) & 255;
    j = (j + state[i]) & 255;
    [state[i], state[j]] = [state[j], state[i]];
    const keyStream = state[(state[i] + state[j]) & 255];
    output[index] = data[index] ^ keyStream;
  }

  return output;
}

function md5Bytes(input: Uint8Array): Uint8Array {
  const originalLength = input.length;
  const withPaddingLength = (((originalLength + 8) >> 6) + 1) << 6;
  const buffer = new Uint8Array(withPaddingLength);
  buffer.set(input);
  buffer[originalLength] = 0x80;

  const bitLength = originalLength * 8;
  for (let index = 0; index < 8; index += 1) {
    buffer[withPaddingLength - 8 + index] = (bitLength >>> (index * 8)) & 0xff;
  }

  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;

  for (let offset = 0; offset < buffer.length; offset += 64) {
    const words = new Uint32Array(16);
    for (let index = 0; index < 16; index += 1) {
      const base = offset + index * 4;
      words[index] =
        buffer[base] |
        (buffer[base + 1] << 8) |
        (buffer[base + 2] << 16) |
        (buffer[base + 3] << 24);
    }

    let a = a0;
    let b = b0;
    let c = c0;
    let d = d0;

    for (let index = 0; index < 64; index += 1) {
      let f = 0;
      let g = 0;

      if (index < 16) {
        f = (b & c) | (~b & d);
        g = index;
      } else if (index < 32) {
        f = (d & b) | (~d & c);
        g = (5 * index + 1) % 16;
      } else if (index < 48) {
        f = b ^ c ^ d;
        g = (3 * index + 5) % 16;
      } else {
        f = c ^ (b | ~d);
        g = (7 * index) % 16;
      }

      const temp = d;
      d = c;
      c = b;
      b = add32(
        b,
        rotateLeft(add32(add32(a, f), add32(MD5_K[index], words[g])), MD5_S[index])
      );
      a = temp;
    }

    a0 = add32(a0, a);
    b0 = add32(b0, b);
    c0 = add32(c0, c);
    d0 = add32(d0, d);
  }

  const output = new Uint8Array(16);
  const state = [a0, b0, c0, d0];
  for (let index = 0; index < state.length; index += 1) {
    output[index * 4] = state[index] & 0xff;
    output[index * 4 + 1] = (state[index] >>> 8) & 0xff;
    output[index * 4 + 2] = (state[index] >>> 16) & 0xff;
    output[index * 4 + 3] = (state[index] >>> 24) & 0xff;
  }

  return output;
}

function add32(left: number, right: number): number {
  return (left + right) >>> 0;
}

function rotateLeft(value: number, shift: number): number {
  return ((value << shift) | (value >>> (32 - shift))) >>> 0;
}

const MONTH_TOKEN_REGEX =
  /(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|SETIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE|ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[\/-](\d{2,4})/g;

const AMOUNT_REGEX = /\$?(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})/g;

const MD5_S = new Uint8Array([
  7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14,
  20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10,
  15, 21, 6, 10, 15, 21, 6, 10, 15, 21
]);

const MD5_K = new Uint32Array([
  0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee, 0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
  0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be, 0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
  0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa, 0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
  0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed, 0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
  0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c, 0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
  0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05, 0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
  0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039, 0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
  0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1, 0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391
]);
