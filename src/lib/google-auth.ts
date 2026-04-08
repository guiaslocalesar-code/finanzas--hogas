import { AppError, type Env, type GoogleTokenResponse } from "./types";

const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

type TokenCacheEntry = {
  accessToken: string;
  expiresAt: number;
  serviceAccountEmail: string;
};

let tokenCache: TokenCacheEntry | null = null;

export async function getGoogleAccessToken(env: Env): Promise<string> {
  validateGoogleEnv(env);
  const now = Math.floor(Date.now() / 1000);

  if (
    tokenCache &&
    tokenCache.serviceAccountEmail === env.SERVICE_ACCOUNT_EMAIL &&
    tokenCache.expiresAt - 60 > now
  ) {
    console.log("[google-auth] Using cached access token", {
      serviceAccountEmail: env.SERVICE_ACCOUNT_EMAIL
    });
    return tokenCache.accessToken;
  }

  try {
    const assertion = await createServiceAccountJwt(env);
    const body = new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion
    });

    console.log("[google-auth] Requesting Google access token", {
      serviceAccountEmail: env.SERVICE_ACCOUNT_EMAIL,
      tokenUrl: GOOGLE_TOKEN_URL
    });

    const response = await fetch(GOOGLE_TOKEN_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: body.toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("[google-auth] Google token request failed", {
        status: response.status,
        body: errorText
      });
      throw new AppError(
        "google-token",
        "Google OAuth token request failed.",
        `status=${response.status} body=${errorText}`
      );
    }

    const data = (await response.json()) as GoogleTokenResponse;

    tokenCache = {
      accessToken: data.access_token,
      expiresAt: now + data.expires_in,
      serviceAccountEmail: env.SERVICE_ACCOUNT_EMAIL
    };

    console.log("[google-auth] Google access token acquired", {
      expiresIn: data.expires_in,
      tokenType: data.token_type
    });

    return data.access_token;
  } catch (error) {
    if (error instanceof AppError) {
      throw error;
    }

    console.error("[google-auth] Unexpected access token error", error);
    throw new AppError(
      "google-token",
      "Unable to obtain Google access token.",
      error instanceof Error ? error.message : String(error)
    );
  }
}

async function createServiceAccountJwt(env: Env): Promise<string> {
  try {
    const now = Math.floor(Date.now() / 1000);

    const header = {
      alg: "RS256",
      typ: "JWT"
    };

    const payload = {
      iss: env.SERVICE_ACCOUNT_EMAIL,
      scope: SHEETS_SCOPE,
      aud: GOOGLE_TOKEN_URL,
      exp: now + 3600,
      iat: now
    };

    const encodedHeader = base64UrlEncodeJson(header);
    const encodedPayload = base64UrlEncodeJson(payload);
    const unsignedToken = `${encodedHeader}.${encodedPayload}`;

    console.log("[google-auth] Creating signed JWT", {
      serviceAccountEmail: env.SERVICE_ACCOUNT_EMAIL,
      scope: SHEETS_SCOPE
    });

    const signature = await signJwt(unsignedToken, env.PRIVATE_KEY);
    return `${unsignedToken}.${signature}`;
  } catch (error) {
    if (error instanceof AppError) {
      throw error;
    }

    console.error("[google-auth] JWT creation failed", error);
    throw new AppError(
      "jwt",
      "Unable to create signed JWT for Google Service Account.",
      error instanceof Error ? error.message : String(error)
    );
  }
}

async function signJwt(unsignedToken: string, privateKey: string): Promise<string> {
  try {
    const normalizedKey = normalizePrivateKey(privateKey);
    const keyData = pemToArrayBuffer(normalizedKey);
    const cryptoKey = await crypto.subtle.importKey(
      "pkcs8",
      keyData,
      {
        name: "RSASSA-PKCS1-v1_5",
        hash: "SHA-256"
      },
      false,
      ["sign"]
    );

    const signature = await crypto.subtle.sign(
      "RSASSA-PKCS1-v1_5",
      cryptoKey,
      new TextEncoder().encode(unsignedToken)
    );

    return base64UrlEncodeBytes(new Uint8Array(signature));
  } catch (error) {
    console.error("[google-auth] JWT signing failed", error);
    throw new AppError(
      "jwt",
      "Unable to sign JWT with the provided private key.",
      error instanceof Error ? error.message : String(error)
    );
  }
}

function normalizePrivateKey(privateKey: string): string {
  return privateKey.replace(/^"(.*)"$/s, "$1").replace(/\\n/g, "\n").trim();
}

function validateGoogleEnv(env: Env): void {
  if (!env.SPREADSHEET_ID?.trim()) {
    throw new AppError("env", "Missing SPREADSHEET_ID environment binding.");
  }

  if (!env.SERVICE_ACCOUNT_EMAIL?.trim()) {
    throw new AppError("env", "Missing SERVICE_ACCOUNT_EMAIL environment binding.");
  }

  if (!env.PRIVATE_KEY?.trim()) {
    throw new AppError("env", "Missing PRIVATE_KEY environment binding.");
  }

  const normalizedKey = normalizePrivateKey(env.PRIVATE_KEY);
  if (!normalizedKey.includes("BEGIN PRIVATE KEY")) {
    throw new AppError(
      "env",
      "PRIVATE_KEY does not look like a valid PKCS#8 private key.",
      "Expected PEM header -----BEGIN PRIVATE KEY-----."
    );
  }
}

function pemToArrayBuffer(pem: string): ArrayBuffer {
  const base64 = pem
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s+/g, "");

  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes.buffer;
}

function base64UrlEncodeJson(value: unknown): string {
  return base64UrlEncodeBytes(new TextEncoder().encode(JSON.stringify(value)));
}

function base64UrlEncodeBytes(bytes: Uint8Array): string {
  let binary = "";

  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
