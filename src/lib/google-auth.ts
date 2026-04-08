import type { Env, GoogleTokenResponse } from "./types";

const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

type TokenCacheEntry = {
  accessToken: string;
  expiresAt: number;
  serviceAccountEmail: string;
};

let tokenCache: TokenCacheEntry | null = null;

export async function getGoogleAccessToken(env: Env): Promise<string> {
  const now = Math.floor(Date.now() / 1000);

  if (
    tokenCache &&
    tokenCache.serviceAccountEmail === env.SERVICE_ACCOUNT_EMAIL &&
    tokenCache.expiresAt - 60 > now
  ) {
    return tokenCache.accessToken;
  }

  const assertion = await createServiceAccountJwt(env);
  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion
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
    throw new Error(`Google OAuth token request failed: ${response.status} ${errorText}`);
  }

  const data = (await response.json()) as GoogleTokenResponse;

  tokenCache = {
    accessToken: data.access_token,
    expiresAt: now + data.expires_in,
    serviceAccountEmail: env.SERVICE_ACCOUNT_EMAIL
  };

  return data.access_token;
}

async function createServiceAccountJwt(env: Env): Promise<string> {
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

  const signature = await signJwt(unsignedToken, env.PRIVATE_KEY);
  return `${unsignedToken}.${signature}`;
}

async function signJwt(unsignedToken: string, privateKey: string): Promise<string> {
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
}

function normalizePrivateKey(privateKey: string): string {
  return privateKey.replace(/^"(.*)"$/s, "$1").replace(/\\n/g, "\n").trim();
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
