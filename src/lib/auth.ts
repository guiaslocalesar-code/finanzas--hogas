import type { Env, SessionUser } from "./types";

const SESSION_DURATION_SECONDS = 60 * 60 * 24 * 14;

export const SESSION_COOKIE_NAME = "fh_session";

export async function hashPassword(password: string, salt: string): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(`${salt}:${password}`)
  );
  return toHex(new Uint8Array(digest));
}

export function generateSalt(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  return toHex(bytes);
}

export async function createSessionToken(env: Env, user: SessionUser): Promise<string> {
  const payload = {
    ...user,
    exp: Math.floor(Date.now() / 1000) + SESSION_DURATION_SECONDS
  };
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signature = await signValue(env, encodedPayload);
  return `${encodedPayload}.${signature}`;
}

export async function verifySessionToken(env: Env, token: string): Promise<SessionUser | null> {
  const [encodedPayload, signature] = token.split(".");

  if (!encodedPayload || !signature) {
    return null;
  }

  const expectedSignature = await signValue(env, encodedPayload);
  if (signature !== expectedSignature) {
    return null;
  }

  try {
    const decoded = JSON.parse(base64UrlDecode(encodedPayload)) as SessionUser & { exp?: number };
    if (!decoded || typeof decoded !== "object") {
      return null;
    }

    if (!decoded.exp || decoded.exp < Math.floor(Date.now() / 1000)) {
      return null;
    }

    return {
      userId: String(decoded.userId || "").trim(),
      username: String(decoded.username || "").trim(),
      displayName: String(decoded.displayName || "").trim(),
      role: decoded.role === "superadmin" ? "superadmin" : "user"
    };
  } catch {
    return null;
  }
}

export function sessionMaxAgeSeconds(): number {
  return SESSION_DURATION_SECONDS;
}

async function signValue(env: Env, value: string): Promise<string> {
  const secret = (env.SESSION_SECRET || env.PRIVATE_KEY || "").trim();
  if (!secret) {
    throw new Error("Missing SESSION_SECRET or PRIVATE_KEY for session signing.");
  }

  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );

  const signature = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value));
  return toHex(new Uint8Array(signature));
}

function toHex(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function base64UrlEncode(value: string): string {
  return btoa(value).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlDecode(value: string): string {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - (normalized.length % 4 || 4)) % 4);
  return atob(padded);
}
