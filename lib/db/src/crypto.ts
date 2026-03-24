import { createCipheriv, createDecipheriv, randomBytes } from "crypto";

function getKey(): Buffer {
  const hex = process.env["ENCRYPTION_KEY"] ?? "";
  if (!hex) throw new Error("ENCRYPTION_KEY env var is not set");
  if (hex.length !== 64) throw new Error("ENCRYPTION_KEY must be 64 hex characters (32 bytes)");
  return Buffer.from(hex, "hex");
}

const ENC_PREFIX = "enc:v1:";

export function encrypt(plaintext: string): string {
  const key = getKey();
  const iv = randomBytes(16);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${ENC_PREFIX}${iv.toString("hex")}:${ciphertext.toString("hex")}:${tag.toString("hex")}`;
}

export function decrypt(value: string): string {
  if (!value.startsWith(ENC_PREFIX)) return value;
  const rest = value.slice(ENC_PREFIX.length);
  const colonIdx1 = rest.indexOf(":");
  const colonIdx2 = rest.indexOf(":", colonIdx1 + 1);
  if (colonIdx1 === -1 || colonIdx2 === -1) throw new Error("Invalid encrypted value format");
  const ivHex = rest.slice(0, colonIdx1);
  const ciphertextHex = rest.slice(colonIdx1 + 1, colonIdx2);
  const tagHex = rest.slice(colonIdx2 + 1);
  const key = getKey();
  const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(ivHex, "hex"));
  decipher.setAuthTag(Buffer.from(tagHex, "hex"));
  const decrypted = Buffer.concat([
    decipher.update(Buffer.from(ciphertextHex, "hex")),
    decipher.final(),
  ]);
  return decrypted.toString("utf8");
}

export function isEncrypted(value: unknown): value is string {
  return typeof value === "string" && value.startsWith(ENC_PREFIX);
}

export function encryptJson(value: unknown): string {
  return encrypt(JSON.stringify(value));
}

export function decryptJson<T = unknown>(value: unknown): T {
  if (value == null) return value as T;
  const raw = typeof value === "string" ? value : JSON.stringify(value);
  if (!raw.startsWith(ENC_PREFIX)) {
    return (typeof value === "string" ? JSON.parse(value) : value) as T;
  }
  return JSON.parse(decrypt(raw)) as T;
}

export function maybeEncrypt(value: string | null | undefined): string | null {
  if (value == null) return null;
  if (isEncrypted(value)) return value;
  return encrypt(value);
}

export function maybeDecrypt(value: string | null | undefined): string | null {
  if (value == null) return null;
  return decrypt(value);
}
