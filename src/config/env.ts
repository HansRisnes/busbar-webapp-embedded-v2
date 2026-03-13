import dotenv from "dotenv";

dotenv.config({ quiet: true });

const DEFAULT_REQUEST_TIMEOUT_MS = 10_000;

export class EnvValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "EnvValidationError";
  }
}

export interface EnvConfig {
  requestTimeoutMs: number;
  lmeXmlUrl?: string;
  lmeUsername?: string;
  lmePassword?: string;
  lmeApiKey?: string;
  isLmeConfigured: boolean;
}

function readTrimmed(name: string): string | undefined {
  const raw = process.env[name];
  if (raw === undefined) {
    return undefined;
  }

  const trimmed = raw.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function parseRequestTimeout(raw: string | undefined): number {
  if (!raw) {
    return DEFAULT_REQUEST_TIMEOUT_MS;
  }

  const parsed = Number(raw);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new EnvValidationError(
      "REQUEST_TIMEOUT_MS must be a positive integer (milliseconds).",
    );
  }

  return parsed;
}

function validateLmeUrl(url: string | undefined): void {
  if (!url) {
    return;
  }

  try {
    new URL(url);
  } catch {
    throw new EnvValidationError("LME_XML_URL must be a valid URL.");
  }
}

function validateCredentialPair(
  username: string | undefined,
  password: string | undefined,
): void {
  if ((username && !password) || (!username && password)) {
    throw new EnvValidationError(
      "LME_USERNAME and LME_PASSWORD must both be set when basic auth is used.",
    );
  }
}

function loadEnv(): EnvConfig {
  const requestTimeoutMs = parseRequestTimeout(readTrimmed("REQUEST_TIMEOUT_MS"));
  const lmeXmlUrl = readTrimmed("LME_XML_URL");
  const lmeUsername = readTrimmed("LME_USERNAME");
  const lmePassword = readTrimmed("LME_PASSWORD");
  const lmeApiKey = readTrimmed("LME_API_KEY");

  validateLmeUrl(lmeXmlUrl);
  validateCredentialPair(lmeUsername, lmePassword);

  return {
    requestTimeoutMs,
    lmeXmlUrl,
    lmeUsername,
    lmePassword,
    lmeApiKey,
    isLmeConfigured: Boolean(lmeXmlUrl),
  };
}

export const env = loadEnv();
