import { env } from "../config/env.js";

const DEFAULT_RETRIES = 2;
const DEFAULT_RETRY_DELAY_MS = 300;

export interface HttpRequestOptions {
  method?: string;
  headers?: Record<string, string>;
  body?: BodyInit | null;
  timeoutMs?: number;
  retries?: number;
  retryDelayMs?: number;
}

export class HttpRequestError extends Error {
  public readonly url: string;
  public readonly status?: number;
  public readonly attempts: number;
  public readonly responseBody?: string;

  constructor(args: {
    message: string;
    url: string;
    attempts: number;
    status?: number;
    responseBody?: string;
  }) {
    super(args.message);
    this.name = "HttpRequestError";
    this.url = args.url;
    this.status = args.status;
    this.attempts = args.attempts;
    this.responseBody = args.responseBody;
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryableStatus(status: number): boolean {
  return status >= 500 && status <= 599;
}

function isNetworkError(error: unknown): boolean {
  if (error instanceof DOMException && error.name === "AbortError") {
    return true;
  }
  return error instanceof TypeError;
}

function parseRetryCount(retries: number | undefined): number {
  if (retries === undefined) {
    return DEFAULT_RETRIES;
  }
  return retries >= 0 ? Math.floor(retries) : 0;
}

function buildRequestInit(options: HttpRequestOptions): RequestInit {
  return {
    method: options.method ?? "GET",
    headers: options.headers,
    body: options.body,
  };
}

async function requestRaw(
  url: string,
  options: HttpRequestOptions = {},
): Promise<Response> {
  const maxRetries = parseRetryCount(options.retries);
  const totalAttempts = maxRetries + 1;
  const timeoutMs = options.timeoutMs ?? env.requestTimeoutMs;
  const retryDelayMs = options.retryDelayMs ?? DEFAULT_RETRY_DELAY_MS;
  const init = buildRequestInit(options);

  let lastError: unknown;

  for (let attempt = 1; attempt <= totalAttempts; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const response = await fetch(url, { ...init, signal: controller.signal });

      if (isRetryableStatus(response.status) && attempt < totalAttempts) {
        await sleep(retryDelayMs * attempt);
        continue;
      }

      if (!response.ok) {
        const responseBody = await response.text().catch(() => undefined);
        throw new HttpRequestError({
          message: `HTTP ${response.status} from ${url}`,
          url,
          status: response.status,
          attempts: attempt,
          responseBody,
        });
      }

      return response;
    } catch (error: unknown) {
      lastError = error;
      if (isNetworkError(error) && attempt < totalAttempts) {
        await sleep(retryDelayMs * attempt);
        continue;
      }

      if (error instanceof HttpRequestError) {
        throw error;
      }

      const reason =
        error instanceof Error ? error.message : "Unknown network error";
      throw new HttpRequestError({
        message: `Request failed for ${url}: ${reason}`,
        url,
        attempts: attempt,
      });
    } finally {
      clearTimeout(timer);
    }
  }

  const reason =
    lastError instanceof Error ? lastError.message : "Unknown request failure";
  throw new HttpRequestError({
    message: `Request failed for ${url}: ${reason}`,
    url,
    attempts: totalAttempts,
  });
}

export async function fetchJson<T>(
  url: string,
  options: HttpRequestOptions = {},
): Promise<T> {
  const response = await requestRaw(url, {
    ...options,
    headers: {
      Accept: "application/json",
      ...(options.headers ?? {}),
    },
  });

  try {
    return (await response.json()) as T;
  } catch (error: unknown) {
    const reason =
      error instanceof Error ? error.message : "Invalid JSON response";
    throw new HttpRequestError({
      message: `Failed to parse JSON from ${url}: ${reason}`,
      url,
      attempts: 1,
      status: response.status,
    });
  }
}

export async function fetchText(
  url: string,
  options: HttpRequestOptions = {},
): Promise<string> {
  const response = await requestRaw(url, options);
  return response.text();
}
