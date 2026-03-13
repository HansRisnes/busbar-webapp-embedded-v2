import { fetchJson } from "../lib/http.js";
import type { CurrencyPair, RatePoint } from "../types/marketData.js";

const USD_NOK_URL =
  "https://data.norges-bank.no/api/data/EXR/B.USD.NOK.SP?format=sdmx-json";
const EUR_NOK_URL =
  "https://data.norges-bank.no/api/data/EXR/B.EUR.NOK.SP?format=sdmx-json";

type JsonObject = Record<string, unknown>;

export class NorgesBankError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "NorgesBankError";
  }
}

function asObject(value: unknown): JsonObject | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  return value as JsonObject;
}

function parseNumeric(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.replace(",", ".").trim();
  if (!normalized) {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeIsoDate(value: string): string | null {
  const trimmed = value.trim();
  const isoDateMatch = /^\d{4}-\d{2}-\d{2}$/.test(trimmed);
  if (!isoDateMatch) {
    return null;
  }

  const parsed = Date.parse(`${trimmed}T00:00:00Z`);
  return Number.isNaN(parsed) ? null : trimmed;
}

function extractLatestObservationFromSdmx(payload: unknown): {
  rate: number;
  date: string;
  observationIndex: number;
} {
  const root = asObject(payload);
  const data = asObject(root?.data);
  const structure = asObject(data?.structure);
  const dimensions = asObject(structure?.dimensions);
  const observationDims = dimensions?.observation;

  if (!Array.isArray(observationDims) || observationDims.length === 0) {
    throw new NorgesBankError("SDMX response is missing observation dimensions.");
  }

  const firstObservationDim = asObject(observationDims[0]);
  const timeValuesRaw = firstObservationDim?.values;
  if (!Array.isArray(timeValuesRaw)) {
    throw new NorgesBankError("SDMX response is missing time values.");
  }

  const dataSets = data?.dataSets;
  if (!Array.isArray(dataSets) || dataSets.length === 0) {
    throw new NorgesBankError("SDMX response is missing data sets.");
  }

  const firstDataSet = asObject(dataSets[0]);
  const series = asObject(firstDataSet?.series);
  if (!series) {
    throw new NorgesBankError("SDMX response is missing series.");
  }

  let latestIndex = -1;
  let latestRawValue: unknown;

  for (const seriesValue of Object.values(series)) {
    const seriesObject = asObject(seriesValue);
    const observations = asObject(seriesObject?.observations);
    if (!observations) {
      continue;
    }

    for (const [indexKey, observationValue] of Object.entries(observations)) {
      const idx = Number(indexKey);
      if (!Number.isInteger(idx) || idx < 0) {
        continue;
      }
      if (idx > latestIndex) {
        latestIndex = idx;
        latestRawValue = observationValue;
      }
    }
  }

  if (latestIndex < 0) {
    throw new NorgesBankError("No observations found in SDMX response.");
  }

  const latestRateCandidate = Array.isArray(latestRawValue)
    ? latestRawValue[0]
    : latestRawValue;
  const rate = parseNumeric(latestRateCandidate);
  if (rate === null || rate <= 0) {
    throw new NorgesBankError(
      `Latest observation exists at index ${latestIndex}, but rate is missing or invalid.`,
    );
  }

  const timeValueObject = asObject(timeValuesRaw[latestIndex]);
  const dateRaw = typeof timeValueObject?.id === "string" ? timeValueObject.id : "";
  const date = normalizeIsoDate(dateRaw);
  if (!date) {
    throw new NorgesBankError(
      `Latest observation index ${latestIndex} is missing a valid ISO date.`,
    );
  }

  return { rate, date, observationIndex: latestIndex };
}

async function getRate(url: string, pair: CurrencyPair): Promise<RatePoint> {
  const payload = await fetchJson<unknown>(url);

  try {
    const latest = extractLatestObservationFromSdmx(payload);
    return {
      pair,
      rate: latest.rate,
      date: latest.date,
      source: "Norges Bank",
    };
  } catch (error: unknown) {
    if (error instanceof NorgesBankError) {
      throw error;
    }
    const message = error instanceof Error ? error.message : "Unknown parse error";
    throw new NorgesBankError(`Failed to parse ${pair} from Norges Bank: ${message}`);
  }
}

export function getUsdNok(): Promise<RatePoint> {
  return getRate(USD_NOK_URL, "USD/NOK");
}

export function getEurNok(): Promise<RatePoint> {
  return getRate(EUR_NOK_URL, "EUR/NOK");
}
