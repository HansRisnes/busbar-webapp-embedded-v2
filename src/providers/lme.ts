import { XMLParser } from "fast-xml-parser";
import { env } from "../config/env.js";
import { fetchText } from "../lib/http.js";
import type { MetalPoint } from "../types/marketData.js";

const ALUMINIUM_REGEX = /aluminium|aluminum/i;
const PRICE_HINT_REGEX = /price|official|settlement|average|cash|value|close|usd/i;
const DATE_HINT_REGEX = /date|time|period|asof|value.?date/i;
const CURRENCY_HINT_REGEX = /currency|ccy|curr|iso/i;
const UNIT_HINT_REGEX = /unit|uom|ton|tonne|metric/i;
const PRICE_TYPE_HINT_REGEX = /price.?type|type|basis|valuation/i;
const EXCLUDE_PRICE_KEY_REGEX = /date|time|period|year|month|day/i;

type JsonObject = Record<string, unknown>;

interface LeafEntry {
  path: string;
  key: string;
  value: unknown;
  valueAsString: string;
  searchableKey: string;
}

interface CandidateNode {
  path: string;
  leaves: LeafEntry[];
  textBlob: string;
}

export class LmeProviderError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "LmeProviderError";
  }
}

function asObject(value: unknown): JsonObject | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  return value as JsonObject;
}

function collectLeafEntries(
  value: unknown,
  path: string,
  key: string,
  out: LeafEntry[],
): void {
  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      collectLeafEntries(item, `${path}[${index}]`, key, out);
    });
    return;
  }

  const objectValue = asObject(value);
  if (objectValue) {
    for (const [childKey, childValue] of Object.entries(objectValue)) {
      collectLeafEntries(childValue, `${path}.${childKey}`, childKey, out);
    }
    return;
  }

  const valueAsString =
    typeof value === "string" ? value.trim() : value === null ? "" : String(value);
  out.push({
    path,
    key,
    value,
    valueAsString,
    searchableKey: `${path}.${key}`.toLowerCase(),
  });
}

function collectCandidateNodes(value: unknown, path: string, out: CandidateNode[]): void {
  if (Array.isArray(value)) {
    value.forEach((item, index) => collectCandidateNodes(item, `${path}[${index}]`, out));
    return;
  }

  const objectValue = asObject(value);
  if (!objectValue) {
    return;
  }

  const leaves: LeafEntry[] = [];
  collectLeafEntries(objectValue, path, "", leaves);

  const keyNames = Object.keys(objectValue).join(" ").toLowerCase();
  const valueNames = leaves
    .map((leaf) => `${leaf.key} ${leaf.valueAsString}`.toLowerCase())
    .join(" ");
  const textBlob = `${path.toLowerCase()} ${keyNames} ${valueNames}`.trim();

  out.push({ path, leaves, textBlob });

  for (const [childKey, childValue] of Object.entries(objectValue)) {
    collectCandidateNodes(childValue, `${path}.${childKey}`, out);
  }
}

function parseNumeric(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
    return null;
  }
  if (/^\d{8}$/.test(trimmed)) {
    return null;
  }

  const compact = trimmed.replace(/\s+/g, "");
  const normalized =
    compact.includes(",") && compact.includes(".")
      ? compact.replace(/,/g, "")
      : compact.includes(",")
        ? compact.replace(/,/g, ".")
        : compact;

  const cleaned = normalized.replace(/[^\d.-]/g, "");
  if (!cleaned) {
    return null;
  }

  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeDate(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const epoch = Date.parse(`${trimmed}T00:00:00Z`);
    return Number.isNaN(epoch) ? null : trimmed;
  }
  if (/^\d{8}$/.test(trimmed)) {
    const normalized = `${trimmed.slice(0, 4)}-${trimmed.slice(4, 6)}-${trimmed.slice(6, 8)}`;
    const epoch = Date.parse(`${normalized}T00:00:00Z`);
    return Number.isNaN(epoch) ? null : normalized;
  }

  const parsed = Date.parse(trimmed);
  if (Number.isNaN(parsed)) {
    return null;
  }
  return new Date(parsed).toISOString().slice(0, 10);
}

function pickPrice(leaves: LeafEntry[]): number | null {
  const hinted = leaves.filter(
    (leaf) =>
      PRICE_HINT_REGEX.test(leaf.searchableKey) &&
      !EXCLUDE_PRICE_KEY_REGEX.test(leaf.searchableKey),
  );
  const fallback = leaves.filter((leaf) => !EXCLUDE_PRICE_KEY_REGEX.test(leaf.searchableKey));

  for (const candidateList of [hinted, fallback]) {
    for (const leaf of candidateList) {
      const parsed = parseNumeric(leaf.value);
      if (parsed !== null && parsed > 0) {
        return parsed;
      }
    }
  }

  return null;
}

function pickDate(leaves: LeafEntry[]): string | null {
  const hinted = leaves.filter((leaf) => DATE_HINT_REGEX.test(leaf.searchableKey));
  for (const leaf of hinted) {
    const date = normalizeDate(leaf.valueAsString);
    if (date) {
      return date;
    }
  }

  for (const leaf of leaves) {
    const date = normalizeDate(leaf.valueAsString);
    if (date) {
      return date;
    }
  }

  return null;
}

function pickCurrency(leaves: LeafEntry[]): string | null {
  const hinted = leaves.filter((leaf) => CURRENCY_HINT_REGEX.test(leaf.searchableKey));
  const fallback = leaves;

  for (const candidateList of [hinted, fallback]) {
    for (const leaf of candidateList) {
      const upper = leaf.valueAsString.toUpperCase();
      if (/^[A-Z]{3}$/.test(upper)) {
        return upper;
      }
      const match = upper.match(/\b(USD|EUR|NOK|GBP)\b/);
      const currency = match?.[1];
      if (currency) {
        return currency;
      }
    }
  }

  return null;
}

function pickUnit(leaves: LeafEntry[]): string | null {
  const hinted = leaves.filter((leaf) => UNIT_HINT_REGEX.test(leaf.searchableKey));
  for (const leaf of hinted) {
    const normalized = leaf.valueAsString.toLowerCase();
    if (!normalized) {
      continue;
    }
    if (/\b(t|ton|tons|tonne|tonnes|mt)\b/.test(normalized)) {
      return "t";
    }
    return leaf.valueAsString;
  }
  return null;
}

function pickPriceType(candidate: CandidateNode): string {
  const normalized = candidate.textBlob.toLowerCase();
  if (/monthly\s*average|month\s*average/.test(normalized)) {
    return "Monthly Average";
  }
  if (/\bsettlement\b/.test(normalized)) {
    return "Settlement";
  }
  if (/\bofficial\b/.test(normalized)) {
    return "Official";
  }

  const hinted = candidate.leaves.filter((leaf) =>
    PRICE_TYPE_HINT_REGEX.test(leaf.searchableKey),
  );
  for (const leaf of hinted) {
    if (leaf.valueAsString) {
      return leaf.valueAsString;
    }
  }

  return "Unknown";
}

function listTopLevelNodes(parsed: unknown): string[] {
  const objectValue = asObject(parsed);
  if (!objectValue) {
    return [];
  }
  return Object.keys(objectValue);
}

function mapAluminiumPoint(parsedXml: unknown): MetalPoint {
  const nodes: CandidateNode[] = [];
  collectCandidateNodes(parsedXml, "$", nodes);

  const aluminiumCandidates = nodes.filter((node) => ALUMINIUM_REGEX.test(node.textBlob));
  if (aluminiumCandidates.length === 0) {
    const topLevelNodes = listTopLevelNodes(parsedXml);
    const nodeSummary = topLevelNodes.length > 0 ? topLevelNodes.join(", ") : "(none)";
    throw new LmeProviderError(
      `Could not find aluminium nodes in LME XML. Top-level nodes found: ${nodeSummary}.`,
    );
  }

  const failures: string[] = [];

  // The exact LME feed schema may vary. This mapper uses heuristics and can be
  // tuned by adjusting key-pattern regexes once real XML responses are verified.
  for (const candidate of aluminiumCandidates) {
    const price = pickPrice(candidate.leaves);
    const date = pickDate(candidate.leaves);

    if (price === null || price <= 0) {
      failures.push(`${candidate.path}: missing numeric price`);
      continue;
    }
    if (!date) {
      failures.push(`${candidate.path}: missing valid date`);
      continue;
    }

    return {
      metal: "LME Aluminium",
      price,
      currency: pickCurrency(candidate.leaves) ?? "USD",
      unit: pickUnit(candidate.leaves) ?? "t",
      priceType: pickPriceType(candidate),
      date,
      source: "LME",
    };
  }

  const details = failures.slice(0, 6).join("; ");
  throw new LmeProviderError(
    `Found aluminium-related nodes but could not map required fields (price/date). Details: ${details}`,
  );
}

function buildLmeHeaders(): Record<string, string> {
  const headers: Record<string, string> = {
    Accept: "application/xml, text/xml;q=0.9, */*;q=0.8",
  };

  if (env.lmeApiKey) {
    headers["X-API-Key"] = env.lmeApiKey;
  }

  if (env.lmeUsername && env.lmePassword) {
    const credentials = Buffer.from(`${env.lmeUsername}:${env.lmePassword}`).toString(
      "base64",
    );
    headers.Authorization = `Basic ${credentials}`;
  }

  return headers;
}

export async function getLmeAluminium(): Promise<MetalPoint> {
  if (!env.lmeXmlUrl) {
    throw new LmeProviderError(
      "LME feed is not configured. Set LME_XML_URL in .env to enable LME aluminium data.",
    );
  }

  try {
    const xml = await fetchText(env.lmeXmlUrl, { headers: buildLmeHeaders() });
    if (!xml.trim()) {
      throw new LmeProviderError("LME XML response was empty.");
    }

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      trimValues: true,
    });

    const parsedXml = parser.parse(xml);
    return mapAluminiumPoint(parsedXml);
  } catch (error: unknown) {
    if (error instanceof LmeProviderError) {
      throw error;
    }
    const reason = error instanceof Error ? error.message : "Unknown LME provider error";
    throw new LmeProviderError(`Failed to fetch/parse LME XML feed: ${reason}`);
  }
}
