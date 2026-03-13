import { getLmeAluminium } from "../providers/lme.js";
import { getEurNok, getUsdNok } from "../providers/norgesBank.js";
import type { MarketSnapshot } from "../types/marketData.js";

function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

export async function getMarketSnapshot(): Promise<MarketSnapshot> {
  const snapshot: MarketSnapshot = {
    fetchedAt: new Date().toISOString(),
    fx: {
      usdNok: null,
      eurNok: null,
    },
    metals: {
      aluminium: null,
    },
    errors: [],
  };

  const [usdResult, eurResult, lmeResult] = await Promise.allSettled([
    getUsdNok(),
    getEurNok(),
    getLmeAluminium(),
  ]);

  if (usdResult.status === "fulfilled") {
    snapshot.fx.usdNok = usdResult.value;
  } else {
    snapshot.errors.push(`USD/NOK: ${toErrorMessage(usdResult.reason)}`);
  }

  if (eurResult.status === "fulfilled") {
    snapshot.fx.eurNok = eurResult.value;
  } else {
    snapshot.errors.push(`EUR/NOK: ${toErrorMessage(eurResult.reason)}`);
  }

  if (lmeResult.status === "fulfilled") {
    snapshot.metals.aluminium = lmeResult.value;
  } else {
    snapshot.errors.push(`LME Aluminium: ${toErrorMessage(lmeResult.reason)}`);
  }

  return snapshot;
}
