export type CurrencyPair = "USD/NOK" | "EUR/NOK";

export interface RatePoint {
  pair: CurrencyPair;
  rate: number;
  date: string;
  source: "Norges Bank";
}

export interface MetalPoint {
  metal: "LME Aluminium";
  price: number;
  currency: string;
  unit: string;
  priceType: string;
  date: string;
  source: "LME";
}

export interface MarketSnapshot {
  fetchedAt: string;
  fx: {
    usdNok: RatePoint | null;
    eurNok: RatePoint | null;
  };
  metals: {
    aluminium: MetalPoint | null;
  };
  errors: string[];
}
