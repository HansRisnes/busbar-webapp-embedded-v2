export const XAP_ALUMINIUM_REFERENCE_USD_PER_TON = 2562;

const XAP_INDEXED_TYPES = new Set([
  'straight_3m', 'straight_3m_dist',
  'straight_1501_2000', 'straight_1501_2000_dist',
  'straight_500_1000', 'straight_500_1000_dist',
  'elbow_horizontal_90', 'elbow_vertical_90',
  'tee_horizontal_offset', 'tee_vertical_offset',
  'z_horizontal_offset', 'z_vertical_offset',
  'board_feed', 'crt_board_feed', 'flexible_set'
]);

export function isXapIndexedType(type){
  return XAP_INDEXED_TYPES.has(type);
}

export function getCurrentAluminiumPrice(aluminium, fxUpdatedAt){
  if (!aluminium || aluminium.source !== 'Manuell') return null;
  if (aluminium.currency !== 'USD' || aluminium.unit !== 't') return null;
  const price = Number(aluminium.price);
  const updatedAt = Date.parse(aluminium.updatedAt || '');
  const fxTime = Date.parse(fxUpdatedAt || '');
  if (!Number.isFinite(price) || price <= 0 || !Number.isFinite(updatedAt)
    || !Number.isFinite(fxTime) || updatedAt < fxTime) return null;
  return price;
}

export function adjustXapUsdPrice(basePriceUsd, type, aluminiumUsdPerTon){
  const base = Number(basePriceUsd);
  if (!Number.isFinite(base) || base < 0) return NaN;
  if (!isXapIndexedType(type)) return base;
  const market = Number(aluminiumUsdPerTon);
  if (!Number.isFinite(market) || market <= 0) return NaN;
  const factor = 1 + 0.8 * (market - XAP_ALUMINIUM_REFERENCE_USD_PER_TON)
    / XAP_ALUMINIUM_REFERENCE_USD_PER_TON;
  return Math.round((base * factor + Number.EPSILON) * 100) / 100;
}
