export const round2 = n => Math.round(n * 100) / 100;

export const fmtNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

export const fmtIntNO = new Intl.NumberFormat('no-NO', {
  maximumFractionDigits: 0
});

export const fmtFxNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

export const fmtTimestampNO = new Intl.DateTimeFormat('no-NO', {
  dateStyle: 'short',
  timeStyle: 'short'
});

export const fmtPercentNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 0,
  maximumFractionDigits: 2
});

export const fmtMarketPercentNO = new Intl.NumberFormat('no-NO', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
