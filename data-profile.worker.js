'use strict';

function toNumber(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || String(value).trim() === '') return NaN;
  const normalized = String(value).trim().replace(/,/g, '');
  return /^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(normalized) ? Number(normalized) : NaN;
}
function quantile(sorted, p) {
  if (!sorted.length) return NaN;
  const position = (sorted.length - 1) * p, lower = Math.floor(position), upper = Math.ceil(position);
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

self.onmessage = (event) => {
  try {
    const { rows, headers } = event.data;
    const numeric = [], categorical = [], mixed = [], ordered = [], missingByField = {}, cardinalityByField = {}, summary = {};
    let missing = 0;
    for (const header of headers) {
      const values = rows.map((row) => row[header]);
      const present = values.filter((value) => value !== '' && value !== null && value !== undefined);
      const numbers = present.map(toNumber);
      const ratio = present.length ? numbers.filter(Number.isFinite).length / present.length : 0;
      const isNumeric = present.length > 0 && ratio >= .9;
      if (present.length > 0 && ratio >= .1 && ratio < .9) mixed.push(header);
      const dateLike = /(date|time|timestamp|year|month|generation|iteration|epoch|step|round|cycle|eval)/i.test(header) || (present.length > 0 && present.slice(0, Math.min(20, present.length)).every((value) => typeof value === 'string' && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(value)));
      const absent = values.length - present.length;
      missingByField[header] = absent; missing += absent;
      if (isNumeric) {
        numeric.push(header);
        const sorted = numbers.filter(Number.isFinite).sort((a, b) => a - b);
        const mean = sorted.reduce((sum, value) => sum + value, 0) / sorted.length;
        const q1 = quantile(sorted, .25), q3 = quantile(sorted, .75), spread = q3 - q1;
        const outliers = sorted.filter((value) => value < q1 - 1.5 * spread || value > q3 + 1.5 * spread).length;
        summary[header] = { count: sorted.length, min: sorted[0], q1, median: quantile(sorted, .5), q3, max: sorted.at(-1), mean, sd: Math.sqrt(sorted.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, sorted.length - 1)), outliers };
      } else categorical.push(header);
      if (!isNumeric) cardinalityByField[header] = new Set(present.map((value) => String(value))).size;
      if (dateLike) ordered.push(header);
    }
    const keys = new Set(); let duplicates = 0;
    for (const row of rows) {
      const key = JSON.stringify(headers.map((header) => row[header]));
      if (keys.has(key)) duplicates++; else keys.add(key);
    }
    self.postMessage({ types: { numeric, categorical, mixed, ordered, missingByField, cardinalityByField, summary }, quality: { missing, duplicates } });
  } catch (error) {
    self.postMessage({ error: error?.message || 'Could not profile the imported data.' });
  }
};
