(() => {
  'use strict';

  const $ = (selector) => document.querySelector(selector);
  const MAX_FILE_BYTES = 25 * 1024 * 1024;
  const MAX_ROWS = 50000;
  const MAX_FIELDS = 1000;
  const MAX_CELLS = 1000000;
  const MAX_PDF_PAGES = 50;
  const MAX_CHART_POINTS = 6000;
  const state = {
    rows: [], headers: [], source: null, types: null, profile: null,
    recommendations: [], selected: '', chart: null, workbook: null,
    pdf: null, documentTables: [], imageFile: null, ocrAvailable: false, ocrBusy: false, imageURL: '', imageGrid: null,
    renderToken: 0, profileToken: 0, profileWorker: null, profileReject: null, parserPromises: new Map(), currentFileName: '', sourceFileInfo: null
  };

  const vizMeta = {
    line: ['Line / time series', 'Ordered measurements and repeated observations.'],
    area: ['Area trend', 'Cumulative or continuous change across an ordered axis.'],
    scatter: ['Scatter plot', 'Association between two measured variables.'],
    bubble: ['Bubble plot', 'Two measurements with a third variable encoded by size.'],
    bar: ['Category comparison', 'Aggregated values by category.'],
    groupedbar: ['Grouped comparison', 'Compare categories across a second grouping field.'],
    hist: ['Histogram', 'Frequency distribution of a measured variable.'],
    ecdf: ['Empirical CDF', 'Cumulative distribution without bin-width choices.'],
    box: ['Box plot', 'Median, quartiles, range, and group-wise spread.'],
    heat: ['Correlation matrix', 'Pairwise Pearson correlations with pairwise-complete values.'],
    pareto: ['Pareto chart', 'Ranked category totals with cumulative contribution.'],
    pixelheat: ['Image intensity map', 'Downsampled spatial brightness profile from an image.'],
    table: ['Data table', 'Inspect the exact extracted records and field values.'],
    mindmap: ['Concept map', 'Show hierarchy from headings, indentation, or record categories.'],
    flow: ['Process sequence', 'Display ordered stages extracted from notes or document text.']
  };

  const palettes = {
    zstack: ['#16766d', '#bd9252', '#344f69', '#bc675b', '#4a9b8b', '#755b87', '#8c9b58', '#d17c4c'],
    colorblind: ['#0072B2', '#E69F00', '#009E73', '#D55E00', '#CC79A7', '#56B4E9', '#F0E442', '#333333'],
    mono: ['#202a32', '#46545f', '#64737e', '#87949d', '#a8b1b7', '#c4cbd0'],
    journal: ['#3B5B92', '#C36B38', '#5B8F78', '#9B6B8C', '#C6A34F', '#657A88', '#A64949', '#7D8C4E']
  };

  function setMessage(message, kind = 'info') {
    const el = $('#import-message');
    el.textContent = message;
    el.dataset.kind = kind;
  }

  function ensureScript(path, ready) {
    if (ready()) return Promise.resolve();
    if (state.parserPromises.has(path)) return state.parserPromises.get(path);
    const promise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = new URL(path, document.baseURI).href;
      script.async = true;
      script.onload = () => ready() ? resolve() : reject(new Error(`Could not initialize ${path}.`));
      script.onerror = () => reject(new Error(`Could not load the required parser: ${path}.`));
      document.head.appendChild(script);
    });
    state.parserPromises.set(path, promise);
    return promise;
  }

  const ensureXLSX = () => ensureScript('./vendor/xlsx-0.20.3.min.js', () => !!window.XLSX);
  const ensureMammoth = () => ensureScript('./vendor/mammoth-1.13.0.min.js', () => !!window.mammoth);
  const ensurePDF = async () => {
    await ensureScript('./vendor/pdfjs-3.11.174.min.js', () => !!window.pdfjsLib);
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdfjs-3.11.174.worker.min.js', document.baseURI).href;
    return window.pdfjsLib;
  };
  const ensureECharts = () => ensureScript('./vendor/echarts-6.1.0.min.js', () => !!window.echarts);

  function parseDelimited(text, delimiter, enforceLimits = false) {
    const rows = [];
    let row = [], field = '', quoted = false, headers = null, truncated = false;
    let rowLimit = MAX_ROWS;
    const finishRow = () => {
      row.push(field); field = '';
      if (row.some((value) => String(value).trim())) {
        if (!headers) {
          headers = uniqueHeaders(row);
          if (enforceLimits && headers.length > MAX_FIELDS) throw new Error(`This source has more than ${MAX_FIELDS.toLocaleString()} fields.`);
          rowLimit = enforceLimits ? Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_CELLS / Math.max(1, headers.length)))) : MAX_ROWS;
        } else if (rows.length >= rowLimit) {
          truncated = true;
          return true;
        } else rows.push(row.slice(0, headers.length));
      }
      row = [];
      return false;
    };
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (ch === '"') quoted = false;
        else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === delimiter) {
        if (enforceLimits && row.length >= MAX_FIELDS) throw new Error(`This source has more than ${MAX_FIELDS.toLocaleString()} fields.`);
        row.push(field); field = '';
      }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        if (finishRow()) break;
      } else field += ch;
    }
    if (!truncated && (row.length || field.length)) finishRow();
    if (quoted && enforceLimits) throw new Error('The delimited text ends inside a quoted field. Check its quotation marks.');
    if (!headers || !rows.length) return null;
    const records = rows.map((values) => Object.fromEntries(headers.map((name, i) => [name, values[i] ?? ''])));
    return { rows: records, headers, truncated };
  }

  function uniqueHeaders(values) {
    const seen = new Map();
    return values.map((value, i) => {
      const base = String(value ?? '').replace(/^\uFEFF/, '').trim() || `field_${i + 1}`;
      const count = (seen.get(base) || 0) + 1;
      seen.set(base, count);
      return count === 1 ? base : `${base}_${count}`;
    });
  }

  function detectDelimited(text) {
    const sample = [];
    let start = 0;
    for (let i = 0; i <= text.length && sample.length < 12; i++) {
      if (i === text.length || text[i] === '\n' || text[i] === '\r') {
        const line = text.slice(start, i);
        if (line.trim()) sample.push(line);
        if (text[i] === '\r' && text[i + 1] === '\n') i++;
        start = i + 1;
      }
    }
    const candidates = [',', '\t', ';', '|'];
    let best = null;
    for (const delimiter of candidates) {
      const parsed = parseDelimited(sample.join('\n'), delimiter);
      if (!parsed) continue;
      const width = parsed.headers.length;
      const score = width > 1 ? width * sample.length : 0;
      if (!best || score > best.score) best = { delimiter, score };
    }
    return best?.delimiter || null;
  }

  function normalizeRows(input) {
    if (!Array.isArray(input) || !input.length) return null;
    let headers, sourceRows, totalRows;
    if (Array.isArray(input[0])) {
      headers = uniqueHeaders(input[0]);
      sourceRows = input.slice(1);
      totalRows = sourceRows.length;
    } else if (input[0] && typeof input[0] === 'object') {
      sourceRows = input.slice(0, MAX_ROWS);
      totalRows = input.length;
      const keys = new Set();
      for (const row of sourceRows) if (row && typeof row === 'object' && !Array.isArray(row)) {
        for (const key of Object.keys(row)) { keys.add(key); if (keys.size > MAX_FIELDS) break; }
        if (keys.size > MAX_FIELDS) break;
      }
      headers = uniqueHeaders([...keys]);
    } else return null;
    if (headers.length > MAX_FIELDS) throw new Error(`This source has more than ${MAX_FIELDS.toLocaleString()} fields.`);
    const rowLimit = Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_CELLS / Math.max(1, headers.length))));
    const selectedRows = sourceRows.slice(0, rowLimit);
    const normalized = selectedRows.map((row) => {
      if (Array.isArray(row)) return Object.fromEntries(headers.map((name, i) => [name, row[i] ?? '']));
      if (row && typeof row === 'object' && !Array.isArray(row)) return Object.fromEntries(headers.map((name) => [name, row[name] ?? '']));
      return Object.fromEntries(headers.map((name) => [name, '']));
    });
    return { rows: normalized, headers, truncated: totalRows > rowLimit };
  }

  function parseSource(text) {
    const sourceText = text.replace(/^\uFEFF/, '').trim();
    if (!sourceText) throw new Error('Add data or text before profiling.');
    if (sourceText.length > MAX_FILE_BYTES) throw new Error('Text input exceeds the 25 MB processing limit. Split the source into smaller parts.');
    if (sourceText.startsWith('{') || sourceText.startsWith('[')) {
      let object;
      try { object = JSON.parse(sourceText); }
      catch { throw new Error('This appears to be JSON but could not be parsed. Check its brackets, commas, and quotes.'); }
      const records = normalizeRows(Array.isArray(object) ? object : (object.data || object.records || object.rows));
      if (records) return { ...records, kind: 'table', label: 'Pasted JSON' };
      if (object && typeof object === 'object' && !Array.isArray(object)) {
        const rows = Object.entries(object).map(([concept, detail]) => ({ concept, detail: typeof detail === 'object' ? JSON.stringify(detail) : String(detail) }));
        return { rows, headers: ['concept', 'detail'], kind: 'hierarchy', label: 'Pasted JSON object' };
      }
    }
    const delimiter = detectDelimited(sourceText);
    if (delimiter) {
      const parsed = parseDelimited(sourceText, delimiter, true);
      if (parsed) {
        return { ...parsed, kind: 'table', label: delimiter === '\t' ? 'Tab-separated data' : 'Delimited data' };
      }
    }
    const lineLimit = Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_CELLS / 4)));
    const lines = [];
    let lineStart = 0, lineTruncated = false;
    for (let i = 0; i <= sourceText.length; i++) {
      if (i === sourceText.length || sourceText[i] === '\n' || sourceText[i] === '\r') {
        const line = sourceText.slice(lineStart, i);
        if (line.trim()) {
          if (lines.length >= lineLimit) { lineTruncated = true; break; }
          lines.push(line);
        }
        if (sourceText[i] === '\r' && sourceText[i + 1] === '\n') i++;
        lineStart = i + 1;
      }
    }
    let sawRoot = false;
    const rows = lines.map((line, index) => {
      const headingDepth = (line.match(/^\s*(#{1,6})\s/) || [])[1]?.length;
      const indent = (line.match(/^\s*/)?.[0] || '').length;
      const level = headingDepth ? headingDepth - 1 : indent >= 4 ? 2 : indent >= 1 || /^\s*[-*•\d.)]/.test(line) ? 1 : (sawRoot ? 1 : 0);
      sawRoot = true;
      return { line: line.trim().replace(/^#{1,6}\s+/, '').replace(/^[-*•\d.)\s]+/, ''), level, original: line.trim(), index };
    });
    return { rows, headers: ['line', 'level', 'original', 'index'], kind: 'notes', label: 'Text or notes', truncated: lineTruncated };
  }

  function setSource(source, fileName = '') {
    const rowLimit = Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_CELLS / Math.max(1, source.headers.length))));
    if (source.truncated && !source.warning) {
      source.warning = `The source exceeds the processing limits; the imported view is capped at ${rowLimit.toLocaleString()} rows and ${MAX_CELLS.toLocaleString()} cells.`;
    }
    if (source.rows.length > rowLimit) {
      source.rows = source.rows.slice(0, rowLimit);
      source.truncated = true;
      source.warning = `${source.warning ? `${source.warning} ` : ''}Import was capped at ${rowLimit.toLocaleString()} rows (${MAX_CELLS.toLocaleString()} cells) to keep browser analysis responsive.`;
    }
    if (source.rows.length > MAX_ROWS) {
      source.rows = source.rows.slice(0, MAX_ROWS);
      source.truncated = true;
    }
    if (source.headers.length > MAX_FIELDS) throw new Error(`This source has more than ${MAX_FIELDS.toLocaleString()} fields. Select a narrower table or sheet before profiling.`);
    state.rows = source.rows;
    state.headers = source.headers;
    state.source = source;
    state.profile = null;
    state.types = null;
    state.currentFileName = fileName || source.label || 'Pasted source';
    syncWorkflow();
    profileStudio().catch((error) => { $('#diagnosis').textContent = error.message || 'Profiling failed.'; $('#data-state').textContent = 'PROFILE FAILED'; });
  }

  function inferTypes() {
    const numeric = [], categorical = [], mixed = [], ordered = [], missingByField = {}, cardinalityByField = {};
    const summary = {};
    for (const header of state.headers) {
      const values = state.rows.map((row) => row[header]);
      const present = values.filter((value) => value !== '' && value !== null && value !== undefined);
      const numbers = present.map((value) => (typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''))));
      const numericRatio = present.length ? numbers.filter(Number.isFinite).length / present.length : 0;
      const isNumeric = present.length > 0 && numericRatio >= 0.9;
      if (present.length > 0 && numericRatio >= .1 && numericRatio < .9) mixed.push(header);
      const dateLike = /(date|time|timestamp|year|month|generation|iteration|epoch|step|round|cycle|eval)/i.test(header) || (present.length > 0 && present.slice(0, Math.min(20, present.length)).every((value) => typeof value === 'string' && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(value)));
      const missing = values.length - present.length;
      missingByField[header] = missing;
      if (isNumeric) {
        numeric.push(header);
        const clean = numbers.filter(Number.isFinite).sort((a, b) => a - b);
        const mean = clean.reduce((sum, value) => sum + value, 0) / clean.length;
        const q1 = quantile(clean, .25), q3 = quantile(clean, .75), spread = q3 - q1;
        const outliers = clean.filter((value) => value < q1 - 1.5 * spread || value > q3 + 1.5 * spread).length;
        summary[header] = { count: clean.length, min: clean[0], q1, median: quantile(clean, .5), q3, max: clean.at(-1), mean, sd: Math.sqrt(clean.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, clean.length - 1)), outliers };
      } else { categorical.push(header); cardinalityByField[header] = new Set(present.map((value) => String(value))).size; }
      if (dateLike) ordered.push(header);
    }
    return { numeric, categorical, mixed, ordered, missingByField, cardinalityByField, summary };
  }

  function quantile(sorted, p) {
    if (!sorted.length) return NaN;
    const pos = (sorted.length - 1) * p, lower = Math.floor(pos), upper = Math.ceil(pos);
    return sorted[lower] + (sorted[upper] - sorted[lower]) * (pos - lower);
  }

  function profileQuality(computed) {
    const n = state.rows.length;
    const missing = computed?.quality?.missing ?? Object.values(state.types.missingByField).reduce((sum, count) => sum + count, 0);
    let duplicates = computed?.quality?.duplicates;
    if (!Number.isFinite(duplicates)) {
      const duplicateKeys = new Set(); duplicates = 0;
      for (const row of state.rows) { const key = JSON.stringify(state.headers.map((name) => row[name])); if (duplicateKeys.has(key)) duplicates++; else duplicateKeys.add(key); }
    }
    const cells = Math.max(1, n * state.headers.length);
    const profile = { rows: n, fields: state.headers.length, numeric: state.types.numeric.length, categorical: state.types.categorical.length, missing, missingPct: 100 * missing / cells, duplicates, duplicatePct: 100 * duplicates / Math.max(1, n) };
    state.profile = profile;
    const chips = [
      [`${n.toLocaleString()} observations`, ''],
      [`${state.headers.length} fields`, ''],
      [`${profile.numeric} numeric`, ''],
      [`${profile.categorical} categorical / text`, ''],
      [`${missing.toLocaleString()} missing cells (${profile.missingPct.toFixed(1)}%)`, missing ? 'caution' : 'good'],
      [`${duplicates.toLocaleString()} duplicate rows`, duplicates ? 'caution' : 'good']
    ];
    $('#quality-profile').innerHTML = chips.map(([label, cls]) => `<span class="quality-chip ${cls}">${escapeHTML(label)}</span>`).join('');
    $('#quality-profile').hidden = false;
    const warnings = [];
    if (state.source.truncated) warnings.push(`The source exceeded a row or cell limit; only the first ${n.toLocaleString()} rows are included in the workspace.`);
    if (!profile.numeric && state.source.kind !== 'notes' && state.source.kind !== 'hierarchy') warnings.push('No numeric fields were detected. Try a structure view or inspect whether numeric values were imported as formatted text.');
    if (state.types.mixed?.length) warnings.push(`Mixed numeric/text values were found in ${state.types.mixed.length} field(s): ${state.types.mixed.slice(0, 8).join(', ')}. Review these fields before selecting a chart.`);
    if (missing) warnings.push(`${missing.toLocaleString()} missing cells are excluded from calculations where required; they are not imputed.`);
    if (duplicates) warnings.push(`${duplicates.toLocaleString()} repeated rows were detected. They remain in the dataset and are not removed.`);
    const outlierFields = Object.entries(state.types.summary).filter(([, stats]) => stats.outliers > 0);
    if (outlierFields.length) warnings.push(`Values outside 1.5×IQR fences were flagged in ${outlierFields.length} numeric field(s). These are screening flags, not data errors.`);
    if (state.source.warning) warnings.push(state.source.warning);
    const summary = Object.entries(state.types.summary).slice(0, 4).map(([name, stats]) => `${name}: min ${formatNum(stats.min)}, median ${formatNum(stats.median)}, max ${formatNum(stats.max)}`);
    const levels = Object.entries(state.types.cardinalityByField || {}).slice(0, 4).map(([name, count]) => `${name}: ${count} levels`);
    const notes = warnings.length ? `<ul>${warnings.map((warning) => `<li class="caution">${escapeHTML(warning)}</li>`).join('')}</ul>` : '<span class="good">No missing or duplicate records detected in the imported rows.</span>';
    $('#diagnosis').innerHTML = `<b>Source profile.</b> ${n.toLocaleString()} rows × ${state.headers.length} fields. Numeric fields: ${state.types.numeric.map(escapeHTML).join(', ') || 'none detected'}. Categorical/text fields: ${state.types.categorical.map(escapeHTML).join(', ') || 'none detected'}.${summary.length ? `<br><b>Numeric ranges.</b> ${summary.map(escapeHTML).join(' · ')}.` : ''}${levels.length ? `<br><b>Category levels.</b> ${levels.map(escapeHTML).join(' · ')}.` : ''}${notes}`;
  }

  function formatNum(value) { return Number.isFinite(value) ? Number(value.toPrecision(4)).toString() : 'n/a'; }
  function escapeHTML(value) { return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char])); }

  function optionFor(select, value, label) {
    const option = document.createElement('option'); option.value = value; option.textContent = label; select.appendChild(option);
  }

  function fillAxisControls() {
    const x = $('#axis-x'), y = $('#axis-y'), group = $('#axis-group');
    x.replaceChildren(); y.replaceChildren(); group.replaceChildren();
    const all = state.headers, numeric = state.types.numeric;
    for (const field of all) optionFor(x, field, field);
    for (const field of numeric.length ? numeric : all) optionFor(y, field, field);
    optionFor(group, '', 'No grouping');
    for (const field of state.types.categorical) optionFor(group, field, field);
    const preferredX = state.types.ordered.find((field) => all.includes(field)) || state.types.categorical[0] || numeric[0] || all[0];
    const preferredY = numeric.find((field) => field !== preferredX) || numeric[0] || all[0];
    x.value = preferredX || ''; y.value = preferredY || ''; group.value = '';
    $('#label-x').value = preferredX || ''; $('#label-y').value = preferredY || '';
  }

  function compatibleViews() {
    const n = state.types.numeric.length, c = state.types.categorical.length, kind = state.source.kind, views = [];
    const add = (key, score, rationale) => views.push({ key, score, rationale });
    add('table', 1, 'preserves imported rows and mixed field types');
    if (n >= 1 && (c >= 1 || n >= 2)) add('bar', 2, 'compares a measure across categories');
    if (n >= 1 && c >= 2) add('groupedbar', 2, 'compares categories within a second grouping variable');
    if (n >= 1 && (n >= 2 || state.types.ordered.length || c >= 1)) add('line', state.types.ordered.length ? 5 : 2, 'uses a sequence or selected horizontal field to show change');
    if (n >= 1 && state.types.ordered.length) add('area', 2, 'shows the scale of change across an ordered field');
    if (n >= 2) add('scatter', 4, 'uses two numeric fields to inspect association and unusual observations');
    if (n >= 3) add('bubble', 2, 'adds a third numeric field as point size');
    if (n >= 1) { add('hist', 4, 'summarizes the frequency distribution of a numeric field'); add('ecdf', 3, 'shows cumulative proportion without choosing histogram bins'); add('box', 3, 'summarizes spread and quartiles, optionally by group'); }
    if (n >= 2 && n <= 20) add('heat', 3, 'shows pairwise Pearson correlation for up to 20 numeric variables');
    if (c && (n || kind === 'image')) add('pareto', 2, 'ranks positive group contributions and cumulative share');
    if (kind === 'image') add('pixelheat', 6, 'the uploaded image was sampled into an x/y brightness grid');
    if (['notes', 'hierarchy', 'document', 'pdf', 'ocr'].includes(kind) || c && !n) { add('mindmap', 4, 'turns headings, indentation, or categories into an inspectable hierarchy'); add('flow', 2, 'turns ordered lines into a process sequence'); }
    const intent = $('#intent').value;
    const boost = { compare: ['bar', 'groupedbar', 'box'], trend: ['line', 'area'], distribution: ['hist', 'ecdf', 'box'], relationship: ['scatter', 'bubble', 'heat'], composition: ['pareto', 'bar'], structure: ['mindmap'], process: ['flow'] }[intent] || [];
    for (const view of views) if (boost.includes(view.key)) view.score += 4;
    return views.sort((a, b) => b.score - a.score || a.key.localeCompare(b.key));
  }

  async function computeProfileInWorker() {
    if (!window.Worker) throw new Error('This browser does not support background profiling.');
    if (state.profileWorker) {
      state.profileWorker.terminate();
      state.profileWorker = null;
      const rejectPrior = state.profileReject; state.profileReject = null;
      rejectPrior?.(new Error('Profiling was superseded by a newer source.'));
    }
    return new Promise((resolve, reject) => {
      let worker;
      try { worker = new Worker(new URL('./data-profile.worker.js', document.baseURI)); }
      catch (error) { reject(error); return; }
      state.profileWorker = worker;
      state.profileReject = reject;
      worker.onmessage = (event) => { worker.terminate(); if (state.profileWorker === worker) { state.profileWorker = null; state.profileReject = null; } event.data?.error ? reject(new Error(event.data.error)) : resolve(event.data); };
      worker.onerror = (event) => { worker.terminate(); if (state.profileWorker === worker) { state.profileWorker = null; state.profileReject = null; } reject(new Error(event.message || 'Background profiling failed.')); };
      worker.postMessage({ rows: state.rows, headers: state.headers });
    });
  }

  async function profileStudio() {
    if (!state.rows.length || !state.headers.length) { $('#diagnosis').textContent = 'No usable rows or fields were found in this source.'; $('#data-state').textContent = 'NO USABLE DATA'; return; }
    if (state.headers.length > MAX_FIELDS) { $('#diagnosis').textContent = `This source has more than ${MAX_FIELDS.toLocaleString()} fields. Select a narrower table or sheet before profiling.`; $('#data-state').textContent = 'TOO MANY FIELDS'; return; }
    const token = ++state.profileToken;
    $('#data-state').textContent = 'PROFILING SOURCE…';
    let computed;
    try { computed = await computeProfileInWorker(); }
    catch (error) { computed = null; }
    if (token !== state.profileToken) return;
    state.types = computed?.types || inferTypes();
    profileQuality(computed);
    fillAxisControls();
    state.recommendations = compatibleViews();
    state.selected = state.recommendations[0]?.key || 'table';
    $('#recommendation-caption').textContent = `${state.currentFileName} · ranked by field types, cardinality, and your selected research question.`;
    $('#figure-settings').hidden = false;
    $('#exports').hidden = false;
    updateExportAvailability();
    $('#viz-catalog').innerHTML = state.recommendations.map((view, index) => `<button type="button" class="viz-option ${index < 3 ? 'recommended' : ''} ${view.key === state.selected ? 'active' : ''}" data-viz="${view.key}"><strong>${vizMeta[view.key][0]}</strong><span>${vizMeta[view.key][1]}<br><em>Recommended fit: ${escapeHTML(view.rationale)}.</em></span></button>`).join('');
    $('#viz-catalog').querySelectorAll('[data-viz]').forEach((button) => button.addEventListener('click', () => {
      state.selected = button.dataset.viz;
      $('#viz-catalog').querySelectorAll('[data-viz]').forEach((item) => item.classList.toggle('active', item === button));
      updateExportAvailability();
      renderFigure();
      syncWorkflow();
    }));
    $('#data-state').textContent = `${state.rows.length.toLocaleString()} ROWS · ${state.headers.length} FIELDS`;
    syncWorkflow();
    await renderFigure();
  }

  function updateExportAvailability() {
    for (const button of document.querySelectorAll('#exports [data-export]')) {
      const format = button.dataset.export;
      const allowed = !['svg', 'png2', 'png4'].includes(format) || state.selected !== 'table';
      const structureFormat = !['mmd', 'opml'].includes(format) || ['mindmap', 'flow'].includes(state.selected);
      button.disabled = !(allowed && structureFormat);
      button.title = button.disabled ? (structureFormat ? 'Choose a chart or structure view to export an image.' : 'Choose Concept map or Process sequence to export a structured outline.') : '';
    }
  }

  function parseMatrix(matrix, kind, label, warning = '') {
    const data = normalizeRows(matrix);
    if (!data) throw new Error('No tabular data was found in the selected sheet.');
    if (data.rows.length > MAX_ROWS) data.rows = data.rows.slice(0, MAX_ROWS);
    const limitWarning = data.truncated ? `The worksheet exceeds the processing limit; imported rows are capped at ${data.rows.length.toLocaleString()} and 1,000,000 cells.` : '';
    setSource({ ...data, kind, label, warning: [warning, limitWarning].filter(Boolean).join(' ') }, label);
  }

  function showWorkbookSheets(workbook) {
    state.workbook = workbook;
    if (!workbook.SheetNames.length) throw new Error('This workbook contains no worksheets.');
    $('#sheet-wrap label').textContent = 'Worksheet';
    const select = $('#sheet-select'); select.replaceChildren();
    for (const sheetName of workbook.SheetNames) optionFor(select, sheetName, sheetName);
    $('#sheet-wrap').hidden = workbook.SheetNames.length < 2;
    parseWorkbookSheet(workbook.SheetNames[0]);
  }

  function parseWorkbookSheet(sheetName) {
    const sheet = state.workbook?.Sheets[sheetName];
    if (!sheet) return;
    const ref = sheet['!ref'];
    if (!ref) throw new Error('The selected worksheet is empty.');
    const full = window.XLSX.utils.decode_range(ref);
    const fieldCount = full.e.c - full.s.c + 1;
    const cappedColumns = Math.min(fieldCount, MAX_FIELDS);
    const dataRowLimit = Math.max(1, Math.min(MAX_ROWS, Math.floor(MAX_CELLS / cappedColumns)));
    const range = { s: { ...full.s }, e: { r: Math.min(full.e.r, full.s.r + dataRowLimit), c: Math.min(full.e.c, full.s.c + MAX_FIELDS - 1) } };
    const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false, range });
    let formulas = 0;
    for (let r = range.s.r; r <= range.e.r; r++) for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[window.XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.f === 'string') formulas++;
    }
    const truncated = full.e.r - full.s.r > dataRowLimit || fieldCount > MAX_FIELDS;
    const warning = [truncated ? `The worksheet exceeds the processing limit; this view is capped at ${dataRowLimit.toLocaleString()} data rows and ${MAX_FIELDS.toLocaleString()} fields.` : '', formulas ? `${formulas} formula cells were found in the imported range. Cached workbook values are read; formulas are not recalculated.` : ''].filter(Boolean).join(' ');
    parseMatrix(matrix, 'table', `${state.currentFileName} · ${sheetName}`, warning);
  }

  function showDocumentTables(tables) {
    state.documentTables = tables;
    $('#sheet-wrap label').textContent = 'Word table';
    const select = $('#sheet-select'); select.replaceChildren();
    tables.forEach((table, index) => optionFor(select, String(index), `Table ${index + 1} · ${table.length - 1} rows`));
    $('#sheet-wrap').hidden = tables.length < 2;
    parseDocumentTable(0);
  }

  function parseDocumentTable(index) {
    const table = state.documentTables[index];
    if (!table) return;
    const data = normalizeRows(table);
    if (!data) throw new Error('The Word table did not contain a header row and data.');
    setSource({ ...data, kind: 'document', label: `${state.currentFileName} · Table ${Number(index) + 1}`, warning: 'Extracted a DOCX table. Check merged cells, units, and column labels against the original document.' }, state.currentFileName);
  }

  function htmlToPlainText(html) {
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    const blocks = [];
    for (const node of parsed.body.children) {
      if (node.tagName === 'TABLE') {
        for (const row of node.rows) blocks.push([...row.cells].map((cell) => cell.textContent.replace(/\s+/g, ' ').trim()).join('\t'));
        blocks.push('');
        continue;
      }
      if (/^H[1-6]$/.test(node.tagName)) blocks.push(`${'#'.repeat(Number(node.tagName[1]))} ${node.textContent.replace(/\s+/g, ' ').trim()}`);
      else if (node.tagName === 'UL' || node.tagName === 'OL') [...node.querySelectorAll(':scope > li')].forEach((item) => blocks.push(`- ${item.textContent.replace(/\s+/g, ' ').trim()}`));
      else if (/^(P|BLOCKQUOTE|PRE)$/.test(node.tagName)) blocks.push(node.textContent.replace(/\s+/g, ' ').trim());
    }
    return blocks.join('\n');
  }

  async function processDocx(file) {
    await ensureMammoth();
    setMessage('Extracting document text and headings…');
    const result = await window.mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
    const documentTree = new DOMParser().parseFromString(result.value, 'text/html');
    const tables = [...documentTree.querySelectorAll('table')].map((table) => [...table.rows].map((row) => [...row.cells].map((cell) => cell.textContent.replace(/\s+/g, ' ').trim()))).filter((table) => table.length > 1 && table.some((row) => row.length > 1));
    const text = htmlToPlainText(result.value);
    if (!text.trim()) throw new Error('No selectable text was extracted from this Word file.');
    $('#data-text').value = text;
    if (tables.length) showDocumentTables(tables);
    else setSource({ ...parseSource(text), kind: 'document', label: file.name, warning: result.messages.length ? `${result.messages.length} document-structure notice(s) occurred during text extraction; review the extracted text.` : '' }, file.name);
    setMessage(tables.length ? `Extracted ${tables.length} table(s). Select a table and verify merged cells, units, and headers. The full text is available in the paste area.` : `Extracted document text. ${result.messages.length ? 'Review extraction notices; complex layout and tables may need manual cleanup.' : 'Check headings and lists before making a concept map.'}`, result.messages.length || tables.length ? 'caution' : 'good');
  }

  async function processPDF(file) {
    const pdfjs = await ensurePDF();
    setMessage('Reading selectable PDF text…');
    const bytes = new Uint8Array(await file.arrayBuffer());
    const pdf = await pdfjs.getDocument({ data: bytes, isEvalSupported: false }).promise;
    state.pdf = pdf;
    const pages = Math.min(pdf.numPages, MAX_PDF_PAGES);
    const chunks = [];
    let itemCount = 0;
    for (let pageNo = 1; pageNo <= pages; pageNo++) {
      const page = await pdf.getPage(pageNo);
      const content = await page.getTextContent();
      const items = content.items.filter((item) => typeof item.str === 'string' && item.str.trim());
      itemCount += items.length;
      const grouped = new Map();
      for (const item of items) {
        const y = Math.round(item.transform?.[5] || 0);
        if (!grouped.has(y)) grouped.set(y, []);
        grouped.get(y).push(item);
      }
      const lines = [...grouped.entries()].sort((a, b) => b[0] - a[0]).map(([, line]) => line.sort((a, b) => a.transform[4] - b.transform[4]).map((item) => item.str).join(' ').trim()).filter(Boolean);
      if (lines.length) chunks.push(`# Page ${pageNo}\n${lines.join('\n')}`);
    }
    const text = chunks.join('\n\n');
    $('#data-text').value = text;
    setSource({ ...parseSource(text || `PDF has ${pdf.numPages} pages but no selectable text.`), kind: 'pdf', label: file.name, warning: pdf.numPages > MAX_PDF_PAGES ? `PDF has ${pdf.numPages} pages; only the first ${MAX_PDF_PAGES} were read in this workspace.` : (!itemCount ? 'No selectable text was found. This may be a scanned document; OCR can attempt text extraction from the first pages.' : 'PDF text extraction preserves reading order approximately; tables and multi-column layouts should be checked against the source.') }, file.name);
    const scanned = itemCount === 0;
    state.ocrAvailable = scanned;
    $('#recognize-image').disabled = !scanned;
    $('#recognize-image').textContent = scanned ? 'OCR scanned PDF pages' : 'PDF has selectable text';
    setMessage(scanned ? `No selectable text found in ${pdf.numPages} page(s). OCR is available for up to 10 pages.` : `Read ${pages} PDF page(s). Verify columns and tables against the original PDF.`, scanned ? 'caution' : 'good');
  }

  async function processImage(file) {
    const bitmap = await createImageBitmap(file);
    const originalWidth = bitmap.width, originalHeight = bitmap.height;
    const scale = Math.min(1, 128 / bitmap.width, 128 / bitmap.height, Math.sqrt(10000 / (bitmap.width * bitmap.height)));
    const width = Math.max(1, Math.round(bitmap.width * scale)), height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    context.drawImage(bitmap, 0, 0, width, height);
    const pixels = context.getImageData(0, 0, width, height).data;
    const rows = [];
    const grid = Array.from({ length: height }, () => Array(width));
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      const offset = (y * width + x) * 4, red = pixels[offset], green = pixels[offset + 1], blue = pixels[offset + 2];
      const max = Math.max(red, green, blue), min = Math.min(red, green, blue), delta = max - min;
      const brightness = .2126 * red + .7152 * green + .0722 * blue;
      let hue = 0;
      if (delta) { if (max === red) hue = 60 * (((green - blue) / delta) % 6); else if (max === green) hue = 60 * ((blue - red) / delta + 2); else hue = 60 * ((red - green) / delta + 4); }
      const saturation = max ? 100 * delta / max : 0;
      const row = { x, y, red, green, blue, brightness: Number(brightness.toFixed(2)), hue: Number(((hue + 360) % 360).toFixed(2)), saturation: Number(saturation.toFixed(2)) };
      rows.push(row); grid[y][x] = row.brightness;
    }
    bitmap.close();
    state.imageFile = file; state.imageGrid = grid;
    state.ocrAvailable = true;
    if (state.imageURL) URL.revokeObjectURL(state.imageURL);
    state.imageURL = URL.createObjectURL(file);
    $('#source-preview').innerHTML = `<b>${escapeHTML(file.name)}</b> · ${originalWidth} × ${originalHeight} px<br><img src="${state.imageURL}" alt="Selected image preview"><button class="btn" id="reprofile-image" type="button">Restore pixel analysis</button>`;
    $('#source-preview').hidden = false;
    $('#reprofile-image').addEventListener('click', () => profileImageSource(file, rows, width, height));
    $('#recognize-image').disabled = false; $('#recognize-image').textContent = 'Read text with OCR';
    const headers = ['x', 'y', 'red', 'green', 'blue', 'brightness', 'hue', 'saturation'];
    setSource({ rows, headers, kind: 'image', label: file.name, imageWidth: width, imageHeight: height, originalWidth, originalHeight, warning: `Image measurements use a ${width} × ${height} downsampled grid (${rows.length.toLocaleString()} pixels). This is a pixel profile, not object or scientific-image interpretation.` }, file.name);
    setMessage(`Sampled ${rows.length.toLocaleString()} pixels from a ${originalWidth} × ${originalHeight} image. OCR can read text, but does not digitize plotted values.`, 'caution');
  }

  function profileImageSource(file, rows, width, height) {
    const headers = ['x', 'y', 'red', 'green', 'blue', 'brightness', 'hue', 'saturation'];
    setSource({ rows, headers, kind: 'image', label: file.name, imageWidth: width, imageHeight: height, warning: `Image measurements use a ${width} × ${height} downsampled grid (${rows.length.toLocaleString()} pixels). This is a pixel profile, not object or scientific-image interpretation.` }, file.name);
  }

  async function processFile(file) {
    if (!file) return;
    if (file.size > MAX_FILE_BYTES) { setMessage('This file exceeds the 25 MB processing limit. Reduce or split it, then try again.', 'caution'); return; }
    const previous = { currentFileName: state.currentFileName, workbook: state.workbook, documentTables: state.documentTables, pdf: state.pdf, imageFile: state.imageFile, ocrAvailable: state.ocrAvailable, imageURL: state.imageURL, sourceFileInfo: state.sourceFileInfo };
    state.currentFileName = file.name;
    state.sourceFileInfo = { name: file.name, type: file.type || 'unknown', sizeBytes: file.size, lastModified: file.lastModified ? new Date(file.lastModified).toISOString() : null, sha256: null };
    state.workbook = null; state.documentTables = []; state.pdf = null; state.imageFile = null; state.ocrAvailable = false;
    $('#source-preview').hidden = true;
    $('#sheet-wrap').hidden = true;
    $('#recognize-image').disabled = true;
    try {
      const ext = file.name.toLowerCase().split('.').pop();
      if (['csv', 'tsv'].includes(ext)) {
        const text = await file.text(), parsed = parseSource(text);
        $('#data-text').value = text;
        setSource({ ...parsed, label: file.name }, file.name);
        setMessage(`Imported ${state.rows.length.toLocaleString()} rows. Check the inferred field types and missing-value report.`, 'good');
      } else if (ext === 'json') {
        const text = await file.text(); $('#data-text').value = text;
        setSource({ ...parseSource(text), label: file.name }, file.name);
        setMessage('Parsed JSON locally. Confirm the selected record structure before interpreting the figures.', 'good');
      } else if (['txt', 'md'].includes(ext)) {
        const text = await file.text(); $('#data-text').value = text;
        setSource({ ...parseSource(text), label: file.name }, file.name);
        setMessage('Imported text. Headings and indentation inform the structure views; automatic scientific interpretation is not performed.');
      } else if (['xlsx', 'xls', 'xlsm', 'ods'].includes(ext)) {
        await ensureXLSX(); setMessage('Reading workbook sheets…');
        const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellFormula: true, dense: true });
        showWorkbookSheets(workbook);
        setMessage(`${workbook.SheetNames.length} worksheet(s) found. Choose a sheet and check cached formula values, dates, and missing cells.`, 'good');
      } else if (ext === 'docx') {
        await processDocx(file);
      } else if (ext === 'pdf') {
        await processPDF(file);
      } else if (['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif'].includes(ext)) {
        await processImage(file);
      } else if (ext === 'doc') {
        throw new Error('Legacy .doc files are not supported. Save the document as .docx or export its text to .txt.');
      } else throw new Error(`The .${escapeHTML(ext)} format is not supported. Try CSV, TSV, JSON, XLSX, XLS, ODS, DOCX, PDF, a common raster image, TXT, or Markdown.`);
      if (window.crypto?.subtle) {
        try {
          const digest = await window.crypto.subtle.digest('SHA-256', await file.arrayBuffer());
          state.sourceFileInfo.sha256 = [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
        } catch { /* Hash is optional; file parsing and the analysis record remain available. */ }
      }
      syncWorkflow();
      if (previous.imageURL && state.imageURL === previous.imageURL) { URL.revokeObjectURL(previous.imageURL); state.imageURL = ''; }
    } catch (error) {
      state.currentFileName = previous.currentFileName; state.workbook = previous.workbook; state.documentTables = previous.documentTables; state.pdf = previous.pdf; state.imageFile = previous.imageFile; state.ocrAvailable = previous.ocrAvailable; state.imageURL = previous.imageURL; state.sourceFileInfo = previous.sourceFileInfo;
      $('#recognize-image').disabled = !state.ocrAvailable;
      $('#diagnosis').textContent = error.message || 'This file could not be analyzed.';
      setMessage('Import failed. The original file has not been changed.', 'caution');
    }
  }

  function median(values) { const a = values.slice().sort((x, y) => x - y); return quantile(a, .5); }
  function valuesFor(field) { return state.rows.map((row) => toNumber(row[field])).filter(Number.isFinite); }
  function toNumber(value) { if (typeof value === 'number') return value; if (value === null || value === undefined || String(value).trim() === '') return NaN; const normalized = String(value).trim().replace(/,/g, ''); return /^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(normalized) ? Number(normalized) : NaN; }
  function selectedFields() { return { x: $('#axis-x').value, y: $('#axis-y').value, group: $('#axis-group').value, aggregation: $('#aggregate').value }; }
  function rowsForChart() {
    const max = MAX_CHART_POINTS, step = Math.max(1, Math.ceil(state.rows.length / max));
    if (step === 1) return state.rows;
    return state.rows.filter((_, index) => index % step === 0);
  }
  function groupRows(field) {
    const groups = new Map();
    for (const row of state.rows) { const label = field ? String(row[field] ?? '(missing)') : 'All observations'; if (!groups.has(label)) groups.set(label, []); groups.get(label).push(row); }
    return [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }));
  }
  function summarize(values, kind) {
    const clean = values.filter(Number.isFinite);
    if (kind === 'count') return clean.length;
    if (!clean.length) return NaN;
    if (kind === 'sum') return clean.reduce((a, b) => a + b, 0);
    if (kind === 'median') return median(clean);
    return clean.reduce((a, b) => a + b, 0) / clean.length;
  }
  function aggregateCategories(xField, yField, groupField, aggregation) {
    const map = new Map();
    for (const row of state.rows) {
      const category = String(row[xField] ?? '(missing)'), group = groupField ? String(row[groupField] ?? '(missing)') : 'Series';
      const key = `${category}\u0000${group}`;
      if (!map.has(key)) map.set(key, { category, group, values: [] });
      const value = aggregation === 'count' ? 1 : toNumber(row[yField]);
      if (Number.isFinite(value)) map.get(key).values.push(value);
    }
    return [...map.values()].map((entry) => ({ ...entry, value: summarize(entry.values, aggregation) })).filter((entry) => Number.isFinite(entry.value));
  }
  function cappedGroups(groups) {
    if (groups.length <= 10) return groups;
    const main = groups.slice(0, 9), rest = groups.slice(9).flatMap(([, rows]) => rows);
    return [...main, ['Other groups', rest]];
  }
  function seriesByGroup(rows, groupField) {
    const groups = new Map();
    for (const row of rows) { const name = groupField ? String(row[groupField] ?? '(missing)') : 'Series'; if (!groups.has(name)) groups.set(name, []); groups.get(name).push(row); }
    return cappedGroups([...groups.entries()].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0])));
  }
  function chartPalette() { return palettes[$('#palette').value] || palettes.zstack; }
  function figureStyle() {
    const px = Number($('#text-size').value) || 14;
    return { color: '#1b2e38', backgroundColor: '#ffffff', textStyle: { fontFamily: 'Arial, Helvetica, sans-serif', fontSize: px, color: '#1b2e38' }, animation: false, grid: { left: 76, right: 34, top: 88, bottom: 72, containLabel: true }, toolbox: { show: false }, tooltip: { trigger: 'axis', confine: true, backgroundColor: 'rgba(23,45,57,.94)', borderWidth: 0, textStyle: { color: '#fff', fontSize: px - 1 } }, color: chartPalette() };
  }
  function pearson(x, y) {
    const pairs = x.map((value, i) => [value, y[i]]).filter((pair) => pair.every(Number.isFinite));
    if (pairs.length < 3) return NaN;
    const mx = pairs.reduce((s, p) => s + p[0], 0) / pairs.length, my = pairs.reduce((s, p) => s + p[1], 0) / pairs.length;
    let cov = 0, vx = 0, vy = 0;
    for (const [a, b] of pairs) { cov += (a - mx) * (b - my); vx += (a - mx) ** 2; vy += (b - my) ** 2; }
    return vx && vy ? cov / Math.sqrt(vx * vy) : NaN;
  }
  function histogram(values) {
    const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
    if (!sorted.length) return { labels: [], counts: [] };
    const q1 = quantile(sorted, .25), q3 = quantile(sorted, .75), range = sorted.at(-1) - sorted[0];
    if (range === 0) return { labels: [formatNum(sorted[0])], counts: [sorted.length], binWidth: 0 };
    const fd = range && q3 > q1 ? Math.ceil(range / (2 * (q3 - q1) / Math.cbrt(sorted.length))) : Math.ceil(Math.sqrt(sorted.length));
    const bins = Math.max(5, Math.min(60, Number.isFinite(fd) && fd > 0 ? fd : 8));
    const width = range ? range / bins : 1, counts = Array(bins).fill(0);
    for (const value of sorted) counts[Math.min(bins - 1, Math.floor((value - sorted[0]) / width))]++;
    const labels = counts.map((_, i) => `${formatNum(sorted[0] + i * width)}–${formatNum(sorted[0] + (i + 1) * width)}`);
    return { labels, counts, binWidth: width };
  }
  function figureTitle() { return $('#figure-title').value.trim() || `${vizMeta[state.selected]?.[0] || 'Research figure'} · ${state.currentFileName}`; }
  function buildChartOption(type, width = 1100, height = 700) {
    const { x, y, group, aggregation } = selectedFields(), rows = rowsForChart(), style = figureStyle(), title = figureTitle();
    const xName = $('#label-x').value.trim() || x || 'Category', yName = $('#label-y').value.trim() || y || 'Value';
    const numericX = state.types.numeric.includes(x), numericY = state.types.numeric.includes(y);
    style.title = { text: title, left: 'center', top: 16, textStyle: { fontSize: Number($('#text-size').value) + 2, fontWeight: 'normal', color: '#172d39' } };
    style.aria = { enabled: true, decal: { show: true } };
    const categoryAxis = (name) => ({ type: 'category', name, nameLocation: 'middle', nameGap: 44, axisLabel: { color: '#536970', hideOverlap: true, rotate: 0 }, axisLine: { lineStyle: { color: '#a9b7b2' } }, axisTick: { alignWithLabel: true }, splitLine: { show: false } });
    const valueAxis = (name) => ({ type: 'value', name, nameLocation: 'middle', nameGap: 52, axisLabel: { color: '#536970' }, axisLine: { show: false }, splitLine: { lineStyle: { color: '#e7ece9' } }, scale: true });
    const groups = seriesByGroup(rows, group);

    if (['line', 'area'].includes(type)) {
      const xValues = rows.map((row) => row[x]).filter((value) => value !== '' && value !== undefined);
      const categoricalX = !numericX || /date|time/i.test(x);
      const categories = [...new Set(xValues.map(String))];
      const orderedRows = rows.slice().sort((a, b) => categoricalX ? categories.indexOf(String(a[x])) - categories.indexOf(String(b[x])) : toNumber(a[x]) - toNumber(b[x]));
      const lines = seriesByGroup(orderedRows, group).slice(0, 10);
      style.grid = { left: 90, right: 40, top: 90, bottom: 90, containLabel: true };
      style.xAxis = categoricalX ? categoryAxis(xName) : { ...valueAxis(xName), boundaryGap: false };
      style.yAxis = valueAxis(yName);
      if (categoricalX) style.xAxis.data = categories;
      style.dataZoom = rows.length > 200 ? [{ type: 'inside', throttle: 80 }, { type: 'slider', height: 18, bottom: 22 }] : [];
      style.legend = lines.length > 1 ? { top: 54, type: 'scroll' } : { show: false };
      style.series = lines.map(([name, items], index) => ({ name, type: 'line', smooth: false, showSymbol: items.length < 80, symbolSize: 6, connectNulls: false, areaStyle: type === 'area' ? { opacity: .16 } : undefined, lineStyle: { width: 2.2 }, data: items.map((row) => [categoricalX ? String(row[x]) : toNumber(row[x]), toNumber(row[y])]) }));
      return style;
    }
    if (type === 'scatter' || type === 'bubble') {
      const plotted = rows.filter((row) => Number.isFinite(toNumber(row[x])) && Number.isFinite(toNumber(row[y])));
      const subsets = seriesByGroup(plotted, group);
      style.grid = { left: 90, right: 42, top: 78, bottom: 82, containLabel: true };
      style.xAxis = valueAxis(xName); style.yAxis = valueAxis(yName);
      style.dataZoom = [{ type: 'inside' }, { type: 'slider', height: 18, bottom: 20 }];
      style.legend = subsets.length > 1 ? { top: 46, type: 'scroll' } : { show: false };
      style.tooltip = { ...style.tooltip, trigger: 'item', formatter: (params) => `${escapeHTML(params.seriesName)}<br>${escapeHTML(xName)}: ${formatNum(params.value[0])}<br>${escapeHTML(yName)}: ${formatNum(params.value[1])}` };
      style.series = subsets.map(([name, items]) => ({ name, type: 'scatter', symbolSize: 8, data: items.map((row) => [toNumber(row[x]), toNumber(row[y]), type === 'bubble' ? toNumber(row[state.types.numeric.find((field) => field !== x && field !== y)]) : 0]), large: items.length > 1000, largeThreshold: 1000, itemStyle: { opacity: .72 } }));
      if (type === 'bubble') {
        const sizeField = state.types.numeric.find((field) => field !== x && field !== y), sizes = valuesFor(sizeField), minSize = Math.min(...sizes), maxSize = Math.max(...sizes);
        style.visualMap = { show: false, dimension: 2, min: Number.isFinite(minSize) ? minSize : 0, max: Number.isFinite(maxSize) && maxSize > minSize ? maxSize : (minSize || 1) + 1, inRange: { symbolSize: [6, 26] } };
      }
      return style;
    }
    if (type === 'bar' || type === 'groupedbar' || type === 'pareto') {
      let aggregates = aggregateCategories(x, y, type === 'groupedbar' ? group : '', aggregation === 'count' ? 'count' : aggregation);
      let paretoTotal = 0;
      if (type === 'pareto') {
        const allContributions = aggregateCategories(x, y, '', aggregation === 'count' ? 'count' : 'sum').filter((entry) => entry.value > 0).sort((a, b) => b.value - a.value);
        paretoTotal = allContributions.reduce((sum, entry) => sum + entry.value, 0);
        aggregates = allContributions.slice(0, 30);
      }
      const allCategories = [...new Set(aggregates.map((entry) => entry.category))];
      const categoryTotals = allCategories.map((category) => [category, aggregates.filter((entry) => entry.category === category).reduce((sum, entry) => sum + Math.abs(entry.value), 0)]);
      const categories = allCategories.length > 40 ? categoryTotals.sort((a, b) => b[1] - a[1]).slice(0, 40).map(([category]) => category) : allCategories;
      const seriesNames = [...new Set(aggregates.map((entry) => entry.group))].slice(0, 10);
      style.grid = { left: 90, right: type === 'pareto' ? 72 : 36, top: 84, bottom: 96, containLabel: true };
      style.xAxis = { ...categoryAxis(xName), data: categories, axisLabel: { rotate: categories.length > 10 ? 35 : 0, interval: 'auto', color: '#536970' } };
      style.yAxis = type === 'pareto' ? [valueAxis(yName), { ...valueAxis('Cumulative share'), min: 0, max: 100, axisLabel: { formatter: '{value}%' }, splitLine: { show: false } }] : valueAxis(aggregation === 'count' ? 'Count' : `${aggregation} ${yName}`);
      style.dataZoom = categories.length > 20 ? [{ type: 'inside' }, { type: 'slider', height: 18, bottom: 20 }] : [];
      if (type === 'pareto') {
        const total = paretoTotal || 1; let cumulative = 0;
        style.legend = { top: 48 };
        style.series = [{ name: 'Contribution', type: 'bar', data: categories.map((category) => aggregates.find((entry) => entry.category === category)?.value ?? 0), barMaxWidth: 42 }, { name: 'Cumulative share', type: 'line', yAxisIndex: 1, data: categories.map((category) => { cumulative += aggregates.find((entry) => entry.category === category)?.value ?? 0; return cumulative / total * 100; }), symbolSize: 6, lineStyle: { width: 2 } }];
      } else {
        style.legend = seriesNames.length > 1 ? { top: 48, type: 'scroll' } : { show: false };
        style.series = seriesNames.map((name) => ({ name, type: 'bar', barMaxWidth: 42, data: categories.map((category) => aggregates.find((entry) => entry.category === category && entry.group === name)?.value ?? null) }));
      }
      return style;
    }
    if (type === 'hist') {
      const values = valuesFor(y), hist = histogram(values);
      style.grid = { left: 80, right: 30, top: 80, bottom: 78, containLabel: true };
      style.xAxis = { ...categoryAxis(`Bins · ${yName}`), data: hist.labels, axisLabel: { rotate: hist.labels.length > 12 ? 35 : 0, hideOverlap: true, color: '#536970' } };
      style.yAxis = valueAxis('Observations'); style.tooltip = { ...style.tooltip, trigger: 'axis' };
      style.series = [{ name: 'Observations', type: 'bar', data: hist.counts, barGap: '0%', itemStyle: { borderColor: '#fff', borderWidth: 1 } }];
      return style;
    }
    if (type === 'ecdf') {
      const subsets = seriesByGroup(rows, group), series = subsets.map(([name, items]) => {
        const values = items.map((row) => toNumber(row[y])).filter(Number.isFinite).sort((a, b) => a - b);
        const stride = Math.max(1, Math.ceil(values.length / MAX_CHART_POINTS)), data = [];
        let lastIndex = -1;
        for (let index = stride - 1; index < values.length; index += stride) { data.push([values[index], (index + 1) / values.length]); lastIndex = index; }
        if (values.length && lastIndex !== values.length - 1) data.push([values.at(-1), 1]);
        return { name, type: 'line', showSymbol: false, step: 'end', data };
      });
      style.grid = { left: 90, right: 35, top: 80, bottom: 78, containLabel: true };
      style.xAxis = valueAxis(yName); style.yAxis = { ...valueAxis('Cumulative proportion'), min: 0, max: 1, axisLabel: { formatter: (value) => `${Math.round(value * 100)}%` } };
      style.legend = series.length > 1 ? { top: 46, type: 'scroll' } : { show: false }; style.series = series; return style;
    }
    if (type === 'box') {
      const subsets = group ? groupRows(group) : [['All observations', state.rows]];
      const data = subsets.slice(0, 20).map(([name, items]) => {
        const values = items.map((row) => toNumber(row[y])).filter(Number.isFinite).sort((a, b) => a - b);
        return { name, values: values.length ? [values[0], quantile(values, .25), quantile(values, .5), quantile(values, .75), values.at(-1)] : null };
      }).filter((item) => item.values);
      style.grid = { left: 76, right: 28, top: 80, bottom: 82, containLabel: true };
      style.xAxis = { ...categoryAxis(group || 'Dataset'), data: data.map((item) => item.name) }; style.yAxis = valueAxis(yName);
      style.tooltip = { ...style.tooltip, trigger: 'item' }; style.series = [{ name: 'Five-number summary', type: 'boxplot', data: data.map((item) => item.values), itemStyle: { borderWidth: 1.5 } }];
      return style;
    }
    if (type === 'heat') {
      const fields = state.types.numeric.slice(0, 20), values = fields.map((rowField, rowIndex) => fields.map((colField, colIndex) => [colIndex, rowIndex, pearson(state.rows.map((row) => toNumber(row[colField])), state.rows.map((row) => toNumber(row[rowField])))]).filter((cell) => Number.isFinite(cell[2]))).flat();
      style.grid = { left: 112, right: 72, top: 72, bottom: 112, containLabel: true };
      style.tooltip = { ...style.tooltip, position: 'top', formatter: (params) => `${escapeHTML(fields[params.value[1]])} × ${escapeHTML(fields[params.value[0]])}<br>r = ${Number(params.value[2]).toFixed(3)}` };
      style.xAxis = { ...categoryAxis(''), data: fields, splitArea: { show: true }, axisLabel: { rotate: 45, color: '#536970' } };
      style.yAxis = { ...categoryAxis(''), data: fields, inverse: true, splitArea: { show: true } };
      style.visualMap = { min: -1, max: 1, calculable: false, orient: 'vertical', right: 0, top: 'middle', inRange: { color: ['#3b5b92', '#f7f8f7', '#bd675b'] }, text: ['+1', '-1'] };
      style.series = [{ type: 'heatmap', data: values, label: { show: fields.length <= 12, formatter: (params) => Number(params.value[2]).toFixed(2), fontSize: Number($('#text-size').value) - 2 }, emphasis: { itemStyle: { shadowBlur: 5, shadowColor: 'rgba(0,0,0,.2)' } } }];
      return style;
    }
    if (type === 'pixelheat') {
      const imageRows = state.rows, gridW = state.source.imageWidth, gridH = state.source.imageHeight;
      style.grid = { left: 52, right: 92, top: 72, bottom: 55 };
      style.tooltip = { ...style.tooltip, formatter: (params) => `x ${params.value[0]}, y ${params.value[1]}<br>Brightness ${Number(params.value[2]).toFixed(1)}` };
      style.xAxis = { type: 'value', min: 0, max: gridW - 1, name: 'Pixel x', nameLocation: 'middle', nameGap: 32, splitLine: { show: false } };
      style.yAxis = { type: 'value', min: 0, max: gridH - 1, inverse: true, name: 'Pixel y', nameLocation: 'middle', nameGap: 34, splitLine: { show: false } };
      style.visualMap = { min: 0, max: 255, calculable: false, orient: 'vertical', right: 0, top: 'middle', inRange: { color: ['#111820', '#3c6470', '#e0bf79', '#fff8e8'] }, text: ['light', 'dark'] };
      style.series = [{ type: 'heatmap', data: imageRows.map((row) => [row.x, row.y, row.brightness]), progressive: 2000, progressiveThreshold: 5000, itemStyle: { borderWidth: 0 } }];
      return style;
    }
    if (type === 'mindmap') {
      const tree = notesTree(); style.grid = { left: 20, right: 120, top: 66, bottom: 20 };
      style.tooltip = { trigger: 'item', triggerOn: 'mousemove', formatter: (params) => escapeHTML(params.data.name) };
      style.series = [{ type: 'tree', data: [tree], top: 55, left: 32, bottom: 20, right: 120, orient: 'LR', symbol: 'emptyCircle', symbolSize: 8, expandAndCollapse: false, roam: true, initialTreeDepth: -1, lineStyle: { color: '#9caaa6', width: 1.3 }, label: { position: 'left', verticalAlign: 'middle', align: 'right', fontSize: 12, width: 200, overflow: 'truncate' }, leaves: { label: { position: 'right', align: 'left' } }, emphasis: { focus: 'descendant' }, animationDuration: 0 }];
      return style;
    }
    if (type === 'flow') {
      const labels = state.rows.map((row) => String(row.line ?? row[state.headers[0]] ?? '')).filter(Boolean).slice(0, 24);
      const usableWidth = Math.max(600, width - 170), usableHeight = Math.max(360, height - 180), nodes = labels.map((name, index) => ({ id: String(index), name: name.slice(0, 90), x: 60 + index * usableWidth / Math.max(1, labels.length - 1), y: usableHeight / 2, symbolSize: [Math.min(180, Math.max(110, width / Math.max(4, labels.length) - 10)), 58], itemStyle: { color: index === 0 ? '#e7f1ed' : '#f6f8f7', borderColor: '#16766d' }, label: { color: '#172d39', overflow: 'truncate', width: 150, lineHeight: 16 } }));
      style.grid = { left: 20, right: 20, top: 70, bottom: 25 };
      style.tooltip = { trigger: 'item', formatter: (params) => escapeHTML(params.data.name) };
      style.series = [{ type: 'graph', layout: 'none', coordinateSystem: null, roam: true, data: nodes, edges: nodes.slice(1).map((node, index) => ({ source: String(index), target: node.id })), edgeSymbol: ['none', 'arrow'], edgeSymbolSize: 9, lineStyle: { color: '#bd9252', width: 2 }, label: { show: true, position: 'inside', fontSize: 11, lineHeight: 15, formatter: (params) => params.data.name }, emphasis: { focus: 'adjacency' } }];
      return style;
    }
    return style;
  }

  function notesTree() {
    const root = { name: state.currentFileName || 'Source notes', children: [] }, stack = [{ level: -1, node: root }];
    for (const row of state.rows.slice(0, 500)) {
      const text = String(row.line ?? row[state.headers[0]] ?? '').trim(); if (!text) continue;
      const level = Math.max(0, Math.min(6, Number(row.level) || 0));
      const node = { name: text.slice(0, 150), children: [] };
      while (stack.length > 1 && stack.at(-1).level >= level) stack.pop();
      stack.at(-1).node.children.push(node); stack.push({ level, node });
    }
    if (!root.children.length) root.children = state.types.categorical.slice(0, 15).map((field) => ({ name: field, children: [...new Set(state.rows.map((row) => String(row[field] ?? '')).filter(Boolean))].slice(0, 10).map((value) => ({ name: value })) }));
    return root;
  }

  function renderTable() {
    const wrap = document.createElement('div'); wrap.className = 'data-table-wrap';
    const table = document.createElement('table'); table.className = 'data-table';
    const head = document.createElement('thead'), tr = document.createElement('tr');
    state.headers.forEach((field) => { const th = document.createElement('th'); th.textContent = field; tr.appendChild(th); }); head.appendChild(tr); table.appendChild(head);
    const body = document.createElement('tbody');
    state.rows.slice(0, 100).forEach((row) => { const line = document.createElement('tr'); state.headers.forEach((field) => { const td = document.createElement('td'); td.textContent = String(row[field] ?? ''); line.appendChild(td); }); body.appendChild(line); });
    table.appendChild(body); wrap.appendChild(table);
    if (state.rows.length > 100) { const note = document.createElement('p'); note.className = 'microcopy'; note.textContent = `Showing the first 100 of ${state.rows.length.toLocaleString()} rows. Data exports include the full imported dataset.`; wrap.appendChild(note); }
    $('#render-area').replaceChildren(wrap);
  }

  async function renderFigure() {
    const token = ++state.renderToken;
    $('#figure-caption').hidden = false;
    const notes = [];
    if (state.rows.length > MAX_CHART_POINTS && ['line', 'area', 'scatter', 'bubble', 'ecdf'].includes(state.selected)) notes.push(`The plotted series is deterministically sampled to about ${MAX_CHART_POINTS.toLocaleString()} points for responsiveness; data exports retain all ${state.rows.length.toLocaleString()} rows.`);
    const groupField = selectedFields().group;
    if (['line', 'area'].includes(state.selected) && groupField && new Set(state.rows.map((row) => String(row[groupField] ?? ''))).size > 10) notes.push('Only the 10 largest groups are displayed in this line view.');
    if (state.selected === 'pareto' && new Set(state.rows.map((row) => String(row[selectedFields().x] ?? ''))).size > 30) notes.push('The figure displays the 30 largest positive categories; cumulative share uses all positive categories.');
    if (['bar', 'groupedbar'].includes(state.selected) && new Set(state.rows.map((row) => String(row[selectedFields().x] ?? ''))).size > 40) notes.push('The figure displays the 40 categories with the largest aggregate values.');
    if (state.selected === 'groupedbar' && groupField && new Set(state.rows.map((row) => String(row[groupField] ?? '(missing)'))).size > 10) notes.push('The 10 largest groups are displayed; smaller groups are combined as “Other groups”.');
    if (state.selected === 'box' && groupField && new Set(state.rows.map((row) => String(row[groupField] ?? ''))).size > 20) notes.push('Only the first 20 groups are shown in this box plot.');
    notes.push('Missing values are excluded where calculations require numeric values; source rows are not modified.');
    $('#figure-caption').textContent = `${vizMeta[state.selected][0]} · ${state.recommendations.find((view) => view.key === state.selected)?.rationale || 'user-selected view'}. ${notes.join(' ')}`;
    $('#recommendation-caption').textContent = `Selected: ${vizMeta[state.selected][0]}. ${state.recommendations.find((view) => view.key === state.selected)?.rationale || ''}`;
    if (state.selected === 'table') { if (state.chart) { state.chart.dispose(); state.chart = null; } renderTable(); return; }
    try {
      await ensureECharts();
      if (token !== state.renderToken) return;
      if (state.chart) { state.chart.dispose(); state.chart = null; }
      const root = document.createElement('div'); root.id = 'figure-chart'; root.setAttribute('role', 'img'); root.setAttribute('aria-label', `${vizMeta[state.selected][0]} for ${state.currentFileName}`);
      $('#render-area').replaceChildren(root);
      const renderer = state.rows.length > 2000 || state.selected === 'pixelheat' ? 'canvas' : 'svg';
      state.chart = window.echarts.init(root, null, { renderer, useDirtyRect: true });
      state.chart.setOption(buildChartOption(state.selected), { notMerge: true, lazyUpdate: true });
      requestAnimationFrame(() => state.chart?.resize());
    } catch (error) {
      $('#render-area').textContent = `Figure could not be rendered: ${error.message}`;
    }
  }

  function csvText(delimiter = ',') {
    const quote = (value) => { const text = String(value ?? ''); return text.includes(delimiter) || /["\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text; };
    return [state.headers, ...state.rows.map((row) => state.headers.map((field) => row[field]))].map((row) => row.map(quote).join(delimiter)).join('\r\n');
  }
  function download(name, content, mime) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob), anchor = document.createElement('a');
    anchor.href = url; anchor.download = name.replace(/[\\/:*?"<>|]+/g, '-'); document.body.appendChild(anchor); anchor.click(); anchor.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
  }
  function exportSettings() {
    return { title: figureTitle(), chart: state.selected, axes: selectedFields(), labels: { x: $('#label-x').value, y: $('#label-y').value }, aggregation: $('#aggregate').value, palette: $('#palette').value, textSize: Number($('#text-size').value), width: Math.min(6000, Math.max(600, Number($('#figure-width').value) || 1600)), height: Math.min(5000, Math.max(400, Number($('#figure-height').value) || 1000)), source: state.currentFileName, rows: state.rows.length, fields: state.headers, missingCells: state.profile.missing, duplicateRows: state.profile.duplicates, deterministicPointLimit: MAX_CHART_POINTS, software: 'Research Data & Figure Studio' };
  }

  function workflowPlan() {
    if (!state.profile || !state.types) return null;
    const profile = state.profile, selected = vizMeta[state.selected]?.[0] || 'Data table';
    const x = $('#axis-x')?.value || '', y = $('#axis-y')?.value || '', group = $('#axis-group')?.value || '';
    const intent = $('#workflow-question')?.value.trim() || $('#intent')?.selectedOptions[0]?.textContent || 'Not supplied';
    const truncation = state.source?.truncated ? ' The imported view was capped; repeat the analysis on the complete source or document the scope of this subset.' : '';
    const mixed = state.types.mixed?.length ? ` Mixed fields need review: ${state.types.mixed.slice(0, 8).join(', ')}.` : '';
    const pixel = state.source?.kind === 'image' ? ' Image-derived values are a downsampled pixel profile and do not represent object recognition or digitized research measurements.' : '';
    return [
      `Confirm the extraction: ${profile.rows.toLocaleString()} rows and ${profile.fields} fields from “${state.currentFileName}”. Review the preview, labels, dates, and units.${truncation}`,
      `Assess data quality before plotting: ${profile.missing.toLocaleString()} missing cells and ${profile.duplicates.toLocaleString()} exact duplicate rows were detected. No rows are removed and no values are imputed.${mixed}`,
      `State the research purpose: “${intent}”. Confirm the independent observation unit, study design, and whether the measurements are comparable; these are not inferred from file structure.`,
      `Inspect the selected ${selected}: horizontal field “${x || 'not applicable'}”, vertical field “${y || 'not applicable'}”, grouping “${group || 'none'}”, aggregation “${$('#aggregate')?.value || 'none'}”. Check that these choices match the question and the source units.`,
      `Review numeric screening flags and distribution shape, then decide and document any exclusions or transformations outside this tool. The detected outlier counts are exploratory 1.5×IQR flags, not error labels or inferential results.`,
      `Export the figure and its configuration record together. Retain the original source, record preprocessing and statistical methods separately, and verify all reported results with an appropriate validated analysis. ${pixel}`
    ];
  }

  function syncWorkflow() {
    const plan = workflowPlan(), list = $('#workflow-plan-list');
    if (!list) return;
    const ready = !!plan;
    $('#workflow-state').textContent = ready ? 'SOURCE PROFILED · REVIEW REQUIRED' : 'WAITING FOR DATA';
    $('#workflow-summary').textContent = ready
      ? `${state.currentFileName} · ${state.profile.rows.toLocaleString()} rows · ${state.profile.fields} fields · current view: ${vizMeta[state.selected]?.[0] || 'Data table'}. Generated from the current workspace profile.`
      : 'Load and profile material in Data Studio to prepare a source-specific plan.';
    if (ready) {
      list.replaceChildren(...plan.map((text) => { const li = document.createElement('li'); li.textContent = text; return li; }));
    } else {
      list.innerHTML = '<li>Import a source in Data Studio and inspect what was extracted.</li><li>Choose the question and analysis unit before interpreting a visualization.</li><li>Review the generated figure and export the analysis record.</li>';
    }
    $('#download-workflow-json').disabled = !ready;
    $('#download-workflow-md').disabled = !ready;
  }

  function makeWorkflowRecord() {
    const settings = exportSettings();
    const checks = Object.fromEntries([...document.querySelectorAll('[data-review]')].map((input) => [input.dataset.review, input.checked]));
    return {
      schema: 'research-workflow-record/v1',
      project: $('#workflow-study').value.trim() || null,
      researchQuestion: $('#workflow-question').value.trim() || null,
      researcherNotes: {
        preprocessingAndExclusions: $('#workflow-processing').value.trim() || null,
        statisticalMethodsAndSoftware: $('#workflow-statistics').value.trim() || null
      },
      generatedAt: new Date().toISOString(),
      source: {
        file: state.currentFileName,
        kind: state.source?.kind || 'pasted text',
        bytes: state.sourceFileInfo?.sizeBytes ?? null,
        mimeType: state.sourceFileInfo?.type ?? null,
        lastModified: state.sourceFileInfo?.lastModified ?? null,
        sha256: state.sourceFileInfo?.sha256 ?? null,
        rows: state.profile.rows,
        fields: state.headers,
        truncated: !!state.source?.truncated
      },
      profile: {
        numericFields: state.types.numeric,
        categoricalFields: state.types.categorical,
        mixedFields: state.types.mixed,
        missingCells: state.profile.missing,
        duplicateRows: state.profile.duplicates,
        outlierScreeningCountsByField: Object.fromEntries(Object.entries(state.types.summary).map(([field, summary]) => [field, summary.outliers]))
      },
      purpose: $('#intent').value,
      figure: { ...settings, recommendationRationale: state.recommendations.find((view) => view.key === state.selected)?.rationale || null },
      analysisPlan: workflowPlan(),
      researcherChecklist: checks,
      caveats: [
        'The profile uses imported values and does not establish scientific validity.',
        'Missing values are not imputed; duplicate rows are retained.',
        'The outlier screen uses 1.5×IQR fences and is not a data-error classifier.',
        'No hypothesis test, confidence interval, causal estimate, or inferential conclusion is generated.',
        'Researcher checklist values are self-reported and are not machine verified.',
        ...(state.source?.warning ? [state.source.warning] : [])
      ],
      software: { name: 'Differential Evolution Visualizer · Research Data & Figure Studio', buildDate: '2026-11-27' }
    };
  }

  function downloadWorkflow(format) {
    const record = makeWorkflowRecord(), safe = slug(record.project || record.source.file || 'research-analysis');
    if (format === 'json') return download(`${safe}-workflow-record.json`, JSON.stringify(record, null, 2), 'application/json;charset=utf-8');
    const lines = [
      `# Research analysis plan${record.project ? ` · ${record.project}` : ''}`, '',
      `- Source: ${record.source.file}`, `- Source SHA-256: ${record.source.sha256 || 'not available (pasted input or browser does not support local hashing)'}`,
      `- Imported rows / fields: ${record.source.rows} / ${record.source.fields.length}`, `- Research question: ${record.researchQuestion || 'Not supplied'}`,
      `- Selected visualization: ${record.figure.chart}`, `- Fields: ${record.figure.axes.x || '—'} / ${record.figure.axes.y || '—'}`, '',
      '## Profile and checks', `- Missing cells: ${record.profile.missingCells}`, `- Duplicate rows retained: ${record.profile.duplicateRows}`,
      `- Numeric fields: ${record.profile.numericFields.join(', ') || 'none detected'}`, `- Mixed fields to review: ${record.profile.mixedFields.join(', ') || 'none detected'}`, '',
      '## Researcher notes', `- Preprocessing / exclusions: ${record.researcherNotes.preprocessingAndExclusions || 'Not supplied'}`, `- Statistical method / software: ${record.researcherNotes.statisticalMethodsAndSoftware || 'Not supplied'}`, '',
      '## Analysis plan', ...record.analysisPlan.map((step, index) => `${index + 1}. ${step}`), '',
      '## Researcher review', ...Object.entries(record.researcherChecklist).map(([key, complete]) => `- [${complete ? 'x' : ' '}] ${key}`), '',
      '## Scope', ...record.caveats.map((caveat) => `- ${caveat}`), '',
      `Generated: ${record.generatedAt}`
    ];
    return download(`${safe}-analysis-plan.md`, lines.join('\n'), 'text/markdown;charset=utf-8');
  }
  async function exportSVGOrPNG(format) {
    await ensureECharts();
    const settings = exportSettings(), container = document.createElement('div');
    container.style.cssText = `position:fixed;left:-10000px;top:0;width:${settings.width}px;height:${settings.height}px;overflow:hidden;background:#fff;`;
    document.body.appendChild(container);
    const renderer = format === 'svg' ? 'svg' : 'canvas';
    const chart = window.echarts.init(container, null, { renderer, width: settings.width, height: settings.height, devicePixelRatio: window.devicePixelRatio || 1 });
    try {
      chart.setOption(buildChartOption(state.selected, settings.width, settings.height), { notMerge: true, animation: false });
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      if (format === 'svg') {
        const svg = container.querySelector('svg');
        if (!svg) throw new Error('Vector output was unavailable for this figure. Try PNG export.');
        const xml = new XMLSerializer().serializeToString(svg);
        download(`${slug(settings.title)}.svg`, `<?xml version="1.0" encoding="UTF-8"?>\n${xml}`, 'image/svg+xml;charset=utf-8');
      } else {
        const scale = format === 'png4' ? 4 : 2;
        const uri = chart.getDataURL({ type: 'png', pixelRatio: scale, backgroundColor: '#ffffff', excludeComponents: ['toolbox'] });
        const response = await fetch(uri); download(`${slug(settings.title)}-${scale}x.png`, await response.blob(), 'image/png');
      }
    } finally { chart.dispose(); container.remove(); }
  }
  function slug(text) { return String(text || 'research-figure').toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 80) || 'research-figure'; }

  async function exportFile(format) {
    if (format === 'svg' || format === 'png2' || format === 'png4') return exportSVGOrPNG(format);
    if (format === 'csv') return download(`${slug(state.currentFileName)}-data.csv`, '\uFEFF' + csvText(','), 'text/csv;charset=utf-8');
    if (format === 'tsv') return download(`${slug(state.currentFileName)}-data.tsv`, '\uFEFF' + csvText('\t'), 'text/tab-separated-values;charset=utf-8');
    if (format === 'json') return download(`${slug(state.currentFileName)}-data.json`, JSON.stringify({ source: state.currentFileName, fields: state.headers, records: state.rows }, null, 2), 'application/json');
    if (format === 'xlsx') {
      await ensureXLSX();
      const workbook = window.XLSX.utils.book_new(), sheet = window.XLSX.utils.json_to_sheet(state.rows, { header: state.headers });
      window.XLSX.utils.book_append_sheet(workbook, sheet, 'Data'); window.XLSX.writeFile(workbook, `${slug(state.currentFileName)}-data.xlsx`); return;
    }
    if (format === 'figure') {
      const figure = { ...exportSettings(), data: state.rows, chartOption: buildChartOption(state.selected) };
      return download(`${slug(figure.title)}-figure.json`, JSON.stringify(figure, null, 2), 'application/json');
    }
    if (format === 'html' || format === 'md') {
      const settings = exportSettings(); await ensureECharts();
      let svg = '';
      if (state.selected !== 'table') {
        const container = document.createElement('div'); container.style.cssText = `position:fixed;left:-10000px;top:0;width:${settings.width}px;height:${settings.height}px;`;
        document.body.appendChild(container); const chart = window.echarts.init(container, null, { renderer: 'svg', width: settings.width, height: settings.height });
        try { chart.setOption(buildChartOption(state.selected, settings.width, settings.height), { animation: false }); await new Promise((resolve) => requestAnimationFrame(resolve)); svg = container.querySelector('svg')?.outerHTML || ''; } finally { chart.dispose(); container.remove(); }
      }
      if (format === 'html') {
      const body = state.selected === 'table' ? `<h2>Imported data preview</h2><pre>${escapeHTML(csvText().slice(0, 100000))}</pre>` : svg;
      const truncationNote = state.selected === 'table' && csvText().length > 100000 ? '<p><small>This HTML report contains a 100,000-character data preview only. Export CSV, TSV, JSON, or Excel for the complete imported data.</small></p>' : '';
      const report = `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${escapeHTML(settings.title)}</title><style>body{font:15px/1.55 Arial,sans-serif;max-width:1300px;margin:40px auto;padding:0 24px;color:#172d39}svg{width:100%;height:auto}table{border-collapse:collapse}small{color:#5c7078}</style><h1>${escapeHTML(settings.title)}</h1><p>${escapeHTML(settings.source)} · ${settings.rows.toLocaleString()} rows · ${settings.fields.length} fields</p>${body}${truncationNote}<p><small>Static report. Figure settings and imported-data summary: ${escapeHTML(JSON.stringify(settings))}</small></p></html>`;
        return download(`${slug(settings.title)}.html`, report, 'text/html;charset=utf-8');
      }
      const markdown = `# ${settings.title}\n\n- Source: ${settings.source}\n- Observations: ${settings.rows}\n- Chart: ${vizMeta[state.selected][0]}\n- Fields: ${settings.fields.join(', ')}\n- Axes: ${settings.axes.x} / ${settings.axes.y}\n- Aggregation: ${settings.aggregation}\n- Missing cells: ${settings.missingCells} (excluded where required; not imputed)\n- Duplicate rows: ${settings.duplicateRows} (retained)\n\nRecommendation is a transparent data-shape heuristic, not a statistical test or scientific interpretation.\n`;
      return download(`${slug(settings.title)}.md`, markdown, 'text/markdown;charset=utf-8');
    }
    if (format === 'mmd') {
      if (!['mindmap', 'flow'].includes(state.selected)) throw new Error('Select Concept map or Process sequence before exporting Mermaid.');
      const tree = notesTree();
      const emit = (node, depth) => node.children.flatMap((child) => [`${'  '.repeat(depth)}${escapeMermaid(child.name)}`, ...emit(child, depth + 1)]);
      const mermaid = `mindmap\n  root((${escapeMermaid(tree.name)}))\n${emit(tree, 2).join('\n')}\n`;
      return download(`${slug(state.currentFileName)}-mindmap.mmd`, mermaid, 'text/plain;charset=utf-8');
    }
    if (format === 'opml') {
      if (!['mindmap', 'flow'].includes(state.selected)) throw new Error('Select Concept map or Process sequence before exporting OPML.');
      const tree = notesTree();
      const emit = (node) => node.children.map((child) => `<outline text="${escapeHTML(child.name)}">${emit(child)}</outline>`).join('');
      return download(`${slug(state.currentFileName)}.opml`, `<?xml version="1.0" encoding="UTF-8"?><opml version="2.0"><head><title>${escapeHTML(tree.name)}</title></head><body><outline text="${escapeHTML(tree.name)}">${emit(tree)}</outline></body></opml>`, 'text/xml;charset=utf-8');
    }
  }
  function escapeMermaid(text) { return String(text).replace(/[()[\]{}]/g, '').replace(/[&<>]/g, '').replace(/[\r\n]/g, ' ').trim().slice(0, 120); }

  async function runOCR() {
    if (state.ocrBusy || !state.ocrAvailable) return;
    state.ocrBusy = true;
    $('#recognize-image').disabled = true;
    $('#data-file').disabled = true;
    try {
      if (!window.Tesseract) await ensureScript('https://cdn.jsdelivr.net/npm/tesseract.js@7.0.0/dist/tesseract.min.js', () => !!window.Tesseract);
      const worker = await window.Tesseract.createWorker('eng', 1, { logger: (message) => { if (message.status) setMessage(`OCR: ${message.status}${message.progress ? ` · ${Math.round(message.progress * 100)}%` : ''}`); } });
      try {
        const pages = [];
        if (state.imageFile) {
          const result = await worker.recognize(state.imageFile); pages.push(result.data.text);
        } else if (state.pdf) {
          for (let pageNo = 1; pageNo <= Math.min(state.pdf.numPages, 10); pageNo++) {
            setMessage(`OCR scanned PDF page ${pageNo} of ${Math.min(state.pdf.numPages, 10)}…`);
            const page = await state.pdf.getPage(pageNo), viewport = page.getViewport({ scale: 1.5 }), canvas = document.createElement('canvas');
            canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
            await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
            const result = await worker.recognize(canvas); pages.push(`# Page ${pageNo}\n${result.data.text}`);
          }
        }
        const text = pages.join('\n\n').trim();
        if (!text) throw new Error('OCR found no readable text. Try a clearer image or higher-resolution scan.');
        $('#data-text').value = text;
        setSource({ ...parseSource(text), kind: 'ocr', label: `${state.currentFileName} · OCR`, warning: 'OCR text can contain recognition errors; verify every extracted value against the image or scan before analysis.' }, `${state.currentFileName} · OCR`);
        setMessage('OCR text extracted. Recognition can misread symbols, decimal points, and labels; inspect before using it.', 'caution');
      } finally { await worker.terminate(); }
    } catch (error) { setMessage(`OCR could not run: ${error.message}`, 'caution'); }
    finally {
      state.ocrBusy = false;
      $('#data-file').disabled = false;
      $('#recognize-image').disabled = !state.ocrAvailable;
    }
  }

  $('#data-file').addEventListener('change', (event) => processFile(event.target.files?.[0]));
  $('#sheet-select').addEventListener('change', () => { if (state.workbook) parseWorkbookSheet($('#sheet-select').value); else if (state.documentTables.length) parseDocumentTable($('#sheet-select').value); });
  $('#recognize-image').addEventListener('click', runOCR);
  $('#analyze').addEventListener('click', () => {
    try {
      const text = $('#data-text').value.trim();
      if (text) { const source = parseSource(text); state.sourceFileInfo = null; setSource(source, 'Pasted or edited source'); setMessage('Text profiled locally. Review the field-type and quality summary before interpreting the figure.', 'good'); }
      else if (!state.rows.length) throw new Error('Select a supported file or paste data/text to begin.');
      else profileStudio();
    } catch (error) { $('#diagnosis').textContent = error.message; }
  });
  $('#intent').addEventListener('change', () => { if (state.rows.length) { state.recommendations = compatibleViews(); profileStudio(); } });
  $('#download-workflow-json').addEventListener('click', () => downloadWorkflow('json'));
  $('#download-workflow-md').addEventListener('click', () => downloadWorkflow('md'));
  $('#workflow-to-studio').addEventListener('click', () => document.querySelector('#tab-btn-studio').click());
  document.querySelectorAll('[data-review]').forEach((input) => input.addEventListener('change', syncWorkflow));
  for (const selector of ['#workflow-question', '#workflow-study']) $(selector).addEventListener('input', syncWorkflow);
  $('#clear-data').addEventListener('click', () => {
    if (state.chart) { state.chart.dispose(); state.chart = null; }
    if (state.profileWorker) { state.profileWorker.terminate(); state.profileWorker = null; state.profileReject?.(new Error('Profiling cancelled.')); state.profileReject = null; }
    if (state.imageURL) URL.revokeObjectURL(state.imageURL);
    Object.assign(state, { rows: [], headers: [], source: null, types: null, profile: null, recommendations: [], selected: '', workbook: null, pdf: null, documentTables: [], imageFile: null, ocrAvailable: false, imageURL: '', imageGrid: null, currentFileName: '', sourceFileInfo: null });
    $('#data-file').value = ''; $('#data-text').value = ''; $('#sheet-wrap').hidden = true; $('#source-preview').hidden = true; $('#recognize-image').disabled = true;
    $('#workflow-question').value = ''; $('#workflow-study').value = ''; $('#workflow-processing').value = ''; $('#workflow-statistics').value = '';
    $('#data-state').textContent = 'NO SOURCE LOADED'; $('#quality-profile').hidden = true; $('#figure-settings').hidden = true; $('#exports').hidden = true; $('#figure-caption').hidden = true;
    $('#viz-catalog').innerHTML = '<div class="catalog-placeholder">Compatible chart types and recommendations will appear after the source is profiled.</div>';
    $('#render-area').innerHTML = '<div class="render-empty"><div><strong>Your figure will appear here</strong><span>Source data stays available for review and export.</span></div></div>';
    $('#recommendation-caption').textContent = 'Upload a source to see its profile and compatible visual forms.';
    document.querySelectorAll('[data-review]').forEach((input) => { input.checked = false; });
    $('#diagnosis').textContent = 'The analysis first checks field types, missing values, duplicates, ranges, and sample size. It does not infer scientific meaning or alter source values.';
    setMessage('For scanned PDFs and images, OCR is optional and runs locally in your browser.');
    syncWorkflow();
  });
  document.querySelectorAll('#exports [data-export]').forEach((button) => button.addEventListener('click', async () => {
    button.disabled = true;
    try { await exportFile(button.dataset.export); }
    catch (error) { setMessage(`Export failed: ${error.message}`, 'caution'); }
    finally { button.disabled = false; }
  }));
  for (const selector of ['#axis-x', '#axis-y', '#axis-group', '#aggregate', '#palette', '#text-size', '#figure-title', '#label-x', '#label-y', '#figure-width', '#figure-height']) {
    $(selector).addEventListener('change', () => { if (state.rows.length) { renderFigure(); syncWorkflow(); } });
  }
  syncWorkflow();
  window.addEventListener('resize', () => state.chart?.resize());
  document.querySelectorAll('.tab[data-tab="studio"]').forEach((button) => button.addEventListener('click', () => setTimeout(() => state.chart?.resize(), 40)));
})();
