/* NAM Market Improvement Tracking — RDF KPI Dashboard
   Reads an Excel workbook (RDF Action Plan + Cases sheets), recomputes every KPI
   from case-level data and tracks weekly / monthly / quarterly improvement. */

/* ------------------------------------------------------------------ helpers */
const $ = id => document.getElementById(id);
const norm = s => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
const EPS = 1e-9;

function toNum(v) {
  if (v == null || v === '' || v === '-') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const n = parseFloat(String(v).replace(/[, %]/g, ''));
  return isFinite(n) ? n : null;
}
const n0 = v => toNum(v) ?? 0;

function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function toDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') {                       // Excel serial
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);   // dd/mm/yyyy
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

const fmtPct = v => v == null ? '—' : (v * 100).toFixed(1) + '%';
const fmtNum = v => v == null ? '—' : v.toFixed(2);
const fmtInt = v => v == null ? '—' : Math.round(v).toLocaleString();
const fmtPP  = v => (v >= 0 ? '+' : '') + (v * 100).toFixed(1) + ' pp';
const fmtDN  = v => (v >= 0 ? '+' : '') + v.toFixed(2);
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/* ------------------------------------------------------------- KPI registry */
/* num/den are per-case contributions; a case is excluded from a KPI when den() is null.
   `bench` is a FIXED global benchmark supplied directly (not calculated from the uploaded
   workbook) — it is used identically everywhere: latest-update cards, trend overlay,
   region comparison and the scorecard. */
const KPIS = [
  { key: 'partsAdvised', label: 'Parts Advised Success rate', short: 'Parts Advised', fmt: 'pct', dir: 'up',
    bench: 0.336,
    def: 'Share of RDF-advised part-list items that match the relevant case parts list.',
    calc: 'SUM(Advised Part Numbers Matched) / SUM(Advised Part Numbers Listed)',
    src: 'BN / BM', reqCols: ['advListed', 'advMatched'],
    den: r => (S.kpiAvailable.partsAdvised && r.advListed > 0) ? r.advListed : null, num: r => r.advMatched },

  { key: 'visits', label: 'Visits per Case', short: 'Visits/case', fmt: 'num', dir: 'down',
    bench: 1.37,
    def: 'Average number of field visits required for each case.',
    calc: 'SUM(Visits) / COUNT(valid Case Number)',
    src: 'AD / A', reqCols: ['visits'],
    den: () => 1, num: r => r.visits },

  { key: 'servicePlan', label: 'Service Plan availability', short: 'Service Plan', fmt: 'pct', dir: 'up',
    bench: 0.738,
    def: 'Share of cases where a Service Plan is available to support diagnosis and dispatch readiness.',
    calc: 'SUM(Service Plan Available) / COUNT(valid Case Number)',
    src: 'AR / A', reqCols: ['spa'],
    den: () => 1, num: r => r.spa },

  { key: 'rdfSuccess', label: 'RDF Success / total', short: 'RDF Success', fmt: 'pct', dir: 'up',
    bench: 0.569,
    def: 'Share of total mapped cases with a successful RDF record or result.',
    calc: 'SUM(RDF) / COUNT(valid Case Number)',
    src: 'AA / A', reqCols: ['rdf'],
    den: () => 1, num: r => r.rdf },

  { key: 'preFirst', label: 'Pre-first-visit ordering success', short: 'Pre-first-visit order', fmt: 'pct', dir: 'up',
    bench: 0.523,
    def: 'Share of cases with at least one advised FRU where an advised part was ordered before the first visit.',
    calc: 'SUM(Advised Part Ordered Before First Visit) / COUNT(cases where # FRUs Advised > 0)',
    src: 'AT / AE', reqCols: ['orderedBefore', 'frusAdvised'],
    den: r => r.frusAdvised > 0 ? 1 : null, num: r => r.orderedBefore },

  { key: 'partsConsumed', label: 'Parts consumed per Case', short: 'Parts consumed/case', fmt: 'num', dir: 'down',
    bench: 1.37,
    def: 'Average quantity of parts consumed for each case.',
    calc: 'SUM(# Parts Consumed) / COUNT(valid Case Number)',
    src: 'AH / A', reqCols: ['partsConsumed'],
    den: () => 1, num: r => r.partsConsumed },

  /* secondary KPIs also present in the workbook */
  { key: 'onsite', label: 'Onsite hours per Case', short: 'Onsite h/case', fmt: 'num', dir: 'down',
    bench: 3.18, secondary: true,
    def: 'Average field-engineer time spent onsite for each case.',
    calc: 'SUM(Onsite Hours) / COUNT(valid Case Number)',
    src: 'BB / A', reqCols: ['onsite'],
    den: () => 1, num: r => r.onsite },

  { key: 'partsReturn', label: 'Parts return rate', short: 'Parts return', fmt: 'pct', dir: 'up',
    bench: 0.73, secondary: true,
    def: 'Returned quantity divided by ordered quantity minus consumed quantity; only positive unused balances count.',
    calc: 'SUM(Returned Parts Qty) / SUM(# Parts Ordered − # Parts Consumed), cases where that difference > 0',
    src: 'BO / (AF − AH)', reqCols: ['partsReturned', 'partsOrdered', 'partsConsumed'],
    den: r => (S.kpiAvailable.partsReturn && (r.partsOrdered - r.partsConsumed) > 0)
      ? (r.partsOrdered - r.partsConsumed) : null,
    num: r => r.partsReturned }
];
const KPI_BY_KEY = Object.fromEntries(KPIS.map(k => [k.key, k]));
/* canonical header names the column-resolution engine looks for, keyed like reqCols above */
const CANON_HEADERS = {
  caseNo: 'Case Number', region: 'IB Region', zone: 'IB Zone', modality: 'Modality', market: 'Market',
  rdf: 'RDF', visits: 'Visits', frusAdvised: '# FRUs Advised', partsOrdered: '# Parts Ordered',
  partsConsumed: '# Parts Consumed', onsite: 'Onsite Hours', spa: 'Service Plan Available',
  orderedBefore: 'Advised Part Ordered Before First Visit', partsReturned: 'Returned Parts Qty (parsed)',
  advListed: 'Advised Part Numbers Listed', advMatched: 'Advised Part Numbers Matched',
  advisedList: 'Parts Advised List', consumedList: 'Parts Consumed List'
};
const fmtOf = k => k.fmt === 'pct' ? fmtPct : fmtNum;
const fmtDelta = k => k.fmt === 'pct' ? fmtPP : fmtDN;

/* Extract {num, qty} entries with a valid 12-digit part number from a free-text parts-list cell.
   Lines are split on newlines/semicolons; a line without a 12-digit number is excluded. Quantity
   is read from an "x2" / "qty 2" / "(2)" style marker on the line, defaulting to 1. */
function parsePartsListCell(v) {
  if (v == null || v === '') return [];
  const lines = String(v).split(/[\n;]+/).map(s => s.trim()).filter(Boolean);
  const out = [];
  for (const line of lines) {
    const m = line.match(/(?<!\d)(\d{12})(?!\d)/);
    if (!m) continue;
    const qm = line.slice(m.index + 12).match(/(?:x|qty|quantity)\s*[:=]?\s*(\d+)/i) || line.match(/\((\d+)\)/);
    out.push({ num: m[1], qty: qm ? (parseInt(qm[1], 10) || 1) : 1 });
  }
  return out;
}
const famKey = num => num.slice(0, 11);
/* An advised part matches when a consumed part shares the exact 12-digit number or the same
   11-digit family; a qualifying match counts the advised line's full quantity. */
function matchAdvisedParts(advisedEntries, consumedEntries) {
  const exact = new Set(consumedEntries.map(e => e.num));
  const fam = new Set(consumedEntries.map(e => famKey(e.num)));
  let total = 0, matched = 0;
  for (const e of advisedEntries) {
    total += e.qty;
    if (exact.has(e.num) || fam.has(famKey(e.num))) matched += e.qty;
  }
  return { total, matched };
}

/* --------------------------------------------------------------- app state */
const S = {
  fileName: '', cases: [], plan: [], planCols: [], planRaw: [],
  dateCols: [], caseHeaders: [], caseColumnMap: {}, sheetNames: [], hasCases: false,
  assistantTables: [],
  colReport: [], missingReport: [], caseStats: null,
  kpiAvailable: { partsAdvised: true, partsReturn: true }
};
const charts = {};

/* --------------------------------------------------------------- file load */
const dz = $('dropzone'), fi = $('fileInput');
dz.addEventListener('click', () => fi.click());
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('over'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('over');
  if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});
fi.addEventListener('change', e => { if (e.target.files.length) handleFiles(e.target.files); });
$('btnChange').addEventListener('click', () => {
  $('dashView').classList.add('hidden'); $('importView').classList.remove('hidden');
  $('btnChange').classList.add('hidden');
});

function status(html, cls) { $('importStatus').innerHTML = `<span class="${cls || ''}">${html}</span>`; }

async function handleFiles(fileList) {
  const files = [...fileList];
  if (files.length > 1) S.assistantTables = [];
  for (let index = 0; index < files.length; index++) {
    await handleFile(files[index], index > 0);
  }
}

function handleFile(file, appendTable) {
  return new Promise(resolve => {
    S.fileName = file.name;
    status('Reading <b>' + esc(file.name) + '</b> …');
    const fr = new FileReader();
    fr.onerror = () => {
      status('Could not read the file.', 'err');
      resolve();
    };
    fr.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        if (/\.csv$/i.test(file.name)) buildReportFromWorkbook(wb, file.name, appendTable);
        else buildFromWorkbook(wb);
      } catch (err) {
        console.error(err);
        status('Failed to parse workbook: ' + esc(err.message), 'err');
      }
      resolve();
    };
    fr.readAsArrayBuffer(file);
  });
}

function buildReportFromWorkbook(wb, fileName, appendTable) {
  const sheetName = wb.SheetNames[0];
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
  let headerIndex = -1;
  for (let i = 0; i < grid.length; i++) {
    const row = grid[i] || [];
    const labels = row.filter(v => v != null && String(v).trim()).map(v => String(v).trim());
    if (labels.includes('IB Region') || labels.includes('KPI')) { headerIndex = i; break; }
  }
  if (headerIndex < 0) {
    headerIndex = grid.reduce((best, row, i) =>
      (row || []).filter(v => v != null && String(v).trim()).length >
      (grid[best] || []).filter(v => v != null && String(v).trim()).length ? i : best, 0);
  }
  const headers = (grid[headerIndex] || []).map(v => v == null ? '' : String(v).trim());
  const rows = grid.slice(headerIndex + 1).map(row => {
    const raw = Object.create(null);
    headers.forEach((header, i) => { if (header) raw[header] = row[i]; });
    return { raw };
  }).filter(r => Object.values(r.raw).some(v => v != null && String(v).trim() && String(v).trim() !== '-'));
  const id = `report-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  if (!appendTable) S.assistantTables = [];
  S.assistantTables.push({ id, name: fileName, headers: headers.filter(Boolean), rows, type: 'report' });
  S.fileName = fileName;
  resetAssistantConversation(true);
  refreshAssistantDatasets(id);
  status(`Report <b>${esc(fileName)}</b>: ${rows.length} summary rows loaded for the NAM Data Assistant. ` +
    'This report contains aggregates, not individual cases.', 'ok');
  updateAssistantScope();
}

/* Find a sheet by fuzzy name */
function findSheet(wb, candidates) {
  for (const c of candidates) {
    const hit = wb.SheetNames.find(n => norm(n) === norm(c));
    if (hit) return hit;
  }
  for (const c of candidates) {
    const hit = wb.SheetNames.find(n => norm(n).includes(norm(c)));
    if (hit) return hit;
  }
  return null;
}

/* Resolve a header name from a header list: exact-normalized, then prefix, then contains */
function resolveCol(headers, wanted) {
  const w = norm(wanted);
  let h = headers.find(x => norm(x) === w); if (h) return h;
  h = headers.find(x => norm(x).startsWith(w)); if (h) return h;
  h = headers.find(x => w.startsWith(norm(x)) && norm(x).length > 6); if (h) return h;
  h = headers.find(x => norm(x).includes(w)); return h || null;
}

/* Score a sheet's header row against the canonical case-level columns, to auto-detect
   which sheet holds the raw case data when it isn't named "Cases". */
function scoreCaseSheet(wb, name) {
  const first = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null })[0] || [];
  const headers = first.filter(h => h != null).map(h => String(h));
  if (headers.length < 5) return 0;
  let hits = 0;
  for (const key in CANON_HEADERS) if (resolveCol(headers, CANON_HEADERS[key])) hits++;
  return hits;
}
function detectCaseSheet(wb, excludeName) {
  let best = null, bestScore = 0;
  for (const name of wb.SheetNames) {
    if (name === excludeName) continue;
    const s = scoreCaseSheet(wb, name);
    if (s > bestScore) { bestScore = s; best = name; }
  }
  return bestScore >= 6 ? best : null;
}

function buildFromWorkbook(wb) {
  resetAssistantConversation(false);
  S.sheetNames = wb.SheetNames;
  const msgs = [];

  /* ---- RDF Action Plan sheet (optional — older multi-tab workbooks only) ---- */
  const planName = findSheet(wb, ['RDF Action Plan', 'Action Plan']);
  S.plan = []; S.planCols = []; S.planRaw = [];
  if (planName) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[planName], { header: 1, raw: true, defval: null });
    S.planRaw = grid;
    parsePlanTable(grid, msgs);
    msgs.push(`Sheet <b>${esc(planName)}</b>: ${S.plan.length} regions, 30-60-90 plan loaded.`);
  } else {
    msgs.push('No "RDF Action Plan" sheet found — skipping the 30-60-90 plan text; all KPI values are still recomputed from case data.');
  }

  /* ---- Case-level sheet: named "Cases" etc., else auto-detected by header content ---- */
  let casesName = findSheet(wb, ['Cases', 'Case List', 'Action Case List', 'Raw']);
  let autoDetected = false;
  if (!casesName) { casesName = detectCaseSheet(wb, planName); autoDetected = true; }
  S.cases = []; S.dateCols = []; S.caseHeaders = []; S.caseColumnMap = {};
  S.hasCases = false; S.colReport = []; S.missingReport = []; S.caseStats = null;
  if (casesName) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[casesName], { raw: true, defval: null });
    if (rows.length) {
      parseCases(rows, msgs);
      msgs.push(`Sheet <b>${esc(casesName)}</b>${autoDetected ? ' (auto-detected as the case-level sheet)' : ''}: ${S.cases.length} valid cases parsed.`);
    }
    S.assistantTables = S.cases.length ? [{
      id: 'cases', name: S.fileName, headers: S.caseHeaders, rows: S.cases, type: 'cases'
    }] : [];
  }
  if (!S.hasCases) {
    msgs.push('<span class="err">No usable case-level sheet found — trends and region comparison are disabled; KPI benchmarks remain the fixed NAM targets.</span>');
  } else if (S.missingReport.length) {
    msgs.push('<span class="err">Not calculable from this workbook: ' +
      S.missingReport.map(m => `${esc(m.kpi)} (${esc(m.reason)})`).join('; ') + '.</span>');
  }

  status(msgs.join('<br>'), 'ok');
  resetAssistantConversation(true);
  refreshAssistantDatasets('cases');
  updateAssistantScope();
  setTimeout(() => initDashboard(), 250);
}

/* Region table inside the plan sheet: header row whose first cell is "Region" */
function parsePlanTable(grid, msgs) {
  let hr = -1;
  for (let i = 0; i < grid.length; i++) {
    const r = grid[i] || [];
    if (norm(r[0]) === 'region' && r.filter(x => x != null).length > 4) { hr = i; break; }
  }
  S.plan = []; S.planCols = [];
  if (hr < 0) { msgs.push('<span class="err">Could not locate the region table header in the action-plan sheet.</span>'); return; }
  const heads = (grid[hr] || []).map(h => h == null ? '' : String(h).trim());
  S.planCols = heads;
  for (let i = hr + 1; i < grid.length; i++) {
    const r = grid[i] || [];
    const name = r[0] == null ? '' : String(r[0]).trim();
    if (!name) continue;
    const o = { _region: name };
    heads.forEach((h, j) => { if (h) o[h] = r[j]; });
    S.plan.push(o);
  }
}

/* Case-level parsing into a compact normalised record. Builds S.colReport / S.missingReport
   for the auto-benchmark "column mapping" and "missing fields" output. */
const BANNED_LABELS = new Set(['selectionstatus', 'modality', 'region', 'market', 'year', 'grandtotal', 'total']);
function parseCases(rows, msgs) {
  const H = Object.keys(rows[0]);
  S.caseHeaders = H.slice();
  const letterOf = h => { const i = H.indexOf(h); return i < 0 ? '' : XLSX.utils.encode_col(i); };
  const C = {
    region: resolveCol(H, CANON_HEADERS.region), zone: resolveCol(H, CANON_HEADERS.zone),
    modality: resolveCol(H, CANON_HEADERS.modality), market: resolveCol(H, CANON_HEADERS.market),
    caseNo: resolveCol(H, CANON_HEADERS.caseNo),
    advListed: resolveCol(H, CANON_HEADERS.advListed),
    advMatched: resolveCol(H, CANON_HEADERS.advMatched),
    advisedList: resolveCol(H, CANON_HEADERS.advisedList),
    consumedList: resolveCol(H, CANON_HEADERS.consumedList),
    visits: resolveCol(H, CANON_HEADERS.visits),
    onsite: resolveCol(H, CANON_HEADERS.onsite),
    partsOrdered: resolveCol(H, CANON_HEADERS.partsOrdered),
    partsConsumed: resolveCol(H, CANON_HEADERS.partsConsumed),
    partsReturned: resolveCol(H, CANON_HEADERS.partsReturned),
    spa: resolveCol(H, CANON_HEADERS.spa),
    rdf: H.find(h => norm(h) === 'rdf') || null,
    frusAdvised: resolveCol(H, CANON_HEADERS.frusAdvised),
    orderedBefore: resolveCol(H, CANON_HEADERS.orderedBefore)
  };
  /* guard: "Service Plan Available" must not resolve to "…Before First Visit" */
  if (C.spa && norm(C.spa).includes('before')) {
    C.spa = H.find(h => norm(h) === 'serviceplanavailable') || C.spa;
  }
  const missing = ['visits', 'spa', 'rdf', 'partsConsumed'].filter(k => !C[k]);
  if (missing.length) {
    msgs.push('<span class="err">Cases sheet is missing required columns: ' + missing.join(', ') + '</span>');
    return;
  }

  /* ---- column-mapping report (for the auto-benchmark panel) ---- */
  S.colReport = Object.keys(CANON_HEADERS).map(key => ({
    key, wanted: CANON_HEADERS[key], header: C[key] || null, letter: C[key] ? letterOf(C[key]) : '',
    found: !!C[key]
  }));
  S.caseColumnMap = { ...C };
  /* KPIs whose source column(s) are entirely absent from this workbook must never be
     silently computed from a default-zero value - mark them unavailable so bench/num/den
     stay null instead of fabricating a 0% result. */
  const useAdvisedTextLists = !(C.advListed && C.advMatched) && !!(C.advisedList && C.consumedList);
  S.kpiAvailable = {
    partsAdvised: !!(C.advListed && C.advMatched) || useAdvisedTextLists,
    partsReturn: !!C.partsReturned && !!C.partsOrdered && !!C.partsConsumed
  };
  S.missingReport = [];
  if (!C.advListed || !C.advMatched) {
    if (!useAdvisedTextLists) {
      S.missingReport.push({
        kpi: 'Parts Advised Success rate',
        reason: 'Columns "Advised Part Numbers Listed" / "Advised Part Numbers Matched" (BM/BN) are not present, and no usable "Parts Advised List" / "Parts Consumed List" text columns were found either.'
      });
    }
    if (!S.kpiAvailable.partsReturn) {
      S.missingReport.push({
        kpi: 'Parts return rate',
        reason: 'Columns "Returned Parts Qty (parsed)", "# Parts Ordered", and "# Parts Consumed" are required.'
      });
    }
  }

  /* candidate date columns */
  const dateCandidates = H.filter(h => /date/i.test(h));
  const dateCols = [];
  for (const h of dateCandidates) {
    const ok = rows.slice(0, 200).filter(r => toDate(r[h])).length;
    if (ok >= 5) dateCols.push(h);
  }
  S.dateCols = dateCols;

  const rawCount = rows.length;
  let droppedNoCaseNo = 0, droppedLabelRow = 0;

  S.cases = rows.map(r => {
    const caseNoRaw = C.caseNo ? r[C.caseNo] : null;
    const caseNo = caseNoRaw == null ? '' : String(caseNoRaw).trim();
    let advListedVal, advMatchedVal;
    if (C.advListed && C.advMatched) {
      advListedVal = n0(r[C.advListed]); advMatchedVal = n0(r[C.advMatched]);
    } else if (useAdvisedTextLists) {
      const { total, matched } = matchAdvisedParts(
        parsePartsListCell(r[C.advisedList]), parsePartsListCell(r[C.consumedList]));
      advListedVal = total; advMatchedVal = matched;
    } else {
      advListedVal = 0; advMatchedVal = 0;
    }
    const raw = Object.create(null);
    H.forEach(h => { raw[h] = r[h]; });
    const o = {
      caseNo,
      region: C.region ? String(r[C.region] ?? '').trim() : '',
      zone: C.zone ? String(r[C.zone] ?? '').trim() : '',
      modality: C.modality ? String(r[C.modality] ?? '').trim() : '',
      market: C.market ? String(r[C.market] ?? '').trim() : '',
      advListed: advListedVal, advMatched: advMatchedVal,
      visits: n0(r[C.visits]), onsite: n0(r[C.onsite]),
      partsOrdered: n0(r[C.partsOrdered]), partsConsumed: n0(r[C.partsConsumed]),
      partsReturned: n0(r[C.partsReturned]),
      spa: n0(r[C.spa]), rdf: n0(r[C.rdf]),
      frusAdvised: n0(r[C.frusAdvised]), orderedBefore: n0(r[C.orderedBefore]),
      d: {}, raw
    };
    dateCols.forEach(h => { o.d[h] = toDate(r[h]); });
    return o;
  }).filter(r => {
    // exclude blank rows and report/filter footer labels ("Selection Status", "Modality", "Region", …)
    if (!r.caseNo) { droppedNoCaseNo++; return false; }
    if (BANNED_LABELS.has(norm(r.caseNo)) || BANNED_LABELS.has(norm(r.region))) { droppedLabelRow++; return false; }
    return true;
  });

  S.caseStats = { rawCount, valid: S.cases.length, droppedNoCaseNo, droppedLabelRow,
    regions: uniq(S.cases.map(r => r.region)).length };

  if (useAdvisedTextLists) {
    const casesWithAdvised = S.cases.filter(r => r.advListed > 0);
    const totalQty = S.cases.reduce((s, r) => s + r.advListed, 0);
    const matchedQty = S.cases.reduce((s, r) => s + r.advMatched, 0);
    const rate = totalQty > EPS ? matchedQty / totalQty : null;
    S.partsAdvisedReport = { casesReviewed: S.cases.length, casesWithAdvised: casesWithAdvised.length,
      totalQty, matchedQty, rate };
    msgs.push('Parts Advised Success rate derived from free-text "Parts Advised List" / "Parts Consumed List" ' +
      `columns (AM/AO): ${S.cases.length} cases reviewed, ${casesWithAdvised.length} with advised parts, ` +
      `total advised qty ${totalQty}, matched qty ${matchedQty}, overall rate ${rate == null ? '—' : (rate * 100).toFixed(1) + '%'}.`);
  }

  S.hasCases = S.cases.length > 0 && dateCols.length > 0;
  if (S.cases.length && !dateCols.length)
    msgs.push('<span class="err">Cases sheet has no usable date column — trends disabled.</span>');
}

/* --------------------------------------------------------- KPI aggregation */
function aggregate(rows) {
  const out = { n: rows.length };
  for (const k of KPIS) {
    let num = 0, den = 0, cnt = 0;
    for (const r of rows) {
      const d = k.den(r);
      if (d == null || d <= 0) continue;
      num += k.num(r); den += d; cnt++;
    }
    out[k.key] = den > EPS ? num / den : null;
    out[k.key + '_n'] = cnt;
    out[k.key + '_num'] = den > EPS ? num : null;
    out[k.key + '_den'] = den > EPS ? den : null;
  }
  return out;
}

/* ------------------------------------------------------------ period logic */
function pad2(n) { return String(n).padStart(2, '0'); }
function periodKey(date, gran) {
  if (gran === 'year') return `${date.getFullYear()}`;
  if (gran === 'month') return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}`;
  if (gran === 'quarter') return `${date.getFullYear()}-Q${Math.floor(date.getMonth() / 3) + 1}`;
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));          // Monday of that week
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function periodLabel(key, gran) {
  if (gran === 'year') return key;
  if (gran === 'month') { const [y, m] = key.split('-'); return `${MON[+m - 1]} ${y}`; }
  if (gran === 'quarter') return key;
  const [y, m, d] = key.split('-'); return `w/c ${d} ${MON[+m - 1]} ${y.slice(2)}`;
}
const fmtDateShort = d => `${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear()}`;
/* exact [start, end] date span covered by a period key, for the "which dates does this cover" hint */
function periodRange(key, gran) {
  if (gran === 'year') { const y = +key; return { start: new Date(y, 0, 1), end: new Date(y, 11, 31) }; }
  if (gran === 'month') { const [y, m] = key.split('-').map(Number); return { start: new Date(y, m - 1, 1), end: new Date(y, m, 0) }; }
  if (gran === 'quarter') {
    const [y, q] = key.split('-Q').map(Number);
    return { start: new Date(y, (q - 1) * 3, 1), end: new Date(y, q * 3, 0) };
  }
  const [y, m, d] = key.split('-').map(Number);
  const start = new Date(y, m - 1, d);
  const end = new Date(y, m - 1, d); end.setDate(end.getDate() + 6);
  return { start, end };
}

/* --------------------------------------------------------------- dashboard */
function initDashboard() {
  $('importView').classList.add('hidden');
  $('dashView').classList.remove('hidden');
  $('btnChange').classList.remove('hidden');
  $('fileInfo').innerHTML = esc(S.fileName) + ' &middot; ' +
    (S.hasCases ? S.cases.length + ' cases' : 'snapshot only') + ' &middot; ' + S.plan.length + ' regions';
  $('footFile').textContent = S.fileName;
  $('footTime').textContent = new Date().toLocaleString();

  /* date basis options */
  const dc = $('fDateCol'); dc.innerHTML = '';
  S.dateCols.forEach(h => dc.add(new Option(h, h)));
  const pref = S.dateCols.find(h => /creation/i.test(h)) || S.dateCols[0];
  if (pref) dc.value = pref;

  fillSelect('fRegion', 'All regions', uniq(S.cases.map(r => r.region)).length
    ? uniq(S.cases.map(r => r.region)) : S.plan.map(p => p._region));
  fillSelect('fZone', 'All zones', uniq(S.cases.map(r => r.zone)));
  fillSelect('fModality', 'All modalities', uniq(S.cases.map(r => r.modality)));

  const opts = KPIS.map(k => `<option value="${k.key}">${esc(k.label)}</option>`).join('');
  $('detailKpi').innerHTML = opts;
  $('regionKpi').innerHTML = opts;

  ['fGran', 'fDateCol', 'fRegion', 'fZone', 'fModality', 'fFrom', 'fTo', 'fMinCases']
    .forEach(id => $(id).addEventListener('change', render));
  $('detailKpi').addEventListener('change', renderDetail);
  ['regionKpi', 'regionSort'].forEach(id => $(id).addEventListener('change', renderRegionChart));
  $('btnReset').addEventListener('click', () => {
    ['fRegion', 'fZone', 'fModality'].forEach(id => $(id).value = '');
    $('fFrom').value = ''; $('fTo').value = ''; $('fMinCases').value = 5; render();
  });
  $('btnExportPeriods').addEventListener('click', exportPeriods);
  $('btnExportRegions').addEventListener('click', exportRegions);

  /* default date window = full range */
  if (S.hasCases) {
    const ds = S.cases.map(r => r.d[dc.value]).filter(Boolean).sort((a, b) => a - b);
    if (ds.length) { $('fFrom').value = iso(ds[0]); $('fTo').value = iso(ds[ds.length - 1]); }
  }
  renderDefs();
  renderPlanTable();
  render();
}
const iso = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const uniq = a => [...new Set(a.filter(x => x && x !== '-'))].sort();
function fillSelect(id, allLabel, vals) {
  const el = $(id); el.innerHTML = ''; el.add(new Option(allLabel, ''));
  vals.forEach(v => el.add(new Option(v, v)));
}

/* ---- current filtered rows ---- */
function filtered() {
  if (!S.hasCases) return [];
  const dcol = $('fDateCol').value;
  const reg = $('fRegion').value, zn = $('fZone').value, mo = $('fModality').value;
  const from = $('fFrom').value ? new Date($('fFrom').value + 'T00:00:00') : null;
  const to = $('fTo').value ? new Date($('fTo').value + 'T23:59:59') : null;
  return S.cases.filter(r => {
    if (reg && r.region !== reg) return false;
    if (zn && r.zone !== zn) return false;
    if (mo && r.modality !== mo) return false;
    const d = r.d[dcol];
    if (!d) return false;
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });
}

function buildPeriods(rows) {
  const gran = $('fGran').value, dcol = $('fDateCol').value;
  const minC = Math.max(1, parseInt($('fMinCases').value) || 1);
  const map = new Map();
  rows.forEach(r => {
    const d = r.d[dcol]; if (!d) return;
    const k = periodKey(d, gran);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  });
  return [...map.entries()].sort((a, b) => a[0] < b[0] ? -1 : 1)
    .map(([k, rs]) => ({ key: k, label: periodLabel(k, gran), range: periodRange(k, gran), rows: rs, ...aggregate(rs) }))
    .filter(p => p.n >= minC);
}

/* improvement classification, direction aware */
function classify(k, cur, prev) {
  if (cur == null || prev == null) return { cls: 'flat', dirTxt: 'no comparison', delta: null };
  const delta = cur - prev;
  const rel = Math.abs(delta) / Math.max(Math.abs(prev), 1e-6);
  if (rel < 0.02) return { cls: 'flat', dirTxt: 'flat', delta };
  const better = k.dir === 'up' ? delta > 0 : delta < 0;
  return { cls: better ? 'imp' : 'dec', dirTxt: better ? 'improving' : 'declining', delta };
}
function vsBench(k, v) {
  if (v == null || k.bench == null) return { cls: 'n', txt: '—', gap: null };
  const gap = k.dir === 'up' ? v - k.bench : k.bench - v;   // positive = good
  const rel = Math.abs(gap) / Math.max(Math.abs(k.bench), 1e-6);
  const cls = gap >= 0 ? 'g' : (rel <= 0.2 ? 'a' : 'r');
  return { cls, gap, txt: (k.fmt === 'pct' ? fmtPP(gap) : fmtDN(gap)) };
}
/* least-squares slope over the series, normalised per period */
function slope(vals) {
  const pts = vals.map((v, i) => [i, v]).filter(p => p[1] != null);
  if (pts.length < 3) return null;
  const n = pts.length, sx = pts.reduce((a, p) => a + p[0], 0), sy = pts.reduce((a, p) => a + p[1], 0);
  const sxy = pts.reduce((a, p) => a + p[0] * p[1], 0), sxx = pts.reduce((a, p) => a + p[0] * p[0], 0);
  const d = n * sxx - sx * sx;
  return Math.abs(d) < EPS ? null : (n * sxy - sx * sy) / d;
}

/* ------------------------------------------------------------------ render */
let CTX = {};
function render() {
  const rows = filtered();
  const periods = buildPeriods(rows);
  const overall = aggregate(rows);
  const cur = periods[periods.length - 1] || null;
  const prev = periods[periods.length - 2] || null;
  CTX = { rows, periods, overall, cur, prev };

  renderScope();
  renderCards();
  renderNarrative();
  renderSparks();
  renderDetail();
  renderPeriodTable();
  renderRegionChart();
  renderRegionTable();
}

const GRAN_NOUN = { week: 'weeks', month: 'months', quarter: 'quarters', year: 'years' };
function renderScope() {
  const { rows, periods, cur } = CTX;
  const gran = $('fGran').value;
  const parts = [];
  parts.push(`<b>${rows.length.toLocaleString()}</b> cases in scope`);
  parts.push(`<b>${periods.length}</b> ${GRAN_NOUN[gran] || 'periods'} of history`);
  const reg = $('fRegion').value, zn = $('fZone').value, mo = $('fModality').value;
  parts.push('Region: <b>' + (reg || 'All') + '</b>');
  if (zn) parts.push('Zone: <b>' + zn + '</b>');
  if (mo) parts.push('Modality: <b>' + mo + '</b>');
  parts.push('Date basis: <b>' + ($('fDateCol').value || '—') + '</b>');
  $('scopeBar').innerHTML = parts.join(' &nbsp;&middot;&nbsp; ');
  $('latestPeriodLbl').textContent = cur ? cur.label + ' · ' + cur.n + ' cases' : 'no period data';
  $('granRangeHint').textContent = cur && cur.range
    ? `Showing ${cur.label} — ${fmtDateShort(cur.range.start)} to ${fmtDateShort(cur.range.end)}`
    : 'No period selected';
}

function renderCards() {
  const { overall, cur, prev } = CTX;
  const host = $('kpiCards'); host.innerHTML = '';
  KPIS.forEach(k => {
    const f = fmtOf(k);
    const curV = cur ? cur[k.key] : overall[k.key];
    const prevV = prev ? prev[k.key] : null;
    const c = classify(k, curV, prevV);
    const b = vsBench(k, curV);
    const arrow = c.delta == null ? '' : (c.cls === 'imp' ? '&#9650;' : c.cls === 'dec' ? '&#9660;' : '&#9644;');
    const dcls = c.cls === 'imp' ? 'up' : c.cls === 'dec' ? 'down' : 'neu';
    const el = document.createElement('div');
    el.className = 'card ' + c.cls;
    el.innerHTML = `
      <div class="k-name">${esc(k.label)}${k.secondary ? ' <span class="hint">(secondary)</span>' : ''}</div>
      <div class="k-val">${f(curV)}</div>
      <div class="${dcls}" style="font-size:12.5px">${arrow} ${c.delta == null ? 'no prior period' :
        fmtDelta(k)(c.delta) + ' vs ' + (prev ? prev.label : 'prev') + ' &middot; ' + c.dirTxt}</div>
      <div class="k-row"><span>Top-3 benchmark</span><span>${f(k.bench)}</span></div>
      <div class="k-row"><span>Gap to benchmark</span><span class="badge ${b.cls}">${b.txt}</span></div>
      <div class="k-row"><span>Period average (all scope)</span><span>${f(overall[k.key])}</span></div>`;
    host.appendChild(el);
  });
}

function renderNarrative() {
  const { periods, cur, prev, overall } = CTX;
  const host = $('narrative');
  if (!cur) { host.innerHTML = '<div class="n-item">No period data in the current scope — widen the date range or lower the minimum cases per period.</div>'; return; }
  const gran = $('fGran').value;
  const gname = gran === 'week' ? 'week' : gran === 'month' ? 'month' : gran === 'year' ? 'year' : 'quarter';
  const items = KPIS.map(k => {
    const f = fmtOf(k), fd = fmtDelta(k);
    const c = classify(k, cur[k.key], prev ? prev[k.key] : null);
    const b = vsBench(k, cur[k.key]);
    const series = periods.slice(-6).map(p => p[k.key]);
    const sl = slope(series);
    let trendTxt = '';
    if (sl != null) {
      const better = k.dir === 'up' ? sl > 0 : sl < 0;
      const flat = Math.abs(sl) / Math.max(Math.abs(overall[k.key] || 1), 1e-6) < 0.01;
      trendTxt = flat ? `Trend over the last ${series.length} ${gname}s is broadly flat.`
        : `Trend over the last ${series.length} ${gname}s is <b class="${better ? 'up' : 'down'}">${better ? 'improving' : 'worsening'}</b> (${fd(sl)} per ${gname}).`;
    }
    const cmp = c.delta == null ? 'No previous period to compare.'
      : `${c.cls === 'imp' ? '<b class="up">Improving</b>' : c.cls === 'dec' ? '<b class="down">Declining</b>' : '<b class="neu">Flat</b>'}: ${fd(c.delta)} vs ${esc(prev.label)}.`;
    const absFmt = v => k.fmt === 'pct' ? (Math.abs(v) * 100).toFixed(1) + ' pp' : Math.abs(v).toFixed(2);
    const bench = b.gap == null ? '' : (b.gap >= 0
      ? `Currently <b class="up">at or above</b> the Top-3 benchmark of ${f(k.bench)} (${b.txt}).`
      : `Currently <b class="down">below</b> the Top-3 benchmark of ${f(k.bench)} by ${absFmt(b.gap)}.`);
    return { k, c, html: `<div class="n-item"><b>${esc(k.label)}</b> — ${f(cur[k.key])} in ${esc(cur.label)}. ${cmp} ${bench} ${trendTxt}</div>` };
  });
  const order = { dec: 0, flat: 1, imp: 2 };
  items.sort((a, b) => (order[a.c.cls] - order[b.c.cls]) || (a.k.secondary ? 1 : 0) - (b.k.secondary ? 1 : 0));
  host.innerHTML = items.map(i => i.html).join('');
}

const GRID = '#2a3140', TICK = '#8b949e';
function baseOpts(k, extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: { legend: { display: false }, tooltip: { callbacks: {} } },
    scales: {
      x: { grid: { color: GRID }, ticks: { color: TICK, maxRotation: 45, minRotation: 0, font: { size: 10 } } },
      y: { grid: { color: GRID }, ticks: { color: TICK, font: { size: 10 },
        callback: v => k && k.fmt === 'pct' ? (v * 100).toFixed(0) + '%' : v } }
    }
  }, extra || {});
}
if (typeof Chart !== 'undefined') { Chart.defaults.animation = false; Chart.defaults.color = '#8b949e'; }
function mkChart(canvas, cfg) {
  const id = canvas.id || canvas._cid || (canvas._cid = 'c' + Math.random());
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(canvas, cfg);
  return charts[id];
}

function renderSparks() {
  const { periods, cur, prev } = CTX;
  const host = $('sparkGrid'); host.innerHTML = '';
  if (!periods.length) { host.innerHTML = '<div class="hint">No period data available.</div>'; return; }
  KPIS.forEach(k => {
    const f = fmtOf(k);
    const c = classify(k, cur ? cur[k.key] : null, prev ? prev[k.key] : null);
    const wrap = document.createElement('div');
    wrap.className = 'spark';
    wrap.innerHTML = `<h4>${esc(k.short)}</h4>
      <div class="s-meta">${f(cur ? cur[k.key] : null)} &middot; bench ${f(k.bench)} &middot;
        <span class="${c.cls === 'imp' ? 'up' : c.cls === 'dec' ? 'down' : 'neu'}">${c.dirTxt}</span></div>
      <div class="box"><canvas></canvas></div>`;
    host.appendChild(wrap);
    const cv = wrap.querySelector('canvas'); cv.id = 'spark_' + k.key;
    const color = c.cls === 'imp' ? '#2ea043' : c.cls === 'dec' ? '#e5484d' : '#4c8dff';
    mkChart(cv, {
      type: 'line',
      data: {
        labels: periods.map(p => p.label),
        datasets: [
          { data: periods.map(p => p[k.key]), borderColor: color, backgroundColor: color + '33',
            borderWidth: 2, pointRadius: 2, tension: .25, fill: true, spanGaps: true },
          { data: periods.map(() => k.bench), borderColor: '#8b949e', borderWidth: 1.5,
            borderDash: [5, 4], pointRadius: 0 }
        ]
      },
      options: baseOpts(k, {
        plugins: { legend: { display: false }, tooltip: { callbacks: {
          label: ctx => (ctx.datasetIndex === 0 ? 'Actual: ' : 'Benchmark: ') + f(ctx.parsed.y) } } },
        scales: Object.assign(baseOpts(k).scales, {
          x: { grid: { display: false }, ticks: { color: TICK, font: { size: 9 }, maxTicksLimit: 6 } }
        })
      })
    });
  });
}

function renderDetail() {
  const { periods } = CTX;
  const k = KPI_BY_KEY[$('detailKpi').value] || KPIS[0];
  const f = fmtOf(k);
  const vals = periods.map(p => p[k.key]);
  const cols = periods.map((p, i) => {
    if (i === 0 || vals[i] == null || vals[i - 1] == null) return '#4c8dff';
    const better = k.dir === 'up' ? vals[i] > vals[i - 1] : vals[i] < vals[i - 1];
    return better ? '#2ea043' : '#e5484d';
  });
  mkChart($('detailChart'), {
    type: 'line',
    data: {
      labels: periods.map(p => p.label),
      datasets: [
        { label: k.label, data: vals, borderColor: '#4c8dff', backgroundColor: 'rgba(76,141,255,.12)',
          borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: cols, tension: .25, fill: true, spanGaps: true },
        { label: 'Top-3 benchmark', data: periods.map(() => k.bench), borderColor: '#e3a008',
          borderWidth: 2, borderDash: [6, 5], pointRadius: 0 }
      ]
    },
    options: baseOpts(k, {
      plugins: {
        legend: { display: true, labels: { color: '#e6edf3', boxWidth: 14 } },
        title: { display: true, color: '#8b949e', text: k.def, font: { size: 11, weight: 'normal' } },
        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + f(ctx.parsed.y) } }
      }
    })
  });
  mkChart($('volumeChart'), {
    type: 'bar',
    data: { labels: periods.map(p => p.label),
      datasets: [{ label: 'Cases', data: periods.map(p => p.n), backgroundColor: '#3d4757' }] },
    options: baseOpts(null, { plugins: { legend: { display: false },
      title: { display: true, color: '#8b949e', text: 'Case volume per period', font: { size: 11, weight: 'normal' } } } })
  });
}

function renderPeriodTable() {
  const { periods } = CTX;
  const t = $('periodTable');
  if (!periods.length) { t.innerHTML = '<tbody><tr><td>No period data.</td></tr></tbody>'; return; }
  let h = '<thead><tr><th>Period</th><th>Cases</th>';
  KPIS.forEach(k => h += `<th>${esc(k.short)}</th>`);
  h += '</tr></thead><tbody>';
  periods.slice().reverse().forEach((p, ri, arr) => {
    const prevP = arr[ri + 1];
    h += `<tr><td><b>${esc(p.label)}</b></td><td class="num">${p.n}</td>`;
    KPIS.forEach(k => {
      const f = fmtOf(k), v = p[k.key];
      const b = vsBench(k, v);
      const c = classify(k, v, prevP ? prevP[k.key] : null);
      const arrow = c.delta == null ? '' :
        `<span class="${c.cls === 'imp' ? 'up' : c.cls === 'dec' ? 'down' : 'neu'}"> ${c.cls === 'imp' ? '▲' : c.cls === 'dec' ? '▼' : '▬'}${fmtDelta(k)(c.delta)}</span>`;
      h += `<td class="num ${b.cls}">${f(v)}${arrow}</td>`;
    });
    h += '</tr>';
  });
  h += '<tr class="bench"><td>Top-3 benchmark</td><td class="num">—</td>';
  KPIS.forEach(k => h += `<td class="num">${fmtOf(k)(k.bench)}</td>`);
  h += '</tr></tbody>';
  t.innerHTML = h;
}

/* ---- region level (recomputed from cases when available, else plan sheet) ---- */
function regionStats() {
  if (S.hasCases && CTX.rows.length) {
    const m = new Map();
    CTX.rows.forEach(r => { if (!m.has(r.region)) m.set(r.region, []); m.get(r.region).push(r); });
    return [...m.entries()].map(([name, rs]) => ({ name, n: rs.length, ...aggregate(rs), source: 'cases' }));
  }
  /* fall back to the action-plan snapshot */
  return S.plan.map(p => {
    const o = { name: p._region, n: n0(p[S.planCols.find(c => norm(c) === 'cases')]), source: 'plan' };
    KPIS.forEach(k => {
      const col = S.planCols.find(c => norm(c) === norm(k.label));
      o[k.key] = col ? toNum(p[col]) : null;
    });
    return o;
  });
}

function renderRegionChart() {
  const k = KPI_BY_KEY[$('regionKpi').value] || KPIS[0];
  const sort = $('regionSort').value;
  const hasBench = k.bench != null;
  let st = regionStats().filter(r => r.name && r[k.key] != null);
  const gapOf = r => !hasBench ? 0 : (k.dir === 'up' ? r[k.key] - k.bench : k.bench - r[k.key]);
  if (sort === 'gap' && hasBench) st.sort((a, b) => gapOf(a) - gapOf(b));
  else if (sort === 'value' || sort === 'gap') st.sort((a, b) => k.dir === 'up' ? a[k.key] - b[k.key] : b[k.key] - a[k.key]);
  else if (sort === 'cases') st.sort((a, b) => b.n - a.n);
  else st.sort((a, b) => a.name.localeCompare(b.name));

  const f = fmtOf(k);
  $('regionChart').parentElement.style.height = Math.max(320, st.length * 22 + 70) + 'px';
  mkChart($('regionChart'), {
    type: 'bar',
    data: {
      labels: st.map(r => `${r.name} (${r.n})`),
      datasets: [
        { label: k.label, data: st.map(r => r[k.key]),
          backgroundColor: st.map(r => !hasBench ? '#4c8dff' : gapOf(r) >= 0 ? '#2ea043'
            : (Math.abs(gapOf(r)) / Math.max(Math.abs(k.bench), 1e-6) <= 0.2 ? '#e3a008' : '#e5484d')),
          borderWidth: 0, barPercentage: .85, categoryPercentage: .9 },
        ...(hasBench ? [{ label: 'Top-3 benchmark', type: 'line', data: st.map(() => k.bench),
          borderColor: '#c9d1d9', borderWidth: 2, borderDash: [6, 4], pointRadius: 0 }] : [])
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#e6edf3', boxWidth: 14 } },
        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + f(ctx.parsed.x) } }
      },
      scales: {
        x: { grid: { color: GRID }, ticks: { color: TICK, callback: v => k.fmt === 'pct' ? (v * 100).toFixed(0) + '%' : v } },
        y: { grid: { display: false }, ticks: { color: TICK, font: { size: 10 }, autoSkip: false } }
      }
    }
  });
}

function renderRegionTable() {
  const st = regionStats().filter(r => r.name).sort((a, b) => b.n - a.n);
  const planIdx = Object.fromEntries(S.plan.map(p => [norm(p._region), p]));
  const focusCol = S.planCols.find(c => /primaryfocus/.test(norm(c)));
  const belowCol = S.planCols.find(c => /metricsbelow/.test(norm(c)));
  const prioCol = S.planCols.find(c => /^priority/.test(norm(c)));
  let h = '<thead><tr><th>Region</th><th>Cases</th>';
  KPIS.forEach(k => h += `<th>${esc(k.short)}</th>`);
  h += '<th>Below benchmark</th><th>Priority</th><th>Primary focus</th></tr></thead><tbody>';
  st.forEach(r => {
    let below = 0;
    let cells = '';
    KPIS.forEach(k => {
      const b = vsBench(k, r[k.key]);
      if (b.gap != null && b.gap < 0) below++;
      cells += `<td class="num ${b.cls}">${fmtOf(k)(r[k.key])}</td>`;
    });
    const p = planIdx[norm(r.name)] || {};
    h += `<tr><td><b>${esc(r.name)}</b></td><td class="num">${r.n}</td>${cells}` +
      `<td class="num">${below} / ${KPIS.length}</td>` +
      `<td>${esc(prioCol ? (p[prioCol] ?? '') : '')}</td>` +
      `<td class="wrap">${esc(focusCol ? (p[focusCol] ?? '') : '')}</td></tr>`;
  });
  h += '<tr class="bench"><td>Top-3 benchmark</td><td class="num">—</td>';
  KPIS.forEach(k => h += `<td class="num">${fmtOf(k)(k.bench)}</td>`);
  h += '<td></td><td></td><td></td></tr></tbody>';
  $('regionTable').innerHTML = h;
}

function renderPlanTable() {
  const t = $('planTable');
  if (!S.plan.length) { $('planPanel').classList.add('hidden'); return; }
  const cols = S.planCols.filter(c => c && (/^region$/i.test(c) || /focus|30|60|90|priority|confidence|benchmark|plan/i.test(c)));
  if (cols.length < 2) { $('planPanel').classList.add('hidden'); return; }
  let h = '<thead><tr>' + cols.map(c => `<th>${esc(c)}</th>`).join('') + '</tr></thead><tbody>';
  S.plan.forEach(p => {
    h += '<tr>' + cols.map(c => {
      const v = p[c] == null ? '' : String(p[c]).trim();
      return `<td class="${v.length > 40 ? 'wrap' : ''}">${esc(v)}</td>`;
    }).join('') + '</tr>';
  });
  $('planTable').innerHTML = h + '</tbody>';
}

function renderDefs() {
  let h = '<thead><tr><th>KPI</th><th>What it measures</th><th>Direction</th><th>Formula used by this dashboard</th><th>Top-3 benchmark</th></tr></thead><tbody>';
  KPIS.forEach(k => {
    h += `<tr><td><b>${esc(k.label)}</b></td><td class="wrap">${esc(k.def)}</td>` +
      `<td>${k.dir === 'up' ? 'Higher is better' : 'Lower is better'}</td>` +
      `<td class="wrap"><code>${esc(k.calc)}</code></td>` +
      `<td class="num">${fmtOf(k)(k.bench)}</td></tr>`;
  });
  $('defTable').innerHTML = h + '</tbody>';
}

/* ------------------------------------------------------------------ export */
function download(name, text) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  a.download = name; a.click(); URL.revokeObjectURL(a.href);
}
const csvCell = v => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;

function exportPeriods() {
  const rows = [['Period', 'Cases', ...KPIS.map(k => k.label)]];
  CTX.periods.forEach(p => rows.push([p.label, p.n, ...KPIS.map(k => p[k.key] == null ? '' : p[k.key])]));
  rows.push(['Top-3 benchmark', '', ...KPIS.map(k => k.bench)]);
  download('rdf_kpi_periods.csv', rows.map(r => r.map(csvCell).join(',')).join('\n'));
}
function exportRegions() {
  const rows = [['Region', 'Cases', ...KPIS.map(k => k.label)]];
  regionStats().forEach(r => rows.push([r.name, r.n, ...KPIS.map(k => r[k.key] == null ? '' : r[k.key])]));
  rows.push(['Top-3 benchmark', '', ...KPIS.map(k => k.bench)]);
  download('rdf_kpi_regions.csv', rows.map(r => r.map(csvCell).join(',')).join('\n'));
}

/* ------------------------------------------------------ local data assistant */
const ASSISTANT_GREETING =
  'Hi! How can I assist you today? I can answer questions based on the Excel workbook currently uploaded to this dashboard.';
const ASSISTANT_HELP =
  'I calculate from the latest uploaded data only. I can count, filter, group, find distinct values, sum, average, median, minimum, maximum, rank, list matching rows, compare regions, calculate percentages/shares, look up a single case by its case number, and calculate the dashboard KPIs.\n\n' +
  'Use any loaded column name. Combine filters with “and”; use “>”, “<”, “=”, or “between X and Y” on any numeric column; group by up to three columns. Examples: “Count cases with RDF 1 and FVF 0”, “Average Visits by Region and Modality”, “Sum Custom Cost where T2 Engineer is Alice”, “Show cases where Visits between 2 and 5”, “What percentage of cases have RDF 1?”, “Top 10 cases by Onsite Hours”, or “Give field Remarks for case 0124999777”.\n\n' +
  'I show the applied scope, filters, sample size, and calculation. If I cannot map a request to loaded data unambiguously, I will ask for clearer wording rather than invent a result.';
const CHAT_ROW_CAP = 25;
const CHAT_ALIASES = {
  caseNo: ['case number', 'case id'],
  region: ['ib region', 'region'],
  zone: ['ib zone', 'zone'],
  modality: ['modality'],
  market: ['market'],
  visits: ['visits', 'visit count'],
  onsite: ['onsite hours', 'onsite'],
  partsOrdered: ['parts ordered'],
  partsConsumed: ['parts consumed'],
  partsReturned: ['returned parts qty', 'parts returned qty', 'parts returned'],
  spa: ['service plan availability', 'service plan available', 'service plan'],
  rdf: ['rdf'],
  frusAdvised: ['frus advised', 'fru advised'],
  orderedBefore: ['advised part ordered before first visit', 'ordered before first visit']
};
const KPI_ALIASES = {
  partsAdvised: ['parts advised success', 'parts advised rate'],
  visits: ['visits per case', 'visit per case'],
  servicePlan: ['service plan availability', 'service plan rate'],
  rdfSuccess: ['rdf success rate', 'rdf success', 'rdf rate'],
  preFirst: ['pre first visit ordering success', 'pre-first-visit ordering success', 'pre first visit'],
  partsConsumed: ['parts consumed per case'],
  onsite: ['onsite hours per case', 'onsite per case'],
  partsReturn: ['parts return rate', 'parts return']
};

function resetAssistantConversation(showGreeting) {
  const host = $('chatMessages');
  if (!host) return;
  host.innerHTML = '';
  if (typeof resetChatHistory === 'function') resetChatHistory();
  lastCaseLookup = null;
  if (showGreeting !== false) {
    addChatMessage('assistant', ASSISTANT_GREETING);
    if (!S.cases.length) {
      addChatMessage('assistant', 'Please upload an Excel workbook first. I will use its parsed case-level data locally in this browser.');
    }
  }
}

function updateAssistantScope() {
  const el = $('chatScope');
  if (!el) return;
  const table = activeAssistantTable();
  el.textContent = table && table.type === 'report'
    ? `${table.rows.length.toLocaleString()} report rows · ${table.name}`
    : S.cases.length
    ? `${S.cases.length.toLocaleString()} cases · full uploaded dataset`
    : 'No workbook loaded';
}

function refreshAssistantDatasets(selectedId) {
  const select = $('chatDataset');
  if (!select) return;
  const previous = selectedId || select.value;
  select.innerHTML = '';
  S.assistantTables.forEach(table => select.add(new Option(
    `${table.type === 'cases' ? 'Case data' : 'Report'}: ${table.name}`, table.id)));
  if (previous && S.assistantTables.some(table => table.id === previous)) select.value = previous;
  else if (S.assistantTables.length) select.value = S.assistantTables[S.assistantTables.length - 1].id;
  updateAssistantScope();
}

function activeAssistantTable() {
  const selected = $('chatDataset') && $('chatDataset').value;
  return S.assistantTables.find(table => table.id === selected) || S.assistantTables[0] || null;
}

function addChatMessage(role, text, result) {
  const host = $('chatMessages');
  const wrap = document.createElement('div');
  wrap.className = `chat-msg ${role}`;
  const bubble = document.createElement('div');
  bubble.className = 'chat-bubble';
  const summary = document.createElement('div');
  summary.className = 'chat-summary';
  summary.textContent = text;
  bubble.appendChild(summary);

  if (result && result.columns && result.rows && result.rows.length) {
    const exportButton = document.createElement('button');
    exportButton.type = 'button';
    exportButton.className = 'chat-export-btn';
    exportButton.textContent = 'Export Excel';
    exportButton.setAttribute('aria-label', 'Export this displayed result table to Excel');
    exportButton.addEventListener('click', () => exportChatResult(result));
    bubble.appendChild(exportButton);
    const tableWrap = document.createElement('div');
    tableWrap.className = 'chat-result-wrap';
    const table = document.createElement('table');
    table.className = 'chat-result';
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    result.columns.forEach(c => {
      const th = document.createElement('th');
      th.textContent = c;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    result.rows.forEach(row => {
      const tr = document.createElement('tr');
      row.forEach(value => {
        const td = document.createElement('td');
        td.textContent = displayCell(value);
        td.title = displayCell(value);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tableWrap.appendChild(table);
    bubble.appendChild(tableWrap);
  }
  if (result && result.meta) {
    const meta = document.createElement('div');
    meta.className = 'chat-meta';
    meta.textContent = result.meta;
    bubble.appendChild(meta);
  }
  wrap.appendChild(bubble);
  host.appendChild(wrap);
  host.scrollTop = host.scrollHeight;
}

function exportChatResult(result) {
  if (!result || !result.columns || !result.rows || !result.rows.length || typeof XLSX === 'undefined') return;
  /* Export the rendered result only: never replace its row cap with the full source dataset. */
  const visibleTable = [result.columns, ...result.rows.map(row => row.map(displayCell))];
  const sheet = XLSX.utils.aoa_to_sheet(visibleTable);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, 'Chat result');
  const fileName = `nam-chat-result-${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(book, fileName, { compression: true });
}

function displayCell(v) {
  if (v == null || v === '') return '—';
  if (v instanceof Date) return iso(v);
  if (typeof v === 'number') return Number.isInteger(v) ? v.toLocaleString() : v.toLocaleString(undefined, { maximumFractionDigits: 2 });
  return String(v);
}

/* ------------------------------------------------------ typo-tolerant matching
   Common misspellings of column names, KPI names, and dimension values (e.g.
   "vists" for "visits", "bostn" for "Boston") are corrected before parsing, so
   the same query logic below still works unchanged. Correction is intentionally
   conservative: only words of 4+ letters are considered, English query/grammar
   words are never touched, and an ambiguous best match is left uncorrected
   rather than guessed. Every correction made is disclosed in the answer's meta
   line so results are never silently reinterpreted. */
const QUERY_STOPWORDS = new Set([
  'the', 'a', 'an', 'is', 'are', 'was', 'were', 'and', 'or', 'for', 'in', 'on', 'at', 'of', 'to', 'with',
  'where', 'have', 'has', 'show', 'list', 'count', 'how', 'many', 'much', 'number', 'what', 'which', 'total',
  'sum', 'average', 'avg', 'mean', 'median', 'minimum', 'maximum', 'min', 'max', 'top', 'bottom', 'highest',
  'lowest', 'best', 'worst', 'distinct', 'unique', 'different', 'types', 'type', 'by', 'compare', 'between',
  'case', 'cases', 'columns', 'column', 'available', 'can', 'you', 'calculate', 'do', 'help', 'supported',
  'percent', 'percentage', 'share', 'proportion', 'last', 'month', 'this', 'that', 'year', 'all', 'than',
  'greater', 'less', 'equal', 'equals', 'yes', 'no', 'true', 'false', 'created', 'creation', 'rate', 'success',
  'per', 'not', 'from', 'value', 'values', 'row', 'rows', 'data', 'workbook', 'excel', 'sheet', 'export',
  'name', 'names', 'please', 'about', 'currently', 'uploaded', 'dashboard', 'january', 'february', 'march',
  'april', 'june', 'july', 'august', 'september', 'october', 'november', 'december'
]);

function levenshtein(a, b) {
  a = String(a); b = String(b);
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  /* Damerau-Levenshtein (optimal string alignment): standard insert/delete/substitute
     plus adjacent-transposition at cost 1. Swapped adjacent letters (e.g. "scpoe" for
     "scope") are the single most common real-world typo, and plain Levenshtein scores
     them as 2 edits — often exceeding the fuzzy-match allowance and silently leaving
     the typo uncorrected. Treating a transposition as 1 edit fixes that class of typo
     across every fuzzy-matched word (columns, region/zone/modality/market values). */
  const d = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) d[i][0] = i;
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[m][n];
}

function fuzzyAllowance(len) {
  return len <= 3 ? 0 : len <= 6 ? 1 : 2;
}

function buildSpellingVocab(table) {
  const words = new Set();
  const addPhrase = phrase => {
    const matches = String(phrase == null ? '' : phrase).toLowerCase().match(/[a-z0-9]+/g);
    if (matches) matches.forEach(w => words.add(w));
  };
  if (table && table.type === 'report') {
    table.headers.forEach(addPhrase);
    const labelHeader = table.headers.find(h => /^(ib region|region|kpi)$/i.test(h));
    if (labelHeader) uniq(table.rows.map(r => rawValue(r, labelHeader))).forEach(addPhrase);
  } else if (S.caseHeaders && S.caseHeaders.length) {
    S.caseHeaders.forEach(header => {
      addPhrase(header);
      columnTerms(header).forEach(addPhrase);
    });
    Object.values(KPI_ALIASES).forEach(list => list.forEach(addPhrase));
    ['region', 'zone', 'modality', 'market'].forEach(key => {
      const header = S.caseColumnMap[key];
      if (!header) return;
      uniq(S.cases.map(r => rawValue(r, header))).forEach(addPhrase);
    });
  }
  return words;
}

function correctSpelling(question, table) {
  const vocab = buildSpellingVocab(table);
  const corrections = [];
  if (!vocab.size) return { corrected: question, corrections };
  const corrected = question.replace(/[a-zA-Z][a-zA-Z0-9]*/g, token => {
    const lower = token.toLowerCase();
    if (token.length < 4 || vocab.has(lower) || QUERY_STOPWORDS.has(lower)) return token;
    const allowance = fuzzyAllowance(lower.length);
    let best = null, bestDist = allowance + 1, ambiguous = false;
    vocab.forEach(word => {
      if (Math.abs(word.length - lower.length) > allowance) return;
      const d = levenshtein(lower, word);
      if (d < bestDist) { bestDist = d; best = word; ambiguous = false; }
      else if (d === bestDist && word !== best) ambiguous = true;
    });
    if (best && !ambiguous && bestDist > 0 && bestDist <= allowance) {
      corrections.push({ from: token, to: best });
      return best;
    }
    return token;
  });
  return { corrected, corrections };
}

function findKpi(question) {
  const q = norm(question);
  for (const key of Object.keys(KPI_ALIASES)) {
    if (KPI_ALIASES[key].some(a => q.includes(norm(a)))) return KPI_BY_KEY[key];
  }
  return null;
}

function columnTerms(header) {
  const terms = [header];
  if (/\b\w+s$/i.test(header)) terms.push(header.replace(/s$/i, ''));
  else terms.push(header + 's');
  Object.keys(S.caseColumnMap).forEach(key => {
    if (S.caseColumnMap[key] === header) terms.push(...(CHAT_ALIASES[key] || []));
  });
  return [...new Set(terms)].sort((a, b) => b.length - a.length);
}

function findColumns(question, options) {
  const excluded = new Set((options && options.exclude) || []);
  const matches = [];
  S.caseHeaders.forEach(header => {
    if (excluded.has(header)) return;
    const term = columnTerms(header).find(candidate => {
      /* Collapse literal spaces in a multi-word column term to "one or more
         whitespace characters" so accidental double-spacing or a stray tab
         between words (e.g. "rdf  scope" with two spaces) still matches the
         column, instead of silently falling through to a shorter column that
         happens to share the same first word (e.g. plain "RDF"). */
      const escaped = regexEscapeTerm(candidate);
      return new RegExp(`(^|[^a-z0-9])${escaped}(?![a-z0-9])`, 'i').test(question);
    });
    if (term) matches.push({ header, score: norm(term).length });
  });
  return matches.sort((a, b) => b.score - a.score).map(match => match.header);
}

function findColumn(question, options) {
  return findColumns(question, options)[0] || null;
}

function rawValue(row, header) {
  return row.raw ? row.raw[header] : null;
}

function regexEscape(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/* Escapes a column term for use in a RegExp, and collapses any literal spaces
   into "one or more whitespace" so accidental double-spacing between the
   words of a multi-word column name (e.g. "rdf  scope") still matches, rather
   than only matching a shorter column that shares the same first word. */
function regexEscapeTerm(s) {
  return regexEscape(s).replace(/ /g, '\\s+');
}

function queryFilters(question, options) {
  const q = String(question).toLowerCase();
  const skipRegions = options && options.skipRegions;
  const skipValueConditions = options && options.skipValueConditions;
  let rows = S.cases.slice();
  const descriptions = [];
  const inlineEqualityHeaders = new Set();
  const dimensionKeys = skipRegions ? ['zone', 'modality', 'market'] : ['region', 'zone', 'modality', 'market'];

  dimensionKeys.forEach(key => {
    const header = S.caseColumnMap[key];
    if (!header) return;
    const vals = uniq(S.cases.map(r => rawValue(r, header))).sort((a, b) => String(b).length - String(a).length);
    const matched = vals.find(v => {
      const text = String(v).toLowerCase();
      return new RegExp(`(^|[^a-z0-9])${regexEscapeTerm(text)}([^a-z0-9]|$)`, 'i').test(q);
    });
    if (matched != null) {
      rows = rows.filter(r => String(rawValue(r, header) ?? '').toLowerCase() === String(matched).toLowerCase());
      descriptions.push(`${header} = ${matched}`);
    }
  });

  /* "between A and B" range filters, checked before plain comparisons so a
     header is not double-matched by both branches. Skipped when computing a
     percentage question's denominator (the base scope before the condition). */
  if (!skipValueConditions) S.caseHeaders.forEach(header => {
    const terms = columnTerms(header);
    for (const term of terms) {
      const rx = new RegExp(`${regexEscapeTerm(term)}\\s*(?:is\\s+)?between\\s*(-?[\\d,]+(?:\\.\\d+)?)\\s*(?:and|-|to)\\s*(-?[\\d,]+(?:\\.\\d+)?)`, 'i');
      const m = question.match(rx);
      if (!m) continue;
      const lo = toNum(m[1]), hi = toNum(m[2]);
      if (lo == null || hi == null) continue;
      const min = Math.min(lo, hi), max = Math.max(lo, hi);
      rows = rows.filter(r => {
        const value = toNum(rawValue(r, header));
        return value != null && value >= min && value <= max;
      });
      descriptions.push(`${header} between ${min} and ${max}`);
      break;
    }
  });

  /* Plain comparisons (>, <, >=, <=, =). A column may appear more than once
     (e.g. "Visits > 1 and Visits < 5"), so every occurrence is applied as an
     AND condition rather than only the first match. */
  if (!skipValueConditions) S.caseHeaders.forEach(header => {
    const terms = columnTerms(header);
    for (const term of terms) {
      const rx = new RegExp(`${regexEscapeTerm(term)}\\s*(>=|<=|>|<|=)\\s*(-?[\\d,]+(?:\\.\\d+)?)`, 'gi');
      const matches = [...question.matchAll(rx)];
      if (!matches.length) continue;
      matches.forEach(m => {
        const target = toNum(m[2]);
        if (target == null) return;
        rows = rows.filter(r => {
          const value = toNum(rawValue(r, header));
          if (value == null) return false;
          return m[1] === '>' ? value > target : m[1] === '<' ? value < target
            : m[1] === '>=' ? value >= target : m[1] === '<=' ? value <= target : value === target;
        });
        descriptions.push(`${header} ${m[1]} ${target}`);
      });
      break;
    }
  });

  /* Support multiple plain-language equalities joined by "and", including shorthand
     such as "cases with RDF 0 and FVF is 1". Each clause is an AND condition. */
  if (!skipValueConditions) S.caseHeaders.forEach(header => {
    const terms = columnTerms(header).sort((a, b) => b.length - a.length);
    for (const term of terms) {
      const rx = new RegExp(`(?:^|\\b(with|where|and|have|has)\\s+)${regexEscapeTerm(term)}(?![a-z0-9])\\s*(?:(=|is|equals)\\s*)?["']?([^,?]+?)["']?(?=\\s+\\band\\s+|[?!.]?\\s*$)`, 'i');
      const m = question.match(rx);
      if (!m) continue;
      const wanted = m[3].trim();
      if (!wanted || /^(?:is|equals)$/i.test(wanted)) continue;
      if (/[<>]=?/.test(wanted) || /^\s*=/.test(wanted)) continue; /* already handled by the comparison filter above */
      if (!m[2]) {
        const remaining = question.slice((m.index || 0) + m[0].length);
        const followingClause = remaining.match(/^\s+and\s+(.+)$/i);
        /* A bare value is valid on its own only when introduced by "have"/"has"
           (unambiguous equality, e.g. "cases have RDF 1") or when it chains to
           another recognized column via "and". This preserves "RDF 0 and 1"
           (introduced by "with") as a same-column distribution question. */
        const unambiguousIntroducer = /^(?:have|has)$/i.test(m[1] || '');
        /* The chained clause only counts as a legitimate different-column AND when it
           resolves to a column other than the one just matched; "RDF 1 and RDF 0" must
           stay a same-column distribution question, not an (impossible) RDF=1 AND RDF=0. */
        const chainedColumn = followingClause ? findColumn(followingClause[1]) : null;
        const hasDifferentChainedColumn = chainedColumn && chainedColumn !== header;
        if (!unambiguousIntroducer && !hasDifferentChainedColumn && !inlineEqualityHeaders.size) continue;
      }
      rows = rows.filter(r => String(rawValue(r, header) ?? '').trim().toLowerCase() === wanted.toLowerCase());
      descriptions.push(`${header} = ${wanted}`);
      inlineEqualityHeaders.add(header);
      break;
    }
  });

  if (!skipValueConditions && !inlineEqualityHeaders.size) {
    const equality = question.match(/\b(?:where|with)\s+(.+?)\s+(?:=|is|equals)\s+["']?([^,"']+?)["']?[?!.]?\s*$/i);
    if (equality) {
      const header = findColumn(equality[1]);
      if (header) {
        const wanted = equality[2].trim().toLowerCase();
        rows = rows.filter(r => String(rawValue(r, header) ?? '').trim().toLowerCase() === wanted);
        descriptions.push(`${header} = ${equality[2].trim()}`);
        inlineEqualityHeaders.add(header);
      }
    }
  }

  let dateHeader = null;
  if (/\b(created|creation)\b/i.test(question))
    dateHeader = S.dateCols.find(h => /creat/i.test(h));
  dateHeader = dateHeader || findColumn(question, { exclude: S.caseHeaders.filter(h => !S.dateCols.includes(h)) });
  if (!dateHeader && (/\b(last month|this month)\b/i.test(question) ||
      /\b(19|20)\d{2}\b/.test(question) ||
      MON.some(m => new RegExp(`\\b${m}`, 'i').test(question))))
    dateHeader = S.dateCols.find(h => /creat/i.test(h)) || S.dateCols[0];
  if (dateHeader) {
    let start = null, end = null;
    const now = new Date();
    if (/\blast month\b/i.test(question)) {
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
    } else if (/\bthis month\b/i.test(question)) {
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    } else {
      const year = question.match(/\b(19|20)\d{2}\b/);
      const month = MON.findIndex(m => new RegExp(`\\b${m}(?:uary|ruary|ch|il|e|y|ust|tember|ober|ember)?\\b`, 'i').test(question));
      if (year) {
        const y = +year[0];
        start = new Date(y, month >= 0 ? month : 0, 1);
        end = month >= 0 ? new Date(y, month + 1, 0, 23, 59, 59) : new Date(y, 11, 31, 23, 59, 59);
      }
    }
    if (start && end) {
      rows = rows.filter(r => {
        const d = r.d[dateHeader] || toDate(rawValue(r, dateHeader));
        return d && d >= start && d <= end;
      });
      descriptions.push(`${dateHeader}: ${iso(start)} to ${iso(end)}`);
    }
  }
  return { rows, descriptions, inlineEqualityCount: inlineEqualityHeaders.size };
}

/* Remembers the most recently looked-up single case so a follow-up question
   that only names a field (e.g. "field remarks", "remote remark", "open field
   remarks") can be answered without repeating the case number, similar to how
   a person would naturally continue a conversation about "this case". Reset
   whenever the chat is cleared or a new workbook is uploaded. */
let lastCaseLookup = null;

/* Single-case lookup, e.g. "give field remarks for case 0124999777" or
   "show case 0124999777". Deliberately scoped to the singular word "case"
   (the \b...\b boundary excludes "cases") so it never collides with the
   existing plural case-count/filter/list questions. The extracted identifier
   must contain a digit, so ordinary words like "cases where..." can never be
   mistaken for a case number. Every occurrence of "case" is checked (not just
   the first) because a field name can itself contain the word, e.g.
   "field Case Owner for case B-1" — the first "case" there has no digit.
   "this case"/"that case"/"same case" refers back to the last case looked up
   in this conversation. */
function findCaseLookup(question) {
  const caseHeader = S.caseColumnMap.caseNo;
  if (!caseHeader || !/\bcase\b/i.test(question)) return null;
  const rx = /\bcase\b\s*(?:number|no\.?|id)?\s*[:#]?\s*["']?([A-Za-z0-9][A-Za-z0-9\-\/]{2,})["']?/gi;
  for (const idMatch of question.matchAll(rx)) {
    const candidate = idMatch[1].replace(/[?!.,]+$/, '');
    if (!/\d/.test(candidate)) continue;
    const row = S.cases.find(r => String(rawValue(r, caseHeader) ?? '').trim().toLowerCase() === candidate.toLowerCase());
    return { candidate, row, matchText: idMatch[0], caseHeader };
  }
  const refersToLastCase = question.match(/\b(?:this|that|same|the)\s+case\b/i);
  if (refersToLastCase && lastCaseLookup) {
    return { candidate: lastCaseLookup.candidate, row: lastCaseLookup.row, matchText: refersToLastCase[0], caseHeader: lastCaseLookup.caseHeader };
  }
  return null;
}

/* Resolves the column a question is asking about, given the text remaining
   after any case-identifier phrase has been removed. Tries a direct match
   first so real column names that happen to contain a generic introducer
   word (e.g. "Field Remarks", "Remote Remarks") still match correctly; only
   falls back to stripping "field"/"value"/"column" as an introducer word when
   nothing matched directly (so "field Case Owner" still resolves to the
   "Case Owner" column, which does not contain the word "field"). */
function extractRequestedField(text, excludeHeader) {
  const options = excludeHeader ? { exclude: [excludeHeader] } : undefined;
  let field = findColumn(text, options);
  if (field) return field;
  const stripped = text.replace(/\b(?:field|value|column)\b/gi, ' ');
  if (stripped !== text) field = findColumn(stripped, options);
  return field;
}

function filterMeta(filters, sampleSize) {
  const scope = filters.descriptions.length ? filters.descriptions.join('; ') : 'Full uploaded dataset';
  return `Scope: ${scope} · Sample size: ${sampleSize.toLocaleString()} case${sampleSize === 1 ? '' : 's'}`;
}

function formatKpiValue(k, value) {
  return value == null ? 'Not calculable from the available columns' : fmtOf(k)(value);
}

function groupedRows(rows, header) {
  const groups = new Map();
  rows.forEach(r => {
    const value = rawValue(r, header);
    const label = value == null || value === '' ? '(blank)' : String(value);
    if (!groups.has(label)) groups.set(label, []);
    groups.get(label).push(r);
  });
  return [...groups.entries()];
}

function groupedRowsBy(rows, headers) {
  const groups = new Map();
  rows.forEach(row => {
    const values = headers.map(header => {
      const value = rawValue(row, header);
      return value == null || value === '' ? '(blank)' : String(value);
    });
    const key = JSON.stringify(values);
    if (!groups.has(key)) groups.set(key, { values, rows: [] });
    groups.get(key).rows.push(row);
  });
  return [...groups.values()];
}

function findReportColumn(table, question) {
  const qn = norm(question);
  const hits = table.headers.filter(header => {
    const h = norm(header);
    return h.length >= 3 && (qn.includes(h) || qn.includes(h + 's'));
  }).sort((a, b) => norm(b).length - norm(a).length);
  return hits[0] || null;
}

function filterReportRows(table, question) {
  const q = String(question).toLowerCase();
  const labelHeader = table.headers.find(header => /^(ib region|region|kpi)$/i.test(header));
  if (!labelHeader) return { rows: table.rows, description: '' };
  const labels = [...new Set(table.rows.map(row => rawValue(row, labelHeader))
    .filter(value => value != null && String(value).trim()))].sort((a, b) => String(b).length - String(a).length);
  const matched = labels.find(value => {
    const label = String(value).toLowerCase();
    return new RegExp(`(^|[^a-z0-9])${regexEscapeTerm(label)}([^a-z0-9]|$)`, 'i').test(q);
  });
  if (matched == null) return { rows: table.rows, description: '' };
  return {
    rows: table.rows.filter(row => String(rawValue(row, labelHeader)).toLowerCase() === String(matched).toLowerCase()),
    description: `${labelHeader} = ${matched}`
  };
}

function answerReportQuestion(question, table) {
  const q = String(question).trim();
  const ql = q.toLowerCase();
  const rows = table.rows;
  if (/\b(list|show|what).*(available )?columns\b|\bcolumn names\b/i.test(q)) {
    return { summary: `${table.name} contains ${table.headers.length} columns.`,
      columns: ['#', 'Column'], rows: table.headers.map((h, i) => [i + 1, h]),
      meta: `Source: ${table.name} · ${rows.length} summary rows` };
  }
  const header = findReportColumn(table, q);
  if (!header) {
    return { summary: `Please name a column from ${table.name} so I can calculate without guessing.`,
      meta: `Available columns: ${table.headers.slice(0, 8).join(', ')}${table.headers.length > 8 ? '…' : ''}` };
  }
  const filtered = filterReportRows(table, q);
  const reportRows = filtered.rows;
  const reportMeta = `Source: ${table.name} · ${filtered.description || `${rows.length} summary rows`}`;
  const values = reportRows.map(row => ({ row, value: rawValue(row, header) }))
    .filter(item => item.value != null && item.value !== '');
  if (/\b(distinct|unique)\b|\bdifferent\s+types?\s+of\b|\btypes?\s+of\b|\bhow many different\b/i.test(q)) {
    const groups = new Map();
    values.forEach(({ value }) => {
      const key = String(value);
      groups.set(key, (groups.get(key) || 0) + 1);
    });
    const result = [...groups.entries()].sort((a, b) => b[1] - a[1]);
    return { summary: `${result.length} distinct ${header} values in ${table.name}.`,
      columns: [header, 'Rows'], rows: result,
      meta: reportMeta };
  }
  const nums = values.map(item => toNum(item.value)).filter(value => value != null);
  const operation = /\baverage|avg|mean\b/i.test(q) ? 'average'
    : /\bmedian\b/i.test(q) ? 'median'
    : /\bsum|total\b/i.test(q) ? 'sum'
    : /\bminimum|min\b/i.test(q) ? 'minimum'
    : /\bmaximum|max\b/i.test(q) ? 'maximum' : null;
  if (operation) {
    if (!nums.length) return { summary: `${header} has no numeric values to ${operation}.`,
      meta: reportMeta };
    const value = operation === 'sum' ? nums.reduce((a, b) => a + b, 0)
      : operation === 'average' ? nums.reduce((a, b) => a + b, 0) / nums.length
      : operation === 'median' ? median(nums)
      : operation === 'minimum' ? Math.min(...nums) : Math.max(...nums);
    return { summary: `${operation[0].toUpperCase() + operation.slice(1)} ${header}: ${displayCell(value)}.`,
      columns: ['Calculation', 'Column', 'Result', 'Valid rows'],
      rows: [[operation, header, value, nums.length]],
      meta: `${reportMeta} · uses report rows, not underlying individual cases` };
  }
  if (/\b(top|bottom|highest|lowest)\b/i.test(q)) {
    if (!nums.length) return { summary: `${header} is not numeric, so it cannot be ranked.`,
      meta: reportMeta };
    const descending = /\b(top|highest)\b/i.test(q);
    const limitMatch = q.match(/\b(?:top|bottom)\s+(\d+)\b/i);
    const limit = Math.min(limitMatch ? +limitMatch[1] : 5, CHAT_ROW_CAP);
    const label = table.headers.find(h => h !== header) || 'Row';
    const ranked = values.map(item => ({ ...item, numeric: toNum(item.value) })).filter(item => item.numeric != null)
      .sort((a, b) => descending ? b.numeric - a.numeric : a.numeric - b.numeric).slice(0, limit);
    return { summary: `${descending ? 'Top' : 'Bottom'} ${ranked.length} rows by ${header}.`,
      columns: [label, header], rows: ranked.map(item => [rawValue(item.row, label), item.numeric]),
      meta: reportMeta };
  }
  if (/\b(count|how many|number of)\b/i.test(q)) {
    return { summary: `${values.length} report rows have a value in ${header}.`,
      columns: ['Measure', 'Result'], rows: [[`Rows with ${header}`, values.length]],
      meta: `${reportMeta} · report rows, not individual cases` };
  }
  if (filtered.description && values.length) {
    const labelHeader = table.headers.find(candidate => /^(ib region|region|kpi)$/i.test(candidate));
    return {
      summary: `${header} for ${rawValue(values[0].row, labelHeader)}: ${displayCell(values[0].value)}.`,
      columns: [labelHeader, header], rows: values.map(item => [rawValue(item.row, labelHeader), item.value]),
      meta: `${reportMeta} · report value, not an individual-case calculation`
    };
  }
  return { summary: `I found ${header} in ${table.name}. Ask for its sum, average, minimum, maximum, distinct values, a top/bottom ranking, or name a region/KPI.`,
    meta: 'Report calculations use the uploaded summary rows; they cannot recreate absent case-level records.' };
}

function answerDataQuestion(rawQuestion) {
  const originalQuestion = String(rawQuestion || '').trim();
  if (!originalQuestion) return { summary: 'Type a question about the uploaded workbook.' };
  const table = activeAssistantTable();
  const spelling = correctSpelling(originalQuestion, table);
  const result = answerDataQuestionCore(spelling.corrected);
  if (result && spelling.corrections.length) {
    const note = 'Interpreted ' + spelling.corrections.map(c => `“${c.from}” as “${c.to}”`).join(', ') + '.';
    result.meta = result.meta ? `${result.meta} · ${note}` : note;
  }
  return result;
}

function answerDataQuestionCore(question) {
  const q = String(question || '').trim();
  const ql = q.toLowerCase();
  const table = activeAssistantTable();
  if (!q) return { summary: 'Type a question about the uploaded workbook.' };
  if (/\b(what|which|how)\b.*\b(can you|can i|do|calculate|help|supported)\b|\bhelp\b/i.test(q)) {
    const headers = table ? table.headers : S.caseHeaders;
    const source = table ? `${table.name} · ${table.rows.length.toLocaleString()} current ${table.type === 'report' ? 'report rows' : 'cases'}` :
      'No workbook loaded';
    return {
      summary: ASSISTANT_HELP,
      columns: ['Available column', 'Use'],
      rows: headers.map(header => [header, 'filter, group, calculate, rank, or list']),
      meta: `Source: ${source} · table export includes only displayed rows`
    };
  }
  if (table && table.type === 'report') return answerReportQuestion(q, table);
  if (!S.cases.length) {
    return { summary: 'Please upload an Excel workbook first. I answer only from the workbook parsed in this browser.',
      meta: 'No data is sent to any server.' };
  }

  if (/\b(list|show|what).*(available )?columns\b|\bcolumn names\b/i.test(q)) {
    return {
      summary: `The case sheet contains ${S.caseHeaders.length} available columns.`,
      columns: ['#', 'Column'],
      rows: S.caseHeaders.map((h, i) => [i + 1, h]),
      meta: `Scope: workbook schema · ${S.cases.length.toLocaleString()} valid cases loaded`
    };
  }

  const lookup = findCaseLookup(q);
  if (lookup) {
    if (!lookup.row) {
      return {
        summary: `No case matches ${lookup.caseHeader} "${lookup.candidate}" in the uploaded workbook.`,
        meta: `Looked up against ${S.cases.length.toLocaleString()} loaded cases · Check the case number and try again`
      };
    }
    lastCaseLookup = { row: lookup.row, candidate: lookup.candidate, caseHeader: lookup.caseHeader };
    const remainder = q.replace(lookup.matchText, ' ');
    const field = extractRequestedField(remainder, lookup.caseHeader);
    if (field) {
      return {
        summary: `${field} for case ${lookup.candidate}: ${displayCell(rawValue(lookup.row, field))}.`,
        columns: [lookup.caseHeader, field],
        rows: [[lookup.candidate, rawValue(lookup.row, field)]],
        meta: `Looked up 1 case by ${lookup.caseHeader} · say "show all fields for case ${lookup.candidate}" for the full record`
      };
    }
    return {
      summary: `Full record for case ${lookup.candidate} (${S.caseHeaders.length} fields).`,
      columns: ['Field', 'Value'],
      rows: S.caseHeaders.map(h => [h, rawValue(lookup.row, h)]),
      meta: `Looked up 1 case by ${lookup.caseHeader} · ask for a specific field, e.g. "field ${S.caseHeaders[0]} for case ${lookup.candidate}"`
    };
  }

  /* A bare field-name follow-up with no "case" keyword at all (e.g. "field
     remarks", "remote remark", "open field remarks") continues the last
     single-case lookup in this conversation, so the user does not have to
     repeat the case number for every follow-up question about it. */

  /* Percentage/share questions (e.g. "What percentage of cases have RDF 1?" or
     "What % of Boston cases have Visits > 3?"). The denominator is the base
     scope (any region/zone/modality/market/date filter) before the specific
     value condition is applied; the numerator adds that condition back. */
  if (/\b(percent|percentage|%|share|proportion)\b/i.test(q) && !/\bpercentage\s+point/i.test(q)) {
    const full = queryFilters(q);
    const base = queryFilters(q, { skipValueConditions: true });
    if (!base.rows.length) {
      return { summary: 'No cases match the requested scope, so a percentage cannot be calculated.',
        meta: filterMeta(base, 0) };
    }
    const pct = (full.rows.length / base.rows.length) * 100;
    const conditionDesc = full.descriptions.filter(d => !base.descriptions.includes(d)).join('; ') || 'requested condition';
    return {
      summary: `${pct.toFixed(1)}% of the ${base.rows.length.toLocaleString()} case${base.rows.length === 1 ? '' : 's'} in scope meet ${conditionDesc} (${full.rows.length.toLocaleString()} of ${base.rows.length.toLocaleString()}).`,
      columns: ['Measure', 'Result'],
      rows: [['Matching cases', full.rows.length], ['Base cases', base.rows.length], ['Percentage', `${pct.toFixed(1)}%`]],
      meta: `Scope: ${base.descriptions.length ? base.descriptions.join('; ') : 'Full uploaded dataset'} · Condition: ${conditionDesc}`
    };
  }

  const compare = /\bcompare\b/i.test(q);
  const mentionedRegions = S.caseColumnMap.region
    ? uniq(S.cases.map(r => r.region)).filter(v => ql.includes(String(v).toLowerCase()))
    : [];
  if (compare && mentionedRegions.length >= 2) {
    const columns = ['KPI', ...mentionedRegions];
    const rows = KPIS.map(k => [
      `${k.label}${k.secondary ? ' (secondary)' : ''}`,
      ...mentionedRegions.map(region => {
        const rs = S.cases.filter(r => r.region.toLowerCase() === String(region).toLowerCase());
        return formatKpiValue(k, aggregate(rs)[k.key]);
      })
    ]);
    const counts = mentionedRegions.map(region =>
      `${region}: ${S.cases.filter(r => r.region.toLowerCase() === String(region).toLowerCase()).length}`).join(', ');
    return {
      summary: `Compared ${mentionedRegions.join(' and ')} across all dashboard KPIs using the exact dashboard formulas.`,
      columns, rows,
      meta: `Scope: full uploaded dataset, filtered by region · Cases: ${counts} · Fixed benchmarks are unchanged`
    };
  }
  if (compare) {
    return { summary: 'Please name at least two loaded regions to compare, for example: “Compare Chicago and Tampa for all KPIs”.',
      meta: 'Supported comparison dimension: IB Region' };
  }

  /* A value-distribution question (for example, "How many cases with RDF 0 and 1?")
     must be handled before KPI matching, because RDF is also the name of a KPI input. */
  const distributionValueClause = q.replace(/^.*?\b(?:with|where|have|has)\b/i, '');
  const distributionHeader = findColumn(distributionValueClause);
  const asksCaseCount = (/\b(count|how many|how much|number of)\b/i.test(q) && /\bcases?\b/i.test(q)) ||
    /\bcases\s+(?:with|where|have|has)\b/i.test(q);
  const asksValues = /\b(?:with|where|have|has)\b/i.test(q) &&
    /(?:\b\d+(?:\.\d+)?\b|\b(?:yes|no|true|false)\b)/i.test(q);
  const hasNumericComparison = /(?:>=|<=|>|<|=)\s*-?\d+(?:\.\d+)?/.test(q) || /\bbetween\b/i.test(q);
  if (asksCaseCount && asksValues && !hasNumericComparison && distributionHeader) {
    const filters = queryFilters(q);
    if (filters.inlineEqualityCount > 1) {
      return {
        summary: `${filters.rows.length.toLocaleString()} cases match the combined conditions.`,
        columns: ['Measure', 'Result'], rows: [['Case count', filters.rows.length]],
        meta: filterMeta(filters, filters.rows.length)
      };
    }
    const requestedValues = q.match(/-?\d+(?:\.\d+)?|\b(?:yes|no|true|false)\b/gi) || [];
    const groups = groupedRows(filters.rows, distributionHeader)
      .map(([name, rs]) => [name, rs.length])
      .filter(([name]) => !requestedValues.length || requestedValues.some(value =>
        String(name).toLowerCase() === value.toLowerCase()))
      .sort((a, b) => String(a[0]).localeCompare(String(b[0]), undefined, { numeric: true }));
    if (!groups.length) {
      return {
        summary: `No cases have the requested ${distributionHeader} value${requestedValues.length > 1 ? 's' : ''}.`,
        meta: filterMeta(filters, filters.rows.length)
      };
    }
    const details = groups.map(([value, count]) => `${value}: ${count}`).join('; ');
    return {
      summary: `Cases by ${distributionHeader} — ${details}.`,
      columns: [distributionHeader, 'Cases'], rows: groups,
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  const kpi = findKpi(q);
  if (kpi) {
    const groupHeader = /\bregion\b/i.test(q) ? S.caseColumnMap.region
      : /\bzone\b/i.test(q) ? S.caseColumnMap.zone
      : /\bmodality\b/i.test(q) ? S.caseColumnMap.modality
      : /\bmarket\b/i.test(q) ? S.caseColumnMap.market : null;
    const filters = queryFilters(q, { skipRegions: !!groupHeader });
    if (groupHeader && /\b(highest|lowest|top|bottom|best|worst)\b/i.test(q)) {
      const stats = groupedRows(filters.rows, groupHeader).map(([name, rs]) => {
        const a = aggregate(rs);
        return { name, value: a[kpi.key], n: a[kpi.key + '_n'] };
      }).filter(x => x.value != null);
      const high = /\b(highest|top)\b/i.test(q) ||
        (/\bbest\b/i.test(q) && kpi.dir === 'up') || (/\bworst\b/i.test(q) && kpi.dir === 'down');
      stats.sort((a, b) => high ? b.value - a.value : a.value - b.value);
      const nMatch = q.match(/\b(?:top|bottom)\s+(\d+)\b/i);
      const limit = Math.min(nMatch ? +nMatch[1] : 1, CHAT_ROW_CAP);
      const selected = stats.slice(0, limit);
      return {
        summary: selected.length
          ? `${selected[0].name} has the ${high ? 'highest' : 'lowest'} ${kpi.label}: ${formatKpiValue(kpi, selected[0].value)}.`
          : `${kpi.label} is not calculable for the requested groups.`,
        columns: [groupHeader, kpi.label, 'KPI sample'],
        rows: selected.map(x => [x.name, formatKpiValue(kpi, x.value), x.n]),
        meta: `${filterMeta(filters, filters.rows.length)} · Formula: ${kpi.calc} · Benchmark: ${fmtOf(kpi)(kpi.bench)}`
      };
    }
    const a = aggregate(filters.rows);
    return {
      summary: `${kpi.label}: ${formatKpiValue(kpi, a[kpi.key])}.`,
      columns: ['Calculation', 'Value'],
      rows: [
        ['Current value', formatKpiValue(kpi, a[kpi.key])],
        ['Fixed NAM Top-3 benchmark', fmtOf(kpi)(kpi.bench)],
        ['Direction', kpi.dir === 'up' ? 'Higher is better' : 'Lower is better']
      ],
      meta: `${filterMeta(filters, a[kpi.key + '_n'] || 0)} · Formula: ${kpi.calc}`
    };
  }

  const byMatch = q.match(/\bby\s+(.+?)(?:\?|$)/i);
  const groupHeaders = byMatch ? findColumns(byMatch[1]).slice(0, 3) : [];
  const groupHeader = groupHeaders[0] || null;
  if (/\b(count|how many|number of)\b/i.test(q) && groupHeader) {
    const filters = queryFilters(q);
    const groups = groupedRowsBy(filters.rows, groupHeaders)
      .map(group => [...group.values, group.rows.length])
      .sort((a, b) => b[groupHeaders.length] - a[groupHeaders.length]);
    return {
      summary: `${filters.rows.length.toLocaleString()} cases grouped by ${groupHeaders.join(' and ')}.`,
      columns: [...groupHeaders, 'Cases'], rows: groups,
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  if (/\b(distinct|unique)\b|\bdifferent\s+types?\s+of\b|\btypes?\s+of\b|\bhow many different\b/i.test(q)) {
    const header = findColumn(q);
    if (!header) return { summary: 'Which column should I list distinct values for?', meta: 'Example: “List distinct Modality values”' };
    const filters = queryFilters(q);
    const values = groupedRows(filters.rows, header).map(([name, rs]) => [name, rs.length]).sort((a, b) => b[1] - a[1]);
    return {
      summary: `${values.length} distinct ${header} values.`,
      columns: [header, 'Cases'], rows: values,
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  const operation = /\baverage|avg|mean\b/i.test(q) ? 'average'
    : /\bmedian\b/i.test(q) ? 'median'
    : /\bsum|total\b/i.test(q) ? 'sum'
    : /\bminimum|min\b/i.test(q) ? 'minimum'
    : /\bmaximum|max\b/i.test(q) ? 'maximum' : null;
  if (operation && !/\b(total|count|number of)\s+cases?\b/i.test(q)) {
    const operationGroups = groupHeaders;
    const operationGroup = operationGroups[0] || null;
    const header = findColumn(q, { exclude: operationGroups });
    if (!header) return { summary: `Which numeric column should I ${operation}?`, meta: 'Example: “Average Onsite Hours by Region”' };
    const filters = queryFilters(q, { skipRegions: operationGroups.includes(S.caseColumnMap.region) });
    const calculate = rs => {
      const values = rs.map(r => toNum(rawValue(r, header))).filter(v => v != null);
      if (!values.length) return { value: null, n: 0 };
      const value = operation === 'sum' ? values.reduce((a, b) => a + b, 0)
        : operation === 'average' ? values.reduce((a, b) => a + b, 0) / values.length
        : operation === 'median' ? median(values)
        : operation === 'minimum' ? Math.min(...values) : Math.max(...values);
      return { value, n: values.length };
    };
    if (operationGroup && operationGroup !== header) {
      const results = groupedRowsBy(filters.rows, operationGroups).map(group => {
        const calc = calculate(group.rows);
        return [...group.values, calc.value, calc.n];
      }).filter(row => row[operationGroups.length] != null).sort((a, b) => b[operationGroups.length] - a[operationGroups.length]);
      return {
        summary: `${operation[0].toUpperCase() + operation.slice(1)} ${header}, grouped by ${operationGroups.join(' and ')}.`,
        columns: [...operationGroups, `${operation} ${header}`, 'Valid values'], rows: results,
        meta: filterMeta(filters, filters.rows.length)
      };
    }
    const values = filters.rows.map(r => toNum(rawValue(r, header))).filter(v => v != null);
    if (!values.length) return { summary: `${header} has no numeric values in the requested scope.`, meta: filterMeta(filters, filters.rows.length) };
    const value = operation === 'sum' ? values.reduce((a, b) => a + b, 0)
      : operation === 'average' ? values.reduce((a, b) => a + b, 0) / values.length
      : operation === 'median' ? median(values)
      : operation === 'minimum' ? Math.min(...values) : Math.max(...values);
    return {
      summary: `${operation[0].toUpperCase() + operation.slice(1)} ${header}: ${displayCell(value)}.`,
      columns: ['Calculation', 'Column', 'Result', 'Valid values'],
      rows: [[operation, header, value, values.length]],
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  if (/\b(show|list)\b.*\bcases?\b/i.test(q)) {
    const filters = queryFilters(q);
    const mentioned = findColumn(q);
    const defaults = [
      S.caseColumnMap.caseNo, S.caseColumnMap.region, S.caseColumnMap.zone,
      S.caseColumnMap.modality, mentioned,
      S.dateCols.find(h => /creat/i.test(h)) || S.dateCols[0]
    ].filter(Boolean);
    const columns = [...new Set(defaults)].slice(0, 6);
    const shown = filters.rows.slice(0, CHAT_ROW_CAP);
    return {
      summary: `${filters.rows.length.toLocaleString()} matching cases${filters.rows.length > shown.length ? `; showing the first ${shown.length}` : ''}.`,
      columns, rows: shown.map(r => columns.map(h => rawValue(r, h))),
      meta: `${filterMeta(filters, filters.rows.length)} · Row display cap: ${CHAT_ROW_CAP}`
    };
  }

  if (/\b(count|how many|number of|total)\b.*\bcases?\b|\bcases?\s+(count|created)\b/i.test(q)) {
    const filters = queryFilters(q);
    return {
      summary: `${filters.rows.length.toLocaleString()} cases match the request.`,
      columns: ['Measure', 'Result'], rows: [['Case count', filters.rows.length]],
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  if (/\b(top|bottom|highest|lowest)\b/i.test(q)) {
    const header = findColumn(q);
    if (!header) return { summary: 'Please name the numeric column to rank.', meta: 'Example: “Top 10 cases by Visits”' };
    const filters = queryFilters(q);
    const descending = /\b(top|highest)\b/i.test(q);
    const nMatch = q.match(/\b(?:top|bottom)\s+(\d+)\b/i);
    const limit = Math.min(nMatch ? +nMatch[1] : 5, CHAT_ROW_CAP);
    const ranked = filters.rows.map(r => ({ r, v: toNum(rawValue(r, header)) })).filter(x => x.v != null)
      .sort((a, b) => descending ? b.v - a.v : a.v - b.v).slice(0, limit);
    const caseHeader = S.caseColumnMap.caseNo;
    return {
      summary: `${descending ? 'Top' : 'Bottom'} ${ranked.length} cases by ${header}.`,
      columns: [caseHeader || 'Case', header, S.caseColumnMap.region || 'Region'],
      rows: ranked.map(x => [caseHeader ? rawValue(x.r, caseHeader) : x.r.caseNo, x.v, x.r.region]),
      meta: filterMeta(filters, filters.rows.length)
    };
  }

  /* A bare field-name follow-up with no "case" keyword at all (e.g. "field
     remarks", "remote remark", "open field remarks") continues the last
     single-case lookup in this conversation, so the user does not have to
     repeat the case number for every follow-up question about it. Placed as
     a last resort (only reached once no other supported pattern matched) and
     skipped for anything mentioning the plural "cases", so dataset-wide
     questions are never misread as a single-case follow-up. */
  if (lastCaseLookup && !/\bcases\b/i.test(q)) {
    const followUpField = extractRequestedField(q, lastCaseLookup.caseHeader);
    if (followUpField) {
      return {
        summary: `${followUpField} for case ${lastCaseLookup.candidate}: ${displayCell(rawValue(lastCaseLookup.row, followUpField))}.`,
        columns: [lastCaseLookup.caseHeader, followUpField],
        rows: [[lastCaseLookup.candidate, rawValue(lastCaseLookup.row, followUpField)]],
        meta: `Continued from the last looked-up case (${lastCaseLookup.candidate}) · say "field X for case <number>" to switch cases`
      };
    }
  }

  return {
    summary: 'I could not determine a supported calculation without guessing. Please clarify the measure, grouping, or filter.',
    meta: 'Try: “Count cases by Modality”, “Average Visits for Boston”, “Show cases where Visits > 3”, or “List available columns”.'
  };
}

function submitChatQuestion(question) {
  const q = String(question || '').trim();
  if (!q) return;
  addChatMessage('user', q);
  pushChatHistory(q);
  $('chatInput').value = '';
  $('chatTyping').classList.remove('hidden');
  $('chatSend').disabled = true;
  setTimeout(() => {
    const result = answerDataQuestion(q);
    $('chatTyping').classList.add('hidden');
    $('chatSend').disabled = false;
    addChatMessage('assistant', result.summary, result);
    $('chatInput').focus();
  }, 80);
}

/* Command-line-style question history: Up/Down arrow keys recall previously
   asked questions, mirroring shell/terminal history navigation. Only active
   while the input is a single line, so Shift+Enter multi-line composition and
   normal in-text caret movement are never hijacked. */
let chatHistoryList = [];
let chatHistoryIndex = -1;
let chatHistoryDraft = '';

function pushChatHistory(question) {
  if (chatHistoryList[chatHistoryList.length - 1] !== question) chatHistoryList.push(question);
  chatHistoryIndex = -1;
  chatHistoryDraft = '';
}

function resetChatHistory() {
  chatHistoryList = [];
  chatHistoryIndex = -1;
  chatHistoryDraft = '';
}

function recallChatHistory(direction) {
  const input = $('chatInput');
  if (!chatHistoryList.length) return;
  if (direction === 'up') {
    if (chatHistoryIndex === -1) chatHistoryDraft = input.value;
    if (chatHistoryIndex < chatHistoryList.length - 1) chatHistoryIndex++;
    input.value = chatHistoryList[chatHistoryList.length - 1 - chatHistoryIndex];
  } else {
    if (chatHistoryIndex <= 0) {
      chatHistoryIndex = -1;
      input.value = chatHistoryDraft;
    } else {
      chatHistoryIndex--;
      input.value = chatHistoryList[chatHistoryList.length - 1 - chatHistoryIndex];
    }
  }
  const end = input.value.length;
  input.setSelectionRange(end, end);
}

function initDataAssistant() {
  const panel = $('chatPanel');
  const launcher = $('chatLauncher');
  launcher.addEventListener('click', () => {
    panel.classList.remove('hidden');
    launcher.classList.add('hidden');
    launcher.setAttribute('aria-expanded', 'true');
    if (!$('chatMessages').children.length) resetAssistantConversation(true);
    updateAssistantScope();
    $('chatInput').focus();
  });
  $('chatClose').addEventListener('click', () => {
    panel.classList.add('hidden');
    launcher.classList.remove('hidden');
    launcher.setAttribute('aria-expanded', 'false');
    launcher.focus();
  });
  $('chatClear').addEventListener('click', () => {
    resetAssistantConversation(true);
    resetChatHistory();
    $('chatInput').focus();
  });
  $('chatForm').addEventListener('submit', e => {
    e.preventDefault();
    submitChatQuestion($('chatInput').value);
  });
  $('chatInput').addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      $('chatForm').requestSubmit();
      return;
    }
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      if (e.target.value.includes('\n')) return; /* let multi-line editing use normal caret movement */
      e.preventDefault();
      recallChatHistory(e.key === 'ArrowUp' ? 'up' : 'down');
    }
  });
  $('chatStarters').addEventListener('click', e => {
    const button = e.target.closest('button');
    if (button) submitChatQuestion(button.textContent);
  });
  $('chatDataset').addEventListener('change', () => {
    updateAssistantScope();
    addChatMessage('assistant', `Switched to ${activeAssistantTable().name}.`);
  });
  resetAssistantConversation(true);
  updateAssistantScope();
}

initDataAssistant();

/* Deliberately small test surface for the local, dependency-free validation harness. */
globalThis.NAMDashboard = {
  state: S, buildFromWorkbook, buildReportFromWorkbook, parseCases, aggregate, answerDataQuestion, exportChatResult, KPIS,
  submitChatQuestion, recallChatHistory, resetChatHistory
};
