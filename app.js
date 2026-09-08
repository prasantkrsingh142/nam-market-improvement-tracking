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
    bench: 0.73, secondary: true, nullLabel: 'Not comparable',
    def: 'NAM-only: parsed returned quantity / (parts ordered − parts consumed); only cases with a positive unused balance count.',
    calc: 'SUM(Returned Qty parsed from Parts Returned List) / SUM(# Parts Ordered − # Parts Consumed), Market = NAM, cases where that difference > 0',
    src: 'AP (parsed) / (AF − AH)', reqCols: ['partsReturned', 'partsOrdered', 'partsConsumed'],
    den: r => (S.kpiAvailable.partsReturn && norm(r.market) === 'nam' && (r.partsOrdered - r.partsConsumed) > 0)
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
  advisedList: 'Parts Advised List', consumedList: 'Parts Consumed List',
  partsReturnedList: 'Parts Returned List'
};
const fmtOf = k => v => v == null ? (k.nullLabel || '—') : (k.fmt === 'pct' ? fmtPct(v) : fmtNum(v));
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

/* Sum the quantity between "* " and "=" on every line of a free-text Parts Returned List cell
   (e.g. "Part A * 2 = 100€"); blank/"-" is 0, malformed lines contribute 0. */
function parseReturnedQtyList(v) {
  if (v == null || v === '' || String(v).trim() === '-') return 0;
  const lines = String(v).split(/\n+/);
  let sum = 0;
  for (const line of lines) {
    const m = line.match(/\*\s*([\d.,]+)\s*=/);
    if (!m) continue;
    const qty = toNum(m[1]);
    if (qty != null) sum += qty;
  }
  return sum;
}

/* --------------------------------------------------------------- app state */
const S = {
  fileName: '', cases: [], plan: [], planCols: [], planRaw: [],
  dateCols: [], sheetNames: [], hasCases: false,
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
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fi.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });
$('btnChange').addEventListener('click', () => {
  $('dashView').classList.add('hidden'); $('importView').classList.remove('hidden');
  $('btnChange').classList.add('hidden');
});

function status(html, cls) { $('importStatus').innerHTML = `<span class="${cls || ''}">${html}</span>`; }

function handleFile(file) {
  S.fileName = file.name;
  status('Reading <b>' + esc(file.name) + '</b> …');
  const fr = new FileReader();
  fr.onerror = () => status('Could not read the file.', 'err');
  fr.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      buildFromWorkbook(wb);
    } catch (err) {
      console.error(err);
      status('Failed to parse workbook: ' + esc(err.message), 'err');
    }
  };
  fr.readAsArrayBuffer(file);
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
  S.cases = []; S.dateCols = []; S.hasCases = false; S.colReport = []; S.missingReport = []; S.caseStats = null;
  S.partsAdvisedReport = null; S.partsReturnReport = null;
  if (casesName) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[casesName], { raw: true, defval: null });
    if (rows.length) {
      parseCases(rows, msgs);
      msgs.push(`Sheet <b>${esc(casesName)}</b>${autoDetected ? ' (auto-detected as the case-level sheet)' : ''}: ${S.cases.length} valid cases parsed.`);
    }
  }
  if (!S.hasCases) {
    msgs.push('<span class="err">No usable case-level sheet found — trends and region comparison are disabled; KPI benchmarks remain the fixed NAM targets.</span>');
  } else if (S.missingReport.length) {
    msgs.push('<span class="err">Not calculable from this workbook: ' +
      S.missingReport.map(m => `${esc(m.kpi)} (${esc(m.reason)})`).join('; ') + '.</span>');
  }

  status(msgs.join('<br>'), 'ok');
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
    partsReturnedList: resolveCol(H, CANON_HEADERS.partsReturnedList),
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
  /* KPIs whose source column(s) are entirely absent from this workbook must never be
     silently computed from a default-zero value - mark them unavailable so bench/num/den
     stay null instead of fabricating a 0% result. */
  const useAdvisedTextLists = !(C.advListed && C.advMatched) && !!(C.advisedList && C.consumedList);
  S.kpiAvailable = {
    partsAdvised: !!(C.advListed && C.advMatched) || useAdvisedTextLists,
    partsReturn: !!(C.partsReturnedList || C.partsReturned) && !!C.partsOrdered && !!C.partsConsumed
  };
  S.missingReport = [];
  if (!C.advListed || !C.advMatched) {
    if (!useAdvisedTextLists) {
      S.missingReport.push({
        kpi: 'Parts Advised Success rate',
        reason: 'Columns "Advised Part Numbers Listed" / "Advised Part Numbers Matched" (BM/BN) are not present, and no usable "Parts Advised List" / "Parts Consumed List" text columns were found either.'
      });
    }
  }
  if (!S.kpiAvailable.partsReturn) {
    S.missingReport.push({
      kpi: 'Parts return rate',
      reason: 'Column "Parts Returned List" (AP) (or the legacy "Returned Parts Qty (parsed)") together with "# Parts Ordered" (AF) and "# Parts Consumed" (AH) are required, and are not all present.'
    });
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
    const o = {
      caseNo,
      region: C.region ? String(r[C.region] ?? '').trim() : '',
      zone: C.zone ? String(r[C.zone] ?? '').trim() : '',
      modality: C.modality ? String(r[C.modality] ?? '').trim() : '',
      market: C.market ? String(r[C.market] ?? '').trim() : '',
      advListed: advListedVal, advMatched: advMatchedVal,
      visits: n0(r[C.visits]), onsite: n0(r[C.onsite]),
      partsOrdered: n0(r[C.partsOrdered]), partsConsumed: n0(r[C.partsConsumed]),
      partsReturned: C.partsReturnedList ? parseReturnedQtyList(r[C.partsReturnedList]) : n0(r[C.partsReturned]),
      spa: n0(r[C.spa]), rdf: n0(r[C.rdf]),
      frusAdvised: n0(r[C.frusAdvised]), orderedBefore: n0(r[C.orderedBefore]),
      d: {}
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

  if (S.kpiAvailable.partsReturn) buildPartsReturnReport(msgs);
}

/* Parts Return Rate computed from scratch, NAM-only, aggregated (never averaged) per IB Region
   plus an overall NAM total. Region with Unused Parts Quantity <= 0 is "Not comparable". */
function buildPartsReturnReport(msgs) {
  const namCases = S.cases.filter(r => norm(r.market) === 'nam');
  const byRegion = new Map();
  namCases.forEach(r => { if (!byRegion.has(r.region)) byRegion.set(r.region, []); byRegion.get(r.region).push(r); });

  const rowFor = rs => {
    const ordered = rs.reduce((s, r) => s + r.partsOrdered, 0);
    const consumed = rs.reduce((s, r) => s + r.partsConsumed, 0);
    const returned = rs.reduce((s, r) => s + r.partsReturned, 0);
    const unused = ordered - consumed;
    const rate = unused > EPS ? returned / unused : null;
    return { cases: rs.length, ordered, consumed, unused, returned, rate, flagged: rate != null && rate > 1 };
  };

  const regions = [...byRegion.entries()]
    .map(([region, rs]) => ({ region, ...rowFor(rs) }))
    .sort((a, b) => a.region.localeCompare(b.region));
  const overall = rowFor(namCases);

  S.partsReturnReport = { casesReviewed: namCases.length, regions, overall };
  const flaggedRegions = regions.filter(r => r.flagged).map(r => r.region);
  msgs.push('Parts return rate (NAM) computed from scratch using Parts Returned List (AP), ' +
    `# Parts Ordered (AF), # Parts Consumed (AH): ${namCases.length} NAM cases reviewed across ${regions.length} regions, ` +
    `overall rate ${overall.rate == null ? 'Not comparable' : (overall.rate * 100).toFixed(1) + '%'}.` +
    (flaggedRegions.length ? ` <span class="err">Data-quality flag (&gt;100%): ${flaggedRegions.map(esc).join(', ')}.</span>` : ''));
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
  if (gran === 'month') return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}`;
  if (gran === 'quarter') return `${date.getFullYear()}-Q${Math.floor(date.getMonth() / 3) + 1}`;
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));          // Monday of that week
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function periodLabel(key, gran) {
  if (gran === 'month') { const [y, m] = key.split('-'); return `${MON[+m - 1]} ${y}`; }
  if (gran === 'quarter') return key;
  const [y, m, d] = key.split('-'); return `w/c ${d} ${MON[+m - 1]} ${y.slice(2)}`;
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
  if ($('btnExportPartsReturn')) $('btnExportPartsReturn').addEventListener('click', exportPartsReturn);

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
    .map(([k, rs]) => ({ key: k, label: periodLabel(k, gran), rows: rs, ...aggregate(rs) }))
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
  renderPartsReturnTable();
}

function renderScope() {
  const { rows, periods, cur } = CTX;
  const gran = $('fGran').value;
  const parts = [];
  parts.push(`<b>${rows.length.toLocaleString()}</b> cases in scope`);
  parts.push(`<b>${periods.length}</b> ${gran === 'week' ? 'weeks' : gran === 'month' ? 'months' : 'quarters'} of history`);
  const reg = $('fRegion').value, zn = $('fZone').value, mo = $('fModality').value;
  parts.push('Region: <b>' + (reg || 'All') + '</b>');
  if (zn) parts.push('Zone: <b>' + zn + '</b>');
  if (mo) parts.push('Modality: <b>' + mo + '</b>');
  parts.push('Date basis: <b>' + ($('fDateCol').value || '—') + '</b>');
  $('scopeBar').innerHTML = parts.join(' &nbsp;&middot;&nbsp; ');
  $('latestPeriodLbl').textContent = cur ? cur.label + ' · ' + cur.n + ' cases' : 'no period data';
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
  const gname = gran === 'week' ? 'week' : gran === 'month' ? 'month' : 'quarter';
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

function renderPartsReturnTable() {
  const panel = $('partsReturnPanel');
  const rep = S.partsReturnReport;
  if (!panel) return;
  if (!rep || !rep.regions.length) { panel.classList.add('hidden'); return; }
  panel.classList.remove('hidden');
  const rateTxt = r => r.rate == null ? 'Not comparable' : (r.rate * 100).toFixed(1) + '%';
  let h = '<thead><tr><th>IB Region</th><th>Cases</th><th>Total Parts Ordered</th>' +
    '<th>Total Parts Consumed</th><th>Unused Parts Quantity</th><th>Total Returned Quantity</th><th>Parts Return Rate</th></tr></thead><tbody>';
  rep.regions.forEach(r => {
    h += `<tr><td><b>${esc(r.region)}</b></td><td class="num">${r.cases}</td>` +
      `<td class="num">${r.ordered.toLocaleString()}</td><td class="num">${r.consumed.toLocaleString()}</td>` +
      `<td class="num">${r.unused.toLocaleString()}</td><td class="num">${r.returned.toLocaleString()}</td>` +
      `<td class="num ${r.flagged ? 'r' : ''}" ${r.flagged ? 'title="Data-quality flag: rate exceeds 100%"' : ''}>${rateTxt(r)}</td></tr>`;
  });
  const o = rep.overall;
  h += `<tr class="bench"><td>Overall NAM</td><td class="num">${o.cases}</td>` +
    `<td class="num">${o.ordered.toLocaleString()}</td><td class="num">${o.consumed.toLocaleString()}</td>` +
    `<td class="num">${o.unused.toLocaleString()}</td><td class="num">${o.returned.toLocaleString()}</td>` +
    `<td class="num ${o.flagged ? 'r' : ''}">${rateTxt(o)}</td></tr></tbody>`;
  $('partsReturnTable').innerHTML = h;
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
function exportPartsReturn() {
  const rep = S.partsReturnReport;
  if (!rep) return;
  const rows = [['IB Region', 'Cases', 'Total Parts Ordered', 'Total Parts Consumed', 'Unused Parts Quantity',
    'Total Returned Quantity', 'Parts Return Rate']];
  const rateCell = r => r.rate == null ? 'Not comparable' : (r.rate * 100).toFixed(1) + '%';
  rep.regions.forEach(r => rows.push([r.region, r.cases, r.ordered, r.consumed, r.unused, r.returned, rateCell(r)]));
  rows.push(['Overall NAM', rep.overall.cases, rep.overall.ordered, rep.overall.consumed, rep.overall.unused,
    rep.overall.returned, rateCell(rep.overall)]);
  download('rdf_parts_return_nam.csv', rows.map(r => r.map(csvCell).join(',')).join('\n'));
}
