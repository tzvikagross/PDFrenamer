// Fuzzy matching: PDF filenames → Excel investor records (one-to-one)

const MATCH_THRESHOLD = 0.55; // Fuse score: 0=perfect, 1=no match

// ── Name extraction ───────────────────────────────────────────────────────────

/**
 * Find the longest common prefix shared by all PDF filenames (without extension).
 * Trims to the last whitespace so we never cut a word in half.
 */
function computeCommonPrefix(filenames) {
  const stems = filenames.map(f => f.replace(/\.pdf$/i, ''));
  if (!stems.length) return '';
  let prefix = stems[0];
  for (const s of stems.slice(1)) {
    let i = 0;
    while (i < prefix.length && i < s.length && prefix[i] === s[i]) i++;
    prefix = prefix.slice(0, i);
  }
  // Trim back to the last space so we land on a clean word boundary
  const cut = prefix.search(/\s\S*$/);
  return cut > 0 ? prefix.slice(0, cut) : prefix;
}

/**
 * Given a PDF filename and the pre-computed common prefix, return just the
 * investor-name portion, normalised for comparison:
 *   1. Strip extension + common prefix + leading serial number
 *   2. Strip a leading "&" (e.g. "& TRACY VAN BUSKIRK, PETER")
 *   3. Flip "LAST, FIRST" → "FIRST LAST"
 *   4. Lowercase and collapse whitespace
 */
function extractInvestorName(filename, commonPrefix) {
  let s = filename.replace(/\.pdf$/i, '').trim();

  // Remove common prefix (case-insensitive)
  if (commonPrefix) {
    const pfx = commonPrefix.toLowerCase();
    if (s.toLowerCase().startsWith(pfx)) {
      s = s.slice(commonPrefix.length).trim();
    }
  }

  // Remove leading serial number  "1 ", "10 ", "30 " …
  s = s.replace(/^\d+\s+/, '').trim();

  // Remove leading "&" or whitespace artifacts
  s = s.replace(/^[&\s]+/, '').trim();

  // "LAST, FIRST [MIDDLE]"  →  "FIRST [MIDDLE] LAST"
  const commaIdx = s.indexOf(',');
  if (commaIdx > 0) {
    const last  = s.slice(0, commaIdx).trim();
    const first = s.slice(commaIdx + 1).trim();
    s = `${first} ${last}`;
  }

  // Normalize punctuation, underscores, extra spaces; lowercase
  return s
    .replace(/[_.()\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Normalize an Excel investor name the same way (no prefix to strip,
 * but we still want consistent casing / punctuation).
 */
function normalizeInvestorName(name) {
  return name
    .replace(/[_.()\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// ── Token overlap score ───────────────────────────────────────────────────────
// Supplement Fuse.js with an explicit token-overlap metric so that partially
// shared company names (e.g. "clal insurance company ltd" / "clal insurance")
// score well even when the lengths differ a lot.

function tokenOverlap(a, b) {
  const tokA = a.split(/\s+/).filter(t => t.length > 1);
  const tokB = b.split(/\s+/).filter(t => t.length > 1);
  if (!tokA.length || !tokB.length) return 0;

  // Use the SHORTER set as denominator (containment / recall).
  // This means "clal insurance" scores 1.0 against "clal insurance company ltd"
  // because 100% of the shorter name's tokens appear in the longer name.
  const [shorter, longer] = tokA.length <= tokB.length ? [tokA, tokB] : [tokB, tokA];

  let matched = 0;
  const usedLonger = new Set();
  for (const ts of shorter) {
    for (let j = 0; j < longer.length; j++) {
      if (usedLonger.has(j)) continue;
      const tl = longer[j];
      const maxDist = Math.min(2, Math.floor(Math.max(ts.length, tl.length) / 4));
      if (ts === tl || editDistance(ts, tl) <= maxDist) {
        matched++;
        usedLonger.add(j);
        break;
      }
    }
  }
  return matched / shorter.length; // containment: 0..1, higher = better
}

function editDistance(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

// ── Main matching function ────────────────────────────────────────────────────

/**
 * @param {Array<{name: string, handle: FileSystemFileHandle}>} pdfEntries
 * @param {Array<{number: string, name: string}>} investors
 * @returns {Array<{pdf, handle, investorNumber, investorName, fuseIndex}>}
 */
function buildMatches(pdfEntries, investors) {
  const empty = pdfEntries.map(e => ({
    pdf: e.name, handle: e.handle,
    investorNumber: null, investorName: null, fuseIndex: null
  }));
  if (!pdfEntries.length || !investors.length) return empty;

  // Step 1: extract normalised investor-name strings from PDF filenames
  const commonPrefix = computeCommonPrefix(pdfEntries.map(e => e.name));
  const pdfKeys = pdfEntries.map(e => extractInvestorName(e.name, commonPrefix));

  // Step 2: normalise Excel names
  const invKeys = investors.map(inv => normalizeInvestorName(inv.name));

  // Step 3: Fuse.js index over normalised Excel names
  const fuseData  = invKeys.map((k, i) => ({ key: k, origIdx: i }));
  const fuse = new Fuse(fuseData, {
    keys: ['key'],
    threshold: 1.0,       // accept everything; we score manually below
    includeScore: true,
    ignoreLocation: true,
    distance: 1000,
    minMatchCharLength: 2,
  });

  // Step 4: build (combinedScore, pdfIdx, invIdx) triples
  // combinedScore: lower = better  (maps to Fuse's convention)
  const triples = [];

  pdfKeys.forEach((query, pi) => {
    const fuseResults = fuse.search(query);

    // Also compute token-overlap scores for ALL investors (Fuse may miss some)
    invKeys.forEach((invKey, ii) => {
      const overlap   = tokenOverlap(query, invKey);  // 0..1, higher = better
      const fuseScore = fuseResults.find(r => r.item.origIdx === ii)?.score ?? 1.0;

      // Convert overlap to a "cost" (like Fuse score), then blend
      const overlapCost = 1 - overlap;                    // lower = better
      const combined    = Math.min(fuseScore, overlapCost); // take the better signal

      triples.push({ score: combined, pi, ii });
    });
  });

  // Step 5: sort ascending (best matches first) and do greedy one-to-one assignment
  triples.sort((a, b) => a.score - b.score);

  const assignedPdf = new Set();
  const assignedInv = new Set();
  const results     = empty;

  for (const { score, pi, ii } of triples) {
    if (score > MATCH_THRESHOLD) break;
    if (assignedPdf.has(pi) || assignedInv.has(ii)) continue;
    results[pi].investorNumber = investors[ii].number;
    results[pi].investorName   = investors[ii].name;
    results[pi].fuseIndex      = ii;
    assignedPdf.add(pi);
    assignedInv.add(ii);
  }

  return results;
}
