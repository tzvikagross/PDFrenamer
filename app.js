// PDF Renamer — main application logic

'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let folderHandle  = null;
let excelHandle   = null;
let investors     = [];   // [{number: "003", name: "John Dow"}, ...]
let matches       = [];   // [{pdf, handle, investorNumber, investorName}, ...]

// ── DOM refs ─────────────────────────────────────────────────────────────────
const folderPathEl  = document.getElementById('folder-path');
const excelPathEl   = document.getElementById('excel-path');
const fundInput     = document.getElementById('fund-num');
const dateInput     = document.getElementById('date-str');
const docTypeSelect = document.getElementById('doc-type');
const tbody         = document.getElementById('grid-body');
const statusBar     = document.getElementById('status-bar');
const renameBtn     = document.getElementById('rename-btn');
const rowCountEl    = document.getElementById('row-count');

// ── Browser capability check ─────────────────────────────────────────────────
(function checkBrowser() {
  if (!window.showDirectoryPicker || !window.showOpenFilePicker) {
    document.getElementById('browser-notice').classList.add('visible');
  }
})();

// ── Folder picker ─────────────────────────────────────────────────────────────
document.getElementById('btn-folder').addEventListener('click', async () => {
  try {
    folderHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    folderPathEl.textContent = folderHandle.name;
    setStatus('Folder selected: ' + folderHandle.name);
    await tryRunMatching();
  } catch (e) {
    if (e.name !== 'AbortError') setStatus('Error selecting folder: ' + e.message, 'error');
  }
});

// ── Excel picker ──────────────────────────────────────────────────────────────
document.getElementById('btn-excel').addEventListener('click', async () => {
  try {
    [excelHandle] = await window.showOpenFilePicker({
      types: [{
        description: 'Excel files',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel': ['.xls'],
        }
      }]
    });
    excelPathEl.textContent = excelHandle.name;
    await loadInvestors();
    await tryRunMatching();
  } catch (e) {
    if (e.name !== 'AbortError') setStatus('Error selecting Excel file: ' + e.message, 'error');
  }
});

// ── Load investors from Excel ─────────────────────────────────────────────────
async function loadInvestors() {
  try {
    const file = await excelHandle.getFile();
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    investors = [];
    for (const row of rows) {
      const rawNum  = String(row[0] ?? '').trim();
      const rawName = String(row[1] ?? '').trim();
      if (!rawNum || !rawName) continue;
      const num = parseInt(rawNum, 10);
      if (isNaN(num)) continue;
      investors.push({ number: String(num).padStart(3, '0'), name: rawName });
    }

    setStatus(`Excel loaded: ${investors.length} investor(s) found.`);
  } catch (e) {
    setStatus('Error reading Excel: ' + e.message, 'error');
    investors = [];
  }
}

// ── Run matching + render ─────────────────────────────────────────────────────
async function tryRunMatching() {
  if (!folderHandle || !investors.length) return;

  setStatus('Scanning folder…');
  const pdfEntries = await collectPdfEntries(folderHandle);

  if (!pdfEntries.length) {
    setStatus('No PDF files found in the selected folder.', 'error');
    renderGrid([]);
    return;
  }

  setStatus('Matching investor names…');
  matches = buildMatches(pdfEntries, investors);
  renderGrid(matches);
  refreshStatusBar();
}

async function collectPdfEntries(dirHandle) {
  const entries = [];
  for await (const [, handle] of dirHandle.entries()) {
    if (handle.kind === 'file' && /\.pdf$/i.test(handle.name)) {
      entries.push({ name: handle.name, handle });
    }
  }
  entries.sort((a, b) => a.name.localeCompare(b.name));
  return entries;
}

// ── Render grid ───────────────────────────────────────────────────────────────
function renderGrid(rows) {
  tbody.innerHTML = '';

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="2" class="empty-state">
      No PDF files loaded yet. Select a folder and an Excel file above.
    </td></tr>`;
    rowCountEl.textContent = '';
    renameBtn.disabled = true;
    return;
  }

  const optionsHtml = buildInvestorOptions();

  rows.forEach((m, idx) => {
    const tr = document.createElement('tr');
    if (!m.investorNumber) tr.classList.add('unmatched');

    const selectedVal = m.investorNumber ? `${m.investorNumber}|${m.investorName}` : '';

    tr.innerHTML = `
      <td class="col-pdf" title="${escHtml(m.pdf)}">${escHtml(m.pdf)}</td>
      <td class="col-match">
        <select data-idx="${idx}" aria-label="Select investor for ${escHtml(m.pdf)}">
          <option value="">— unmatched —</option>
          ${optionsHtml}
        </select>
      </td>`;

    const sel = tr.querySelector('select');
    sel.value = selectedVal;
    sel.addEventListener('change', () => onInvestorChange(sel, tr, idx));

    tbody.appendChild(tr);
  });

  renameBtn.disabled = false;
  checkDuplicates();
  refreshStatusBar();
}

function buildInvestorOptions() {
  return investors.map(inv => {
    const val   = `${inv.number}|${inv.name}`;
    const label = `${inv.number} — ${inv.name}`;
    return `<option value="${escHtml(val)}">${escHtml(label)}</option>`;
  }).join('');
}

function onInvestorChange(sel, tr, idx) {
  const val = sel.value;
  if (!val) {
    matches[idx].investorNumber = null;
    matches[idx].investorName   = null;
    tr.classList.add('unmatched');
    tr.classList.remove('duplicate');
  } else {
    const [num, ...nameParts] = val.split('|');
    matches[idx].investorNumber = num;
    matches[idx].investorName   = nameParts.join('|');
    tr.classList.remove('unmatched');
  }
  checkDuplicates();
  refreshStatusBar();
}

// ── Duplicate detection ───────────────────────────────────────────────────────
/**
 * Scans all rows for investor numbers assigned to more than one PDF.
 * Marks those rows with the "duplicate" class.
 * Returns true if any duplicates exist.
 */
function checkDuplicates() {
  const rows = Array.from(tbody.querySelectorAll('tr[data-idx], tr:not(.empty-state)'));

  // Build investorNumber → [tr, ...] map
  const numToRows = new Map();
  matches.forEach((m, idx) => {
    if (!m.investorNumber) return;
    const tr = tbody.querySelector(`select[data-idx="${idx}"]`)?.closest('tr');
    if (!tr) return;
    if (!numToRows.has(m.investorNumber)) numToRows.set(m.investorNumber, []);
    numToRows.get(m.investorNumber).push(tr);
  });

  // Clear previous duplicate marks
  tbody.querySelectorAll('tr.duplicate').forEach(tr => tr.classList.remove('duplicate'));

  let hasDuplicates = false;
  for (const [, rowList] of numToRows) {
    if (rowList.length > 1) {
      rowList.forEach(tr => tr.classList.add('duplicate'));
      hasDuplicates = true;
    }
  }

  return hasDuplicates;
}

// ── Status bar + row count ────────────────────────────────────────────────────
function refreshStatusBar() {
  if (!matches.length) return;

  const total      = matches.length;
  const matched    = matches.filter(m => m.investorNumber).length;
  const unmatched  = total - matched;
  const hasDups    = checkDuplicates();

  // Row count summary
  let countText = `${total} file(s) · ${matched} matched · ${unmatched} unmatched`;
  if (hasDups) {
    const dupInvNums = getDuplicateInvestorNumbers();
    countText += ` · ⚠ ${dupInvNums.size} duplicate assignment(s)`;
  }
  rowCountEl.textContent = countText;

  // Status message
  if (hasDups) {
    setStatus(
      `Duplicate investor assignments detected (highlighted in amber). Please fix before renaming.`,
      'error'
    );
  } else if (unmatched > 0 && matched > 0) {
    setStatus(
      `${matched} file(s) will be renamed · ${unmatched} unmatched file(s) will be left unchanged.`
    );
  } else if (matched === 0) {
    setStatus('No files matched yet. Use the dropdowns to assign investors.', 'error');
  } else {
    setStatus(`All ${matched} file(s) matched and ready to rename.`, 'success');
  }
}

function getDuplicateInvestorNumbers() {
  const counts = new Map();
  matches.forEach(m => {
    if (!m.investorNumber) return;
    counts.set(m.investorNumber, (counts.get(m.investorNumber) || 0) + 1);
  });
  const dups = new Set();
  for (const [num, count] of counts) {
    if (count > 1) dups.add(num);
  }
  return dups;
}

// ── Rename button ─────────────────────────────────────────────────────────────
renameBtn.addEventListener('click', async () => {
  if (!validateInputs()) return;

  // Block on duplicates
  if (checkDuplicates()) {
    setStatus(
      'Cannot rename: duplicate investor assignments exist (rows highlighted in amber). Please resolve before continuing.',
      'error'
    );
    return;
  }

  const toRename = matches.filter(m => m.investorNumber);
  const skipped  = matches.length - toRename.length;

  if (!toRename.length) {
    setStatus('No matched files to rename.', 'error');
    return;
  }

  renameBtn.disabled = true;
  setStatus(`Renaming ${toRename.length} file(s)…${skipped ? ` (${skipped} unmatched will be skipped)` : ''}`);

  try {
    const { count, errors } = await renameFiles(matches, fundInput.value.trim(), dateInput.value.trim(), docTypeSelect.value);

    if (errors.length) {
      setStatus(`Renamed ${count} file(s). ${errors.length} error(s): ${errors.join('; ')}`, 'error');
    } else {
      const skipMsg = skipped ? ` · ${skipped} unmatched file(s) left unchanged.` : '';
      setStatus(`Successfully renamed ${count} file(s).${skipMsg}`, 'success');
    }

    renderGrid(matches);
    refreshStatusBar();
  } catch (e) {
    setStatus('Rename failed: ' + e.message, 'error');
  } finally {
    renameBtn.disabled = false;
  }
});

// ── Validation ────────────────────────────────────────────────────────────────
function validateInputs() {
  let ok = true;

  const fund = fundInput.value.trim();
  if (!fund || !/^\d+$/.test(fund)) {
    fundInput.classList.add('invalid');
    ok = false;
  } else {
    fundInput.classList.remove('invalid');
  }

  const date = dateInput.value.trim();
  if (!/^\d{8}$/.test(date)) {
    dateInput.classList.add('invalid');
    ok = false;
  } else {
    dateInput.classList.remove('invalid');
  }

  if (!ok) setStatus('Please fix the highlighted fields before renaming.', 'error');
  return ok;
}

fundInput.addEventListener('input', () => fundInput.classList.remove('invalid'));
dateInput.addEventListener('input', () => dateInput.classList.remove('invalid'));

// ── Helpers ───────────────────────────────────────────────────────────────────
function setStatus(msg, type = '') {
  statusBar.textContent = msg;
  statusBar.className   = type || '';
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
