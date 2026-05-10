// File rename logic

function buildNewName(fundNum, investorNum, dateStr, docType, investorName) {
  const fund = String(parseInt(fundNum, 10)).padStart(3, '0');
  const inv  = String(parseInt(investorNum, 10)).padStart(3, '0');
  const name = investorName.replace(/\s+/g, '');
  return `${fund}-${inv}-${dateStr}-${docType}-${name}.PDF`;
}

/**
 * @param {Array} matches  — from buildMatches(), possibly user-edited
 * @param {string} fundNum
 * @param {string} dateStr
 * @param {string} docType
 * @returns {Promise<{count: number, errors: string[]}>}
 */
async function renameFiles(matches, fundNum, dateStr, docType) {
  let count = 0;
  const errors = [];

  for (const m of matches) {
    if (!m.investorNumber || !m.handle) continue;

    const newName = buildNewName(fundNum, m.investorNumber, dateStr, docType, m.investorName);

    try {
      await m.handle.move(newName);
      m.pdf = newName; // update in-place so the grid can refresh
      count++;
    } catch (err) {
      errors.push(`${m.pdf} → ${newName}: ${err.message}`);
    }
  }

  return { count, errors };
}
