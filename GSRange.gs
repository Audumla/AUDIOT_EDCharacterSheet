var GSRange = (function () {
  // ---------- A1 / sheet-name helpers ----------
  function quoteSheetName(name) {
    const s = String(name);
    return /\s/.test(s) ? "'" + s.replace(/'/g, "''") + "'" : s;
  }

  function peelSheet(s) {
    s = String(s).trim();
    const i = s.lastIndexOf('!');
    if (i < 0) return { sheet: null, body: s };
    let raw = s.slice(0, i), sheet = raw;
    if (raw.startsWith("'") && raw.endsWith("'")) {
      sheet = raw.slice(1, -1).replace(/''/g, "'");
    }
    return { sheet, body: s.slice(i + 1) };
  }

  function splitA1(s) {
    s = String(s).trim();
    const i = s.lastIndexOf('!');
    if (i < 0) return { sheetName: null, addrOnly: s };
    let raw = s.slice(0, i);
    if (raw.startsWith("'") && raw.endsWith("'")) {
      raw = raw.slice(1, -1).replace(/''/g, "'");
    }
    return { sheetName: raw, addrOnly: s.slice(i + 1) };
  }

  // ---------- Column conversions ----------
  function lettersToIndex(letters) {
    let n = 0;
    const s = String(letters).toUpperCase();
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }

  function indexToLetters(n) {
    let s = "";
    n = Math.floor(Number(n));
    if (n <= 0) throw new Error("indexToLetters expects 1-based positive integer");
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = (n - 1) / 26 | 0;
    }
    return s;
  }

  function colLettersRange(startCol, endCol) {
    const n = endCol - startCol + 1;
    const out = new Array(n);
    for (let i = 0; i < n; i++) out[i] = indexToLetters(startCol + i);
    return out;
  }

  /**
   * Resolve a Range-like input to a real Range.
   * Accepts:
   *   - Range object  → returned as-is
   *   - A1 string     → "Sheet!A1:B2" or "'My Sheet'!C5" or "A1" (active sheet)
   * @param {GoogleAppsScript.Spreadsheet.Range|string} arg
   * @param {{spreadsheetId?: string}} [opts]
   * @returns {GoogleAppsScript.Spreadsheet.Range}
   */
  function resolveRange(arg, opts) {
    if (GSUtils.Types.isRangeLike(arg)) return arg;
    if (typeof arg === 'string') {
      const ss = (opts && opts.spreadsheetId)
        ? SpreadsheetApp.openById(opts.spreadsheetId)
        : SpreadsheetApp.getActive();
      const { sheetName, addrOnly } = GSUtils.A1.splitA1(arg);
      const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
      if (!sh) throw new Error("Sheet not found: " + sheetName);
      return sh.getRange(addrOnly);
    }
    throw new Error("resolveRange expects a Range or A1 string (e.g., 'Data!A2:D20').");
  }

  /**
   * Convert a Range to absolute A1 with optional $row/$col locks.
   * Quotes sheet name only when it contains whitespace.
   * @param {GoogleAppsScript.Spreadsheet.Range} range
   * @param {boolean} [lockRow=true]
   * @param {boolean} [lockCol=true]
   * @return {string}
   */
  function a1FromRange(range, lockRow = true, lockCol = true) {
    const sh = range.getSheet();
    const sheet = GSUtils.A1.quoteSheetName(sh.getName());

    const r0 = range.getRow();
    const c0 = range.getColumn();
    const r1 = r0 + range.getNumRows() - 1;
    const c1 = c0 + range.getNumColumns() - 1;

    const start = (lockCol ? "$" : "") + GSUtils.Col.indexToLetters(c0) + (lockRow ? "$" : "") + r0;
    const end   = (lockCol ? "$" : "") + GSUtils.Col.indexToLetters(c1) + (lockRow ? "$" : "") + r1;

    const a1 = (r0 === r1 && c0 === c1) ? start : (start + ":" + end);
    return sheet + "!" + a1;
  }

  /**
   * getDisplayArray(rangeOrA1, {trimEmptyRows=true})
   * - Returns a 2D array of DISPLAY VALUES (strings).
   * - Removes rows that are entirely blank/whitespace by default.
   * - No image/formula handling (pure display).
   *
   * @param {GoogleAppsScript.Spreadsheet.Range|string} rangeOrA1
   * @param {{trimEmptyRows?: boolean}} [opts]
   * @return {string[][]}
   */
  function getDisplayArray(rangeOrA1, opts) {
    opts = opts || {};
    const trimEmptyRows = opts.trimEmptyRows !== false;

    const rng = resolveRange(rangeOrA1, opts);
    const displays = rng.getDisplayValues(); // strings
    const H = displays.length;
    const W = H ? displays[0].length : 0;
    if (!H || !W) return [];

    if (!trimEmptyRows) return displays;

    const out = [];
    for (let r = 0; r < H; r++) {
      const row = displays[r];
      let nonEmpty = false;
      for (let c = 0; c < W; c++) {
        const v = row[c];
        if (v != null && String(v).trim() !== "") { nonEmpty = true; break; }
      }
      if (nonEmpty) out.push(row);
    }
    return out;
  }

  /**
   * getValuesArrayWithImageRefs(rangeOrA1)
   * - Returns a 2D array of VALUES, but:
   *   - If a cell is an IMAGE (either =IMAGE(...) or in-cell image),
   *     it becomes a reference formula string: ='<Sheet Name>'!A1
   *   - Otherwise, returns the raw value.
   *
   * @param {GoogleAppsScript.Spreadsheet.Range|string} rangeOrA1
   * @return {any[][]}
   */
  function getValuesArrayWithImageRefs(rangeOrA1, opts) {
    const rng = resolveRange(rangeOrA1, opts);
    const sh  = rng.getSheet();
    const sName = sh.getName();

    const values   = rng.getValues();
    const formulas = rng.getFormulas();

    const H = values.length;
    const W = H ? values[0].length : 0;
    if (!H || !W) return [];

    const startRow = rng.getRow();
    const startCol = rng.getColumn();
    const colLetters = GSUtils.Col.colLettersRange(startCol, startCol + W - 1);
    const sheetPrefix = "=" + GSUtils.A1.quoteSheetName(sName) + "!";
    const rxImage = /^=IMAGE\(/i;

    const out = new Array(H);
    for (let r = 0; r < H; r++) {
      const row = new Array(W);
      const rowNum = startRow + r;
      for (let c = 0; c < W; c++) {
        const v = values[r][c];
        const f = formulas[r][c];

        // (1) Formula-based image
        if (f && rxImage.test(f)) {
          row[c] = sheetPrefix + colLetters[c] + rowNum; // e.g. ='My Sheet'!C5
          continue;
        }

        // (2) In-cell image (CellImage object) — quick detection
        if (v && typeof v === 'object') {
          const looksLikeCellImage =
            typeof v.getUrl === 'function' ||
            typeof v.getAltTextTitle === 'function';
          if (looksLikeCellImage) {
            row[c] = sheetPrefix + colLetters[c] + rowNum;
            continue;
          }
        }

        // (3) Plain value
        row[c] = v;
      }
      out[r] = row;
    }
    return out;
  }

  // ---------- Type guards ----------
  function isRangeLike(x) {
    return x && typeof x.getA1Notation === 'function' && typeof x.getValues === 'function';
  }

  return {
    quoteSheetName, 
    peelSheet, 
    splitA1, 
    lettersToIndex, 
    indexToLetters, 
    colLettersRange,
    resolveRange, 
    a1FromRange,
    getDisplayArray,
    getValuesArrayWithImageRefs,
    isRangeLike
  };
})();
