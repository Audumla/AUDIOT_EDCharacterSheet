var GSRange = (function () {
  // ============================================================
  // Column conversions
  // ============================================================

  /**
   * Convert column letters (e.g., "A", "AA") to a 1-based column index.
   * @param {string} letters Column letters.
   * @returns {number} 1-based column index.
   */
  function lettersToIndex(letters) {
    let n = 0;
    const s = String(letters).toUpperCase();
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }

  /**
   * Convert a 1-based column index to column letters (A=1, Z=26, AA=27).
   * @param {number} n 1-based column index.
   * @returns {string} Column letters.
   */
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

  /**
   * Precompute letters for an inclusive column range.
   * @param {number} startCol 1-based start column.
   * @param {number} endCol 1-based end column.
   * @returns {string[]} Array of column letters.
   */
  function colLettersRange(startCol, endCol) {
    const n = endCol - startCol + 1;
    const out = new Array(n);
    for (let i = 0; i < n; i++) out[i] = indexToLetters(startCol + i);
    return out;
  }

  // ============================================================
  // A1 / sheet helpers
  // ============================================================

  /**
   * Quote a sheet name if it contains whitespace; also escapes single quotes.
   * @param {string} name Sheet name.
   * @returns {string} Quoted-or-raw sheet name (no trailing !).
   */
  function quoteSheetName(name) {
    const s = String(name);
    return /\s/.test(s) ? "'" + s.replace(/'/g, "''") + "'" : s;
  }

  /**
   * Split an A1 string into { sheetName, addrOnly }.
   * Accepts "'My Sheet'!A1:B2", "Sheet!A1", or "A1:B2" (no sheet).
   * @param {string} s A1 string.
   * @returns {{sheetName:string|null, addrOnly:string}}
   */
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

  // ============================================================
  // Box parsing / composing (unified, $-aware)
  // ============================================================

  /**
 * Parse an A1 string (with optional sheet and $ locks) into a “box”.
 * Supports:
 *  - Cell/cell area: $A$1:$D$5, B3:C10, a1
 *  - Column-only:    $A:$D, A:A, $B:$B, $A, A
 *  - Row-only:       $3:$7, 5:5, $12, 7
 *
 * @param {string} a1 Full A1 string (with or without sheet prefix).
 * @returns {{
 *   sheetName:string|null,
 *   c1:number,r1:number,c2:number,r2:number,
 *   lc1:boolean,lr1:boolean,lc2:boolean,lr2:boolean
 * }}
 */
function parseBox(a1) {
  const { sheetName, addrOnly } = splitA1(a1);
  const s = String(addrOnly).trim();

  // 1) Cell or area: [$]COL[$]ROW ( : [$]COL[$]ROW )?
  //    e.g. $A$1:$D$5, B3:C10, A1
  let m = /^(\$?)([A-Z]+)(\$?)(\d+)(?::(\$?)([A-Z]+)(\$?)(\d+))?$/i.exec(s);
  if (m) {
    let lc1 = m[1] === '$';
    let c1  = lettersToIndex(m[2].toUpperCase());
    let lr1 = m[3] === '$';
    let r1  = parseInt(m[4], 10);

    const hasRight = !!m[5];
    let lc2 = hasRight ? (m[5] === '$') : lc1;
    let c2  = hasRight ? lettersToIndex(m[6].toUpperCase()) : c1;
    let lr2 = hasRight ? (m[7] === '$') : lr1;
    let r2  = hasRight ? parseInt(m[8], 10) : r1;

    if (c1 > c2) { [c1,c2] = [c2,c1]; [lc1,lc2] = [lc2,lc1]; }
    if (r1 > r2) { [r1,r2] = [r2,r1]; [lr1,lr2] = [lr2,lr1]; }

    return { sheetName, c1, r1, c2, r2, lc1, lr1, lc2, lr2 };
  }

  // 2) Column-only range: [$]COL : [$]COL  e.g. $A:$D, A:A, $B:$B
  m = /^(\$?)([A-Z]+):(\$?)([A-Z]+)$/i.exec(s);
  if (m) {
    let lc1 = m[1] === '$';
    let c1  = lettersToIndex(m[2].toUpperCase());
    let lc2 = m[3] === '$';
    let c2  = lettersToIndex(m[4].toUpperCase());
    if (c1 > c2) { [c1,c2] = [c2,c1]; [lc1,lc2] = [lc2,lc1]; }
    // rows span whole sheet: 1..∞ (unlocked rows by definition)
    return { sheetName, c1, r1: 1, c2, r2: Number.POSITIVE_INFINITY, lc1, lr1: false, lc2, lr2: false };
  }

  // 3) Single column: [$]COL  e.g. $A, D
  m = /^(\$?)([A-Z]+)$/i.exec(s);
  if (m) {
    const lc1 = m[1] === '$';
    const c   = lettersToIndex(m[2].toUpperCase());
    return { sheetName, c1: c, r1: 1, c2: c, r2: Number.POSITIVE_INFINITY, lc1, lr1: false, lc2: lc1, lr2: false };
  }

  // 4) Row-only range: [$]ROW : [$]ROW  e.g. $3:$7, 5:5
  m = /^(\$?)(\d+):(\$?)(\d+)$/.exec(s);
  if (m) {
    let lr1 = m[1] === '$';
    let r1  = parseInt(m[2], 10);
    let lr2 = m[3] === '$';
    let r2  = parseInt(m[4], 10);
    if (r1 > r2) { [r1,r2] = [r2,r1]; [lr1,lr2] = [lr2,lr1]; }
    // columns span whole sheet: 1..∞ (unlocked columns by definition)
    return { sheetName, c1: 1, r1, c2: Number.POSITIVE_INFINITY, r2, lc1: false, lr1, lc2: false, lr2 };
  }

  // 5) Single row: [$]ROW  e.g. $12, 7
  m = /^(\$?)(\d+)$/.exec(s);
  if (m) {
    const lr1 = m[1] === '$';
    const r   = parseInt(m[2], 10);
    return { sheetName, c1: 1, r1: r, c2: Number.POSITIVE_INFINITY, r2: r, lc1: false, lr1, lc2: false, lr2: lr1 };
  }

  throw new Error(`Unsupported A1 [${a1}]`);
}

/**
 * Compose an A1 string from a box (preserves $ locks).
 * Handles:
 *  - cell/cell area
 *  - column-only (1..∞ rows)  → A:A or A:D (with $ as provided)
 *  - row-only (1..∞ cols)     → 3:7 or 5:5 (with $ as provided)
 * @param {{
 *   sheetName:string|null,
 *   c1:number,r1:number,c2:number,r2:number,
 *   lc1:boolean,lr1:boolean,lc2:boolean,lr2:boolean
 * }} box
 * @returns {string}
 */
function composeBox(box) {
  const hasInfiniteRows = box.r2 === Number.POSITIVE_INFINITY;
  const hasInfiniteCols = box.c2 === Number.POSITIVE_INFINITY;
  const sheet = box.sheetName ? quoteSheetName(box.sheetName) + "!" : "";

  // Column-only form (rows 1..∞)
  if (!hasInfiniteCols && hasInfiniteRows) {
    const left  = (box.lc1 ? "$" : "") + indexToLetters(box.c1);
    const right = (box.lc2 ? "$" : "") + indexToLetters(box.c2);
    return sheet + (box.c1 === box.c2 ? (left + ":" + right) : (left + ":" + right)); // A:A or A:D
  }

  // Row-only form (cols 1..∞)
  if (!hasInfiniteRows && hasInfiniteCols) {
    const left  = (box.lr1 ? "$" : "") + String(box.r1);
    const right = (box.lr2 ? "$" : "") + String(box.r2);
    return sheet + (box.r1 === box.r2 ? (left + ":" + right) : (left + ":" + right)); // 3:3 or 3:7
  }

  // Finite rectangle (cell or area)
  const cell = (c, r, lc, lr) => (lc ? "$" : "") + indexToLetters(c) + (lr ? "$" : "") + r;
  const L = cell(box.c1, box.r1, box.lc1, box.lr1);
  const R = cell(box.c2, box.r2, box.lc2, box.lr2);
  const body = (box.c1 === box.c2 && box.r1 === box.r2) ? L : (L + ":" + R);
  return sheet + body;
}


  // ============================================================
  // Range resolution / formatting
  // ============================================================

  /**
   * Resolve a Range-like input to a Range.
   * Accepts:
   *   - Range object → returned as-is
   *   - A1 string    → resolved against opts.ss (or active spreadsheet)
   * @param {GoogleAppsScript.Spreadsheet.Range|string} arg
   * @param {{ss?: GoogleAppsScript.Spreadsheet.Spreadsheet}} [opts]
   * @returns {GoogleAppsScript.Spreadsheet.Range}
   */
  function resolveRange(arg, opts) {
    if (arg && typeof arg.getA1Notation === 'function' && typeof arg.getValues === 'function') {
      return arg; // Range
    }
    if (typeof arg === 'string') {
      const { sheetName, addrOnly } = splitA1(arg);
      const ss = (opts && opts.ss) ? opts.ss : SpreadsheetApp.getActive();
      const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
      if (!sh) throw new Error("Sheet not found: " + sheetName);
      return sh.getRange(addrOnly);
    }
    throw new Error("resolveRange expects a Range or A1 string (e.g., 'Data!A2:D20').");
  }

  /**
   * Convert a Range to absolute A1 (optionally $-lock rows/cols).
   * @param {GoogleAppsScript.Spreadsheet.Range} range Range to convert.
   * @param {boolean} [lockRow=true] Whether to $-lock rows.
   * @param {boolean} [lockCol=true] Whether to $-lock columns.
   * @returns {string} A1 with sheet (quoted if needed).
   */
  function a1FromRange(range, lockRow = true, lockCol = true) {
    const sh = range.getSheet();
    const sheet = quoteSheetName(sh.getName());
    const r0 = range.getRow();
    const c0 = range.getColumn();
    const r1 = r0 + range.getNumRows() - 1;
    const c1 = c0 + range.getNumColumns() - 1;

    const start = (lockCol ? "$" : "") + indexToLetters(c0) + (lockRow ? "$" : "") + r0;
    const end   = (lockCol ? "$" : "") + indexToLetters(c1) + (lockRow ? "$" : "") + r1;
    const a1    = (r0 === r1 && c0 === c1) ? start : (start + ":" + end);
    return sheet + "!" + a1;
  }

  /**
   * Build A1 from an onEdit(e) event object (absolute A1 with locks).
   * @param {{range: GoogleAppsScript.Spreadsheet.Range}} e Event object with e.range.
   * @returns {string} A1 with sheet.
   */
  function a1FromEvent(e) {
    if (!e || !e.range) throw new Error("Missing event or e.range");
    return a1FromRange(e.range);
  }

  /**
   * Ensure an A1 string has a sheet prefix; if missing, uses the active sheet of ss.
   * @param {string} a1 A1 string (with/without sheet).
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Spreadsheet (defaults to active).
   * @returns {string} A1 string guaranteed to have a sheet prefix.
   */
  function ensureSheetOnA1(a1, ss) {
    const { sheetName, addrOnly } = splitA1(a1);
    if (sheetName) return a1;
    const _ss = ss || SpreadsheetApp.getActive();
    const shName = _ss.getActiveSheet().getName();
    return quoteSheetName(shName) + "!" + addrOnly;
  }

  // ============================================================
  // Array extraction
  // ============================================================

  /**
   * Get a 2D array of display values (strings). Optionally drops all-blank rows.
   * @param {GoogleAppsScript.Spreadsheet.Range|string} rangeOrA1 Range or A1.
   * @param {{trimEmptyRows?:boolean, ss?:GoogleAppsScript.Spreadsheet.Spreadsheet}} [opts]
   * @returns {string[][]} 2D array of display strings.
   */
  function getDisplayArray(rangeOrA1, opts) {
    const rng = resolveRange(rangeOrA1, opts);
    const displays = rng.getDisplayValues();
    const H = displays.length, W = H ? displays[0].length : 0;
    if (!H || !W) return [];
    if (opts && opts.trimEmptyRows === false) return displays;

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
   * Get a 2D array of values; image cells become self-references like ='<Sheet>'!A1.
   * Detects both `=IMAGE(...)` formulas and in-cell `CellImage` objects.
   * @param {GoogleAppsScript.Spreadsheet.Range|string} rangeOrA1 Range or A1.
   * @param {{ss?:GoogleAppsScript.Spreadsheet.Spreadsheet}} [opts]
   * @returns {any[][]} 2D array of raw values, with images replaced by local A1 references.
   */
  function getValuesArrayWithImageRefs(rangeOrA1, opts) {
    const rng = resolveRange(rangeOrA1, opts);
    const sh  = rng.getSheet();
    const sName = sh.getName();

    const values   = rng.getValues();
    const formulas = rng.getFormulas();

    const H = values.length, W = H ? values[0].length : 0;
    if (!H || !W) return [];

    const startRow = rng.getRow();
    const startCol = rng.getColumn();
    const cols = colLettersRange(startCol, startCol + W - 1);
    const sheetPrefix = "=" + quoteSheetName(sName) + "!";
    const rxImage = /^=IMAGE\(/i;

    const out = new Array(H);
    for (let r = 0; r < H; r++) {
      const row = new Array(W);
      const rowNum = startRow + r;
      for (let c = 0; c < W; c++) {
        const v = values[r][c];
        const f = formulas[r][c];

        if (f && rxImage.test(f)) {
          row[c] = sheetPrefix + cols[c] + rowNum;
          continue;
        }
        if (v && typeof v === 'object' &&
           (typeof v.getUrl === 'function' || typeof v.getAltTextTitle === 'function')) {
          row[c] = sheetPrefix + cols[c] + rowNum;
          continue;
        }
        row[c] = v;
      }
      out[r] = row;
    }
    return out;
  }

  // ============================================================
  // A1 geometry / extension
  // ============================================================

  /**
   * Extend or shrink an A1 range by deltas on each side (preserves $ locks).
   * @param {string} a1 Input A1 (with or without sheet, with/without $).
   * @param {{top?:number,bottom?:number,left?:number,right?:number,rows?:number,cols?:number}} delta
   *        Shorthands: rows → bottom; cols → right. Side-specific takes precedence.
   * @param {{clampToSheet?:boolean, ss?:GoogleAppsScript.Spreadsheet.Spreadsheet}} [opts]
   *        If clampToSheet, clamps to sheet bounds (uses opts.ss or active).
   * @returns {string} Extended A1 (sheet preserved if provided).
   */
  function extendA1(a1, delta, opts) {
    if (typeof a1 !== 'string') throw new Error("extendA1 expects an A1 string");
    delta = delta || {};
    opts  = opts  || {};

    const box = parseBox(a1);

    const dTop    = (delta.top    != null) ? delta.top    : 0;
    const dBottom = (delta.bottom != null) ? delta.bottom : (delta.rows || 0);
    const dLeft   = (delta.left   != null) ? delta.left   : 0;
    const dRight  = (delta.right  != null) ? delta.right  : (delta.cols || 0);

    let c1 = box.c1 - dLeft;
    let r1 = box.r1 - dTop;
    let c2 = box.c2 + dRight;
    let r2 = box.r2 + dBottom;

    let maxRows = Number.POSITIVE_INFINITY;
    let maxCols = Number.POSITIVE_INFINITY;

    if (opts.clampToSheet) {
      const ss = opts.ss || SpreadsheetApp.getActive();
      const sh = box.sheetName ? ss.getSheetByName(box.sheetName) : ss.getActiveSheet();
      if (!sh) throw new Error("extendA1: sheet not found: " + box.sheetName);
      maxRows = sh.getMaxRows();
      maxCols = sh.getMaxColumns();
    }

    c1 = Math.max(1, Math.min(c1, maxCols));
    c2 = Math.max(1, Math.min(c2, maxCols));
    r1 = Math.max(1, Math.min(r1, maxRows));
    r2 = Math.max(1, Math.min(r2, maxRows));

    if (c1 > c2) c2 = c1;
    if (r1 > r2) r2 = r1;

    return composeBox({ ...box, c1, r1, c2, r2 });
  }

  // ============================================================
  // Range relationship helpers (geometry checks)
  // ============================================================

  /**
   * Normalize input (A1 or Range/RangeList) into boxes with indices.
   * @param {(string|GoogleAppsScript.Spreadsheet.Range|GoogleAppsScript.Spreadsheet.RangeList)[]} arr
   * @returns {{idx:number, box:{sheet:string|null,c1:number,r1:number,c2:number,r2:number}}[]}
   * @private
   */
  function _expandArrayWithBoxes_(arr) {
    const out = [];
    for (let i = 0; i < arr.length; i++) {
      const it = arr[i];
      if (it && typeof it.getRanges === 'function') {
        const rs = it.getRanges();
        for (let k = 0; k < rs.length; k++) out.push({ idx: i, box: _boxFromRange_(rs[k]) });
      } else if (it && typeof it.getA1Notation === 'function') {
        out.push({ idx: i, box: _boxFromRange_(it) });
      } else {
        const b = parseBox(String(it));
        out.push({ idx: i, box: { sheet: b.sheetName, c1: b.c1, r1: b.r1, c2: b.c2, r2: b.r2 } });
      }
    }
    return out;
  }

  /**
   * Convert a Range into a simple box.
   * @param {GoogleAppsScript.Spreadsheet.Range} rng
   * @returns {{sheet:string|null,c1:number,r1:number,c2:number,r2:number}}
   * @private
   */
  function _boxFromRange_(rng) {
    const sh = rng.getSheet().getName();
    const r1 = rng.getRow(), c1 = rng.getColumn();
    const r2 = r1 + rng.getNumRows() - 1;
    const c2 = c1 + rng.getNumColumns() - 1;
    return { sheet: sh, c1, r1, c2, r2 };
  }

  /**
   * Test the relation between two boxes.
   * @param {'auto'|'intersect'|'within'|'contains'|'equal'} mode
   * @param {{sheet:string|null,c1:number,r1:number,c2:number,r2:number}} A
   * @param {{sheet:string|null,c1:number,r1:number,c2:number,r2:number}} B
   * @returns {boolean}
   * @private
   */
  function _relationTest_(mode, A, B) {
    switch (mode) {
      case 'intersect': return _boxesIntersect_(A, B);
      case 'within':    return _boxWithin_(A, B);
      case 'contains':  return _boxWithin_(B, A);
      case 'equal':     return _boxesEqual_(A, B);
      case 'auto': {
        const isSingleA = (A.c1 === A.c2) && (A.r1 === A.r2);
        return isSingleA ? _boxWithin_(A, B) : _boxesIntersect_(A, B);
      }
      default: throw new Error('Unknown mode: ' + mode);
    }
  }

  /** @private */
  function _boxesIntersect_(A, B) {
    const colOverlap = !(A.c2 < B.c1 || B.c2 < A.c1);
    const rowOverlap = !(A.r2 < B.r1 || B.r2 < A.r1);
    return colOverlap && rowOverlap;
  }

  /** @private */
  function _boxWithin_(A, B) {
    return A.c1 >= B.c1 && A.c2 <= B.c2 && A.r1 >= B.r1 && A.r2 <= B.r2;
  }

  /** @private */
  function _boxesEqual_(A, B) {
    const sheetsOk = (!A.sheet || !B.sheet || A.sheet === B.sheet);
    return sheetsOk && A.c1 === B.c1 && A.c2 === B.c2 && A.r1 === B.r1 && A.r2 === B.r2;
  }

  /**
   * TRUE if any A relates to any B (default mode: intersect).
   * @param {(string|GoogleAppsScript.Spreadsheet.Range|GoogleAppsScript.Spreadsheet.RangeList)[]} aList
   * @param {(string|GoogleAppsScript.Spreadsheet.Range|GoogleAppsScript.Spreadsheet.RangeList)[]} bList
   * @param {{mode?: 'auto'|'intersect'|'within'|'contains'|'equal'}} [opts]
   * @returns {boolean}
   */
  function rangesIntersectAny(aList, bList, opts) {
    const mode = (opts && opts.mode) || 'intersect';
    aList = (typeof aList === "string" ) ? [aList] : aList;
    const Aexp = _expandArrayWithBoxes_(aList);
    const Bexp = _expandArrayWithBoxes_(bList);

    for (let i = 0; i < Aexp.length; i++) {
      const A = Aexp[i].box;
      for (let j = 0; j < Bexp.length; j++) {
        const B = Bexp[j].box;
        if (A.sheet && B.sheet && A.sheet !== B.sheet) continue;
        if (_relationTest_(mode, A, B)) return true;
      }
    }
    return false;
  }

  /**
   * Return all matching pairs with indices and normalized boxes.
   * @param {(string|GoogleAppsScript.Spreadsheet.Range|GoogleAppsScript.Spreadsheet.RangeList)[]} aList
   * @param {(string|GoogleAppsScript.Spreadsheet.Range|GoogleAppsScript.Spreadsheet.RangeList)[]} bList
   * @param {{mode?: 'auto'|'intersect'|'within'|'contains'|'equal'}} [opts]
   * @returns {{ai:number, bi:number, aBox:{sheet:string|null,c1:number,r1:number,c2:number,r2:number}, bBox:{sheet:string|null,c1:number,r1:number,c2:number,r2:number}}[]}
   */
  function rangesIntersectPairs(aList, bList, opts) {
    aList = typeof aList === "string" ? [aList] : aList;
    bList = typeof bList === "string" ? [bList] : bList;
    const mode = (opts && opts.mode) || 'intersect';
    const Aexp = _expandArrayWithBoxes_(aList);
    const Bexp = _expandArrayWithBoxes_(bList);

    const out = [];
    for (let i = 0; i < Aexp.length; i++) {
      const Ai = Aexp[i];
      for (let j = 0; j < Bexp.length; j++) {
        const Bj = Bexp[j];
        if (Ai.box.sheet && Bj.box.sheet && Ai.box.sheet !== Bj.box.sheet) continue;
        if (_relationTest_(mode, Ai.box, Bj.box)) out.push({ ai: Ai.idx, bi: Bj.idx, aBox: Ai.box, bBox: Bj.box });
      }
    }
    return out;
  }

  /**
   * Relation test between two inputs (A1 or Range). 'auto' = single-cell A within B, else intersect.
   * @param {string|GoogleAppsScript.Spreadsheet.Range} a
   * @param {string|GoogleAppsScript.Spreadsheet.Range} b
   * @param {{mode?: 'auto'|'intersect'|'within'|'contains'|'equal'}} [opts]
   * @returns {boolean}
   */
  function inOrIntersects(a, b, opts) {
    const mode = (opts && opts.mode) || 'auto';
    const A = (typeof a === 'string')
      ? ( () => { const bx = parseBox(a); return { sheet: bx.sheetName, c1: bx.c1, r1: bx.r1, c2: bx.c2, r2: bx.r2 }; })()
      : _boxFromRange_(a);
    const B = (typeof b === 'string')
      ? ( () => { const bx = parseBox(b); return { sheet: bx.sheetName, c1: bx.c1, r1: bx.r1, c2: bx.c2, r2: bx.r2 }; })()
      : _boxFromRange_(b);

    if (A.sheet && B.sheet && A.sheet !== B.sheet) return false;
    return _relationTest_(mode, A, B);
  }

  /**
   * Convenience: any overlap.
   * @param {string|GoogleAppsScript.Spreadsheet.Range} a
   * @param {string|GoogleAppsScript.Spreadsheet.Range} b
   * @returns {boolean}
   */
  function rangesIntersect(a, b) { return inOrIntersects(a, b, { mode: 'intersect' }); }

  /**
   * Convenience: A ⊆ B.
   * @param {string|GoogleAppsScript.Spreadsheet.Range} a
   * @param {string|GoogleAppsScript.Spreadsheet.Range} b
   * @returns {boolean}
   */
  function rangeWithin(a, b)     { return inOrIntersects(a, b, { mode: 'within'    }); }

  /**
   * Convenience: A ⊇ B.
   * @param {string|GoogleAppsScript.Spreadsheet.Range} a
   * @param {string|GoogleAppsScript.Spreadsheet.Range} b
   * @returns {boolean}
   */
  function rangeContains(a, b)   { return inOrIntersects(a, b, { mode: 'contains'  }); }

  // ============================================================
  // Public API
  // ============================================================

  return {
    // columns
    lettersToIndex,
    indexToLetters,
    colLettersRange,

    // A1 helpers
    quoteSheetName,
    splitA1,
    parseBox,
    composeBox,
    ensureSheetOnA1,

    // range resolution/format
    resolveRange,
    a1FromRange,
    a1FromEvent,

    // extraction
    getDisplayArray,
    getValuesArrayWithImageRefs,

    // geometry
    extendA1,
    rangesIntersectAny,
    rangesIntersectPairs,
    inOrIntersects,
    rangesIntersect,
    rangeWithin,
    rangeContains,
  };
})();
