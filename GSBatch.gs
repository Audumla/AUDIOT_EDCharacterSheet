/** Namespace: GSBatch */
var GSBatch = (function () {
  /* ===================== Public API ===================== */

  /** Create a new batch (pass Spreadsheet or ID; default active). */
  function newBatch(spreadsheet) {
    const ss = _resolveSpreadsheet_(spreadsheet);
    return { ss, spreadsheetId: ss.getId(), requests: [] };
  }

  /** Merge another batch or raw requests into this batch. */
  function merge(intoBatch, other) {
    if (!intoBatch || !intoBatch.requests) throw new Error('intoBatch is not a batch');
    if (!other) return intoBatch;
    if (Array.isArray(other)) intoBatch.requests.push(...other);
    else if (other.requests) intoBatch.requests.push(...other.requests);
    else intoBatch.requests.push(other);
    return intoBatch;
  }

  /** Commit in ONE call (all queued write requests). */
  function commit(batch, { includeResponse = false, clearAfter = true } = {}) {
    if (!batch || !batch.requests || !batch.requests.length) return null;
    const res = Sheets.Spreadsheets.batchUpdate({
      requests: batch.requests,
      includeSpreadsheetInResponse: !!includeResponse
    }, batch.spreadsheetId);
    if (clearAfter) batch.requests.length = 0;
    return res;
  }

  /* ---------- Values (single cell + ranges) ---------- */

  function addValues(batch, a1OrGridRange, values, opts = {}) {
    const grid = (typeof a1OrGridRange === 'string')
      ? _a1ToGridRange_(batch.ss, a1OrGridRange)
      : a1OrGridRange;

    const shaped = _to2D_(values);
    const H = shaped.length, W = shaped[0]?.length || 0;
    if (!H || !W) throw new Error('values must be non-empty');

    if (_isSingleCellGrid_(grid)) {
      if (opts.autoSize !== false && (H > 1 || W > 1)) {
        grid.endRowIndex    = grid.startRowIndex    + H;
        grid.endColumnIndex = grid.startColumnIndex + W;
      } else if (H !== 1 || W !== 1) {
        throw new Error('A1 is a single cell but values are not 1x1. Pass {autoSize:true}.');
      }
    } else {
      const RH = grid.endRowIndex - grid.startRowIndex;
      const RW = grid.endColumnIndex - grid.startColumnIndex;
      if ((H !== RH || W !== RW) && !opts.allowMismatch) {
        throw new Error(`Values size ${H}x${W} does not match range ${RH}x${RW}. Set allowMismatch:true to write upper-left block.`);
      }
    }

    let needsNumFmt = false;
    const rows = shaped.map(r => ({
      values: r.map(v => {
        const cd = _jsToCellData_(v, { dateTimePattern: opts.dateTimePattern });
        if (cd.userEnteredFormat && cd.userEnteredFormat.numberFormat) needsNumFmt = true;
        return cd;
      })
    }));
    const fields = 'userEnteredValue' + (needsNumFmt ? ',userEnteredFormat.numberFormat' : '');

    batch.requests.push({ updateCells: { range: grid, rows, fields } });
    return batch;
  }

  /** Single-cell write (scalar, Date, "=FORMULA", or null to clear). */
  function addCell(batch, a1OrGridRange, value, opts = {}) {
    return addValues(batch, a1OrGridRange, [[value]], opts);
  }

  /** Clear only values (preserve formatting). */
  function clearValues(batch, a1OrGridRange) {
    const grid = (typeof a1OrGridRange === 'string')
      ? _a1ToGridRange_(batch.ss, a1OrGridRange)
      : a1OrGridRange;
    batch.requests.push({
      updateCells: { range: grid, rows: [{ values: [{}] }], fields: 'userEnteredValue' }
    });
    return batch;
  }

  /** Apply number/date format across a range. */
  function formatNumber(batch, a1OrGridRange, { type = 'NUMBER', pattern } = {}) {
    const grid = (typeof a1OrGridRange === 'string')
      ? _a1ToGridRange_(batch.ss, a1OrGridRange)
      : a1OrGridRange;
    batch.requests.push({
      repeatCell: {
        range: grid,
        cell: { userEnteredFormat: { numberFormat: { type, pattern } } },
        fields: 'userEnteredFormat.numberFormat'
      }
    });
    return batch;
  }

  /* ---------- Inserts (same semantics as Sheets) ---------- */

  function insertRows(batch, startRow, nRows, { inheritFromBefore = false } = {}) {
    const { sheetId } = _ensureSheet_(batch.ss);
    batch.requests.push({
      insertDimension: {
        range: { sheetId, dimension: 'ROWS', startIndex: startRow, endIndex: startRow + nRows },
        inheritFromBefore
      }
    });
    return batch;
  }

  function insertColumns(batch, startCol, nCols, { inheritFromBefore = false } = {}) {
    const { sheetId } = _ensureSheet_(batch.ss);
    batch.requests.push({
      insertDimension: {
        range: { sheetId, dimension: 'COLUMNS', startIndex: startCol, endIndex: startCol + nCols },
        inheritFromBefore
      }
    });
    return batch;
  }

  function insertRange(batch, a1OrGridRange, shiftDimension) {
    const grid = (typeof a1OrGridRange === 'string')
      ? _a1ToGridRange_(batch.ss, a1OrGridRange)
      : a1OrGridRange;
    batch.requests.push({ insertRange: { range: grid, shiftDimension } });
    return batch;
  }

  /* ---------- Loader (single batchGet for many ranges) ---------- */

  /**
   * Batch load many A1 ranges using ONE API call.
   *
   * INPUT SHAPES
   * 1) Array of pairs/objects:
   *    - [{name:"people", range:"'Data'!A2:C"}, {name:"rates", range:"Rates!B2:B"}]
   *    - or [["people", "'Data'!A2:C"], ["rates", "Rates!B2:B"]]
   *    → returns [{name, range, values}, ...]
   *
   * 2) Object with .range props:
   *    - { people:{range:"'Data'!A2:C"}, rates:{range:"Rates!B2:B"} }
   *    → returns a cloned object with same props, each prop → { range, values, ... }
   *
   * OPTS:
   *   - spreadsheet: Spreadsheet object (defaults to batch.ss if opts.batch given, else active)
   *   - batch: pass an existing GSBatch batch to reuse its spreadsheet
   *   - render: 'display'|'raw'|'formula'  (default 'display')
   *   - dateTime: 'SERIAL_NUMBER'|'FORMATTED_STRING' (default 'SERIAL_NUMBER')
   */
  function loadRanges(input, opts = {cfg : Configuration, logger : DEFAULT_LOGGER}) {
    const ss = opts.batch?.ss || _resolveSpreadsheet_(opts.spreadsheet);
    const ssId = ss.getId();

    const valueRenderOption = _renderModeToVRO_(opts.render || 'display');
    const dateTimeRenderOption = opts.dateTime || 'SERIAL_NUMBER';

    const norm = _normalizeInputForBatchLoad_(input, ss);
    if (!norm.items.length) return Array.isArray(input) ? [] : JSON.parse(JSON.stringify(input || {}));

    const ranges = norm.items.map(it => _ensureSheetOnA1_(it.range, ss));
    opts.logger.trace("Loading Ranges " + JSON.stringify(ranges));

    const res = Sheets.Spreadsheets.Values.batchGet(ssId, {
      ranges,
      valueRenderOption,
      dateTimeRenderOption
    });
    const valueRanges = (res && res.valueRanges) || [];

    const withResults = norm.items.map((it, i) => {
      const vr = valueRanges[i] || {};
      return { name: it.name, range: vr.range || ranges[i], values: vr.values || [], _idx: it._idx };
    });

    if (norm.kind === 'array') {
      return withResults.map(x => ({ name: x.name, range: x.range, values: x.values }));
    } else {
      const out = JSON.parse(JSON.stringify(input || {}));
      for (const rec of withResults) {
        const prev = out[rec.name] || {};
        out[rec.name] = Object.assign({}, prev, { range: rec.range, values: rec.values });
      }
      return out;
    }
  }

  /* ===================== Helpers (private) ===================== */

  function _resolveSpreadsheet_(spreadsheet) {
    if (!spreadsheet) return SpreadsheetApp.getActive();
    if (typeof spreadsheet === 'string') return SpreadsheetApp.openById(spreadsheet);
    return spreadsheet; // assume Spreadsheet object
  }

  function _isSingleCellGrid_(g) {
    return (g.endRowIndex - g.startRowIndex === 1) && (g.endColumnIndex - g.startColumnIndex === 1);
  }

  function _jsToCellData_(v, { dateTimePattern } = {}) {
    const out = { userEnteredValue: null };
    if (v == null) { out.userEnteredValue = null; return out; }
    if (typeof v === 'string' && v.length && v[0] === '=') { out.userEnteredValue = { formulaValue: v }; return out; }
    if (typeof v === 'boolean') { out.userEnteredValue = { boolValue: v }; return out; }
    if (typeof v === 'number' && Number.isFinite(v)) { out.userEnteredValue = { numberValue: v }; return out; }
    if (v instanceof Date) {
      const serial = _dateToSerial_(v);
      out.userEnteredValue = { numberValue: serial };
      out.userEnteredFormat = { numberFormat: { type: 'DATE_TIME', pattern: dateTimePattern || 'yyyy-mm-dd hh:mm:ss' } };
      return out;
    }
    out.userEnteredValue = { stringValue: String(v) };
    return out;
  }

  function _dateToSerial_(d) {
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    return d.getTime() / MS_PER_DAY + 25569; // 1899-12-30 epoch
  }

  function _a1ToGridRange_(ss, a1) {
    const { sheetName, addrOnly } = _splitA1_(a1);
    const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sh) throw new Error('Sheet not found: ' + sheetName);
    const sheetId = sh.getSheetId();

    const s = addrOnly.replace(/\$/g, '').toUpperCase();
    let m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(s);
    if (m) {
      const c1 = _colToIndex_(m[1]), r1 = +m[2];
      const c2 = _colToIndex_(m[3]), r2 = +m[4];
      return {
        sheetId,
        startRowIndex: Math.min(r1, r2) - 1,
        endRowIndex: Math.max(r1, r2),
        startColumnIndex: Math.min(c1, c2) - 1,
        endColumnIndex: Math.max(c1, c2)
      };
    }
    m = /^([A-Z]+)(\d+)$/.exec(s);
    if (m) {
      const c = _colToIndex_(m[1]), r = +m[2];
      return {
        sheetId,
        startRowIndex: r - 1, endRowIndex: r,
        startColumnIndex: c - 1, endColumnIndex: c
      };
    }
    throw new Error('Range must be a finite rectangle or single cell: ' + a1);
  }

  function _splitA1_(s) {
    s = String(s).trim();
    const i = s.lastIndexOf('!');
    if (i < 0) return { sheetName: null, addrOnly: s };
    let raw = s.slice(0, i);
    if (raw.startsWith("'") && raw.endsWith("'")) raw = raw.slice(1, -1).replace(/''/g, "'");
    return { sheetName: raw, addrOnly: s.slice(i + 1) };
  }

  function _colToIndex_(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function _to2D_(v) {
    if (Array.isArray(v) && Array.isArray(v[0])) return v;
    if (Array.isArray(v)) return [v];
    return [[v]];
  }

  function _ensureSheet_(ss, nameOrId) {
    if (typeof nameOrId === 'number') return { sheetId: nameOrId };
    const sh = ss.getSheetByName(nameOrId);
    if (!sh) throw new Error('Sheet not found: ' + nameOrId);
    return { sheetId: sh.getSheetId() };
  }

  /* ---------- Loader helpers ---------- */

  function _renderModeToVRO_(mode) {
    switch ((mode || '').toLowerCase()) {
      case 'display': return 'FORMATTED_VALUE';
      case 'raw':     return 'UNFORMATTED_VALUE';
      case 'formula': return 'FORMULA';
      default:        return 'FORMATTED_VALUE';
    }
  }

  function _normalizeInputForBatchLoad_(input, ss) {
    if (Array.isArray(input)) {
      const items = [];
      for (let i = 0; i < input.length; i++) {
        const it = input[i];
        if (Array.isArray(it)) {
          items.push({ name: String(it[0]), range: String(it[1]), _idx: i });
        } else if (it && typeof it === 'object' && 'name' in it && 'range' in it) {
          items.push({ name: String(it.name), range: String(it.range), _idx: i });
        } else {
          throw new Error("Array form must be [{name, range}, ...] or [[name, range], ...]");
        }
      }
      return { kind: 'array', items };
    }
    if (input && typeof input === 'object') {
      const items = [];
      for (const prop of Object.keys(input)) {
        const v = input[prop];
        if (v && typeof v === 'object' && 'range' in v) {
          items.push({ name: prop, range: String(v.range), _idx: prop });
        }
      }
      return { kind: 'object', items };
    }
    throw new Error("loadRanges expects an Array or an Object with .range properties.");
  }

  function _ensureSheetOnA1_(a1, ss) {
    const { sheetName, addrOnly } = _splitA1_(a1);
    if (sheetName) return a1;
    const shName = ss.getActiveSheet().getName();
    return _quoteIfNeeded_(shName) + "!" + addrOnly;
  }

  function _quoteIfNeeded_(name) {
    return /\s/.test(name) ? "'" + String(name).replace(/'/g, "''") + "'" : name;
  }

  /* Expose */
  return {
    /** create/compose/commit (writes) */
    new: newBatch,
    merge,
    commit,

    /** value ops */
    add: {
      values: addValues,
      cell: addCell,
      clear: clearValues,
      formatNumber: formatNumber
    },

    /** insert ops */
    insert: {
      rows: insertRows,
      columns: insertColumns,
      range: insertRange
    },

    /** reads */
    load: {
      ranges: loadRanges
    }
  };
})();
