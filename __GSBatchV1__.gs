/** Namespace: GSBatch */
var GSBatchV1 = (function () {
  var batchID = 1;
  /* ===================== Public API ===================== */

  /** Create a new batch (pass Spreadsheet or ID; default active). */
  function newBatch(spreadsheet) {
    const ss = _resolveSpreadsheet_(spreadsheet);
    return { ss, spreadsheetId: ss.getId(), requests: [], batchID : batchID++ };
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

  function size(batch) {
      return GSUtils.Str.byteLen(JSON.stringify(batch.requests));
  }

  /** Commit in ONE call (all queued write requests). */
  function commit(batch,opts = {}) {
    const { includeResponse =  false, clearAfter = true } = opts;
    console.info(JSON.stringify(batch.requests));
    
    if (batch && batch.requests && batch.requests.length) {
      const res = Sheets.Spreadsheets.batchUpdate({
        requests: batch.requests,
        includeSpreadsheetInResponse: !!includeResponse
      }, batch.spreadsheetId);
      if (clearAfter) batch.requests.length = 0;
//      SpreadsheetApp.flush();
      return res;
    }
    else {
      return undefined;
    }
  }

  /* ---------- Values (single cell + ranges) ---------- */

  function addValues(batch, a1OrGridRange, values, opts = {}) {

    const {allowMismatch = false,autoSize = true} = opts;

    const grid = (typeof a1OrGridRange === 'string')
      ? _a1ToGridRange_(batch.ss, a1OrGridRange)
      : a1OrGridRange;

    const shaped = GSUtils.Arr.to2D(values);
    const H = shaped.length, W = shaped[0]?.length || 0;
    if (!H || !W) throw new Error('values must be non-empty');

    if (autoSize) {
      grid.endRowIndex    = grid.startRowIndex    + H;
      grid.endColumnIndex = grid.startColumnIndex + W;
      
    } else {
      const RH = grid.endRowIndex - grid.startRowIndex;
      const RW = grid.endColumnIndex - grid.startColumnIndex;
      if ((H !== RH || W !== RW) && !allowMismatch) {
        throw new Error(`Values size ${H}x${W} does not match range ${RH}x${RW} (${grid}). Set allowMismatch:true to write upper-left block.`);
      }
    }

    let needsNumFmt = false;
    const rows = shaped.map(r => ({
      values: r.map(v => {
        const cd = _jsToCellData_(v, opts);
        if (cd.userEnteredFormat && cd.userEnteredFormat.numberFormat) needsNumFmt = true;
        return cd;
      })
    }));
    const fields = 'userEnteredValue' + (needsNumFmt ? ',userEnteredFormat.numberFormat' : '');

    batch.requests.push({ updateCells: { range: grid, rows, fields } });

    return batch;
  }

  /** Single-cell write (scalar, Date, "=FORMULA", or null to clear). */
  function addCell(batch, a1OrGridRange, value, opts) {
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
   *   - render: 'display'|'raw'|'formula'  (default 'display')
   *   - dateTime: 'SERIAL_NUMBER'|'FORMATTED_STRING' (default 'SERIAL_NUMBER')
   *   - trim: Remove any rows that have an empty first column from the return values
   */
  function loadRanges(input, opts = {}) {

    const { trim = true } = opts;
    const valueRenderOption = renderMode(opts.render || 'raw');
    const dateTimeRenderOption = opts.dateTime || 'SERIAL_NUMBER';

    const norm = _normalizeInputForBatchLoad_(input, EDContext.context.ss);
    if (!norm.items.length) return Array.isArray(input) ? [] : JSON.parse(JSON.stringify(input || {}));

    const ranges = norm.items.map(it => GSRange.ensureSheetOnA1(it.range, EDContext.context.ss));
    EDLogger.trace("Loading Ranges " + JSON.stringify(ranges));

    const res = Sheets.Spreadsheets.Values.batchGet(EDContext.context.ssid, {
      ranges,
      valueRenderOption,
      dateTimeRenderOption
    });
    const valueRanges = (res && res.valueRanges) || [];

    const withResults = norm.items.map((it, i) => {

      const vr = valueRanges[i] || {};
      const fValues = trim ? vr.values.filter(r => r[0] && r[0].length > 0) : vr.values;
      return { name: it.name, range: vr.range || ranges[i], values: fValues || [], _idx: it._idx };
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


  function _jsToCellData_(v, { dtfmt } = {}) {
    const out = { userEnteredValue: null };
    if (v == null) { out.userEnteredValue = null; return out; }
    if (typeof v === 'string' && v.length && v[0] === '=') { out.userEnteredValue = { formulaValue: v }; return out; }
    if (typeof v === 'boolean') { out.userEnteredValue = { boolValue: v }; return out; }
    if (typeof v === 'number' && Number.isFinite(v)) { out.userEnteredValue = { numberValue: v }; return out; }
    if (v instanceof Date) {
      const serial = GSUtils.Date.dateToSerial(v);
      out.userEnteredValue = { numberValue: serial };
      out.userEnteredFormat = { numberFormat: { type: 'DATE_TIME', pattern: dtfmt || 'yyyy-mm-dd hh:mm:ss' } };
      return out;
    }
    out.userEnteredValue = { stringValue: String(v) };
    return out;
  }

  function _a1ToGridRange_(ss, a1) {
    const { sheetName, addrOnly } = GSRange.splitA1(a1);
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


  function _colToIndex_(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function _ensureSheet_(ss, nameOrId) {
    if (typeof nameOrId === 'number') return { sheetId: nameOrId };
    const sh = ss.getSheetByName(nameOrId);
    if (!sh) throw new Error('Sheet not found: ' + nameOrId);
    return { sheetId: sh.getSheetId() };
  }

  /* ---------- Loader helpers ---------- */

  function renderMode(mode) {
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

/* ---------- Deletes (same semantics as Sheets) ---------- */

/**
 * Delete N whole rows starting at startRow (0-based).
 * @param {*} batch
 * @param {number} startRow 0-based start row index
 * @param {number} nRows    number of rows to delete
 * @param {{sheet?: string|number}} [opts] pass a sheet name or ID; default = active sheet
 */
function deleteRows(batch, startRow, nRows, opts = {}) {
  const { sheetId } = _ensureSheet_(batch.ss, opts.sheet ?? batch.ss.getActiveSheet().getName());
  batch.requests.push({
    deleteDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: startRow, endIndex: startRow + nRows }
    }
  });
  return batch;
}

/**
 * Delete N whole columns starting at startCol (0-based).
 * @param {*} batch
 * @param {number} startCol 0-based start column index
 * @param {number} nCols    number of columns to delete
 * @param {{sheet?: string|number}} [opts]
 */
function deleteColumns(batch, startCol, nCols, opts = {}) {
  const { sheetId } = _ensureSheet_(batch.ss, opts.sheet ?? batch.ss.getActiveSheet().getName());
  batch.requests.push({
    deleteDimension: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: startCol, endIndex: startCol + nCols }
    }
  });
  return batch;
}

/**
 * Delete a rectangular range (shifts cells to fill). Use shiftDimension:
 *  - 'ROWS'    → shift cells up
 *  - 'COLUMNS' → shift cells left
 * @param {*} batch
 * @param {string|Object} a1OrGridRange e.g. "Sheet!B3:D7" or a GridRange
 * @param {'ROWS'|'COLUMNS'} shiftDimension
 */
function deleteRange(batch, a1OrGridRange, shiftDimension) {
  const grid = (typeof a1OrGridRange === 'string')
    ? _a1ToGridRange_(batch.ss, a1OrGridRange)
    : a1OrGridRange;
  batch.requests.push({ deleteRange: { range: grid, shiftDimension } });
  return batch;
}

/**
 * Delete whole rows referenced by an A1 row range, e.g. "Sheet!3:7" or "3:3".
 * (Convenience wrapper around deleteRows.)
 */
function deleteRowsA1(batch, a1RowRange) {
  const { sheetName, addrOnly } = GSRange.splitA1(a1RowRange);
  const ss = batch.ss;
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  const sheetId = sh.getSheetId();

  const m = /^(\d+):(\d+)$/.exec(String(addrOnly).trim());
  if (!m) throw new Error('Expected a row range like "3:7"');
  const r1 = Math.min(+m[1], +m[2]) - 1;     // 0-based
  const r2 = Math.max(+m[1], +m[2]);         // exclusive end in API terms

  batch.requests.push({
    deleteDimension: {
      range: { sheetId, dimension: 'ROWS', startIndex: r1, endIndex: r2 }
    }
  });
  return batch;
}

  /* Expose */
  return {
    /** create/compose/commit (writes) */
    newBatch,
    merge,
    commit,
    size,

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
    
    /** remove ops */
    remove: { 
      rows: deleteRows, 
      columns: deleteColumns, 
      range: deleteRange, 
      rowsA1: deleteRowsA1 
    },

    /** reads */
    load: {
      ranges: loadRanges,
      renderMode
    }
  };
})();
