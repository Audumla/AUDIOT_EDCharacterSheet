/** Namespace: GSBatch — Values-first batching + single-batchUpdate + fallback (SpreadsheetApp only). */
var GSBatch = (function () {
  /* ===================== String/enum constants ===================== */
  const MODE = {
    VALUES: 'values',
    SINGLE: 'singleBatchUpdate',
    SIMPLE: 'simple'
  };

  const KIND = {
    UPDATE: 'update',
    CLEAR:  'clear',
    GET:    'get',
    APPEND: 'append',
    STRUCT: 'struct'
  };

  const VALUE_INPUT = {
    USER: 'USER_ENTERED',
    RAW:  'RAW'
  };

  const INSERT_DATA = {
    INSERT_ROWS: 'INSERT_ROWS'
  };

  const MAJOR_DIM = {
    ROWS: 'ROWS',
    COLS: 'COLUMNS'
  };

  const RENDER = {
    UNFORMATTED: 'UNFORMATTED_VALUE',
    FORMATTED:   'FORMATTED_VALUE',
    FORMULA:     'FORMULA'
  };

  const DATETIME = {
    SERIAL: 'SERIAL_NUMBER',
    STRING: 'FORMATTED_STRING'
  };

  const DIMENSION = {
    ROWS: 'ROWS',
    COLUMNS: 'COLUMNS'
  };

  const FIELDS = {
    USER_VAL: 'userEnteredValue',
    USER_FMT_NUM: 'userEnteredFormat.numberFormat'
  };

  const SHIFT = {
    ROWS: SpreadsheetApp.Dimension.ROWS,
    COLS: SpreadsheetApp.Dimension.COLUMNS
  };

  const DEFAULTS = {
    VALUE_INPUT: VALUE_INPUT.USER,
    RENDER: RENDER.UNFORMATTED,
    DATETIME: DATETIME.SERIAL,
    MAJOR_DIM: MAJOR_DIM.ROWS,
    INSERT_DATA: INSERT_DATA.INSERT_ROWS
  };

  /* ===================== Public API ===================== */
  var batchID = 1;

  function newBatch(spreadsheet, opts = {}) {
    const ss = _resolveSpreadsheet_(spreadsheet);
    const mode = (opts.mode === MODE.SINGLE || opts.mode === MODE.SIMPLE) ? opts.mode : MODE.VALUES;
    return {
      ss,
      spreadsheetId: ss.getId(),
      ops: [],
      mode,
      batchID: batchID++
    };
  }

  function merge(intoBatch, other) {
    if (!intoBatch || !intoBatch.ops) throw new Error('intoBatch is not a GSBatch batch');
    if (!other) return intoBatch;
    if (Array.isArray(other)) intoBatch.ops.push(...other);
    else if (other.ops) intoBatch.ops.push(...other.ops);
    else if (other.kind) intoBatch.ops.push(other);
    else throw new Error('merge expects GSBatch batch, ops array, or single op');
    return intoBatch;
  }

  function size(batch) {
    return GSUtils.Str.byteLen(JSON.stringify(batch.ops));
  }

  function commit(batch, opts = {}) {
    const { clearAfter = true } = opts;
    if (!batch || !batch.ops || !batch.ops.length) return [];

    if (batch.mode === MODE.SINGLE) {
      const requests = _opsToSingleBatchUpdateRequests_(batch);
      const res = Sheets.Spreadsheets.batchUpdate({ requests }, batch.spreadsheetId);
      if (clearAfter) batch.ops.length = 0;
      return [res];
    }

    if (batch.mode === MODE.SIMPLE) {
      const res = _commitFallback_(batch);
      if (clearAfter) batch.ops.length = 0;
      return res;
    }

    // ----- MODE.VALUES (Advanced Sheets Values.*; struct → batchUpdate) -----
    const sid = batch.spreadsheetId;
    const results = [];
    let group = null;

    const flush = () => {
      if (!group) return;
      switch (group.type) {
        case KIND.UPDATE: {
          const data = group.items.map(it => ({
            range: it.range,
            values: it.values,
            majorDimension: it.majorDimension || DEFAULTS.MAJOR_DIM
          }));
          const body = {
            valueInputOption: group.options.valueInputOption || DEFAULTS.VALUE_INPUT,
            includeValuesInResponse: !!group.options.includeValuesInResponse,
            data
          };
          results.push(Sheets.Spreadsheets.Values.batchUpdate(body, sid));
          break;
        }
        case KIND.CLEAR: {
          const body = { ranges: group.items.map(it => it.range) };
          results.push(Sheets.Spreadsheets.Values.batchClear(body, sid));
          break;
        }
        case KIND.GET: {
          const body = {
            ranges: group.items.map(it => it.range),
            valueRenderOption: group.options.valueRenderOption || DEFAULTS.RENDER,
            dateTimeRenderOption: group.options.dateTimeRenderOption || DEFAULTS.DATETIME
          };
          results.push(Sheets.Spreadsheets.Values.batchGet(sid, body));
          break;
        }
        case KIND.STRUCT: {
          const requests = group.items.flatMap(it => it.requests);
          if (requests.length) results.push(Sheets.Spreadsheets.batchUpdate({ requests }, sid));
          break;
        }
      }
      group = null;
    };

    const compat = (grp, op) => {
      if (!grp) return false;
      if (grp.type !== op.kind && !(grp.type === KIND.STRUCT && op.kind === KIND.STRUCT)) return false;
      switch (grp.type) {
        case KIND.UPDATE:
          return (grp.options.valueInputOption || DEFAULTS.VALUE_INPUT) === (op.options.valueInputOption || DEFAULTS.VALUE_INPUT) &&
                 !!grp.options.includeValuesInResponse === !!op.options.includeValuesInResponse;
        case KIND.CLEAR:
          return true;
        case KIND.GET:
          return (grp.options.valueRenderOption || DEFAULTS.RENDER) === (op.options.valueRenderOption || DEFAULTS.RENDER) &&
                 (grp.options.dateTimeRenderOption || DEFAULTS.DATETIME) === (op.options.dateTimeRenderOption || DEFAULTS.DATETIME);
        case KIND.STRUCT:
          return true;
        default:
          return false;
      }
    };

    for (const op of batch.ops) {
      switch (op.kind) {
        case KIND.UPDATE:
          if (!group || !compat(group, op)) { flush(); group = { type: KIND.UPDATE, options: op.options || {}, items: [] }; }
          group.items.push(op);
          break;
        case KIND.CLEAR:
          if (!group || !compat(group, op)) { flush(); group = { type: KIND.CLEAR, options: {}, items: [] }; }
          group.items.push(op);
          break;
        case KIND.GET:
          if (!group || !compat(group, op)) { flush(); group = { type: KIND.GET, options: op.options || {}, items: [] }; }
          group.items.push(op);
          break;
        case KIND.APPEND:
          flush(); // append must be per-call in Values API
          results.push(Sheets.Spreadsheets.Values.append({
            values: op.values,
            majorDimension: op.majorDimension || DEFAULTS.MAJOR_DIM
          }, sid, op.range, {
            valueInputOption: op.options.valueInputOption || DEFAULTS.VALUE_INPUT,
            insertDataOption: op.options.insertDataOption || DEFAULTS.INSERT_DATA,
            includeValuesInResponse: !!op.options.includeValuesInResponse
          }));
          break;
        case KIND.STRUCT:
          if (!group || !compat(group, op)) { flush(); group = { type: KIND.STRUCT, options: {}, items: [] }; }
          group.items.push(op);
          break;
        default:
          throw new Error('Unknown op kind: ' + op.kind);
      }
    }
    flush();

    if (clearAfter) batch.ops.length = 0;
    return results;
  }

  /* ---------- Values ops (enqueue) ---------- */

  function addValues(batch, a1, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('values must be non-empty 2D');
    batch.ops.push({
      kind: KIND.UPDATE,
      range: GSRange.ensureSheetOnA1(a1, batch.ss),
      values: shaped,
      majorDimension: opts.majorDimension || DEFAULTS.MAJOR_DIM,
      options: {
        valueInputOption: opts.valueInputOption || DEFAULTS.VALUE_INPUT,
        includeValuesInResponse: !!opts.includeValuesInResponse
      }
    });
    return batch;
  }

  function addCell(batch, a1, value, opts = {}) {
    return addValues(batch, a1, [[value]], opts);
  }

  function clearValues(batch, a1) {
    batch.ops.push({ kind: KIND.CLEAR, range: GSRange.ensureSheetOnA1(a1, batch.ss) });
    return batch;
  }

  function append(batch, a1, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('append values must be non-empty 2D');
    batch.ops.push({
      kind: KIND.APPEND,
      range: GSRange.ensureSheetOnA1(a1, batch.ss),
      values: shaped,
      majorDimension: opts.majorDimension || DEFAULTS.MAJOR_DIM,
      options: {
        valueInputOption: opts.valueInputOption || DEFAULTS.VALUE_INPUT,
        insertDataOption: opts.insertDataOption || DEFAULTS.INSERT_DATA,
        includeValuesInResponse: !!opts.includeValuesInResponse
      }
    });
    return batch;
  }

  /* ---------- Structural ops (enqueue) ---------- */

  function insertRows(batch, startRow, nRows, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName(), inheritFromBefore = false } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { insertDimension: {
      range: { sheetId, dimension: DIMENSION.ROWS, startIndex: startRow, endIndex: startRow + nRows },
      inheritFromBefore
    }});
    return batch;
  }

  function insertColumns(batch, startCol, nCols, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName(), inheritFromBefore = false } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { insertDimension: {
      range: { sheetId, dimension: DIMENSION.COLUMNS, startIndex: startCol, endIndex: startCol + nCols },
      inheritFromBefore
    }});
    return batch;
  }

  function insertRangeShift(batch, a1OrGridRange, shiftDimension) {
    const grid = (typeof a1OrGridRange === 'string') ? _a1ToGridRange_(batch.ss, a1OrGridRange) : a1OrGridRange;
    _enqueueStruct(batch, { insertRange: { range: grid, shiftDimension } });
    return batch;
  }

  function deleteRows(batch, startRow, nRows, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName() } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { deleteDimension: {
      range: { sheetId, dimension: DIMENSION.ROWS, startIndex: startRow, endIndex: startRow + nRows }
    }});
    return batch;
  }

  function deleteColumns(batch, startCol, nCols, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName() } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { deleteDimension: {
      range: { sheetId, dimension: DIMENSION.COLUMNS, startIndex: startCol, endIndex: startCol + nCols }
    }});
    return batch;
  }

  function deleteRange(batch, a1OrGridRange, shiftDimension) {
    const grid = (typeof a1OrGridRange === 'string') ? _a1ToGridRange_(batch.ss, a1OrGridRange) : a1OrGridRange;
    _enqueueStruct(batch, { deleteRange: { range: grid, shiftDimension } });
    return batch;
  }

  function insertValuesPushDown(batch, a1TopLeft, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('values must be non-empty 2D');

    const ensured = GSRange.ensureSheetOnA1(a1TopLeft, batch.ss);
    const { sheetName, addrOnly } = GSRange.splitA1(ensured);
    const sh = sheetName ? batch.ss.getSheetByName(sheetName) : batch.ss.getActiveSheet();
    if (!sh) throw new Error('Sheet not found: ' + (sheetName || '(active)'));
    const sheet = sh.getName();
    const sheetId = sh.getSheetId();

    const topLeft = _a1ToGridRange_(batch.ss, sheet + '!' + addrOnly);
    const startRow = topLeft.startRowIndex;
    const H = shaped.length;

    insertRows(batch, startRow, H, { sheet, inheritFromBefore: true });
    addValues(batch, sheet + '!' + addrOnly, shaped, { valueInputOption: opts.valueInputOption || DEFAULTS.VALUE_INPUT });

    if (opts.removeFromBottom !== false) {
      const last = sh.getMaxRows();
      const delStart = Math.max(0, last - H);
      _enqueueStruct(batch, { deleteDimension: {
        range: { sheetId, dimension: DIMENSION.ROWS, startIndex: delStart, endIndex: last }
      }});
    }
    return batch;
  }

  /* ---------- Queued reads ---------- */

  function queueGetRanges(batch, ranges, opts = {}) {
    if (typeof ranges === "string") ranges = [ranges];
    const normRanges = [].concat(ranges || []).map(a1 => GSRange.ensureSheetOnA1(a1, batch.ss));
    if (!normRanges.length) return batch;
    const options = {
      valueRenderOption: opts.valueRenderOption || DEFAULTS.RENDER,
      dateTimeRenderOption: opts.dateTimeRenderOption || DEFAULTS.DATETIME
    };
    for (const a1 of normRanges) {
      batch.ops.push({ kind: KIND.GET, range: a1, options });
    }
    return batch;
  }

  // ---- Immediate (non-queued) batchGet (Advanced Sheets) ----
  function loadRangesNow(input, opts = {}) {
    const {
      render = 'raw',
      dateTime = DATETIME.SERIAL,
      trim = true,
      name = "UNKNOWN"
    } = opts;

    if (typeof input === "string") input = [[name,input]];
    const norm = _normalizeInputForBatchLoad_(input);
    if (!norm.items.length) return Array.isArray(input) ? [] : (input ? { ...input } : {});

    const ranges = norm.items.map(it => String(it.range));

    const res = Sheets.Spreadsheets.Values.batchGet(EDContext.context.ssid, {
      ranges,
      valueRenderOption: renderMode(render),
      dateTimeRenderOption: dateTime
    });

    const vrs = (res && res.valueRanges) || [];
    const withResults = norm.items.map((it, i) => {
      const vr = vrs[i] || {};
      const raw = vr.values || [];
      const values = trim ? raw.filter(r => r && r.length && String(r[0]).length > 0) : raw;
      return { name: it.name, range: vr.range || ranges[i], values, _idx: it._idx };
    });

    if (norm.kind === 'array') {
      return withResults.map(x => ({ name: x.name, range: x.range, values: x.values }));
    } else {
      const out = {};
      for (const k of Object.keys(input || {})) out[k] = { ...input[k] };
      for (const rec of withResults) {
        const prev = out[rec.name] || {};
        out[rec.name] = { ...prev, range: rec.range, values: rec.values };
      }
      return out;
    }
  }

  // ---- Immediate fallback loader (SpreadsheetApp only) ----
  function loadRangesNowFallback(input, opts = {}) {
    const { render = 'raw', trim = true, name = 'UNKNOWN' } = opts;
    if (typeof input === "string") input = [[name,input]];
    const norm = _normalizeInputForBatchLoad_(input);
    if (!norm.items.length) return Array.isArray(input) ? [] : (input ? { ...input } : {});

    const withResults = norm.items.map(it => {
      const rng = GSRange.resolveRange(GSRange.ensureSheetOnA1(it.range, EDContext.context.ss), { ss: EDContext.context.ss });
      const values =
        render === 'display' ? rng.getDisplayValues()
      : render === 'formula' ? (function(){ const f=rng.getFormulas(), v=rng.getValues(); return f.map((r,i)=>r.map((c,j)=>c || v[i][j])); })()
      : rng.getValues();
      const outVals = trim ? (values || []).filter(r => r && r.length && String(r[0]).length > 0) : (values || []);
      return { name: it.name, range: GSRange.a1FromRange(rng), values: outVals, _idx: it._idx };
    });

    if (norm.kind === 'array') {
      return withResults.map(x => ({ name: x.name, range: x.range, values: x.values }));
    } else {
      const out = {};
      for (const rec of withResults) out[rec.name] = { range: rec.range, values: rec.values };
      return out;
    }
  }

  /* ===================== Helpers ===================== */

  function _normalizeInputForBatchLoad_(input) {
    if (Array.isArray(input)) {
      const items = [];
      for (let i = 0; i < input.length; i++) {
        const it = input[i];
        if (Array.isArray(it))      items.push({ name: String(it[0]), range: String(it[1]), _idx: i });
        else if (it && typeof it === 'object' && 'name' in it && 'range' in it)
                                   items.push({ name: String(it.name), range: String(it.range), _idx: i });
        else throw new Error("Array form must be [{name, range}, ...] or [[name, range], ...]");
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
    throw new Error("loadRangesNow expects an Array or an Object with .range properties.");
  }

  function renderMode(mode) {
    switch ((mode || '').toLowerCase()) {
      case 'display': return RENDER.FORMATTED;
      case 'raw':     return RENDER.UNFORMATTED;
      case 'formula': return RENDER.FORMULA;
      default:        return RENDER.UNFORMATTED;
    }
  }

  function _resolveSpreadsheet_(spreadsheet) {
    if (!spreadsheet) return SpreadsheetApp.getActive();
    if (typeof spreadsheet === 'string') return SpreadsheetApp.openById(spreadsheet);
    return spreadsheet;
  }

  function _ensureSheet_(ss, nameOrId) {
    if (typeof nameOrId === 'number') return { sheetId: nameOrId };
    const sh = ss.getSheetByName(nameOrId);
    if (!sh) throw new Error('Sheet not found: ' + nameOrId);
    return { sheetId: sh.getSheetId() };
  }

  function _sheetById_(ss, id) {
    const sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) if (sheets[i].getSheetId() === id) return sheets[i];
    throw new Error('Sheet id not found: ' + id);
  }

  function _a1ToGridRange_(ss, a1) {
    const box = GSRange.parseBox(GSRange.ensureSheetOnA1(a1, ss));
    const sh = box.sheetName ? ss.getSheetByName(box.sheetName) : ss.getActiveSheet();
    if (!sh) throw new Error('Sheet not found: ' + (box.sheetName || '(active)'));
    const sheetId = sh.getSheetId();

    return {
      sheetId,
      startRowIndex: box.r1 - 1,
      endRowIndex:   (Number.isFinite(box.r2) ? box.r2 : sh.getMaxRows()),
      startColumnIndex: box.c1 - 1,
      endColumnIndex:   (Number.isFinite(box.c2) ? box.c2 : sh.getMaxColumns())
    };
  }

  function _enqueueStruct(batch, request) {
    batch.ops.push({ kind: KIND.STRUCT, requests: [request] });
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

  function _opsToSingleBatchUpdateRequests_(batch) {
    const requests = [];
    const ss = batch.ss;

    for (const op of batch.ops) {
      switch (op.kind) {
        case KIND.UPDATE: {
          const grid = _a1ToGridRange_(ss, op.range);
          const H = op.values.length, W = op.values[0].length;
          grid.endRowIndex = grid.startRowIndex + (op.majorDimension === MAJOR_DIM.COLS ? W : H);
          grid.endColumnIndex = grid.startColumnIndex + (op.majorDimension === MAJOR_DIM.COLS ? H : W);

          let needsNumFmt = false;
          const rows = op.values.map(r => ({
            values: r.map(v => {
              const cd = _jsToCellData_(v, {});
              if (cd.userEnteredFormat && cd.userEnteredFormat.numberFormat) needsNumFmt = true;
              return cd;
            })
          }));
          const fields = FIELDS.USER_VAL + (needsNumFmt ? (',' + FIELDS.USER_FMT_NUM) : '');
          requests.push({ updateCells: { range: grid, rows, fields } });
          break;
        }
        case KIND.CLEAR: {
          const grid = _a1ToGridRange_(ss, op.range);
          requests.push({ updateCells: {
            range: grid,
            rows: [{ values: [{}] }],
            fields: FIELDS.USER_VAL
          }});
          break;
        }
        case KIND.APPEND: {
          const box = GSRange.parseBox(op.range);
          const sh  = box.sheetName ? ss.getSheetByName(box.sheetName) : ss.getActiveSheet();
          const sheetId = sh.getSheetId();
          const H = op.values.length;
          const W = op.values[0].length;

          const maxRows = sh.getMaxRows();
          requests.push({ insertDimension: {
            range: { sheetId, dimension: DIMENSION.ROWS, startIndex: maxRows, endIndex: maxRows + H },
            inheritFromBefore: true
          }});
          const topLeft = _a1ToGridRange_(ss, op.range);
          const startRowIndex = maxRows;
          const startColumnIndex = topLeft.startColumnIndex;
          const grid = {
            sheetId,
            startRowIndex,
            endRowIndex: startRowIndex + H,
            startColumnIndex,
            endColumnIndex: startColumnIndex + W
          };
          let needsNumFmt = false;
          const rows = op.values.map(r => ({
            values: r.map(v => {
              const cd = _jsToCellData_(v, {});
              if (cd.userEnteredFormat && cd.userEnteredFormat.numberFormat) needsNumFmt = true;
              return cd;
            })
          }));
          const fields = FIELDS.USER_VAL + (needsNumFmt ? (',' + FIELDS.USER_FMT_NUM) : '');
          requests.push({ updateCells: { range: grid, rows, fields } });
          break;
        }
        case KIND.GET:
          // reads aren’t supported in batchUpdate
          break;
        case KIND.STRUCT:
          requests.push(...op.requests);
          break;
        default:
          throw new Error('Unknown op kind: ' + op.kind);
      }
    }
    return requests;
  }

  function _transpose_(m) {
    const H = m.length, W = m[0].length;
    const out = Array.from({ length: W }, () => Array(H));
    for (let r = 0; r < H; r++) for (let c = 0; c < W; c++) out[c][r] = m[r][c];
    return out;
  }

  /* Expose */
  return {
    // enums (exported so you can reuse in your code/tests if you want)
    MODE, KIND, VALUE_INPUT, RENDER, DATETIME, MAJOR_DIM, DIMENSION,

    newBatch,
    merge,
    commit,
    size,

    add: {
      values: addValues,
      cell: addCell,
      clear: clearValues,
      append: append
    },

    insert: {
      rows: insertRows,
      columns: insertColumns,
      range: insertRangeShift,
      values: insertValuesPushDown
    },
    remove: {
      rows: deleteRows,
      columns: deleteColumns,
      range: deleteRange
    },

    load: {
      ranges: queueGetRanges,
      rangesNow: loadRangesNow,
      rangesNowFallback: loadRangesNowFallback,
      renderMode: renderMode,
    }
  };
})();
