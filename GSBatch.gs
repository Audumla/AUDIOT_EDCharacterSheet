/** Namespace: GSBatch — Values-first batching + single-batchUpdate mode + push-down insert + SpreadsheetApp fallback. */
var GSBatch = (function () {
  /* ===================== Constants ===================== */
  const C = {
    // Modes
    MODE: {
      VALUES: 'values',               // merge Values.* where possible (plus struct groups)
      SINGLE: 'singleBatchUpdate',    // convert everything into 1 batchUpdate request
      SIMPLE: 'simple',               // SpreadsheetApp-only (no Advanced Sheets scopes)
    },

    // Op kinds
    KIND: {
      UPDATE: 'update',
      CLEAR:  'clear',
      GET:    'get',
      APPEND: 'append',
      STRUCT: 'struct',
    },

    // Render options
    RENDER: {
      UNFORMATTED: 'UNFORMATTED_VALUE',   // raw
      FORMATTED:   'FORMATTED_VALUE',     // display
      FORMULA:     'FORMULA',             // formulas
    },

    // Date/time render
    DTR: {
      SERIAL:    'SERIAL_NUMBER',
      FORMATTED: 'FORMATTED_STRING',
    },

    // Major dimension
    MAJOR_DIM: {
      ROWS: 'ROWS',
      COLS: 'COLUMNS',
    },

    // Dimension strings for struct ops
    DIM: {
      ROWS: 'ROWS',
      COLUMNS: 'COLUMNS',
    },

    // Shift dimension
    SHIFT: {
      ROWS: 'ROWS',
      COLUMNS: 'COLUMNS',
    },

    // ValueInputOption
    VIO: {
      RAW:  'RAW',
      USER: 'USER_ENTERED',
    },

    // InsertDataOption
    IDO: {
      INSERT_ROWS: 'INSERT_ROWS',
      OVERWRITE:   'OVERWRITE',
    }
  };

  let __batchID__ = 1;
  let defaultMode = C.MODE.VALUES;

  /* ===================== Public API ===================== */

  /**
   * Create a new batch.
   * @param {Spreadsheet|string} spreadsheet Spreadsheet or ID. Defaults to active.
   * @param {{mode?:'values'|'singleBatchUpdate'|'simple'}} [opts]
   */
  function newBatch(opts = {}) {
    const ss = EDContext.context.ss;
    const mode = opts.mode || defaultMode;
    return {
      ss,
      spreadsheetId: ss.getId(),
      ops: [],
      mode,
      batchID: __batchID__++,
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

  /**
   * Commit queued ops.
   * - 'values': merge Values.* calls where compatible; struct ops grouped per batchUpdate
   * - 'singleBatchUpdate': convert EVERYTHING into a single batchUpdate request (one HTTP call)
   * - 'simple': execute sequentially using SpreadsheetApp only (no Advanced Sheets scopes)
   */
  function commit(batch, opts = {}) {
    const { clearAfter = true } = opts;
    if (!batch || !batch.ops || !batch.ops.length) return [];

    if (batch.mode === C.MODE.SINGLE) {
      const requests = _opsToSingleBatchUpdateRequests_(batch);
      const res = Sheets.Spreadsheets.batchUpdate({ requests }, batch.spreadsheetId);
      if (clearAfter) batch.ops.length = 0;
      return [res];
    }

    if (batch.mode === C.MODE.SIMPLE) {
      const res = _commitSimple_(batch);
      if (clearAfter) batch.ops.length = 0;
      return res;
    }

    // ----- 'values' mode (merge by API, preserve order) -----
    const sid = batch.spreadsheetId;
    const results = [];
    let group = null;

    const flush = () => {
      try {
        if (!group) return;
        switch (group.type) {
          case C.KIND.UPDATE: {
            const data = group.items.map(it => ({
              range: it.range,
              values: it.values,
              majorDimension: it.majorDimension || C.MAJOR_DIM.ROWS
            }));
            const body = {
              valueInputOption: group.options.valueInputOption || C.VIO.USER,
              includeValuesInResponse: !!group.options.includeValuesInResponse,
              data
            };
            results.push(Sheets.Spreadsheets.Values.batchUpdate(body, sid));
            break;
          }
          case C.KIND.CLEAR: {
            const body = { ranges: group.items.map(it => it.range) };
            results.push(Sheets.Spreadsheets.Values.batchClear(body, sid));
            break;
          }
          case C.KIND.GET: {
            const body = {
              ranges: group.items.map(it => it.range),
              valueRenderOption: group.options.valueRenderOption || C.RENDER.UNFORMATTED,
              dateTimeRenderOption: group.options.dateTimeRenderOption || C.DTR.SERIAL
            };
            results.push(Sheets.Spreadsheets.Values.batchGet(sid, body));
            break;
          }
          case C.KIND.STRUCT: {
            const requests = group.items.flatMap(it => it.requests);
            if (requests.length) {
              results.push(Sheets.Spreadsheets.batchUpdate({ requests }, sid));
            }
            break;
          }
        }
        group = null;
      }
      catch (e) {
        EDLogger.error(`Batch write failed [${JSON.stringify(group)}]`)
        throw e;
      }
    };

    const compatible = (grp, op) => {
      if (!grp) return false;
      if (grp.type !== op.kind && !(grp.type === C.KIND.STRUCT && op.kind === C.KIND.STRUCT)) return false;
      switch (grp.type) {
        case C.KIND.UPDATE:
          return (grp.options.valueInputOption || C.VIO.USER) === (op.options.valueInputOption || C.VIO.USER) &&
                 !!grp.options.includeValuesInResponse === !!op.options.includeValuesInResponse;
        case C.KIND.CLEAR:
          return true;
        case C.KIND.GET:
          return (grp.options.valueRenderOption || C.RENDER.UNFORMATTED) === (op.options.valueRenderOption || C.RENDER.UNFORMATTED) &&
                 (grp.options.dateTimeRenderOption || C.DTR.SERIAL) === (op.options.dateTimeRenderOption || C.DTR.SERIAL);
        case C.KIND.STRUCT:
          return true;
        default:
          return false;
      }
    };

    for (const op of batch.ops) {
      switch (op.kind) {
        case C.KIND.UPDATE:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.UPDATE, options: op.options || {}, items: [] }; }
          group.items.push(op);
          break;
        case C.KIND.CLEAR:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.CLEAR, options: {}, items: [] }; }
          group.items.push(op);
          break;
        case C.KIND.GET:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.GET, options: op.options || {}, items: [] }; }
          group.items.push(op);
          break;
        case C.KIND.APPEND:
          flush(); // append must be per-call in Values API
          results.push(Sheets.Spreadsheets.Values.append({
            values: op.values,
            majorDimension: op.majorDimension || C.MAJOR_DIM.ROWS
          }, sid, op.range, {
            valueInputOption: op.options.valueInputOption || C.VIO.USER,
            insertDataOption: op.options.insertDataOption || C.IDO.INSERT_ROWS,
            includeValuesInResponse: !!op.options.includeValuesInResponse
          }));
          break;
        case C.KIND.STRUCT:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.STRUCT, options: {}, items: [] }; }
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

  /* ---------- Values ops ---------- */

  function addValues(batch, a1, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('values must be non-empty 2D');
    batch.ops.push({
      kind: C.KIND.UPDATE,
      range: _ensureA1_(batch.ss, a1),
      values: shaped,
      majorDimension: opts.majorDimension || C.MAJOR_DIM.ROWS,
      options: {
        valueInputOption: opts.valueInputOption || C.VIO.USER,
        includeValuesInResponse: !!opts.includeValuesInResponse
      }
    });
    return batch;
  }

  function addCell(batch, a1, value, opts = {}) {
    return addValues(batch, a1, [[value]], opts);
  }

  function clearValues(batch, a1) {
    batch.ops.push({ kind: C.KIND.CLEAR, range: _ensureA1_(batch.ss, a1) });
    return batch;
  }

  function append(batch, a1, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('append values must be non-empty 2D');
    batch.ops.push({
      kind: C.KIND.APPEND,
      range: _ensureA1_(batch.ss, a1),
      values: shaped,
      majorDimension: opts.majorDimension || C.MAJOR_DIM.ROWS,
      options: {
        valueInputOption: opts.valueInputOption || C.VIO.USER,
        insertDataOption: opts.insertDataOption || C.IDO.INSERT_ROWS,
        includeValuesInResponse: !!opts.includeValuesInResponse
      }
    });
    return batch;
  }

  /* ---------- Structural ops (batchUpdate) ---------- */

  function insertRows(batch, startRow0, nRows, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName(), inheritFromBefore = false } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { insertDimension: {
      range: { sheetId, dimension: C.DIM.ROWS, startIndex: startRow0, endIndex: startRow0 + nRows },
      inheritFromBefore
    }});
    return batch;
  }

  function insertColumns(batch, startCol0, nCols, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName(), inheritFromBefore = false } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { insertDimension: {
      range: { sheetId, dimension: C.DIM.COLUMNS, startIndex: startCol0, endIndex: startCol0 + nCols },
      inheritFromBefore
    }});
    return batch;
  }

  function insertRangeShift(batch, a1OrGridRange, shiftDimension) {
    const grid = (typeof a1OrGridRange === 'string') ? _a1ToGridRange_(batch.ss, a1OrGridRange) : a1OrGridRange;
    _enqueueStruct(batch, { insertRange: { range: grid, shiftDimension } });
    return batch;
  }

  function deleteRows(batch, startRow0, nRows, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName() } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { deleteDimension: {
      range: { sheetId, dimension: C.DIM.ROWS, startIndex: startRow0, endIndex: startRow0 + nRows }
    }});
    return batch;
  }

  function deleteColumns(batch, startCol0, nCols, opts = {}) {
    const { sheet = batch.ss.getActiveSheet().getName() } = opts;
    const { sheetId } = _ensureSheet_(batch.ss, sheet);
    _enqueueStruct(batch, { deleteDimension: {
      range: { sheetId, dimension: C.DIM.COLUMNS, startIndex: startCol0, endIndex: startCol0 + nCols }
    }});
    return batch;
  }

  function deleteRange(batch, a1OrGridRange, shiftDimension) {
    const grid = (typeof a1OrGridRange === 'string') ? _a1ToGridRange_(batch.ss, a1OrGridRange) : a1OrGridRange;
    _enqueueStruct(batch, { deleteRange: { range: grid, shiftDimension } });
    return batch;
  }

  /**
   * Insert values with push-down semantics for logging:
   *  - Insert H rows at target (push existing rows down)
   *  - Write values into the newly freed block
   *  - Delete H rows from the BOTTOM of the sheet (data there is lost)
   *
   * @param {object} batch
   * @param {string} a1TopLeft  e.g. "Log!A2"
   * @param {any[][]|any} values
   * @param {{valueInputOption?:'RAW'|'USER_ENTERED', removeFromBottom?:boolean}} opts
   */
  function insertValuesPushDown(batch, a1TopLeft, values, opts = {}) {
    const shaped = GSUtils.Arr.to2D(values);
    if (!shaped.length || !(shaped[0] || []).length) throw new Error('values must be non-empty 2D');

    const { sheetName, addrOnly } = GSRange.splitA1(a1TopLeft);
    const sh = sheetName ? batch.ss.getSheetByName(sheetName) : batch.ss.getActiveSheet();
    if (!sh) throw new Error('Sheet not found: ' + (sheetName || '(active)'));
    const sheet = sh.getName();
    a1TopLeft = sheet + '!' + addrOnly;

    const topLeft = _a1ToGridRange_(batch.ss, a1TopLeft); // single cell
    const startRow = topLeft.startRowIndex;
    const H = shaped.length;

    insertRows(batch, startRow, H, { sheet, inheritFromBefore: true });
    addValues(batch, a1TopLeft, shaped, { valueInputOption: opts.valueInputOption || C.VIO.USER });

    if (opts.removeFromBottom !== false) {
      const sheetId = sh.getSheetId();
      const last = sh.getMaxRows();
      const delStart = Math.max(0, last - H);
      _enqueueStruct(batch, { deleteDimension: {
        range: { sheetId, dimension: C.DIM.ROWS, startIndex: delStart, endIndex: last }
      }});
    }

    return batch;
  }

  /* ---------- Queued reads ---------- */

  /**
   * Queue one or more GETs. Options mirror Sheets API; SIMPLE mode will return a single
   * result containing `valueRanges` in the same order as queued.
   * You can pass `trimEmptyRows:true` to have SIMPLE mode drop rows whose first cell is blank.
   */
  function queueGetRanges(batch, ranges, opts = {}) {
    if (typeof ranges === "string") ranges = [ranges];
    const normRanges = [].concat(ranges || []).map(a1 => _ensureA1_(batch.ss, a1));
    if (!normRanges.length) return batch;
    const options = {
      valueRenderOption: opts.valueRenderOption || C.RENDER.UNFORMATTED,
      dateTimeRenderOption: opts.dateTimeRenderOption || C.DTR.SERIAL,
      trimEmptyRows: !!opts.trimEmptyRows
    };
    for (const a1 of normRanges) {
      batch.ops.push({ kind: C.KIND.GET, range: a1, options });
    }
    return batch;
  }

  // ---- Immediate (non-queued) batchGet ----
  function loadRangesNow(input, opts = {}) {
    const {
      render = 'raw',                // 'raw' | 'display' | 'formula'
      dateTime = 'SERIAL_NUMBER',    // 'SERIAL_NUMBER' | 'FORMATTED_STRING'
      trim = true,                   // drop rows whose first cell is empty-ish
      name = "UNKNOWN"
    } = opts;

    if (typeof input === "string") input = [[name, input]];
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
      const values = trim ? raw.filter(r => r && r.length && String(r[0]).trim() !== '') : raw;
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

  // Normalizes the same shapes you were using:
  //  - [{name, range}, ...]  or  [[name, range], ...]
  //  - { foo:{range: "A1"}, bar:{range:"B1"} }
  function _normalizeInputForBatchLoad_(input) {
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
    throw new Error("loadRangesNow expects an Array or an Object with .range properties.");
  }

  // Map our friendly render names to API flags
  function renderMode(mode) {
    switch ((mode || '').toLowerCase()) {
      case 'display': return C.RENDER.FORMATTED;
      case 'raw':     return C.RENDER.UNFORMATTED;
      case 'formula': return C.RENDER.FORMULA;
      default:        return C.RENDER.UNFORMATTED;
    }
  }

  /* ===================== Helpers (private) ===================== */

  function _resolveSpreadsheet_(spreadsheet) {
    if (!spreadsheet) return SpreadsheetApp.getActive();
    if (typeof spreadsheet === 'string') return SpreadsheetApp.openById(spreadsheet);
    return spreadsheet;
  }

  function _ensureA1_(ss, a1) { return GSRange.ensureSheetOnA1(String(a1), ss); }

  function _ensureSheet_(ss, nameOrId) {
    if (typeof nameOrId === 'number') return { sheetId: nameOrId };
    const sh = ss.getSheetByName(nameOrId);
    if (!sh) throw new Error('Sheet not found: ' + nameOrId);
    return { sheetId: sh.getSheetId() };
  }

  function _a1ToGridRange_(ss, a1) {
    const { sheetName, addrOnly } = GSRange.splitA1(a1);
    const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
    if (!sh) throw new Error('Sheet not found: ' + sheetName);
    const sheetId = sh.getSheetId();

    const s = String(addrOnly).replace(/\$/g, '').toUpperCase();
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
      return { sheetId, startRowIndex: r - 1, endRowIndex: r, startColumnIndex: c - 1, endColumnIndex: c };
    }
    throw new Error('Range must be a finite rectangle or single cell: ' + a1);
  }

  function _colToIndex_(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function _enqueueStruct(batch, request) {
    batch.ops.push({ kind: C.KIND.STRUCT, requests: [request] });
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

  // Convert entire ops queue into *one* batchUpdate requests array (order preserved).
  function _opsToSingleBatchUpdateRequests_(batch) {
    const requests = [];
    const ss = batch.ss;

    for (const op of batch.ops) {
      switch (op.kind) {
        case C.KIND.UPDATE: {
          const grid = _a1ToGridRange_(ss, op.range);
          const H = op.values.length, W = op.values[0].length;
          grid.endRowIndex = grid.startRowIndex + (op.majorDimension === C.MAJOR_DIM.COLS ? W : H);
          grid.endColumnIndex = grid.startColumnIndex + (op.majorDimension === C.MAJOR_DIM.COLS ? H : W);

          let needsNumFmt = false;
          const rows = op.values.map(r => ({
            values: r.map(v => {
              const cd = _jsToCellData_(v, {});
              if (cd.userEnteredFormat && cd.userEnteredFormat.numberFormat) needsNumFmt = true;
              return cd;
            })
          }));
          const fields = 'userEnteredValue' + (needsNumFmt ? ',userEnteredFormat.numberFormat' : '');
          requests.push({ updateCells: { range: grid, rows, fields } });
          break;
        }
        case C.KIND.CLEAR: {
          const grid = _a1ToGridRange_(ss, op.range);
          requests.push({ updateCells: {
            range: grid,
            rows: [{ values: [{}] }],
            fields: 'userEnteredValue'
          }});
          break;
        }
        case C.KIND.APPEND: {
          const { sheetName } = GSRange.splitA1(op.range);
          const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
          const sheetId = sh.getSheetId();
          const H = op.values.length;
          const W = op.values[0].length;

          const maxRows = sh.getMaxRows();
          requests.push({ insertDimension: {
            range: { sheetId, dimension: C.DIM.ROWS, startIndex: maxRows, endIndex: maxRows + H },
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
          const fields = 'userEnteredValue' + (needsNumFmt ? ',userEnteredFormat.numberFormat' : '');
          requests.push({ updateCells: { range: grid, rows, fields } });
          break;
        }
        case C.KIND.GET: {
          // Reads are not supported in a single batchUpdate; skip here.
          break;
        }
        case C.KIND.STRUCT: {
          requests.push(...op.requests);
          break;
        }
        default:
          throw new Error('Unknown op kind: ' + op.kind);
      }
    }
    return requests;
  }

  /* ===================== SIMPLE mode executor (SpreadsheetApp only) ===================== */
  /**
   * Mirrors the grouping and **return shapes** of the Google APIs
   * so client code can be completely mode-agnostic.
   *
   * Returns an array of “API-like responses” in the same order that
   * the 'values' mode would have emitted them (groups + per-append).
   */
  function _commitSimple_(batch) {
    const ss = batch.ss;
    const sid = batch.spreadsheetId;

    // Group ops exactly like 'values' mode so the response array lines up
    const out = [];
    let group = null;

    const flush = () => {
      if (!group) return;

      switch (group.type) {
        case C.KIND.UPDATE: {
          // Simulate Values.batchUpdate: execute writes, build BatchUpdateValuesResponse
          let totalRows = 0, totalCols = 0, totalCells = 0;
          const responses = [];

          for (const it of group.items) {
            const info = _saWriteValues_(ss, it.range, it.values, it.majorDimension || C.MAJOR_DIM.ROWS);
            totalRows  += info.rows;
            totalCols  += info.cols;
            totalCells += info.rows * info.cols;
            responses.push({
              spreadsheetId: sid,
              updatedRange: info.range,
              updatedRows: info.rows,
              updatedColumns: info.cols,
              updatedCells: info.rows * info.cols
            });
          }

          out.push({
            spreadsheetId: sid,
            totalUpdatedRows: totalRows,
            totalUpdatedColumns: totalCols,
            totalUpdatedCells: totalCells,
            responses
          });
          break;
        }

        case C.KIND.CLEAR: {
          // Simulate Values.batchClear
          const cleared = [];
          for (const it of group.items) {
            const a1Full = GSRange.ensureSheetOnA1(it.range, ss);
            GSRange.resolveRange(a1Full, { ss }).clearContent();
            cleared.push(a1Full);
          }
          out.push({ spreadsheetId: sid, clearedRanges: cleared });
          break;
        }

        case C.KIND.GET: {
          // Simulate Values.batchGet
          const valueRanges = [];
          for (const it of group.items) {
            const a1Full = GSRange.ensureSheetOnA1(it.range, ss);
            const vr = _saReadValues_(ss, a1Full, it.options && it.options.valueRenderOption);
            let vals = vr.values || [];
            if (it.options && it.options.trimEmptyRows) {
              vals = vals.filter(row => row && row.length && String(row[0]).trim() !== '');
            }
            valueRanges.push({ range: a1Full, values: vals });
          }
          out.push({ valueRanges });
          break;
        }

        case C.KIND.STRUCT: {
          // Simulate Spreadsheets.batchUpdate replies (best-effort)
          const replies = [];
          for (const it of group.items) {
            for (const req of (it.requests || [])) {
              replies.push(_saDoStructRequest_(ss, req));
            }
          }
          out.push({ replies });
          break;
        }
      }
      group = null;
    };

    const compatible = (grp, op) => {
      if (!grp) return false;
      if (grp.type !== op.kind && !(grp.type === C.KIND.STRUCT && op.kind === C.KIND.STRUCT)) return false;
      switch (grp.type) {
        case C.KIND.UPDATE:
          return (grp.options.valueInputOption || C.VIO.USER) === (op.options.valueInputOption || C.VIO.USER) &&
                 !!grp.options.includeValuesInResponse === !!op.options.includeValuesInResponse;
        case C.KIND.CLEAR:
          return true;
        case C.KIND.GET:
          return (grp.options.valueRenderOption || C.RENDER.UNFORMATTED) === (op.options.valueRenderOption || C.RENDER.UNFORMATTED) &&
                 (grp.options.dateTimeRenderOption || C.DTR.SERIAL) === (op.options.dateTimeRenderOption || C.DTR.SERIAL) &&
                 (!!grp.options.trimEmptyRows) === (!!op.options.trimEmptyRows);
        case C.KIND.STRUCT:
          return true;
        default:
          return false;
      }
    };

    for (const op of batch.ops) {
      switch (op.kind) {
        case C.KIND.UPDATE:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.UPDATE, options: op.options || {}, items: [] }; }
          group.items.push(op);
          break;

        case C.KIND.CLEAR:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.CLEAR, options: {}, items: [] }; }
          group.items.push(op);
          break;

        case C.KIND.GET:
          if (!group || !compatible(group, op)) {
            flush();
            group = { type: C.KIND.GET, options: op.options || {}, items: [] };
          }
          group.items.push(op);
          break;

        case C.KIND.APPEND:
          // Append is per-call in Values API — execute and push a single append-like response
          flush();
          {
            const info = _saAppendValues_(ss, op.range, op.values);
            out.push({
              updates: {
                spreadsheetId: sid,
                updatedRange: info.range,
                updatedRows: info.rows,
                updatedColumns: info.cols,
                updatedCells: info.rows * info.cols
              }
            });
          }
          break;

        case C.KIND.STRUCT:
          if (!group || !compatible(group, op)) { flush(); group = { type: C.KIND.STRUCT, options: {}, items: [] }; }
          group.items.push(op);
          break;

        default:
          throw new Error('Unknown op kind in simple mode: ' + op.kind);
      }
    }
    flush();

    return out;
  }

  // === SIMPLE helpers (SpreadsheetApp) ===
  function _saWriteValues_(ss, a1, values, majorDim) {
    const a1Full = GSRange.ensureSheetOnA1(a1, ss);
    const rng = GSRange.resolveRange(a1Full, { ss });
    const arr2d = GSUtils.Arr.to2D(values);

    let rows = arr2d.length, cols = arr2d[0].length;

    if (majorDim === C.MAJOR_DIM.COLS) {
      // transpose to rows for setValues
      const H = arr2d.length, W = arr2d[0].length;
      const out = Array.from({ length: W }, () => Array(H));
      for (let r = 0; r < H; r++) for (let c = 0; c < W; c++) out[c][r] = arr2d[r][c];
      rows = out.length; cols = out[0].length;
      rng.offset(0, 0, rows, cols).setValues(out);
    } else {
      rng.offset(0, 0, rows, cols).setValues(arr2d);
    }
    return { range: a1Full, rows, cols };
  }

  function _saReadValues_(ss, a1Full, valueRenderOption) {
    const rng = GSRange.resolveRange(a1Full, { ss });
    switch (valueRenderOption || C.RENDER.UNFORMATTED) {
      case C.RENDER.FORMATTED: return { range: a1Full, values: rng.getDisplayValues() };
      case C.RENDER.FORMULA: {
        const f = rng.getFormulas(), v = rng.getValues();
        const values = f.map((row, i) => row.map((cell, j) => cell || v[i][j]));
        return { range: a1Full, values };
      }
      default:
        return { range: a1Full, values: rng.getValues() };
    }
  }

  function _saAppendValues_(ss, a1, values) {
    const a1Full = GSRange.ensureSheetOnA1(a1, ss);
    const { sheetName } = GSRange.splitA1(a1Full);
    const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();

    const box = GSRange.parseBox(a1Full);
    const anchorCol = box.c1;

    const lastRow = sh.getLastRow();
    const startRow = Math.max(lastRow + 1, 1);
    const arr2d = GSUtils.Arr.to2D(values);
    const H = arr2d.length, W = arr2d[0].length;

    const target = sh.getRange(startRow, anchorCol, H, W);
    target.setValues(arr2d);
    const updatedRange = sh.getName() + '!' + target.getA1Notation();
    return { range: updatedRange, rows: H, cols: W };
  }

  function _saDoStructRequest_(ss, req) {
    // Returns a "reply" object vaguely matching Sheets batchUpdate replies
    if (req.insertDimension) {
      const { sheetId, dimension, startIndex, endIndex } = req.insertDimension.range;
      const sh = ss.getSheets().find(s => s.getSheetId() === sheetId);
      const count = Math.max(0, (endIndex|0) - (startIndex|0));
      if (dimension === C.DIM.ROWS) sh.insertRows(startIndex + 1, count);
      else                          sh.insertColumns(startIndex + 1, count);
      return { insertDimension: { sheetId, dimension, startIndex, endIndex } };
    }
    if (req.deleteDimension) {
      const { sheetId, dimension, startIndex, endIndex } = req.deleteDimension.range;
      const sh = ss.getSheets().find(s => s.getSheetId() === sheetId);
      const count = Math.max(0, (endIndex|0) - (startIndex|0));
      if (dimension === C.DIM.ROWS) sh.deleteRows(startIndex + 1, count);
      else                          sh.deleteColumns(startIndex + 1, count);
      return { deleteDimension: { sheetId, dimension, startIndex, endIndex } };
    }
    if (req.insertRange) {
      const g = req.insertRange.range;
      const sh = ss.getSheets().find(s => s.getSheetId() === g.sheetId);
      const a1 = sh.getRange(g.startRowIndex + 1, g.startColumnIndex + 1, (g.endRowIndex - g.startRowIndex), (g.endColumnIndex - g.startColumnIndex)).getA1Notation();
      const full = sh.getName() + '!' + a1;
      const dim = (req.insertRange.shiftDimension === C.SHIFT.ROWS) ? SpreadsheetApp.Dimension.ROWS : SpreadsheetApp.Dimension.COLUMNS;
      GSRange.resolveRange(full, { ss }).insertCells(dim);
      return { insertRange: { range: g, shiftDimension: req.insertRange.shiftDimension } };
    }
    if (req.deleteRange) {
      const g = req.deleteRange.range;
      const sh = ss.getSheets().find(s => s.getSheetId() === g.sheetId);
      const a1 = sh.getRange(g.startRowIndex + 1, g.startColumnIndex + 1, (g.endRowIndex - g.startRowIndex), (g.endColumnIndex - g.startColumnIndex)).getA1Notation();
      const full = sh.getName() + '!' + a1;
      const dim = (req.deleteRange.shiftDimension === C.SHIFT.ROWS) ? SpreadsheetApp.Dimension.ROWS : SpreadsheetApp.Dimension.COLUMNS;
      GSRange.resolveRange(full, { ss }).deleteCells(dim);
      return { deleteRange: { range: g, shiftDimension: req.deleteRange.shiftDimension } };
    }
    if (req.updateCells) {
      const g = req.updateCells.range;
      const sh = ss.getSheets().find(s => s.getSheetId() === g.sheetId);
      const H = (g.endRowIndex - g.startRowIndex);
      const W = (g.endColumnIndex - g.startColumnIndex);
      const rows = (req.updateCells.rows || []);
      const vals = Array.from({ length: H }, (_, r) =>
        Array.from({ length: W }, (_, c) => {
          const cd = (rows[r] && rows[r].values && rows[r].values[c]) || {};
          const v  = (cd.userEnteredValue || {});
          if (v.formulaValue != null) return String(v.formulaValue);
          if (v.numberValue  != null) return Number(v.numberValue);
          if (v.boolValue    != null) return !!v.boolValue;
          if (v.stringValue  != null) return String(v.stringValue);
          return null;
        })
      );
      sh.getRange(g.startRowIndex + 1, g.startColumnIndex + 1, H, W).setValues(vals);
      return { updateCells: { updatedRows: H, updatedColumns: W } };
    }
    // Unrecognized request — return empty reply to preserve index
    return {};
  }

  /* Expose */
  return {
    // constants (re-export for consumers)
    MODE: C.MODE,
    KIND: C.KIND,
    RENDER: C.RENDER,
    DTR: C.DTR,
    MAJOR_DIM: C.MAJOR_DIM,
    DIM: C.DIM,
    SHIFT: C.SHIFT,
    VIO: C.VIO,
    IDO: C.IDO,

    newBatch,
    merge,
    commit,
    size,
    defaultMode,  // readable default; set when calling newBatch({mode})

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
      renderMode: renderMode,
    }
  };
})();
