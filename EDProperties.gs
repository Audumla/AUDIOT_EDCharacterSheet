

var EDProperties = (function () {

  // Non-enumerable key to hold tracking data on target objects
  const __TRACK__KEY = '__tracking__'; // { paths:Set<string>, baseline: Map<string, any> }

  /**
   * Populate data from rows.
   *
   * Modes:
   *  - "pair"   : legacy [name, value] → nested object leaves as simple values.
   *  - "object" : build nested object leaves from headers (dynamic keys, any #cols).
   *               The path column is chosen by nameHeader/nameCol and is ALSO kept
   *               in the leaf object (e.g., { property:"A|B|C", ... }).
   *  - "array"  : return an array of row objects using headers as keys
   *               (no nesting, tracking ignored).
   *
   * Common options:
   *  - headers        : string[] of header names (optional)
   *  - headersRow     : number (0-based) row index to read headers from (optional)
   *  - nameHeader     : header string to use as the path (object mode)
   *  - nameCol        : number (0-based) column index for the path (object mode)
   *  - prefix         : string to prepend to path (object/pair) 
   *  - track          : boolean to snapshot created leaf paths (pair/object)
   *  - name           : the name of the property within the target to store the unpacked results under
   */
  function unpack(target, rows, opts = {}) {
    if (!Array.isArray(rows)) throw new Error('rows must be a 2D array like [[name, ...], ...]');
    const mode = (opts.mode || 'pair').toLowerCase();
    const track = !!opts.track;
    const prefix = opts?.prefix;
    //const name = opts?.name ?? undefined;
    const id = opts?.id ?? "UNKNOWN";

    // ---------- helpers ----------
    const splitPath = (s) =>
      String(s).split(/\||\./).map(t => t.trim()).filter(Boolean).map(GSUtils.Obj.safePropName);

    const ensurePath = (root, parts, defaultLeaf = {}) => {
      if (!root || typeof root !== 'object') {
        throw new Error('unpack: target must be an object');
      }
      if (!Array.isArray(parts) || !parts.length) {
        return { node: root, leafKey: undefined };
      }

      let node = root;

      // create/normalize intermediates (everything except the final key)
      for (let i = 0; i < parts.length - 1; i++) {
        const k = parts[i];
        const cur = node[k];
        if (cur == null) {
          node[k] = {};
        } else if (typeof cur !== 'object' || Array.isArray(cur)) {
          node[k] = { __value: cur };
        }
        node = node[k];
      }

      // ensure/normalize the final leaf
      const leafKey = parts[parts.length - 1];
      const curLeaf = node[leafKey];

      if (curLeaf == null) {
        node[leafKey] = defaultLeaf; // use the provided default directly
      } 

      return { node : node[leafKey], leafKey };
    };


    const setAtPath = (root, pathStr, val) => {
      const parts = splitPath(pathStr);
      if (!parts.length) return val;
      const { node, leafKey } = ensurePath(root, parts);
      node[leafKey] = val;
      return val;
    };

    EDLogger.trace(`Unpacking [ ${id} ]${prefix ? `[ prefix: ${prefix} ]` : ``}[ mode: ${mode} ][ Rows: ${rows.length}]`);

    if (mode === 'array') {
      const { headers, dataStart } = _resolveHeaders(rows, opts);
      if (!headers || !headers.length) return [];

      const out = ensurePath(target,splitPath(prefix ? prefix : id),[]).node;
      for (let r = dataStart; r < rows.length; r++) {
        const row = rows[r] || [];
        const obj = {};
        for (let c = 0; c < headers.length; c++) {
          const keyRaw = headers[c];
          if (keyRaw == null || keyRaw === '') continue;
          const key = GSUtils.Obj.safePropName(String(keyRaw).trim());
          const val = (row.length > c) ? row[c] : null;
          obj[key] = GSUtils.Str.coerce(val);
        }
        if (Object.keys(obj).length) out.push(obj);
      }

      if (track) {
        // Track by row index ("0","1",...)
        const paths = out.map((_, i) => String(i));
        _attachTrackingBaseline_(out, paths);
      }
 
      return out;
    }

    // --- object & pair modes (nested assignment into target) ---
    if (!target || typeof target !== 'object') throw new Error('target must be an object');

    const createdPaths = [];

    if (mode === 'object') {

      const { headers, dataStart } = _resolveHeaders(rows, opts);
      if (!headers || !headers.length) return target;
      const nameCol = _resolveNameCol(headers, opts);
      if (nameCol < 0 || nameCol >= headers.length) throw new Error('Invalid name column for object mode');

      for (let r = dataStart; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;

        const rawName = row[nameCol];
        if (rawName == null || String(rawName).trim() === '') continue;

        let pathStr = String(rawName).trim();
        pathStr = prefix ? (prefix + "." + pathStr) : pathStr;

        const partsRaw = String(pathStr).split(/\||\./).map(s => s.trim()).filter(Boolean);
        if (!partsRaw.length) continue;

        const parts = partsRaw.map(GSUtils.Obj.safePropName);

        // Walk/create intermediate nodes
        let node = target;
        for (let p = 0; p < parts.length - 1; p++) {
          const key = parts[p];
          const cur = node[key];
          if (cur == null) node[key] = {};
          else if (typeof cur !== 'object' || Array.isArray(cur)) node[key] = { __value: cur };
          node = node[key];
        }
        const leafKey = parts[parts.length - 1];

        // Build leaf object from ALL headers/columns (including the path column)
        const leaf = {};
        for (let c = 0; c < headers.length; c++) {
          const headerRaw = headers[c];
          if (headerRaw == null || headerRaw === '') continue;
          const key = GSUtils.Obj.safePropName(String(headerRaw).trim());
          const val = (row.length > c) ? row[c] : null;
          leaf[key] = GSUtils.Str.coerce(val);
        }

        node[leafKey] = leaf; // overwrite
        if (track) createdPaths.push(pathStr);

      }
    } else {
      // mode: 'pair'
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const nameRaw = row[0];
        if (nameRaw == null) continue;

        let name = String(nameRaw).trim();
        if (!name) continue;
        name = prefix ? prefix + "." + name : name;

        const partsRaw = name.split(/\||\./).map(s => s.trim()).filter(Boolean);
        if (!partsRaw.length) continue;

        const parts = partsRaw.map(GSUtils.Obj.safePropName);

        // Walk/create intermediate nodes
        let node = target;
        for (let p = 0; p < parts.length - 1; p++) {
          const key = parts[p];
          const cur = node[key];
          if (cur == null) node[key] = {};
          else if (typeof cur !== 'object' || Array.isArray(cur)) node[key] = { __value: cur };
          node = node[key];
        }

        const leafKey = parts[parts.length - 1];

        const simpleValue = GSUtils.Str.coerce(row.length > 1 ? row[1] : null);
        node[leafKey] = simpleValue; // overwrite
        if (track) createdPaths.push(name);
      }
    }

    if (track) _attachTrackingBaseline_(target, createdPaths);
    return target;
  }

  /** Returns true if any tracked leaf (object or array) has changed from its baseline. */
  function isTrackedModified(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return false;

    if (Array.isArray(root)) {
      for (const p of tr.paths) {
        const idx = Number(p);
        const cur = root[idx];
        if (!GSUtils.Obj.deepEqualSimple(tr.baseline.get(p), cur)) return true;
      }
      return false;
    }

    // object path lookup
    for (const p of tr.paths) {
      if (!GSUtils.Obj.deepEqualSimple(
        tr.baseline.get(p),
        getAtPath(root, p, { sep: '|', transform: GSUtils.Obj.safePropName })
      )) return true;
    }
    return false;
  }

  /** Returns [{ property, before, after }] for tracked leaves that differ (object or array roots). */
  function getTrackedChanges(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return [];
    const out = [];

    if (Array.isArray(root)) {
      for (const p of tr.paths) {
        const idx = Number(p);
        const before = tr.baseline.get(p);
        const after = root[idx];
        if (!GSUtils.Obj.deepEqualSimple(before, after)) {
          out.push({ property: p, before, after });
        }
      }
      return out;
    }

    for (const p of tr.paths) {
      const before = tr.baseline.get(p);
      const after  = getAtPath(root, p, { sep: '|', transform: GSUtils.Obj.safePropName });
      if (!GSUtils.Obj.deepEqualSimple(before, after)) {
        out.push({ property: p, before, after });
      }
    }
    return out;
  }

  /** Make the current values the new baseline (acknowledge/commit). Works for arrays too. */
  function commitTrackedBaseline(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return;

    if (Array.isArray(root)) {
      for (const p of tr.paths) {
        const idx = Number(p);
        tr.baseline.set(p, GSUtils.Obj.deepCloneSimple(root[idx]));
      }
      return;
    }

    for (const p of tr.paths) {
      tr.baseline.set(p, GSUtils.Obj.deepCloneSimple(
        getAtPath(root, p, { sep: '|', transform: GSUtils.Obj.safePropName })
      ));
    }
  }


  // --- helpers for header resolution ---

  function _resolveHeaders(rows, opts) {
    if (Array.isArray(opts.headers) && opts.headers.length) {
      return { headers: opts.headers.map(h => (h == null ? '' : String(h))), dataStart: 0 };
    }
    if (Number.isInteger(opts.headersRow)) {
      const idx = opts.headersRow;
      const hdr = rows[idx] || [];
      const headers = hdr.map(h => (h == null ? '' : String(h)));
      return { headers, dataStart: idx + 1 };
    }
    const headers = (rows[0] || []).map(h => (h == null ? '' : String(h)));
    return { headers, dataStart: 1 };
  }

  function _resolveNameCol(headers, opts) {
    if (Number.isInteger(opts.nameCol)) return opts.nameCol;
    if (typeof opts.nameHeader === 'string' && opts.nameHeader.trim() !== '') {
      const wanted = opts.nameHeader.trim().toLowerCase();
      const idx = headers.findIndex(h => String(h).trim().toLowerCase() === wanted);
      if (idx >= 0) return idx;
    }
    const fallbackIdx = headers.findIndex(h => /^(name|property|path)$/i.test(String(h).trim()));
    return fallbackIdx >= 0 ? fallbackIdx : 0;
  }

  /** Start/refresh tracking for given leaf paths: snapshot current values as baseline. */
  function _attachTrackingBaseline_(root, paths) {
    // Ensure tracking container exists
    let tr = Object.prototype.hasOwnProperty.call(root, __TRACK__KEY) ? root[__TRACK__KEY] : null;
    if (!tr) {
      tr = { paths: new Set(), baseline: new Map() };
      Object.defineProperty(root, __TRACK__KEY, { value: tr, enumerable: false, configurable: true, writable: true });
    }
    // Add paths and snapshot their current values
    for (const p of paths) tr.paths.add(p);
    for (const p of paths) tr.baseline.set(p, GSUtils.Obj.deepCloneSimple(getAtPath(root, p, { sep: '|', transform: GSUtils.Obj.safePropName })));
  }


  /** Stop tracking entirely (removes the hidden property). */
  function clearTracking(root) {
    if (root && Object.prototype.hasOwnProperty.call(root, __TRACK__KEY)) {
      delete root[__TRACK__KEY];
    }
  }

  // -------- Path helpers --------
  /**
   * Get a value from object using a separator-delimited path.
   * By default uses '|' separator and `safePropName` segment transform.
   */
  function getAtPath(obj, path, opts = {}) {
    const xform = opts.transform || GSUtils.Obj.safePropName;
    const parts = String(path).split(/\||\./).map(s => s.trim()).filter(Boolean);
    let cur = obj;
    for (const p of parts) {
      if (cur == null) return undefined;
      cur = cur[xform(p)];
    }
    return cur;
  }  

  /**
   * Reverse of unpack: recreate a 2-D values array from live data in the context.
   *
   * @param {object} spec
   *   Either a definition object (with fields like { unpack, values, name, prefix })
   *   or a plain spec: { mode, template, name, prefix }
   *   - mode      : EDProperties.path.UNPACK.* ('pair'|'array'|'object'|'none')
   *   - template  : 2-D array used as the shape/reference (defaults to spec.values)
   *   - name      : path/name where unpack originally landed (fallback to spec.prefix || spec.name)
   *   - prefix    : legacy prefix (considered when name missing)
   * @param {object} [opts]
   *   - target?: object                 default = EDContext.context
   *   - includeExtras?: boolean         default = false (only keys/cols from template)
   *   - objectValueCol?: number         1-based index of the "value" column (object mode)
   *   - arrayHeaderRow?: boolean        force treat first row as header (array mode). Default: auto.
   * @returns {any[][]} 2-D array suitable for writing back to the sheet
   */
  function repack(spec, opts = {}) {
  if (!spec || typeof spec !== 'object') throw new Error('repack: spec/definition required');

  const target   = opts.target || EDContext.context;
  const modeRaw  = (spec.mode || spec.unpack || 'pair');
  const mode     = String(modeRaw).toLowerCase(); // accept 'PAIR' | 'ARRAY' | 'OBJECT'
  const sourcePath = spec.prefix || '';           // unpack writes under `prefix`
  const template = _ensure2DArray(spec.template || spec.values || []);

  // Where the data lives now (same place unpack wrote it)
  const liveRoot = sourcePath
    ? getAtPath(target, sourcePath, { sep: '|', transform: GSUtils.Obj.safePropName })
    : target;

  switch (mode) {
    case 'pair':   return _repackPair(template, liveRoot, !!opts.includeExtras);
    case 'array':  return _repackArray(template, liveRoot, !!opts.includeExtras, opts.arrayHeaderRow);
    case 'object': return _repackObject(template, liveRoot);
    case 'none':   return template.slice();
    default:       return template.slice();
  }
}
  /**
   * Repack many specs/defs at once. Each item may be a def or {mode,template,name,prefix}.
   * @param {Array<object>} specs
   * @param {object} [opts] same as repack(...)
   * @returns {Array<{name:string, rows:any[][]}>}
   */
  function repackMany(specs, opts = {}) {
    if (!Array.isArray(specs)) throw new Error('repackMany: specs must be an array');
    const out = [];
    for (const s of specs) {
      const nm = s?.name || s?.prefix || '';
      const rows = repack(s, opts);
      out.push({ name: nm, rows });
    }
    return out;
  }

  /* =================== per-mode reverse builders =================== */

  // pair: template Nx2 (key,value). If includeExtras=false, only keys from col1 of template.
  function _repackPair(template, liveObj, includeExtras) {
    const keysFromTemplate = template.map(r => String((r && r[0]) ?? '')).filter(Boolean);
    const obj = (liveObj && typeof liveObj === 'object' && !Array.isArray(liveObj)) ? liveObj : {};
    const keys = includeExtras ? _unique([...keysFromTemplate, ...Object.keys(obj)]) : keysFromTemplate;

    const out = [];

    for (const k of keys) {
      out.push([k, getAtPath(obj,k)])
    };
    return out;
  }

  // array:
  //  - If live is array-of-arrays → return it (trim to template width if includeExtras=false).
  //  - If live is array-of-objects → build rows by columns (header from template if present/forced).
  function _repackArray(template, live, includeExtras, arrayHeaderRowOpt) {
    const hasTpl = template.length > 0;
    const tplHeaderIsStrings = hasTpl && _rowIsAllStrings(template[0]);
    const treatHeader = (arrayHeaderRowOpt == null) ? tplHeaderIsStrings : !!arrayHeaderRowOpt;

    const templateCols = (treatHeader && hasTpl) ? template[0].map(String) : [];
    // live array-of-arrays:
    if (Array.isArray(live) && live.length && Array.isArray(live[0])) {
      if (includeExtras) return live.slice();
      // trim width to template header cols (or template first row width if no header)
      const width = templateCols.length || (template[0] ? template[0].length : 0);
      if (!width) return live.slice();
      return live.map(r => r.slice(0, width));
    }

    // live array-of-objects (or empty)
    const arr = Array.isArray(live) ? live : [];
    const objs = arr.filter(x => x && typeof x === 'object' && !Array.isArray(x));

    let cols = [];
    if (!includeExtras && templateCols.length) {
      cols = templateCols.slice();
    } else if (includeExtras) {
      const set = new Set();
      for (const o of objs) for (const k of Object.keys(o)) set.add(k);
      cols = Array.from(set);
      if (!cols.length && templateCols.length) cols = templateCols.slice();
    } else {
      cols = (objs[0] && Object.keys(objs[0])) || templateCols.slice();
    }

    const out = [];
    if (treatHeader && cols.length) out.push(cols.slice());
    for (const o of objs) out.push(cols.map(c => o[c]));
    return out;
  }

/**
 * Object mode repack:
 * - First row treated as headers if all strings (copied through unchanged).
 * - Column 1 is the path/key (relative to `prefix`).
 * - For columns 2..N, pull values from the leaf object at that path using
 *   header names → safePropName(header) as keys (mirrors unpack(object)).
 */
function _repackObject(template, liveRoot) {
  const hasTpl = template.length > 0;
  const headerIsStrings = hasTpl && _rowIsAllStrings(template[0]);
  const headers = headerIsStrings ? template[0].map(h => (h == null ? '' : String(h))) : [];
  const width   = hasTpl ? (template[0] ? template[0].length : 0) : 0;

  const out = [];

  // Copy header row through unchanged
  let startRow = 0;
  if (headerIsStrings) {
    out.push(template[0].slice());
    startRow = 1;
  }

  for (let r = startRow; r < template.length; r++) {
    const row = template[r] || [];
    const keyPath = String((row[0] ?? '')).trim();
    // If no key in col1, just echo the row back
    if (!keyPath) { out.push(row.slice()); continue; }

    // Look up the leaf object at that key (relative to prefix)
    const leaf = getAtPath(liveRoot ?? EDContext.context, keyPath, {
      sep: '|',
      transform: GSUtils.Obj.safePropName
    });

    // Rebuild row: col1 stays the key; cols 2..N map from leaf by header
    const newRow = row.slice(0, width || row.length);
    if (headerIsStrings && headers.length > 1 && leaf && typeof leaf === 'object') {
      for (let c = 1; c < headers.length; c++) {
        const hdr = headers[c];
        if (!hdr) { newRow[c] = undefined; continue; }
        const prop = GSUtils.Obj.safePropName(hdr);
        newRow[c] = leaf[prop];
      }
    } else if (!headerIsStrings) {
      // No explicit headers: keep whatever was in template beyond col1
      // (you can extend here if you want a heuristic)
    }

    out.push(newRow);
  }

  return out;
}

  /* ========================== small helpers ========================== */

  function _ensure2DArray(v) {
    if (v == null) return [];
    if (Array.isArray(v) && (v.length === 0 || Array.isArray(v[0]))) return v;
    return [Array.isArray(v) ? v : [v]];
  }

  function _rowIsAllStrings(row) {
    if (!row || !row.length) return false;
    for (const c of row) if (c != null && typeof c !== 'string') return false;
    return true;
  }

  function _unique(arr) { const s = new Set(); const o = []; for (const x of arr) if (!s.has(x)) { s.add(x); o.push(x);} return o; }

  function _autoValueCol(template, overrideIdx) {
    if (overrideIdx && overrideIdx > 0) return overrideIdx;
    const w = template[0] ? template[0].length : 0;
    if (w >= 2) return 2;
    if (w >= 1) return 1;
    return 1;
  }


  function mappingByCell(loc, opts = {}) {
    let monitored = EDContext.context.event.mappings.find(e => e.cell == loc);
    if (monitored) {
      EDLogger.info(`Monitored Cell Triggered [${monitored.event}][${this._cell}]`)
      if (CHECK_TYPE == monitored?.type) {
        EDLogger.info(`Activating [${monitored.event}]`)
        EDLogger.notify(`🎲Rolling!!🎲`,{ title: `Performing ${monitored.event}`})
        //EDLogger.trace(JSON.stringify(EDProperties.event.getProperties()))
            // perform the event and then reset the cell
//          const rng = GSBatch.load.rangesNow([EDContext.context.config.EVENT_PROPERTIES_])
        GSBatch.add.cell(EDContext.context.batch,monitored.cell,0);
      }
      EDLogger.debug(JSON.stringify(monitored));
    }
    return monitored;
  }

  function getEventProperties() {
    if (!EDContext.context.event?.properties) {
      EDConfig.load(EDContext.context.config.event, {mode : GSBatch.MODE.SIMPLE});
    }
    return EDContext.context.event?.properties;
  }

  // Public API
  return {
    path : {
      unpack,
      repack,          
      repackMany,      
      isTrackedModified,
      getTrackedChanges,
      commitTrackedBaseline,
      clearTracking,
      getAtPath,
      UNPACK
    },

    event : {
      byCell : mappingByCell,
      getProperties : getEventProperties
    }



  };
})();
