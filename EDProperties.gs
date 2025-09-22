var EDProperties = (function () {

  const UNPACK = {
    object : "object",
    array  : "array",
    pair   : "pair",
    none   : "none",
  }

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
   */
  function unpack(target, rows, opts = {}) {
    if (!Array.isArray(rows)) throw new Error('rows must be a 2D array like [[name, ...], ...]');
    const mode = (opts.mode || 'pair').toLowerCase();
    const track = !!opts.track;
    const prefix = opts?.prefix;
    const name = opts?.name ?? "NO NAME";

    EDLogger.debug(`Unpacking [ ${name} ][ mode: ${mode} ][ Rows: ${rows.length}]`);

    if (mode === 'array') {
      const { headers, dataStart } = _resolveHeaders(rows, opts);
      if (!headers || !headers.length) return [];

      const out = [];
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
        getAtPath(root, p, { sep: '|', transform: safePropName })
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
      const after  = getAtPath(root, p, { sep: '|', transform: safePropName });
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
        getAtPath(root, p, { sep: '|', transform: safePropName })
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
    for (const p of paths) tr.baseline.set(p, GSUtils.Obj.deepCloneSimple(getAtPath(root, p, { sep: '|', transform: safePropName })));
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
    const xform = opts.transform || safePropName;
    const parts = String(path).split(/\||\./).map(s => s.trim()).filter(Boolean);
    let cur = obj;
    for (const p of parts) {
      if (cur == null) return undefined;
      cur = cur[xform(p)];
    }
    return cur;
  }  

  // Public API
  return {
    path : {
      unpack,
      isTrackedModified,
      getTrackedChanges,
      commitTrackedBaseline,
      clearTracking,
      getAtPath,
      UNPACK
    },



  };
})();
