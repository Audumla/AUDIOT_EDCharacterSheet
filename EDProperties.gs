var EDProperties = (function () {
  // Non-enumerable key to hold tracking data on target objects
  const __TRACK__KEY = '__tracking__'; // { paths:Set<string>, baseline: Map<string, any> }

  /**
   * Populate an object from [name, value] rows, supporting nested names and 4-col objects.
   * If opts.track === true, records created leaf paths and snapshots a baseline
   * so you can test later whether those properties were modified.
   *
   * Row shapes:
   *   [name, value] →
   *     sets leaf to a simple coerced value.
   *   [name, propertyMask, value, valueMask] →
   *     sets leaf to an object:
   *       { property: <name>, propertyMask: <coerced>, value: <coerced>, valueMask: <coerced> }
   *
   * Rules:
   *  - name may contain '|' to form nested objects (e.g., "stats|dex|mod").
   *  - Coercion: "true"/"false" → boolean; numeric strings → number; otherwise trimmed string.
   *  - Leaf assignment OVERWRITES any previous value.
   *  - If a non-object exists at an intermediate path, it is promoted to { __value: prev }.
   *
   * @param {Object} target
   * @param {Array<Array<any>>} rows
   * @param {Object} [logger]  // expected to have trace/debug methods; optional
   * @param {{track?: boolean}} [opts]
   * @returns {Object}
   */
  function unpackProperties(target, rows, opts = DEFAULT_OPTS) {
    opts = resolveOpts(opts);
    
    if (!target || typeof target !== 'object') throw new Error('target must be an object');
    if (!Array.isArray(rows)) throw new Error('rows must be a 2D array like [[name, ...], ...]');

    opts.logger.trace("Unpacking " + JSON.stringify(rows));

    const track = !!opts.track;
    const createdPaths = []; // leaf paths we set during this call

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const nameRaw = row[0];
      if (nameRaw == null) continue;

      const name = String(nameRaw).trim();
      if (!name) continue;

      const partsRaw = name.split(/\||\./).map(s => s.trim()).filter(Boolean);
      if (!partsRaw.length) continue;

      // Normalize path segments
      const parts = partsRaw.map(GSUtils.Obj.safePropName);

      // Walk/create intermediate nodes
      let node = target;
      for (let p = 0; p < parts.length - 1; p++) {
        const key = parts[p];
        const cur = node[key];
        if (cur == null) {
          node[key] = {};
        } else if (typeof cur !== 'object' || Array.isArray(cur)) {
          node[key] = { __value: cur }; // promote scalar/array to object
        }
        node = node[key];
      }

      const leafKey = parts[parts.length - 1];

      // 4-column object mode
      if (row.length >= 4) {
        node[leafKey] = {
          property: name,                             // keep the full path string verbatim
          propertyMask: GSUtils.Str.coerce(row[1]),
          value:        GSUtils.Str.coerce(row[2]),
          valueMask:    GSUtils.Str.coerce(row[3]),
        };
        if (track) createdPaths.push(name);
        continue;
      }

      // 2-column simple value
      const simpleValue = GSUtils.Str.coerce(row.length > 1 ? row[1] : null);
      node[leafKey] = simpleValue; // overwrite
      if (track) createdPaths.push(name);
    }

    if (track) _attachTrackingBaseline_(target, createdPaths);
    return target;
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

  /** Returns true if any tracked leaf has changed from its baseline. */
  function isTrackedModified(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return false;
    for (const p of tr.paths) {
      if (!GSUtils.Obj.deepEqualSimple(tr.baseline.get(p),
        getAtPath(root, p, { sep: '|', transform: safePropName }))) {
        return true;
      }
    }
    return false;
  }

  /**
   * Returns an array of { property, before, after } for tracked leaves that differ.
   * If nothing is tracked, returns [].
   */
  function getTrackedChanges(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return [];
    const out = [];
    for (const p of tr.paths) {
      const before = tr.baseline.get(p);
      const after  = getAtPath(root, p, { sep: '|', transform: safePropName });
      if (!GSUtils.Obj.deepEqualSimple(before, after)) out.push({ property: p, before, after });
    }
    return out;
  }

  /** Make the current values the new baseline (acknowledge/commit). */
  function commitTrackedBaseline(root) {
    const tr = root && root[__TRACK__KEY];
    if (!tr) return;
    for (const p of tr.paths) {
      tr.baseline.set(p, GSUtils.Obj.deepCloneSimple(getAtPath(root, p, { sep: '|', transform: safePropName })));
    }
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
  function getAtPath(obj, path, opts = DEFAULT_OPTS) {
    opts = resolveOpts(opts);
    const sep = opts.sep || '|';
    const xform = opts.transform || safePropName;

    const parts = String(path).split(sep).map(s => s.trim()).filter(Boolean);
    let cur = obj;
    for (const p of parts) {
      if (cur == null) return undefined;
      cur = cur[xform(p)];
    }
    return cur;
  }  

  // Public API
  return {
    unpackProperties,
    isTrackedModified,
    getTrackedChanges,
    commitTrackedBaseline,
    clearTracking,
    getAtPath
  };
})();
