/** Namespace: EDConfig */
var EDConfig = (function () {
  /* ================== config ================== */
  var GCACHE_PREFIX = 'EDSheet.config:'; // ScriptProperties key prefix

  /* ================== public: load ================== */
  /**
   * Load one or more config GROUP OBJECTS (queued batchGet).
   *
   * Examples:
   *   EDConfig.load(EDConfiguration.core);
   *   EDConfig.load([EDConfiguration.core, EDConfiguration.event], { flushCache:true });
   *   EDConfig.load(customGroupObj);
   *
   * @param {object|object[]} groups
   * @param {object} [opts]
   *   - flushCache?: boolean = false       // bypass any existing cache on read
   *   - render?: 'raw'|'display'|'formula' = 'raw'
   *   - dateTime?: 'SERIAL_NUMBER'|'FORMATTED_STRING' = 'SERIAL_NUMBER'
   *   - trim?: boolean = true              // drop rows whose first cell is empty-ish (API-side)
   *   - ignoreLoaded = false               // ignore the loaded flag forcing a refresh of the data either from cache or the sheet
   * @returns {{loaded:number, fromCache:number, skipped:number}}
   */
  function load(groups, opts = {}) {
    const {
      flushCache = false,
      render = 'raw',                    // 'raw' | 'display' | 'formula'
      dateTime = 'SERIAL_NUMBER',        // 'SERIAL_NUMBER' | 'FORMATTED_STRING'
      trim = true,
      ignoreLoaded = false               // NEW: do not skip defs/groups with loaded===true
    } = opts;

    const groupList = Array.isArray(groups) ? groups.filter(Boolean) : [groups].filter(Boolean);
    const props = EDContext.context.cache;

    const pending = [];                  // [{ name, range, def, groupKey, groupName, cacheable }]
    let fromCache = 0, skipped = 0;

    // Per-group accumulator for eventual ScriptProperties write
    const G = {};                        // groupKey => { cacheable, name, defs:{ defName: rows[] }, meta:{ defName:{unpack,prefix} }, groupObj }

    for (const group of groupList) {
      if (!group || typeof group !== 'object') continue;

      if (!ignoreLoaded && group.loaded === true) {
        for (const k of Object.keys(group)) if (GSUtils.Obj.isLeaf(group[k])) skipped++;
        continue;
      }

      const groupName      = group.name || '';
      const groupCacheable = !!group.cache;
      const groupKey       = _groupCacheKey(group);
      const cachedPayload  = (!flushCache && groupCacheable) ? _getGroupCache(props, groupKey) : null;
      if (cachedPayload) EDLogger.debug(`Found Cache [ ${groupKey}]`);

      if (!G[groupKey]) G[groupKey] = { cacheable: groupCacheable, name: groupName, defs: {}, meta: {}, groupObj: group };

      const leafKeys = Object.keys(group).filter(k => GSUtils.Obj.isLeaf(group[k]));
      const iterFromCache = (!leafKeys.length && cachedPayload && cachedPayload.defs);

      // If we have no leaf defs but do have a cache, iterate the cache entries.
      // Otherwise iterate the existing group leaves.
      const defItems = iterFromCache
        ? Object.keys(cachedPayload.defs).map(defName => {
            // Try to reconstruct a reasonable property key on the group:
            // If defName is "groupName.key", use "key"; else use defName as-is.
            const tail = (groupName && defName.startsWith(groupName + '.'))
              ? defName.slice(groupName.length + 1)
              : defName;

            // Reuse existing leaf if present; otherwise synthesize a stub and attach.
            let def = group[tail];
            if (!GSUtils.Obj.isLeaf(def)) {
              def = group[tail] = { name: defName, loaded: false };
            }
            return { key: tail, def, defName };
          })
        : leafKeys.map(key => {
            const def = group[key];
            const defName = def.name || (groupName ? `${groupName}.${key}` : key);
            return { key, def, defName };
          });

      for (const { key, def, defName } of defItems) {
        if (!ignoreLoaded && def.loaded === true) { skipped++; continue; }

        const cachedRows = (cachedPayload && cachedPayload.defs) ? cachedPayload.defs[defName] : undefined;
        if (cachedRows) {
          def.values = cachedRows;
          def.loaded = true;

          // Use cached meta (fallback to def.*); rehydrate the def for downstream code.
          const { unpack, prefix } = _pickDefMeta(def, cachedPayload && cachedPayload.meta, defName);
          def.unpack = unpack;
          def.prefix = prefix;

          if (unpack && unpack !== EDProperties.path.UNPACK.none) {
            EDProperties.path.unpack(EDContext.context, cachedRows, { mode: unpack, prefix, id: defName });
          }

          G[groupKey].defs[defName] = cachedRows;
          G[groupKey].meta[defName] = { unpack, prefix };

          fromCache++;
          continue;
        }

        // No cache rows for this def → if we don't even have a range (because we synthesized),
        // we can't batch-get it; skip it gracefully.
        if (!def.range) { continue; }

        // Normalize A1 to include sheet for stability
        const rangeA1 = GSRange.ensureSheetOnA1(String(def.range), EDContext.context.ss);

        pending.push({
          name: defName,
          range: rangeA1,
          def,
          groupKey,
          groupName,
          cacheable: groupCacheable
        });
      }


    }

    if (!pending.length) {
      return { loaded: 0, fromCache, skipped };
    }

    // Queue ONE batchGet via GSBatchValues (keeps original order)
    const b = GSBatch.newBatch(opts);
    const ranges = pending.map(p => p.range);
    EDLogger.debug("Loading Definitions from Sheet " + JSON.stringify(ranges));

    GSBatch.load.ranges(b, ranges, {
      valueRenderOption: GSBatch.load.renderMode(render),
      dateTimeRenderOption: dateTime,
      trim
    });

    const apiResults = GSBatch.commit(b) || [];
    const batchGetRes = apiResults.find(r => r && r.valueRanges) || apiResults[0] || {};
    const valueRanges = Array.isArray(batchGetRes.valueRanges) ? batchGetRes.valueRanges : [];

    let loaded = 0;
    for (let i = 0; i < pending.length; i++) {
      const rec  = pending[i];
      const vr   = valueRanges[i] || {};
      const rows = (vr.values || []).slice();

      rec.def.values = rows;
      rec.def.loaded = true;
      try { EDLogger.debug(`Loaded Definition [ ${rec.name} ] rows=${rows.length}`); } catch(e){}

      const mode = rec.def.unpack;
      if (mode && mode !== EDProperties.path.UNPACK.none) {
        const unpackName = rec.def.prefix;
        EDProperties.path.unpack(EDContext.context, rows, { mode, prefix: unpackName, id: rec.name });
      }

      if (rec.cacheable) {
        if (!G[rec.groupKey]) G[rec.groupKey] = { cacheable: true, name: rec.groupName, defs: {}, meta: {}, groupObj: null };
        G[rec.groupKey].defs[rec.name] = rows;
        // Persist unpack metadata alongside rows
        G[rec.groupKey].meta[rec.name] = { unpack: rec.def.unpack, prefix: rec.def.prefix };
      }

      loaded++;
    }

    // Merge + write group caches (one ScriptProperty per group), preserving both defs and meta
    for (const groupKey of Object.keys(G)) {
      const meta = G[groupKey];
      if (!meta.cacheable) continue;

      const existing = _getGroupCache(props, groupKey) || { defs: {}, meta: {} };
      const mergedDefs = Object.assign({}, existing.defs, meta.defs);
      const mergedMeta = Object.assign({}, existing.meta, meta.meta);

      const payload = { defs: mergedDefs, meta: mergedMeta, ts: Date.now(), name: meta.name || '' };
      _setGroupCache(props, groupKey, payload);
    }

    return { loaded, fromCache, skipped };
  }


  /* ================== public: intersect ================== */
  /**
   * Check if an A1 range intersects or equals ANY range inside one or more defs/groups.
   *
   * Accepts mixed inputs:
   *  - A single def:             { name:"...", range:"Sheet!A1:B2", ... }
   *  - An array of defs:         [{range:"..."}, ...]
   *  - A group object:           { settings:{range:"..."}, mappings:{range:"..."} }
   *  - Nested groups (any depth) — anything that eventually has a .range string
   *
   * @param {string} a1
   * @param {...any} defsOrGroups
   * @returns {{ any:boolean, matches:Array<{range:string, relation:'equal'|'intersect', sourcePath:string}> }}
   */
  function intersect(a1, ...defsOrGroups) {
    // normalize the probe A1 (adds sheet if missing)
    const probe = GSRange.ensureSheetOnA1(String(a1), EDContext.context.ss);

    const targets = _collectA1Ranges_(defsOrGroups);
    const matches = [];

    for (const t of targets) {
      const trg = t.range; // already normalized by _collectA1Ranges_
      if (GSRange.rangesEqual(probe, trg)) {
        matches.push({ range: trg, relation: 'equal', sourcePath: t.path });
        continue;
      }
      if (GSRange.rangesIntersect(probe, trg)) {
        matches.push({ range: trg, relation: 'intersect', sourcePath: t.path });
      }
    }
    return { any: matches.length > 0, matches };
  }


  /* ================== public: updateCache ================== */
  /**
   * Reset group cache from current .values on each leaf def.
   * Writes ONE ScriptProperty per group.
   *
   * @param {object|object[]} groups
   * @param {object} [opts]
   *   - includeNonCacheable?: boolean = false  // if true, writes even when group.cache !== true
   *   - clearIfNoValues?: boolean = false      // if true, clears group cache when no values
   *   - markLoaded?: boolean = false           // set def.loaded = true when values present
   * @returns {{set:number, cleared:number, skipped:number}}
   */
  function updateCache(groups, opts = {}) {
    const {
      includeNonCacheable = false,
      clearIfNoValues = false,
      markLoaded = false
    } = opts;

    const list = Array.isArray(groups) ? groups.filter(Boolean) : [groups].filter(Boolean);
    const props = EDContext.context.cache;

    let set = 0, cleared = 0, skipped = 0;

    for (const group of list) {
      if (!group || typeof group !== 'object') continue;

      const cacheable = !!group.cache;
      if (!includeNonCacheable && !cacheable) { skipped++; continue; }

      const groupKey = _groupCacheKey(group);
      const defsObj = {};
      const metaObj = {}; // NEW: persist per-def unpack/prefix
      let anyValues = false;

      for (const key of Object.keys(group)) {
        const def = group[key];
        if (!GSUtils.Obj.isLeaf(def)) continue;

        const defName = def.name || (group.name ? `${group.name}.${key}` : key);
        if (def && def.values != null) {
          defsObj[defName] = def.values;
          metaObj[defName] = { unpack: def.unpack, prefix: def.prefix };
          if (markLoaded) def.loaded = true;
          anyValues = true;
          EDLogger.trace(`Updating Cache Data [ ${defName} ] [ ${groupKey}) ]`);
        }
      }

      if (anyValues) {
        const existing = _getGroupCache(props, groupKey) || { defs: {}, meta: {} };
        const payload = {
          defs: Object.assign({}, existing.defs, defsObj),
          meta: Object.assign({}, existing.meta, metaObj),
          ts: Date.now(),
          name: group.name || ''
        };
        _setGroupCache(props, groupKey, payload);
        set++;
      } else if (clearIfNoValues) {
        _clearGroupCache(props, groupKey);
        EDLogger.warn(`Cleared Cache [ ${groupKey} ]`);
        cleared++;
      } else {
        skipped++;
      }
    }

    return { set, cleared, skipped };
  }

  /**
   * Convenience: load commonly-needed groups.
   * @param {object} [opts]
   *   - flushCache?: boolean = false       // bypass existing group cache on read
   *   - render?: 'raw'|'display'|'formula' = 'raw'
   *   - dateTime?: 'SERIAL_NUMBER'|'FORMATTED_STRING' = 'SERIAL_NUMBER'
   *   - trim?: boolean = true              // drop rows whose first cell is empty-ish (API-side)
   * @returns {{loaded:number, fromCache:number, skipped:number}}
   */
  function initialize(opts = {}) {
    const {
      boot = true
    } = opts;
    let loadDefs = true; // fixed: avoid implicit global

    if (boot) {
      const defs = load([EDContext.context.config.boot], opts);
      loadDefs = !(defs.loaded == 0 && defs.fromCache == 0);
    }

    if (loadDefs) {
      return load([EDContext.context.config.sheet, EDContext.context.config.core], opts);
    } else {
      EDLogger.error("Could not load Range Definitions");
      return { loaded: 0, fromCache: 0, skipped: 0 };
    }
  }


  /**
   * Collect A1 ranges from arbitrary config-like objects.
   *
   * @param {any[]} inputs  // array of roots to scan
   * @param {Object} [opts]
   *   - diveValues?: boolean     // default false – do not scan .values payloads
   *   - skipKeys?: string[]      // default ['name','unpack','loaded','cache','prefix']
   *   - normalize?: boolean      // default true – ensure sheet prefix on A1
   *   - ss?: Spreadsheet         // used by ensureSheetOnA1 when normalize=true
   *   - unique?: boolean         // default true – de-duplicate results
   * @returns {{range:string, path:string}[]}
   */
  function _collectA1Ranges_(inputs, opts = {}) {
    const {
      diveValues = false,
      skipKeys = ['name', 'unpack', 'loaded', 'cache', 'prefix'],
      normalize = true,
      ss = EDContext.context?.ss,
      unique = true,
    } = opts;

    const out = [];
    const seen = new Set();

    const isA1 = (s) => {
      try { GSRange.parseBox(String(s)); return true; } catch (e) { return false; }
    };

    const normA1 = (s) => normalize ? GSRange.ensureSheetOnA1(String(s), ss) : String(s);

    const push = (range, path) => {
      const a1 = normA1(range);
      if (!unique || !seen.has(a1)) {
        if (unique) seen.add(a1);
        out.push({ range: a1, path });
      }
    };

    const visit = (node, path) => {
      if (node == null) return;

      // String leaf?
      if (typeof node === 'string') {
        if (isA1(node)) push(node, path);
        return;
      }

      // Node with explicit .range?
      if (typeof node === 'object' && typeof node.range === 'string' && isA1(node.range)) {
        // Record the explicit range; do not descend into .values unless asked.
        push(node.range, path ? path + '.range' : 'range');
      }

      // Arrays
      if (Array.isArray(node)) {
        // If path indicates .values and diveValues=false, skip scanning payload
        if (!diveValues && /(^|\.|])values(\.|$|\[)/.test(path || '')) return;
        for (let i = 0; i < node.length; i++) {
          visit(node[i], path ? `${path}[${i}]` : `[${i}]`);
        }
        return;
      }

      // Objects
      if (typeof node === 'object') {
        for (const k of Object.keys(node)) {
          if (!diveValues && k === 'values') continue;      // skip heavy payloads by default
          if (skipKeys && skipKeys.indexOf(k) !== -1) continue; // skip metadata keys
          visit(node[k], path ? `${path}.${k}` : k);
        }
      }
    };

    for (let i = 0; i < (inputs?.length || 0); i++) {
      visit(inputs[i], `arg${i}`);
    }
    return out;
  }


  /* ================== private: cache helpers ================== */

  function _groupCacheKey(group) {
    // Prefer explicit group name for stability
    if (group && typeof group === 'object' && group.name) {
      return GCACHE_PREFIX + group.name;
    }
    // Fallback: stable signature from leaf def names+ranges
    var sig = [];
    try {
      for (const k of Object.keys(group)) {
        const d = group[k];
        if (GSUtils.Obj.isLeaf(d)) sig.push((d.name || k) + '@' + (d.range || ''));
      }
      sig.sort();
    } catch (e) {}
    return GCACHE_PREFIX + 'anon:' + GSUtils.Str.hash32(sig.join('|'));
  }

  function _getGroupCache(props, cacheKey) {
    try {
      const raw = props.getProperty(cacheKey);
      if (!raw) return null;

      const payload = JSON.parse(raw);
      if (!payload || typeof payload !== 'object') return null;
      if (!payload.defs || typeof payload.defs !== 'object') return null;
      return payload;
    } catch (e) {
      EDLogger.warn('Bad Cache [ ' + cacheKey + ' ] [ ' + e.msg + ' ]');
      throw e;
    }
  }

  function _setGroupCache(props, cacheKey, obj) {
    try {
      const s = JSON.stringify(obj);
      props.setProperty(cacheKey, s);
      // ➕ precise byte size message while keeping overall style
      EDLogger.debug(`Set Cache [ ${cacheKey} ] [ ${GSUtils.Str.byteLen(s)} bytes ] `);
    } catch (e) {
      EDLogger.error('Set Cache Failed ' + e);
    }
  }

  function _clearGroupCache(props, cacheKey) {
    try { props.deleteProperty(cacheKey); } catch (e) {}
  }

  // Centralized meta selection (cache-first; fall back to def.* so old caches work)
  function _pickDefMeta(def, metaMap, defName) {
    const m = (metaMap && metaMap[defName]) || {};
    return {
      unpack: (m.unpack != null ? m.unpack : def && def.unpack),
      prefix: (m.prefix != null ? m.prefix : def && def.prefix),
    };
  }


  /* ================== public: edit hook ================== */
  function configEdited(a1) {
    const { any } = intersect(
      GSRange.ensureSheetOnA1(String(a1), EDContext.context.ss),
      EDContext.context.config.sheet,
      EDContext.context.config.core,
      EDContext.context.config.boot
    );

    if (any) {
      EDLogger.notify(`Config cell edited [${a1}]`, {title : "Configuration Updated"});
      initialize({ flushCache: true, ignoreLoaded: true });
      return true;
    }
    return false;
  }

  /* ================== public API ================== */
  return {
    load,
    initialize,
    updateCache,
    configEdited
    // (intersect is intentionally not exported publicly in original style; keep as-is)
  };
})();
