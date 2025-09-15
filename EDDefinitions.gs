var EDDefs = (function () {
  // ------- private constants / services -------
  const ScriptValues = PropertiesService.getUserProperties();

  const DEFINITION_INDEX = {
    name   : 0,
    range  : 1,
    cache  : 2,
    unpack : 3,
  };

  // ------- private helpers -------
  function toBoolFlag(v) { return v === true || v === 1 || v === "1"; }
  function byteLen(str)  { return Utilities.newBlob(str).getBytes().length; }


  function loadDefinitions(defs, loadNonCached = false, useCache = false, opts) {
    opts = resolveOpts(opts);
    const { cfg, logger } = opts;

    const filtered = (defs || []).filter(row =>
      row?.[DEFINITION_INDEX.range] && String(row[DEFINITION_INDEX.range]).length > 0 &&
      (toBoolFlag(row[DEFINITION_INDEX.cache]) || loadNonCached)
    );

    const loadDefs = useCache
      ? filtered.filter(def =>
          !toBoolFlag(def?.[DEFINITION_INDEX.cache]) ||
          getCachedData(def?.[DEFINITION_INDEX.name], opts) === undefined
        )
      : filtered;

    let count = 0;
    if (loadDefs.length > 0) {
      logger.trace("Loading Definitions " + JSON.stringify(loadDefs));
      const loadedDefinitions = GSBatch.load.ranges(loadDefs, { ...opts }); // keep all opts

      for (let i = 0; i < loadedDefinitions.length; i++) {
        const definition = loadedDefinitions[i];          // { name, range, values, ... }
        const src = loadDefs[i];                          // original row definition

        cfg[definition.name] = { ...definition };
        cfg[definition.name].unpack = toBoolFlag(src?.[DEFINITION_INDEX.unpack]);

        if (toBoolFlag(src?.[DEFINITION_INDEX.cache])) {
          setCacheData(cfg[definition.name], opts);
        }
        count++;
      }
    }

    // Unpack after all loads (including cached ones present in cfg)
    for (const def of filtered) {
      const name = def?.[DEFINITION_INDEX.name];
      const node = name ? cfg[name] : undefined;
      if (node?.unpack && node?.values) {
        unpackProperties(cfg, node.values, logger);
        logger.debug(`Unpacked Properties [${node.name}]`);
      }
    }

    return count;
  }

  function initializeDefinitions(loadNonCached = false, useCache = false, opts) {
    opts = resolveOpts(opts);
    const { cfg, logger } = opts;

    logger.info("Initializing Definition Tables");

    if (!useCache && cfg?.DEFINITION_RANGES?.name) {
      clearCacheData(cfg.DEFINITION_RANGES.name, opts);
    }

    if (cfg?.DEFINITION_RANGES?.name) {
      getCachedData(cfg.DEFINITION_RANGES.name, opts);
    }

    const rows = cfg?.DEFINITION_RANGES?.values || [];
    if (rows.length === 1) {
      loadDefinitions(rows, loadNonCached, useCache, opts);
    }
    if (rows.length > 1) {
      loadDefinitions(rows.slice(1), loadNonCached, useCache, opts);
    }
  }

  function checkCachedDataChanged(range, opts) {
    opts = resolveOpts(opts);
    const { cfg, logger } = opts;

    if (cfg?.DEFINITION_RANGES?.name) {
      getCachedData(cfg.DEFINITION_RANGES.name, opts);
    }

    const defRanges = (cfg?.DEFINITION_RANGES?.values || [])
      .filter(row => row?.[DEFINITION_INDEX.range])
      .map(row => row[DEFINITION_INDEX.range]);

    if (GSUtils.Arr.rangesIntersectAny(range, defRanges)) {
      logger.debug(Utilities.formatString("Cached Data Edited [ %s ]", range));
      initializeDefinitions(false, false, opts);
      return true;
    } else {
      logger.trace(Utilities.formatString("Edited Cell Not In Cache [ %s ]", range));
      return false;
    }
  }

  function setCacheData(definition, opts) {
    opts = resolveOpts(opts);
    const { logger } = opts;

    const obj   = JSON.stringify(definition);
    const bytes = byteLen(obj);

    if (bytes > 9000) {
      logger.error(`Cache payload too large for [${definition.name}] (${bytes} bytes). Skipping cache write.`);
      return;
    }

    ScriptValues.setProperty(definition.name, obj);
    logger.debug(`Stored Cache Data [${definition.name}:${definition.range}] (${bytes} bytes)`);
  }

  function clearCacheData(name, opts) {
    opts = resolveOpts(opts);
    const { cfg, logger } = opts;

    ScriptValues.deleteProperty(name);
    if (cfg && cfg[name]) {
      cfg[name].values = undefined;
      logger.debug(`Cleared Cached Data [${name}]`);
    }
  }

  function getCachedData(name, opts) {
    opts = resolveOpts(opts);
    const { cfg, logger } = opts;

    let data = cfg ? cfg[name] : undefined;

    if (data === undefined || data.values === undefined) {
      const property = ScriptValues.getProperty(name);
      if (property != null) {
        try {
          data = JSON.parse(property);
          if (cfg) cfg[name] = data;
          logger.debug(`Found Cached Data [${name}]`);
        } catch (e) {
          clearCacheData(name, opts);
          logger.error(`Invalid Cached Data [${name}] [ ${e.message} ]`);
        }
      }

      if ((data === undefined || data.values === undefined) && cfg && cfg[name]?.range !== undefined) {
        const row = [
          cfg[name].name,
          cfg[name].range,
          cfg[name].cache ? cfg[name].cache : true,
          cfg[name].unpack ? cfg[name].unpack : false,
        ];
        loadDefinitions([row], false, false, opts);
        data = cfg[name];
      } else if (data === undefined || data.values === undefined) {
        return undefined;
      }
    }

    return data?.values;
  }

  // ------- public API -------
  return {
    resolveOpts,
    loadDefinitions,
    initializeDefinitions,
    checkCachedDataChanged,
    setCacheData,
    clearCacheData,
    getCachedData,
  };
})();
