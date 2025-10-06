var EDDefs_ = (function () {
  // ------- private constants / services -------
  const cacheValues = PropertiesService.getScriptProperties();

  const DEFINITION_INDEX = {
    name   : 0,
    range  : 1,
    cache  : 2,
    unpack : 3,
  };

  function loadDefinition(def, opt = {}) {
    const cfg = EDContext.context.config;
  }

  function getDefinition(name) {
    let values = undefined
    if (EDContext.context[name]?.loaded ?? false) {
      loadDefinition(name)
    }
    else {
      values = EDContext.context[name].values;
    }
  }

  function loadDefinitions(defs, loadNonCached = false, useCache = false) {
    const cfg = EDContext.context.config;

    const filtered = (defs || []).filter(row =>
      row?.[DEFINITION_INDEX.range] && String(row[DEFINITION_INDEX.range]).length > 0 &&
      (GSUtils.Str.toBool(row[DEFINITION_INDEX.cache]) || loadNonCached)
    );

    const rangeOnly = (defs || []).filter(row =>
      row?.[DEFINITION_INDEX.range] && String(row[DEFINITION_INDEX.range]).length > 0 &&
      (!GSUtils.Str.toBool(row[DEFINITION_INDEX.cache]) && !loadNonCached)
    );

    const loadDefs = useCache
      ? filtered.filter(def =>
          !GSUtils.Str.toBool(def?.[DEFINITION_INDEX.cache]) ||
          EDDefs_.getCachedData(def?.[DEFINITION_INDEX.name]) === undefined
        )
      : filtered;

    let count = 0;
    if (loadDefs.length > 0) {
      EDLogger.trace("Loading Definitions from Sheet " + JSON.stringify(loadDefs));
      const loadedDefinitions = GSBatch.load.rangesNow(loadDefs); 

      for (let i = 0; i < loadedDefinitions.length; i++) {
        const definition = loadedDefinitions[i];          // { name, range, values, ... }
        const src = loadDefs[i];                          // original row definition

        cfg[definition.name] = { ...definition };
        cfg[definition.name].loaded = true;

        if (GSUtils.Str.toBool(src?.[DEFINITION_INDEX.cache])) {
          EDDefs_.setCacheData(cfg[definition.name]);
        }
        count++;
      }
    }

    if (rangeOnly.length > 0) {
      EDLogger.trace("Adding Definitions " + JSON.stringify(rangeOnly));
      for (const def of rangeOnly) {
        cfg[def[DEFINITION_INDEX.name]] = { 
          name  : def[DEFINITION_INDEX.name],
          range : def[DEFINITION_INDEX.range],
          values : undefined,
          unpack : def[DEFINITION_INDEX.unpack],
          loaded : false

        };        
      }
    }

    // Unpack after all loads (including cached ones present in cfg)
    for (const def of filtered) {
      const name = def?.name || def?.[DEFINITION_INDEX.name];
      const unpack = def?.unpack?.toLowerCase() || def?.[DEFINITION_INDEX.unpack]?.toLowerCase();
      const node = name ? cfg[name] : undefined;
      if (unpack != EDProperties.path.UNPACK.none && node?.values) {
        EDProperties.path.unpack(EDContext.context, node.values, {mode : unpack});

      }
    }

    return count;
  }

  function initializeDefinitions(loadNonCached = false, useCache = false,) {
    
    const cfg = EDContext.context.config;

    EDLogger.info("Initializing Definition Tables");

    if (!useCache && cfg?.DEFINITION_RANGES_?.name) {
      EDDefs_.clearCacheData(cfg.DEFINITION_RANGES_.name);
    }

    if (cfg?.DEFINITION_RANGES_?.name) {
      EDDefs_.getCachedData(cfg.DEFINITION_RANGES_.name);
    }

    const rows = cfg?.DEFINITION_RANGES_?.values || [];
    if (rows.length === 1) {
      EDDefs_.loadDefinitions(rows, loadNonCached, useCache);
    }
    if (rows.length > 1) {
      EDDefs_.loadDefinitions(rows.slice(1), loadNonCached, useCache);
    }
  }

  function checkCachedDataChanged(range) {
    const cfg = EDContext.context.config;

    if (cfg?.DEFINITION_RANGES_?.name) {
      EDDefs_.getCachedData(cfg.DEFINITION_RANGES_.name);
    }

    const defRanges = (cfg?.DEFINITION_RANGES_?.values || [])
      .filter(row => row?.[DEFINITION_INDEX.range])
      .map(row => row[DEFINITION_INDEX.range]);

    var changedDefs = GSRange.rangesIntersectPairs(range, defRanges)
    if (changedDefs.length > 0) {
//      EDLogger.debug(Utilities.formatString("Cached Data Edited [ %s ]", range));
      EDDefs_.initializeDefinitions(false, false);
      return true;
    } else {
//      EDLogger.info(Utilities.formatString("Edited Cell Not In Cache [ %s ]", range));
      return false;
    }
  }

  function setCacheData(definition) {

    const obj   = JSON.stringify(definition);
    const bytes = GSUtils.Str.byteLen(obj);

    if (bytes > 9000) {
      EDLogger.error(`Cache payload too large for [${definition.name}] (${bytes} bytes). Skipping cache write.`);
      return;
    }

    cacheValues.setProperty(definition.name, obj);
    EDLogger.debug(`Stored Cache Data [${definition.name}:${definition.range ? definition.range : definition.values}] (${bytes} bytes)`);
  }

  function clearCacheData(name) {

    const cfg = EDContext.context.config;

    cacheValues.deleteProperty(name);
    if (cfg && cfg[name]) {
      cfg[name].values = undefined;
      EDLogger.debug(`Cleared Cached Data [${name}]`);
    }
  }

  function getCachedData(name) {

    const cfg = EDContext.context.config;

    let data = cfg ? cfg[name] : undefined;

    if (data === undefined || data.values === undefined) {
      const property = cacheValues.getProperty(name);
      if (property != null) {
        try {
          data = JSON.parse(property);
          cfg[name] = data;
          cfg[name].loaded = true;
          EDLogger.debug(`Found Cached Data [${name}]`);
        } catch (e) {
          EDDefs_.clearCacheData(name);
          EDLogger.error(`Invalid Cached Data [${name}] [ ${e.message} ]`);
        }
      }

      if ((data === undefined || data.values === undefined) && cfg && cfg[name]?.range !== undefined) {
        const row = [
          cfg[name].name,
          cfg[name].range,
          cfg[name].cache ? cfg[name].cache : true,
          cfg[name].unpack ? cfg[name].unpack : EDProperties.path.UNPACK.none,
        ];
        EDDefs_.loadDefinitions([row], false, false);
        data = cfg[name];
      } else if (data === undefined || data.values === undefined) {
        return undefined;
      }
    }

    return data?.values;
  }

  // ------- public API -------
  return {

    loadDefinitions,
    initializeDefinitions,
    checkCachedDataChanged,
    setCacheData,
    clearCacheData,
    getCachedData,
    cacheValues,

  };
})();
