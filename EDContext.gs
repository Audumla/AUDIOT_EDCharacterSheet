var EDContext = (function () {

  const STATUS = {
    UNINITIALIZED : {name : "UNINITIALIZED", idx : 1 },
    PENDING       : {name : "PENDING", idx : 2 },
    PROCESSING    : {name : "PROCESSING", idx : 3 },
    COMPLETED     : {name : "COMPLETED", idx : 5 },
    FAILED        : {name : "FAILED", idx : 7 },
    IGNORED       : {name : "IGNORED", idx : 99 },
  }

  const EDConfiguration = {

    DEFINITION_RANGES : {
      range  : "References!$A2:$D20",
      name   : "DEFINITION_RANGES",
      unpack : EDProperties.path.UNPACK.array,
      values : undefined
    },

    EVENT_DEFINITIONS : {},
    SETTINGS : {
      values : undefined
    },
    EVENT_PROPERTIES : {},
    CHARACTER_PROPERTIES : {}

  }

  const context = {

    cfg : EDConfiguration,
    logger  : undefined,

    // spreadsheet context
    ss : undefined,
    ssid : undefined,
    tz : undefined,

    // batch context
    batch : undefined,

    // runtime flags/state
    status : undefined,

    // timing
    startTime : undefined,

    event : {
      activeID : 0,
      status : STATUS.UNINITIALIZED
    }
  }

  /**
   * Create a fresh runtime opts bag.
   * Pass overrides for testing or special cases. Nothing global is mutated.
   *
   * @param {{
   *   ss?: GoogleAppsScript.Spreadsheet.Spreadsheet,
   *   cfg?: any,
   *   logger?: any,
   *   batch?: any,
   *   status?: any,
   *   startTime?: Date
   * }} [overrides]
   */
  function initializeContext(overrides) {
    
    overrides = overrides || {};
    const ss = overrides.ss || SpreadsheetApp.getActive();

    // deps
    context.cfg =    overrides.cfg    || EDConfiguration;
    context.logger = overrides.logger || EDLogger;

    // spreadsheet context
    context.ss = ss;
    context.ssid = ss.getId();
    context.tz = ss.getSpreadsheetTimeZone();

    // batch context
    context.batch = overrides.batch || GSBatch.newBatch(ss);

    // runtime flags/state
    context.status = overrides.status || STATUS.PENDING;

    // timing
    context.startTime = overrides.startTime || new Date();
  }

  // ============================================================
  // Public API
  // ============================================================

  return {
    initializeContext,
    context,
    STATUS
  }
})()