const UNPACK = {
  object : "object",
  array  : "array",
  pair   : "pair",
  none   : "none",
}

var EDContext = (function () {

  const STATUS = {
    INITIALIZING  : "INITIALIZING",
    PROCESSING    : "PROCESSING",
    COMPLETED     : "COMPLETED",
    FAILED        : "FAILED", 
    IGNORED       : "IGNORED",
    UNKNOWN       : "UNKNOWN",
  }

  const EDConfiguration = {

    boot : {
      cache  : true,
      loaded : false,
      name   : "config.boot",

      definitions : {
        name   : "config.boot.definitions",
        range  : "References!$A$1:$D20",
        unpack : UNPACK.object,
        values : undefined,

      }
    },

    sheet : {
      loaded : false,
      cache  : true,
      name   : "config.sheet",
/*
      status : {
        range  : "References!$F$2:$G$3",
        name   : "config.sheet.status",
        unpack : UNPACK.pair,
        values : undefined,
      }
*/
    },

    event : {
      loaded : false,
      cache  : false,
      name   : "config.event",
/*
      properties : {
        range  : "'Test Events'!$A:$D",
        name   : "config.event.properties",
        unpack : UNPACK.array,
        values : undefined,
        prefix : "event.properties"
      }
      */
    },

    core : {
      cache  : true,
      loaded : false,
      name   : "config.core",
      /*
      settings : {
        range  : "References!$F$8:$G$17",
        name   : "config.core.settings",
        unpack : UNPACK.pair,
        values : undefined,
      },

      mappings : {
        range  : "References!$R$1:$S$3",
        name   : "config.core.mappings",
        unpack : UNPACK.array,
        values : undefined,
        prefix : "mappings"
      },

      monitored : {
        range  : "References!$I$1:$P$19",
        name   : "config.core.monitored",
        unpack : UNPACK.array,
        values : undefined,
        prefix : "mappings"
      },

      masks : {
        range  : "'Processed Rules'!$E$11:$G$44",
        name   : "config.core.masks",
        unpack : UNPACK.array,
        values : undefined,
        prefix : "masks"
      }
      */

    },

  }

  const context = {
    cache : PropertiesService.getDocumentProperties(),

    config : EDConfiguration,
    logger  : undefined,

    // spreadsheet context
    ss : undefined,
    ssid : undefined,

    // batch context
    batch : undefined,

    date : {
      tz : undefined,
      format : "yyyy-MM-dd hh:mm:ss",
      time : undefined
    },

    event : {
      id : { 
        value : 0,
      },

      status : {
        state : STATUS.INITIALIZING
      }
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

//    GSBatch = GSBatchV2;
    overrides = overrides || {};
    const ss = overrides.ss || SpreadsheetApp.getActive();
    // deps
    context.config = overrides.cfg    || EDConfiguration;
    context.logger = overrides.logger || EDLogger;


    // spreadsheet context
    context.ss = ss;
    context.ssid = ss.getId();
    context.date.tz = ss.getSpreadsheetTimeZone();

    // batch context
    context.batch = overrides.batch || GSBatch.newBatch(ss);

    // runtime flags/state
    context.event.status.state = overrides.event?.status?.state || STATUS.INITIALIZING;

    // timing
    context.date.time = overrides.date?.time || new Date();
  }

  // ============================================================
  // Public API
  // ============================================================

  return {
    initializeContext,
    context,
    logger : EDLogger,
    STATUS,
    
  }
})()