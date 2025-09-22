var EDLogger = (function () {
  // ---------------- constants ----------------
  var LEVEL = {
    PERFORMANCE : { level: 1,  name: "PERF"        },
    TRACE       : { level: 5,  name: "TRACE"       },
    DEBUG       : { level: 6,  name: "DEBUG"       },
    WARNING     : { level: 7,  name: "WARNING"     },
    INFO        : { level: 8,  name: "INFO"        },
    CHARACTER   : { level: 9,  name: "CHARACTER"   },
    ERROR       : { level: 15, name: "ERROR"       },
    DISABLED    : { level: 99, name: "DISABLED"    },
  };

  // ---------------- state ----------------
  var _buffer = [];      // [{ logLevel, ts, msg }]
  // Logger levels stored as the string name of the level so they can be updated via external scripts
  var settings = {
    sheet   : {
      level : LEVEL.INFO.name,
      maxEntries : 500,
      batchMerge : true
    },
    console : {
      level : LEVEL.TRACE.name
    }
  }

  // ---------------- helpers ----------------
  function _entry(levelObj, msg) {
   // console.info(msg);
    _buffer.push({ logLevel: levelObj, ts: new Date(), msg: msg });
  }

  // ---------------- sinks ----------------

  // console sink
  function _flushConsole(entries,opts) {
    var minLevel = LEVEL[settings.console.level ?? LEVEL.TRACE.name].level;
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (e.logLevel.level < minLevel) continue;
      var dt = GSUtils.Date.formatDate(e.ts);
      const lmsg = (typeof e.msg === "string" ) ? e.msg : JSON.stringify(e.msg);
      var line = Utilities.formatString("【 EVENT : %s 】【 %s 】【 %s 】【 %s 】",EDContext.context.event.activeID, dt, e.logLevel.name, lmsg);
      // Use console.info for uniformity
      console.info(line);
    }
    return undefined;

  }

  // sheet sink (skeleton; fill with your GSBatch call if desired)
  function _flushSheet(entries, opts) {

    var minLevel = LEVEL[settings.sheet.level ?? LEVEL.TRACE.name].level;

    var rows = [];
    for (var i = entries.length-1; i >= 0; i--) {
      var e = entries[i];
      if (e.logLevel.level < minLevel) continue;
      rows.push([
        EDContext.context.event.activeID,
        GSUtils.Date.formatDate(e.ts),
        e.logLevel.name,
        e.msg
      ]);
    }

    const lBatch = (opts?.singleBatch ?? settings.sheet.batchMerge) ? EDContext.context.batch : GSBatch.newBatch(EDContext.context.ss);

    if (rows.length > 0) {

      var logRange = GSRange.extendA1(EDContext.context.cfg.SHEET_LOG.range, { top:0, bottom:rows.length-1, left:0, right:4 } );
      GSBatch.insert.range(lBatch,logRange,SpreadsheetApp.Dimension.ROWS);
      GSBatch.add.values(lBatch,logRange,rows,EDContext.context);
      GSBatch.remove.rows(lBatch,settings.sheet.maxEntries,rows.length,{sheet : EDContext.context.definitions.sheetName});
    }
    
    return lBatch;

  }

  var _sinks = [_flushConsole,_flushSheet];


  // ---------------- public api ----------------

  // emit
  function perf(msg)      { _entry(LEVEL.PERFORMANCE, msg); }
  function trace(msg)     { _entry(LEVEL.TRACE, msg); }
  function debug(msg)     { _entry(LEVEL.DEBUG, msg); }
  function warn(msg)      { _entry(LEVEL.WARNING, msg); }
  function info(msg)      { _entry(LEVEL.INFO, msg); }
  function character(msg) { _entry(LEVEL.CHARACTER, msg); }
  function error(msg)     { _entry(LEVEL.ERROR, msg); }

  // flush buffer to sinks
  function flush(opts = {}) {
    var batches = [];
    var entries = _buffer;
    try {
      _sinks.forEach(s => batches.push(s(entries,opts)));
    } finally {
      clear();
    }
    return batches;
  }

  // housekeeping
  function clear() { _buffer.length = 0; }
  function size()  { return _buffer.length; }
  function peek(n) { return _buffer.slice(Math.max(0, _buffer.length - (n || 50))); }

  // export
  return {
    LEVEL: LEVEL,

    // emit
    trace: trace,
    debug: debug,
    warn: warn,
    info: info,
    character: character,
    error: error,
    perf: perf,

    // control
    flush: flush,
    clear: clear,
    size:  size,
    peek:  peek,

    settings : settings,
  };
})();
