var EDLogger = (function () {
  // ---------------- constants ----------------
  var LEVEL = {
    TRACE     : { level: 0,  name: "TRACE"     },
    DEBUG     : { level: 1,  name: "DEBUG"     },
    WARNING   : { level: 2,  name: "WARNING"   },
    INFO      : { level: 3,  name: "INFO"      },
    CHARACTER : { level: 5,  name: "CHARACTER" },
    ERROR     : { level: 10, name: "ERROR"     },
    DISABLED  : { level: 99, name: "DISABLED"  },
  };

  // ---------------- state ----------------
  var _buffer = [];      // [{ logLevel, ts, msg }]
  // Logger levels stored as the string name of the level so they can be updated via external scripts
  var settings = {
    sheet   : {
      level : LEVEL.INFO.name,
      maxEntries : 500
    },
    console : {
      level : LEVEL.TRACE.name
    }
  }

  // ---------------- helpers ----------------
  function _entry(levelObj, msg) {
//    console.info(msg);
    _buffer.push({ logLevel: levelObj, ts: new Date(), msg: msg });
  }

  // ---------------- sinks ----------------

  // console sink
  function _flushConsole(entries) {
    var minLevel = LEVEL[settings.console.level ?? LEVEL.TRACE.name].level;
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (e.logLevel.level < minLevel) continue;
      var dt = GSUtils.Date.formatDate(e.ts);
      const lmsg = (typeof e.msg === "string" ) ? e.msg : JSON.stringify(e.msg);
      var line = Utilities.formatString("【 EVENT : %s 】【 %s 】【 %s 】【 %s 】",EDContext.context.eventID, dt, e.logLevel.name, lmsg);
      // Use console.info for uniformity
      console.info(line);
    }

  }

  // sheet sink (skeleton; fill with your GSBatch call if desired)
  function _flushSheet(entries) {
    var minLevel = LEVEL[settings.sheet.level ?? LEVEL.TRACE.name].level;

    // TODO: if you want to batch to a sheet, filter first:
    var rows = [];
    for (var i = entries.length-1; i >= 0; i--) {
      var e = entries[i];
      if (e.logLevel.level < minLevel) continue;
      rows.push([
        EDContext.context.eventID,
        GSUtils.Date.formatDate(e.ts),
        e.logLevel.name,
        e.msg
      ]);
    }



    // Example placeholder (uncomment/replace with your implementation):

    if (EDContext.context.batch && rows.length > 0) {

      var logRange = GSRange.extendA1(EDContext.context.cfg.SHEET_LOG.range, { top:0, bottom:rows.length-1, left:0, right:4 } );
      GSBatch.insert.range(EDContext.context.batch,logRange,SpreadsheetApp.Dimension.ROWS);
      GSBatch.add.values(EDContext.context.batch,logRange,rows,EDContext.context);
      GSBatch.remove.rows(EDContext.context.batch,settings.sheet.maxEntries,rows.length);
    }

    // If you want a simple direct write:
    // if (opts && opts.sheetA1 && rows.length) {
    //   var rng = GSRange.resolveRange(opts.sheetA1, opts);
    //   var sheet = rng.getSheet();
    //   sheet.getRange(rng.getRow(), rng.getColumn(), rows.length, rows[0].length).setValues(rows);
    // }

  }

  var _sinks = [_flushConsole,_flushSheet];


  // ---------------- public api ----------------

  // emit
  function trace(msg)     { _entry(LEVEL.TRACE, msg); }
  function debug(msg)     { _entry(LEVEL.DEBUG, msg); }
  function warn(msg)      { _entry(LEVEL.WARNING, msg); }
  function info(msg)      { _entry(LEVEL.INFO, msg); }
  function character(msg) { _entry(LEVEL.CHARACTER, msg); }
  function error(msg)     { _entry(LEVEL.ERROR, msg); }

  // flush buffer to sinks
  function flush() {
    var entries = _buffer;
    try {
      _sinks.forEach(s => s(entries));
    } finally {
      clear();
    }
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

    // control
    flush: flush,
    clear: clear,
    size:  size,
    peek:  peek,

    settings : settings,
  };
})();
