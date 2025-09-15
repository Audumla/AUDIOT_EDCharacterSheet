var Logger = (function () {
  // ---------------- constants ----------------
  var LEVEL = {
    TRACE     : { level: 0,  name: "TRACE"     },
    DEBUG     : { level: 1,  name: "DEBUG"     },
    WARNING   : { level: 2,  name: "WARNING"   },
    INFO      : { level: 3,  name: "INFO"      },
    CHARACTER : { level: 5,  name: "CHARACTER" },
    ERROR     : { level: 10, name: "ERROR"     },
  };

  // ---------------- state ----------------
  var _buffer = [];      // [{ logLevel, ts, msg }]

  // ---------------- helpers ----------------
  function _entry(levelObj, msg) {
    _buffer.push({ logLevel: levelObj, ts: new Date(), msg: String(msg) });
  }

  // ---------------- sinks ----------------

  // console sink
  function _flushConsole(entries, opts = DEFAULT_OPTS) {
    var enabled = opts.cfg?.component?.processor?.scriptLoggerEnabled?.value ?? true;
    if (enabled) {
      var minLevel = opts.cfg?.component?.processor?.scriptLoggerLevel?.value ?? 0;
      for (var i = 0; i < entries.length; i++) {
        var e = entries[i];
        if (e.logLevel.level < minLevel) continue;
        var dt = GSUtils.Date.formatDate(e.ts, (opts && opts.tz));
        var line = Utilities.formatString("【 %s 】【 %s 】【 %s 】", dt, e.logLevel.name, e.msg);
        // Use console.info for uniformity
        console.info(line);
      }
    }

  }

  // sheet sink (skeleton; fill with your GSBatch call if desired)
  function _flushSheet(entries, opts = DEFAULT_OPTS) {
    var enabled = opts.cfg?.component?.processor?.scriptLoggerEnabled?.value ?? true;
    if (enabled) {

      var minLevel = opts.cfg?.component?.processor?.sheetLoggerLevel?.value ?? 0;

      // TODO: if you want to batch to a sheet, filter first:
      var rows = [];
      for (var i = 0; i < entries.length; i++) {
        var e = entries[i];
        if (e.logLevel.level < minLevel) continue;
        rows.push([
          GSUtils.Date.formatDate(e.ts, (opts && opts.tz)),
          e.logLevel.name,
          e.msg
        ]);
      }



      // Example placeholder (uncomment/replace with your implementation):
      if (opts && opts.batch && rows.length) {
        GSBatch.insert.range(opts.batch, GSRange.extendA1(opts.cfg.SHEET_LOG.range, { top:-1, bottom:rows.length, left:0, right:3 } ),SpreadsheetApp.Dimension.ROWS);
        GSBatch.add.values(opts.batch,opts.cfg.SHEET_LOG.range,rows,{autoSize:true,...opts});
      }

      // If you want a simple direct write:
      // if (opts && opts.sheetA1 && rows.length) {
      //   var rng = GSRange.resolveRange(opts.sheetA1, opts);
      //   var sheet = rng.getSheet();
      //   sheet.getRange(rng.getRow(), rng.getColumn(), rows.length, rows[0].length).setValues(rows);
      // }
    }
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
  function flush(opts = DEFAULT_OPTS) {
    var entries = _buffer;
    try {
      _sinks.forEach(s => s(entries,opts));
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
    peek:  peek
  };
})();
