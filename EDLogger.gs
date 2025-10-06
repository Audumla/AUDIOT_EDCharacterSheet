var EDLogger = (function () {
  // ---------------- constants ----------------
  var LEVEL = {
    PERFORMANCE : { level: 1,  name: "PERF "       },
    TRACE       : { level: 5,  name: "TRACE"       },
    DEBUG       : { level: 6,  name: "DEBUG"       },
    WARNING     : { level: 7,  name: "WARN "       },
    INFO        : { level: 8,  name: "INFO "       },
    CHARACTER   : { level: 9,  name: "CHAR "       },
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
    //console.info(msg);
    const entry = { logLevel: levelObj, ts: new Date(), msg: msg ,state : EDContext.context.event.status.state}
    _buffer.push(entry);
    _sinks.forEach(s => s(entry));
  }

  // ---------------- sinks ----------------

  function _consoleSink(entry,opts) {
    var minLevel = LEVEL[settings.console.level ?? LEVEL.TRACE.name].level;
    if (entry.logLevel.level < minLevel) return;
    var dt = GSUtils.Date.formatDate(entry.ts);
    const lmsg = (typeof entry.msg === "string" ) ? entry.msg : JSON.stringify(entry.msg);
    var line = Utilities.formatString("【 %s 】【 %s 】【 %s 】【 %s 】",entry.logLevel.name, dt, entry.state, lmsg);
    console.info(line);

  }

  // console sink
  function _flushConsole(entries,opts) {
    entries.forEach(e => _consoleSink(e));
  }

  function _flushSheet(entries, opts) {

    if (EDContext.context.event.status.state != EDContext.STATUS.IGNORED) {
      var minLevel = LEVEL[settings.sheet.level ?? LEVEL.TRACE.name].level;

      var rows = [];
      for (var i = entries.length-1; i >= 0; i--) {
        var e = entries[i];
        if (e.logLevel.level < minLevel) continue;
        rows.push([
          EDContext.context.event.id.value,
          GSUtils.Date.formatDate(e.ts),
          e.logLevel.name.trim(),
          e.msg.trim()
        ]);
      }

      if (rows.length > 0) {
        const lBatch = (opts?.singleBatch ?? settings.sheet.batchMerge) ? EDContext.context.batch : GSBatch.newBatch(EDContext.context.ss);
        var clogs = GSBatch.load.rangesNow(settings.sheet.range);
        var logs = (clogs.length == 0) ? [] : clogs[0].values;
        const length = Math.min(logs.length+rows.length,settings.sheet.maxEntries);
        logs.unshift(...rows);
        logs.length = length;
        GSBatch.add.values(lBatch,settings.sheet.range,logs,EDContext.context);
        return lBatch;
      }
    }    
  
    return undefined;
  

  }

  var _bufferSinks = [_flushSheet];
  var _sinks = [_consoleSink];


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
      _bufferSinks.forEach(s => batches.push(s(entries,opts)));
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
