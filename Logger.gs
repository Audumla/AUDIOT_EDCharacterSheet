
LOG_LEVEL = {
  TRACE :     {level : 0,  name : "TRACE" },
  DEBUG :     {level : 1,  name : "DEBUG" },
  WARNING :   {level : 2,  name : "WARNING" },
  INFO :      {level : 3,  name : "INFO" },
  CHARACTER : {level : 5,  name : "CHARACTER" },
  ERROR :     {level : 10, name : "ERROR" },
};

class SheetLogger {

  flush(logEntries,opts) {
    var loggerLevel = 0;
    try {
      if (opts?.batch) {
        loggerLevel = LOG_LEVEL[opts.cfg.component.processor.sheetLoggerLevel.value].level;
        GSBatch.insert.range(opts.batch,opts.cfg.SHEET_LOG.range,logEntries.length,)
      }
    }
    finally {
      // log to spreadsheet
    }
  }

}


class ConsoleLogger {
  
  logToConsole(logEntry,cfg) {
    const dt = Utilities.formatDate(logEntry.ts, cfg.TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss");
    const fmsg = Utilities.formatString("【 %s 】【 %s 】【 %s 】", dt,logEntry.logLevel.name,logEntry.msg); 
    console.info(fmsg);
  }

  flush(logEntries, opts = { cfg : Configuration }) {
    var loggerLevel = 0;
    try {
      loggerLevel = LOG_LEVEL[opts.cfg.component.processor.scriptLoggerLevel.value].level;
    }
    finally {
      logEntries.filter( e => e.logLevel.level >= loggerLevel).forEach( e => this.logToConsole(e,opts.cfg));
    }
  }

}

class BatchLogger {
  constructor(opts = {cfg : Configuration}) {
    this._opts = opts;
    this._messageLog = [];
    this._loggers = [];

    this._loggers.push(new ConsoleLogger());
    this._loggers.push(new SheetLogger());
  }
  
  log(logLevel, msg) {
    const logEntry = {
      logLevel : logLevel,
      ts : new Date(), 
      msg : msg
    };

    this._messageLog.push(logEntry);
    //console.info(JSON.stringify(logEntry));

  }

  trace(msg) {
    this.log(LOG_LEVEL.TRACE,msg);
  }

  error(msg) {
    this.log(LOG_LEVEL.ERROR,msg);
  }

  debug(msg) {
    this.log(LOG_LEVEL.DEBUG,msg);
  }

  info(msg) {
    this.log(LOG_LEVEL.INFO,msg);
  }

  warn(msg) {
    this.log(LOG_LEVEL.WARNING,msg);
  }

  flush() {
    this._loggers.forEach(l => l.flush(this._messageLog, this._opts));
  }

}

const DEFAULT_LOGGER = new BatchLogger(Configuration);
const logger = DEFAULT_LOGGER;