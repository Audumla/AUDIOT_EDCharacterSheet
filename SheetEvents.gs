  
const DEFAULT_OPTS = {
}

function initializeOpts() {
  DEFAULT_OPTS.cfg    = Configuration;
  DEFAULT_OPTS.logger = Logger;
  DEFAULT_OPTS.ss     = SpreadsheetApp.getActive();
  DEFAULT_OPTS.ssid   = DEFAULT_OPTS.ss.getId();
  DEFAULT_OPTS.batch  = GSBatch.newBatch(DEFAULT_OPTS.ss);
  DEFAULT_OPTS.tz     = DEFAULT_OPTS.ss.getSpreadsheetTimeZone();
  return DEFAULT_OPTS;
}

function resolveOpts(opts = {}, dOpts = DEFAULT_OPTS) {
  opts.cfg    = opts.cfg ? opts.cfg : dOpts.cfg;
  opts.logger = opts.logger ? opts.logger : dOpts.logger;
  opts.ss     = opts.ss ? opts.ss : dOpts.ss;
  opts.ssid   = opts.ssid ? opts.ssid : dOpts.ssid;
  opts.batch  = opts.batch ? opts.batch : dOpts.batch;
  opts.tz     = opts.tz ? opts.tz : dOpts.tz;
  return opts;
}


/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopen
 */
function onOpenTriggered(e) {
  const opts = initializeOpts();
  try {
    opts.logger.debug("Sheet Loaded ["+SpreadsheetApp.getActive().getId()+"]")
    EDDefs.initializeDefinitions(false,false,opts);
  } finally {
    opts.logger.flush(opts);
  }

}

function onEditTriggered(e) {
  const opts = initializeOpts();
  try {
    const editEvent = new CellEditedEvent(GSRange.a1FromEvent(e));
    editEvent.fireEvent(opts);
  } finally {
    opts.logger.flush(opts);
    GSBatch.commit(opts.batch);
  }
}
  

function onSelectionChange__(e) {
  	
  if (e) return;
  const r = e.range;

  const snap = {
    sid: r.getSheet().getSheetId(), 
    a1:  a1FromEvent(e),
    f:   r.getFormula() || ""      // previous formula ("" if value)
  };

  setCacheData(Configuration.ProcessingKeys.LAST_SELECTED_CONTENT,snap);
  const monitored = JSON_TO_ARRAY(getCacheData(Configuration.data.MONITORED_CELLS.name));
  const found = findRowByColumn(monitored,0,snap.a1);
  if (found != undefined) {
    let overseer = new Overseer(Configuration,SpreadsheetApp.getActiveSpreadsheet());
    r.setValue(found[2]);
    const properties = JSON_TO_ARRAY(getCacheData(Configuration.data.STATIC_PROPERTIES.name));
    const property = found[4];
    const propLocation = findRowByColumn(properties,0,property);
    if (propLocation != undefined && propLocation[1] != "" ) {
      SpreadsheetApp.getActive().getRange(propLocation[1]).setValue(found[5]);
    }
    SpreadsheetApp.flush(); 
    r.setValue(found[3]);
    overseer.incrementExecID();
    if (propLocation != undefined && propLocation[1] != "" ) {
      SpreadsheetApp.getActive().getRange(propLocation[1]).setValue(found[6]);
    }
  }
  
}

function onChange(e) {
   

}