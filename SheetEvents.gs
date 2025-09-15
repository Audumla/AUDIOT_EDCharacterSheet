  // Preserve any extra fields on opts; only fill cfg/logger if missing
  function resolveOpts(opts = {}, dOpts = {
    cfg    : Configuration,
    logger : DEFAULT_LOGGER,
    batch  : GSBatch.new(SpreadsheetApp.getActive())
  }) {
    opts.cfg    = opts.cfg    || dOpts.cfg;
    opts.logger = opts.logger || dOpts.logger;
    opts.batch  = opts.batch  || dOpts.batch;
    return opts;
  }


/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopen
 */
function onOpenTriggered(e) {
  const opts = resolveOpts();
  try {
    opts.logger.debug("Sheet Loaded ["+SpreadsheetApp.getActive().getId()+"]")
    EDDefs.initializeDefinitions(false,false,opts);
  } finally {
    opts.logger.flush();
  }

}

function onEditTriggered(e) {
  const opts = resolveOpts();
  try {
    const editEvent = new CellEditedEvent(a1FromEvent(e),opts);
    editEvent.fireEvent();
  } finally {
    opts.logger.flush();
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