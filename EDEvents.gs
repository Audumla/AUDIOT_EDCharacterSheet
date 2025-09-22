const SELECTED_TYPE = "SELECTED";
const CHECK_TYPE = "CHECK";
EVENT_ID = "Event.ActiveID";
EVENT_STATUS = "Event.Status";

class EDEvent {
  constructor() {
    this._name = "Event";
  }

  openEvent() {
    EDContext.context.event.status = EDContext.STATUS.PROCESSING;
    EDLogger.info(`${this._name} [${EDContext.context.event.status.name}]`);
  }

  closeEvent() {  
    var batches = new Set();
    batches.add(EDContext.context.batch);
    try {
      if (EDContext.context.event.status != EDContext.STATUS.IGNORED) {
        EDLogger.trace(`Incremented EventID [${EDContext.context.event.activeID} : ${EDContext.context.event.activeID+1}]`)
        const eventIDr = GSUtils.Arr.findRowByColumn(EDContext.context.cfg.SETTINGS.values,0,EVENT_ID);
        eventIDr[1] = EDContext.context.event.activeID+1;
        EDDefs.setCacheData(EDContext.context.cfg.SETTINGS,EDContext.context);
        const eventIDloc = GSUtils.Arr.findRowByColumn(EDContext.context.cfg.PROPERTY_MAPPINGS.values,0,EVENT_ID);
        GSBatch.add.cell(EDContext.context.batch,eventIDloc[1],eventIDr[1]);
      }
      else {
        EDContext.context.event.activeID = 0;
      }
      EDLogger.info(`${this._name} [${EDContext.context.event.status.name}] [${(new Date() - EDContext.context.startTime)/1000} secs][${GSBatch.size(EDContext.context.batch)} bytes]`);
      EDLogger.flush().forEach(b => {if (b != undefined) batches.add(b)});
    }
    finally {
      batches.forEach(b => GSBatch.commit(b));
    }

  }

  trigger() {
    DEFAULT_OPTS = EDContext.initializeContext();
    const ctx = EDContext.context;
    try {
      this.openEvent(ctx);
      ctx.event.status = this.fireEvent(ctx);
      
    } catch(e) {
      ctx.event.status = EDContext.STATUS.FAILED;
      ctx.logger.error(e.stack);
    } finally {
      this.closeEvent(ctx);
    }
  }
}

class SheetOpenedEvent extends EDEvent {
  constructor() {
    super();
    this._name = "Sheet Opened Event";
  }

  fireEvent() {
    EDLogger.info(`Sheet Loaded [${EDContext.context.ssid}]`)
    EDDefs.initializeDefinitions(false,false,EDContext.context);
    return EDContext.STATUS.COMPLETED;
  }

}

class CellEvent extends EDEvent {
  constructor(cell) {
    super();
    this._cell = cell;
    this._name = "Cell Event";
  }

  isCellMonitored() {
    const row = GSUtils.Arr.findRowByColumn(EDContext.context.cfg.EVENT_DEFINITIONS.values,2,this._cell);
    var mon = undefined
    if (row != undefined) {
      const [name,type,cell,property,active,inactive] = row;
      mon = {name,type,cell,property,active,inactive};
      EDLogger.info(`Monitored Cell Triggered [${mon.name}][${this._cell}]`)
    }
    return mon;
  }

}

class CellEditedEvent extends CellEvent {
  
  constructor(cell) {
    super(cell);
    this._name = "Cell Edited Event";
  }

  fireEvent() {
    var status = EDContext.STATUS.COMPLETED;
    EDDefs.initializeDefinitions(false,true,this);
    const mon = this.isCellMonitored();
    if (mon == undefined && EDDefs.checkCachedDataChanged(this._cell,this)) {
      EDLogger.info(`Cached Data Edited [${this._cell}]`);
    }
    else {
      if (CHECK_TYPE == mon?.type) {
          EDLogger.info(`Activating [${mon.name}]`)
                // perform the event and then reset the cell
          EDDefs.loadDefinitions([EDContext.context.cfg.EVENT_PROPERTIES],true,false);
          GSBatch.add.cell(EDContext.context.batch,mon.cell,0);
      }
      else {
        EDLogger.info(`Ignored Edited Cell [ ${this._cell} ]`);
        status = EDContext.STATUS.IGNORED;
      }
    }  

    return status
  }
}