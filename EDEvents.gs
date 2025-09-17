const SELECTED_TYPE = "SELECTED";
const CHECK_TYPE = "CHECK";
EVENT_ID = "EventID";

class EDEvent {
  constructor() {
    this._name = "Event";
  }

  openEvent() {
    EDContext.context.status = EDContext.STATUS.PROCESSING;
    EDLogger.info(`${this._name} [${EDContext.context.status.name}]`);
  }

  closeEvent() {  
    if (EDContext.context.status != EDContext.STATUS.IGNORED) {
      EDLogger.trace(`Incremented EventID [${EDContext.context.eventID} : ${EDContext.context.eventID+1}]`)
      const eventIDr = GSUtils.Arr.findRowByColumn(EDContext.context.cfg.SETTINGS.values,0,EVENT_ID);
      eventIDr[1] = EDContext.context.eventID+1;
      EDDefs.setCacheData(EDContext.context.cfg.SETTINGS,EDContext.context);
      const eventIDloc = GSUtils.Arr.findRowByColumn(EDContext.context.cfg.PROPERTY_MAPPINGS.values,0,EVENT_ID);
      GSBatch.add.cell(EDContext.context.batch,eventIDloc[1],eventIDr[1]);
    }
    else {
      EDContext.context.eventID = 0;
    }
    EDLogger.info(`${this._name} [${EDContext.context.status.name}] [${(new Date() - EDContext.context.startTime)/1000} secs]`);
    EDLogger.flush();
    GSBatch.commit(EDContext.context);

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
    if (!EDDefs.checkCachedDataChanged(this._cell,this)) {
      EDDefs.initializeDefinitions(false,true,this);
      const mon = this.isCellMonitored();
      if (CHECK_TYPE == mon?.type) {
        // perform the event and then reset the cell
        EDLogger.info(`Monitored Cell Triggered [${mon.name}][${this._cell}]`)
        GSBatch.add.cell(EDContext.context.batch,mon.cell,0);
      }
      else {
        EDLogger.info(`Ignored Edited Cell [ ${this._cell} ]`);
        status = EDContext.STATUS.IGNORED;
     }
    }
    else {
      EDLogger.info(`Cached Data Edited [${this._cell}]`);
    }

    return status
  }
}