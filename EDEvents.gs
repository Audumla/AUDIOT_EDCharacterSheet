const SELECTED_TYPE = "SELECTED";
const CHECK_TYPE = "CHECK";
EVENT_ID = "Event.ID.Value";
EVENT_STATUS = "Event.Status.State";

class EDEvent {
  constructor() {
    this._name = "Event";
  }

  openEvent() {
    EDContext.initializeContext();
    EDLogger.info(this._name);

  }

  closeEvent(state = EDContext.STATUS.UNKNOWN) {  
    EDContext.context.event.status.state = state;
    var batches = new Set();

    try {
      try {
        if (EDContext.context.event.status.state != EDContext.STATUS.IGNORED) {
  //        EDLogger.trace(JSON.stringify(EDContext.context.event));        
          EDLogger.trace(`Incremented EventID [${EDContext.context.event.id.value} : ${EDContext.context.event.id.value+1}]`);
          EDContext.context.event.id.value++;
          EDContext.context.config.sheet.status.values = EDProperties.path.repack(EDContext.context.config.sheet.status)
          EDConfig.updateCache(EDContext.context.config.sheet);
          GSBatch.add.cell(EDContext.context.batch,EDContext.context.event.id.cell,EDContext.context.event.id.value);
        }
        else {
          EDContext.context.event.id.value = 0;
        }
        EDLogger.info(`${this._name} [${EDContext.context.event.status.state}] [${(new Date() - EDContext.context.date.time)/1000} secs]`);
      }
      finally {
        GSBatch.commit(EDContext.context.batch);
      }

      EDLogger.flush().forEach(b => {if (b != undefined) batches.add(b)});
    }
    finally {
      batches.forEach(b => GSBatch.commit(b));
    }

  }

  trigger() {
    var state = EDContext.STATUS.UNKNOWN;
    try {
      this.openEvent();
      state = this.fireEvent();
      
    } catch(e) {
      state = EDContext.STATUS.FAILED;
      EDLogger.error(e.stack);
    } finally {
      this.closeEvent(state);
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
    EDConfig.initialize({flushCache : true })
    EDContext.context.event.status.state = EDContext.STATUS.PROCESSING;
    EDTriggers.setMenu()
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
    const mon = GSUtils.Arr.findFirst(EDContext.context.mappings,"cell",this._cell);
    if (mon != undefined) {
      EDLogger.info(`Monitored Cell Triggered [${mon.event}][${this._cell}]`)
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
    var status = EDContext.STATUS.IGNORED;
    EDConfig.initialize();
    EDContext.context.event.status.state = EDContext.STATUS.PROCESSING;
    
    const monitored = this.isCellMonitored();
    if (monitored) {
      if (CHECK_TYPE == monitored?.type) {
          EDLogger.info(`Activating [${monitored.event}]`)
                // perform the event and then reset the cell
//          const rng = GSBatch.load.rangesNow([EDContext.context.config.EVENT_PROPERTIES_])

          EDLogger.debug(JSON.stringify(monitored));
          GSBatch.add.cell(EDContext.context.batch,monitored.cell,0);
          status = EDContext.STATUS.COMPLETED;
      }
    }
    else {
      if (EDConfig.checkConfigEdited(this._cell)) {
          status = EDContext.STATUS.COMPLETED;
      }
    }

    if (status == EDContext.STATUS.IGNORED) {
        EDLogger.info(`Ignored Edited Cell [ ${this._cell} ]`);
    }

    return status
  }
}