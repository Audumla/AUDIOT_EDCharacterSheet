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

  setStatus(status) {
    EDContext.context.event.status.state = status;
  }

  closeEvent(state = EDContext.STATUS.UNKNOWN) {  
    this.setStatus(state);
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
      }
      finally {
        GSBatch.commit(EDContext.context.batch);

      }

      EDLogger.flush().forEach(b => {if (b != undefined) batches.add(b)});
    }
    finally {
      batches.forEach(b => GSBatch.commit(b));
      EDLogger.info(`${this._name} [${(new Date() - EDContext.context.date.time)/1000} secs]`);
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
    GSBatch.defaultMode = GSBatch.MODE.SIMPLE;
  }

  fireEvent() {
    EDLogger.info(`Sheet Loaded [${EDContext.context.ssid}]`)
    EDConfig.initialize({flushCache : true })
    this.setStatus(EDContext.STATUS.PROCESSING);
    EDTriggers.writeStatus(EDTriggers.checkInstalled());
    return EDContext.STATUS.COMPLETED;
  }

}

class CellEvent extends EDEvent {
  constructor(cell) {
    super();
    this._cell = cell;
    this._name = "Cell Event";
  }

}

class CellEditedEvent extends CellEvent {
  
  constructor(cell) {
    super(cell);
    this._name = "Cell Edited Event";
  }

  fireEvent() {
    var status = EDContext.STATUS.IGNORED;
    EDConfig.initialize({boot : false});

    this.setStatus(EDContext.STATUS.PROCESSING);
    
    if (EDProperties.event.byCell(this._cell) || EDConfig.configEdited(this._cell)) {
      status = EDContext.STATUS.COMPLETED;
    }

    if (status == EDContext.STATUS.IGNORED) {
        EDLogger.info(`Ignored Edited Cell [ ${this._cell} ]`);
    }

    return status
  }
}


class InstallTriggerEvent extends EDEvent {
  constructor() {
    super();
    this._name = "Install Triggers Event";
  }

  fireEvent() {
    
    try {
      EDConfig.initialize();
      this.setStatus(EDContext.STATUS.PROCESSING);

      var r = EDTriggers.install();
      var st = EDTriggers.checkInstalled();
      EDTriggers.writeStatus(st);
      EDLogger.notify(st.ok ? 'Triggers installed' : 'Trigger install failed', {title : 'ED Tools'});
      return EDContext.STATUS.COMPLETED;

    } finally {
      return EDContext.STATUS.FAILED;
    }

  }


}