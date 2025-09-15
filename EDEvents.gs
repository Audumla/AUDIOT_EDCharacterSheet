const SELECTED_TYPE = "SELECTED";
const CHECK_TYPE = "CHECK";

class CellEvent {
  constructor(cell) {
    this._cell = cell;
  }

  isCellMonitored(opts = DEFAULT_OPTS) {
    const row = GSUtils.Arr.findRowByColumn(this.cfg.EVENT_DEFINITIONS.values,2,this._cell);
    var mon = undefined
    if (row != undefined) {
      const [name,type,cell,property,active,inactive] = row;
      mon = {name,type,cell,property,active,inactive};
      this.logger.trace(`Monitored Cell Triggered [${name}][${cell}]`)

    }
    return mon;
  }

}

class CellEditedEvent extends CellEvent {
  
  constructor(cell) {
    super(cell);
  }

  fireEvent(opts = DEFAULT_OPTS) {
    opts.logger.trace(`Cell Edited [${this._cell}]`);
    if (!EDDefs.checkCachedDataChanged(this._cell,this)) {
      EDDefs.initializeDefinitions(false,true,this);
      const mon = this.isCellMonitored();
      if (CHECK_TYPE == mon?.type) {
        // perform the event and then reset the cell

        GSBatch.add.cell(this.batch,mon.cell,0);

      }
      else {
        opts.logger.trace(`Cell not Monitored [${this._cell}]`);
      }

    }
  }
}