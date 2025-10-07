/**
 * === Trigger Utilities ===
 * - Checks/installs installable triggers
 * - Writes status to the configured Event.Status.Cell
 */
var EDTriggers = (function () {
  var OPEN_FN = 'onOpenTriggered';   // your installable onOpen handler name
  var EDIT_FN = 'onEditTriggered';   // your installable onEdit handler name

  function checkInstalled() {
    // NOTE: Requires auth -> DO NOT call from simple onOpen
    var ts = ScriptApp.getProjectTriggers();
    var openOK = ts.some(function (t) {
      return t.getHandlerFunction() === OPEN_FN &&
             t.getEventType() === ScriptApp.EventType.ON_OPEN;
    });
    var editOK = ts.some(function (t) {
      return t.getHandlerFunction() === EDIT_FN &&
             t.getEventType() === ScriptApp.EventType.ON_EDIT;
    });
    return { open: openOK, edit: editOK, ok: openOK && editOK };
  }

  function writeStatus(installed) {
    try {
      var a1 = EDContext.context.triggers.status.cell;
      if (!a1) { EDLogger.warn('[EDTriggers] No Trigger Status Cell configured'); return; }
      var rng = GSRange.resolveRange(a1, { ss: EDContext.context.ss });
      var msg = installed.ok ? 'INSTALLED' : 'NOT INSTALLED';
      rng.setValue(msg);
      EDLogger.debug('ED Triggers Status : ' + msg);
    } catch (e) {
      EDLogger.error('EDTriggers error: ' + (e && e.stack || e));
    }
  }

  function install() {
    var ss = SpreadsheetApp.getActive();
    var ssId = ss && ss.getId ? ss.getId() : null;
    var res = { createdOpen:false, createdEdit:false };
    try {
      var st = checkInstalled();
      if (!st.open && ssId) {
        ScriptApp.newTrigger(OPEN_FN).forSpreadsheet(ssId).onOpen().create();
        res.createdOpen = true;
      }
      if (!st.edit && ssId) {
        ScriptApp.newTrigger(EDIT_FN).forSpreadsheet(ssId).onEdit().create();
        res.createdEdit = true;
      }
      EDLogger.trace('EDTriggers: open=' + !st.open + ', edit=' + !st.edit);
    } catch (e) {
      EDLogger.error('EDTriggers error: ' + (e && e.stack || e));
    }
    return res;
  }

  function setMenu() {
    SpreadsheetApp.getUi()
      .createMenu('ED Tools')
      .addItem('Install Triggers', 'installTriggers') 
      .addItem('Attack', "fireMeleeAttack")
      .addToUi();
    EDLogger.debug('ED Menu added'); // keep original message text
  }

  return {
    checkInstalled: checkInstalled,
    writeStatus: writeStatus,
    install: install,
    setMenu: setMenu
  };
})();

function perfTrigger(event) {
  GSPerf.start();
  GSPerf.monitor(event).trigger();
  GSPerf.stop();  
}

/** Sheet Triggers */
function onOpen(e) {
  EDTriggers.setMenu(); 
}

function installTriggers() {
  perfTrigger(new InstallTriggerEvent());
}

function onOpenTriggered(e) {
  perfTrigger(new SheetOpenedEvent());
}

function onEditTriggered(e) {

  perfTrigger(new CellEditedEvent(GSRange.a1FromEvent(e)));
}

function fireMeleeAttack(){
  new CellEditedEvent("Visual!$AA$15").trigger();
}

