/**
 * === Trigger Utilities ===
 * - Checks/installs installable triggers
 * - Writes status to the configured Event.Status.Cell
 */
var EDTriggers = (function () {
  var OPEN_FN = 'onOpenTriggered';   // your installable onOpen handler name
  var EDIT_FN = 'onEditTriggered';   // your installable onEdit handler name

  
  function checkInstalled() {
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
      if (!a1) { EDLogger.warning('[EDTriggers] No Trigger Status Cell configured'); return; }
      var rng = GSRange.resolveRange(a1, { ss: EDContext.context.ss });
      var msg = installed.ok
        ? 'INSTALLED'
        : 'NOT INSTALLED';
      rng.setValue(msg);
      EDLogger.info('[EDTriggers] Status → ' + a1 + ' : ' + msg);
    } catch (e) {
      EDLogger.error('[EDTriggers] writeStatus error: ' + (e && e.stack || e));
    }
  }

  function install() {
    var ss = SpreadsheetApp.getActive();
    var res = { createdOpen:false, createdEdit:false };
    try {
      var st = checkInstalled();
      if (!st.open) { ScriptApp.newTrigger(OPEN_FN).forSpreadsheet(ss).onOpen().create(); res.createdOpen = true; }
      if (!st.edit) { ScriptApp.newTrigger(EDIT_FN).forSpreadsheet(ss).onEdit().create(); res.createdEdit = true; }
      EDLogger.info('[EDTriggers] install: open=' + !st.open + ', edit=' + !st.edit);
    } catch (e) {
      EDLogger.error('[EDTriggers] install error: ' + (e && e.stack || e));
    }
    return res;
  }

  function setMenu() {
    try {
      var st = EDTriggers.checkInstalled();
      EDTriggers.writeStatus(st);

      // 2) Add menu
      SpreadsheetApp.getUi()
        .createMenu('ED Tools')
        .addItem(st.ok ? 'Reinstall Triggers' : 'Install Triggers', 'ED_Menu_InstallTriggers')
        .addItem('Recheck Trigger Status', 'ED_Menu_RecheckTriggers')
        .addToUi();

      EDLogger.info('[onOpen] menu added; triggers ok=' + st.ok);
    } catch (err) {
      EDLogger.error('[onOpen] ' + (err && err.stack || err));
    }
  }


  return {
    checkInstalled: checkInstalled,
    writeStatus: writeStatus,
    install: install,
    setMenu: setMenu

  };
})();

function installTriggers() {
  try {
    var r = EDTriggers.install();
    var st = EDTriggers.checkInstalled();
    EDTriggers.writeStatus(st);
    SpreadsheetApp.getActive().toast(st.ok ? 'Triggers installed' : 'Install failed', 'ED Tools', 5);
  } catch (e) {
    EDLogger.error('[ED_Menu_InstallTriggers] ' + (e && e.stack || e));
  }
}

function checkTriggers() {
  try {
    var st = EDTriggers.checkInstalled();
    EDTriggers.writeStatus(st);
    SpreadsheetApp.getActive().toast(st.ok ? 'Triggers OK' : 'Triggers missing', 'ED Tools', 5);
  } catch (e) {
    EDLogger.error('[ED_Menu_RecheckTriggers] ' + (e && e.stack || e));
  }
}

function onOpenTriggered(e) {
  GSPerf.start();
  GSPerf.monitor(new SheetOpenedEvent()).trigger();
  GSPerf.stop();
}

function onEditTriggered(e) {
  GSPerf.start();
  GSPerf.monitor(new CellEditedEvent(GSRange.a1FromEvent(e))).trigger();
  GSPerf.stop();
}
