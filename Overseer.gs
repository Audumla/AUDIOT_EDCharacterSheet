function Overseer(configuration = null, activeSheet = null) {

  return { 
    ss : activeSheet == null ? SpreadsheetApp.getActiveSpreadsheet() : activeSheet,
    context : "INITIALIZATION",

    log : function(type, message) {
      if ((type == DEBUGLOG && this.debug) || (type != DEBUGLOG)) {
        this.messageLog.unshift([this.context,type,message]);
      }
    },

    initialize : function() {
      // load values that can be cached as script properties

    },

    getExecID : function() {
      return getCacheData(Configuration.ProcessingKeys.EXEC_ID_COUNT);
    },

    incrementExecID : function() {
      const id = this.getExecID()+1;
      const idLocation = getCacheData(Configuration.ProcessingKeys.EXEC_ID_LOCATION);

      setCacheData(Configuration.ProcessingKeys.EXEC_ID_COUNT,id);
      setByA1(idLocation,id);
    },
    
    getUI : function() {
      try {
        const ui = SpreadsheetApp.getUi();
        return ui;
      }
      catch (err) {
        this.log("ERROR",err)
      }

    },

    alert : function(text, buttons = null) {
      try {
        this.getUI().alert("Message",text,buttons == null ? this.getUI().ButtonSet.OK : buttons);
      }
      catch (err) {
        this.log("ERROR",err)
      }
    },

    prompt : function(prompt,response) {
      let answer = this.getUI().prompt(prompt,this.getUI().ButtonSet.OK_CANCEL);
      let button = answer.getSelectedButton();
      if (button != this.getUI().Button.CANCEL) {
        response = response.replace("RESPONSE",answer.getResponseText());
      }
      else {
        response = "CANCEL"
      }

      return response;
    },

  }
}
