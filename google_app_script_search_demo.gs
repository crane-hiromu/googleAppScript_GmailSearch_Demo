var EMAIL_MANAGER = EMAIL_MANAGER || {};

EMAIL_MANAGER = {

  INDEX: {
    START: 0,
    LIMIT: 20
  },

  init: function() {
    this.validateName = '任意の名前';
    this.workStartReg = new RegExp('出社連絡');
    this.workEndReg = new RegExp('日報');
    this.currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  appendRow: function() {
    this.setUpAppendRowStatus();
    this.appendRowForSheet();
  },

  setUpAppendRowStatus: function(){
    this.workStartArray = [];
    this.workEndArray = [];
    this.threads = GmailApp.search(this.validateName, this.INDEX.START, this.INDEX.LIMIT);
    this.messages = GmailApp.getMessagesForThreads(this.threads);
  },

  // 動かすためにそのままのコードで長いのでリファクタしてください
  appendRowForSheet: function() {
    for(var i = 0; i < this.messages.length; i++){
      var mailObject = this.messages[i][0];
      var mailTitle = mailObject.getSubject();
      var mailDate = mailObject.getDate();

      if(mailTitle.match(this.workStartReg))　{     
        this.setDateAndTitleForArray(this.workStartArray, mailDate, mailTitle);

      } else if(mailTitle.match(this.workEndReg)) {
        this.setDateAndTitleForArray(this.workEndArray, mailDate ,mailTitle);
      }

    }

    if(!this.threads.length>0) { return }

    this.appendDataToColumns('A', 'B', this.workStartArray);
    this.appendDataToColumns('C', 'D', this.workEndArray);
  },

  setDateAndTitleForArray: function(array, date, title) {
    array.push({
      'date': date,
      'title': title
    });
  },

  appendDataToColumns: function(line1, line2, array) {
    for(var i = 0; i < array.length;i++){
      var dateCell = this.currentSheet.getRange(line1 + (i+1));
      var titleCell = this.currentSheet.getRange(line2 + (i+1));

      dateCell.setValue("");
      titleCell.setValue("");

      dateCell.setValue(String(array[i]['date']));
      titleCell.setValue(array[i]['title']);
    }
  }

}

function generateSpreadSheet() {
  EMAIL_MANAGER.init();
  EMAIL_MANAGER.appendRow();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {name : "Generate Roster By Gmail" , functionName : "generateSpreadSheet"},
  ];
  spreadsheet.addMenu("Roster Automation Tool", entries);
}