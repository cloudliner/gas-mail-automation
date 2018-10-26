function automaticDelete() {
  var start = new Date();
  var email = Session.getActiveUser().getEmail();
  var executes = new Array();

  try {
    var spreadsheet = SpreadsheetApp.openById('1VneCqoMD28HW93COYPxmV3FrjSL5IlVvwz4MngtSqHE');
    
    var sheetSettings = spreadsheet.getSheetByName('Settings');
    var rangeSettings = sheetSettings.getRange(2, 1, sheetSettings.getLastRow() - 1, sheetSettings.getLastColumn());
    var rowSettings = rangeSettings.getValues();
    
    rowSettings.forEach(function(rowSetting) {
      var settingName = rowSetting[0];
      var generalCondition = rowSetting[1];
      var delayDays = rowSetting[2];

      var maxDate = new Date();
      maxDate.setDate(maxDate.getDate() - delayDays);
      
      var y = maxDate.getFullYear();
      var m = maxDate.getMonth() + 1;
      var d = maxDate.getDate() + 1;
      
      // 1回で最大100件削除
      var threads = GmailApp.search(generalCondition + ' before:' + y + '/' + m + '/' + d, 0, 100);
      var isExecuted = false;
      
      threads.forEach(function(thread) {
        var toBeDeleted = true;
        var messageSubject = thread.getFirstMessageSubject();
        var from = thread.getMessages()[0].getFrom();
        var date = thread.getLastMessageDate();
        
        if (maxDate < date) {
          toBeDeleted = false;
        }
        if (thread.isUnread()) {
          toBeDeleted = false;
        }
        if (thread.isImportant()) {
          toBeDeleted = false;
        }
        if (thread.hasStarredMessages()) {
          toBeDeleted = false;
        }
        
        if (toBeDeleted) {
          thread.moveToTrash();
          if (!isExecuted) {
            isExecuted = true;
            Logger.log('<span style="font-weight: bold;">%s &gt;----</span><br/>', settingName);
          }
          Logger.log('Subject: %s, From: %s, Date: %s<br/>', messageSubject, from, date);
        }
        
        var now = new Date();
        var pastTime = (now - start)/1000;
        if (280 < pastTime) {
          throw 'TimeOutException';
        }
      });
      
      if (isExecuted) {
        executes.push(settingName);
        Logger.log('<span style="font-weight: bold;">----&gt; %s</span><br/>', settingName);
      }
    });
    if (executes.length !== 0) {
      var executesTitle = executes.join(', ');
      var body = Logger.getLog();
      MailApp.sendEmail(email, 'GAS-Log: Automatic Delete: ' + executesTitle, body,
                        { htmlBody: body, noReply: true });
    }
  } catch(e) {
    var errorTitle = 'Error';
    if (e === 'TimeOutException') {
      errorTitle = 'TimeOut';
      Logger.log(e);
    } else {
      Logger.severe('%s: %s (line: %s, file: "%s") Stack: "%s"<br/>',
                    e.name||'', e.message||'', e.lineNumber||'', e.fileName||'', e.stack||'');
    }
    var body = Logger.getLog();
    MailApp.sendEmail(email, 'GAS-Log: Automatic Delete: ' + errorTitle, body,
                      { htmlBody: body, noReply: true });
  }
}
