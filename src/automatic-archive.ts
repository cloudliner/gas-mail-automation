const searchMaxGrobal = 10;
const archiveSpreadsheetId = '11GRPQyKcAVvmeLnFARytOmOr3vc0yBB4gP2NRFn7nmk';

function automaticArchive() {
  var start = new Date();
  var email = Session.getActiveUser().getEmail();
  var executes = new Array();
  try {
    var spreadsheet = SpreadsheetApp.openById(archiveSpreadsheetId);
    
    var sheetSettings = spreadsheet.getSheetByName('Settings');
    var rangeSettings = sheetSettings.getRange(2, 1, sheetSettings.getLastRow() - 1, sheetSettings.getLastColumn());
    var rowSettings = rangeSettings.getValues();
    
    rowSettings.forEach(function(rowSetting) {
      var settingName = rowSetting[0];
      var generalCondition = rowSetting[1] as string;
      var delayDays = rowSetting[2] as number;
      var labelNames = rowSetting[3] as string;
      var detailConditionSheetName = rowSetting[4] as string;
      var labelsToBeRemovedSheetName = rowSetting[5] as string;
      var ignoreImportance = rowSetting[6];
      var searchMax = rowSetting[7] as number;
      
      if (!searchMax) {
        searchMax = searchMaxGrobal;
      }

      var maxDate = new Date(Date.now() - 86400000 * delayDays);
      
      var labels = getLabelObjectList(labelNames);
      
      var hour = start.getHours();
      var minute = start.getMinutes();
      if (hour % 24 === 0 && minute < 15) {
        var threads = GmailApp.search(generalCondition);
        Logger.log('<span style="font-weight: bold;">Running night batch for %s(%s)...</span><br/>', settingName, threads.length);
      } else {
        var threads = GmailApp.search(generalCondition, 0, searchMax);
      }
      
      var sheetConditions = spreadsheet.getSheetByName(detailConditionSheetName);
      var rangeConditions = sheetConditions.getRange(2, 1, sheetConditions.getLastRow() - 1, sheetConditions.getLastColumn());
      var rowConditions = rangeConditions.getValues();
      
      var conditions = new Array(); 
      
      rowConditions.forEach(function(rowCondition) {
        var subjectStr = rowCondition[0] as string;
        var subjectRegex = null;
        if (subjectStr && subjectStr.trim().length != 0) {
          subjectRegex = new RegExp(subjectStr);
        }
        var fromStr = rowCondition[1] as string;
        var fromRegex = null;
        if (fromStr && fromStr.trim().length != 0) {
          fromRegex = new RegExp(fromStr);
        }
        conditions.push([subjectRegex, fromRegex]);
      });

      var removeLabels = new Array();

      var sheetLabels = spreadsheet.getSheetByName(labelsToBeRemovedSheetName);
      if (sheetLabels != null) {
        var rangeLabels = sheetLabels.getRange(2, 1, sheetLabels.getLastRow() - 1, sheetLabels.getLastColumn());
        var rowLabels = rangeLabels.getValues();
        
        rowLabels.forEach(function(rowLabel) {
          var labelNameToBeRemoved = rowLabel[0] as string;
          var labelObjectToBeRemoved = GmailApp.getUserLabelByName(labelNameToBeRemoved);
          if (labelObjectToBeRemoved !== null) {
            removeLabels.push(labelObjectToBeRemoved);
          }
        });
      }
      
      var isExecuted = false;
      
      threads.forEach(function(thread) {
        var toBeArchived = false;
        var messageSubject = thread.getFirstMessageSubject();
        var from = thread.getMessages()[0].getFrom();
        var date = thread.getLastMessageDate();
        
        conditions.forEach(function(condition) {
          var subjectRegex = condition[0];
          var fromRegex = condition[1];
          if (subjectRegex !== null) {
            if (messageSubject !== null && messageSubject.match(subjectRegex) !== null) {
              if (fromRegex === null || from.match(fromRegex) !== null) {
                toBeArchived = true;
                labels.forEach(function(label) {
                  thread.addLabel(label);
                });
              }
            }
          } else {
            if (fromRegex === null || from.match(fromRegex) !== null) {
              toBeArchived = true;
              labels.forEach(function(label) {
                thread.addLabel(label);
              });
            }
          }
        });
        
        if (maxDate < date && thread.isUnread()) {
          toBeArchived = false;
        }
        if (!ignoreImportance && thread.isImportant()) {
          toBeArchived = false;
        }
        if (thread.hasStarredMessages()) {
          toBeArchived = false;
        }
        
        if (toBeArchived) {
          thread.moveToArchive();
          removeLabels.forEach(function(label) {
            thread.removeLabel(label);
          });
          if (!isExecuted) {
            isExecuted = true;
            Logger.log('<span style="font-weight: bold;">%s &gt;----</span><br/>', settingName);
          }
          Logger.log('Subject: %s, From: %s, Date: %s<br/>', messageSubject, from, date);
        }
        
        var now = Date.now();
        var pastTime = (now - start.getTime())/1000;
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
      MailApp.sendEmail(email, 'GAS-Log: Automatic Archive: ' + executesTitle, body,
                        { htmlBody: body, noReply: true });
    }
  } catch(e) {
    var errorTitle = 'Error';
    if (e === 'TimeOutException') {
      errorTitle = 'TimeOut';
      Logger.log(e);
    } else {
      Logger.log('%s: %s (line: %s, file: "%s") Stack: "%s"<br/>',
                    e.name||'', e.message||'', e.lineNumber||'', e.fileName||'', e.stack||'');
    }
    var body = Logger.getLog();
    MailApp.sendEmail(email, 'GAS-Log: Automatic Archive: ' + errorTitle, body,
                      { htmlBody: body, noReply: true });
  }
}
