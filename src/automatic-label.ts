const toLimitGrobal = 10;
const labelSpreadsheetId = '1EYnNthMez3zFZlkkk9xt3-BxPv3y9oW4y5l8qfnwXDM';

function automaticLabel() {
  var start = new Date();
  var email = Session.getActiveUser().getEmail();
  var executes = new Array();
  try {
    var spreadsheet = SpreadsheetApp.openById(labelSpreadsheetId);
    
    var sheetSettings = spreadsheet.getSheetByName('Settings');
    var rangeSettings = sheetSettings.getRange(2, 1, sheetSettings.getLastRow() - 1, sheetSettings.getLastColumn());
    var rowSettings = rangeSettings.getValues();
   
    var generalCondition = 'is:inbox';
    var hour = start.getHours();
    var minute = start.getMinutes();
    if (hour % 24 === 2 && minute < 15) {
      var threads = GmailApp.search(generalCondition, 0, 200);
      Logger.log('<span style="font-weight: bold;">Running night batch for (%s)...</span><br/>', threads.length);
    } else {
      var threads = GmailApp.search(generalCondition, 0, 50);
    }

    var customers = new Array();
    rowSettings.forEach(function(rowSetting) {
      var customerLabelName = rowSetting[0] as string;
      var productLabelNames = rowSetting[1] as string;
      var addressStr = rowSetting[2] as string;
      var subjectStr = rowSetting[3] as string;
      var toLimitLocal = rowSetting[4];
      if (!toLimitLocal) {
        toLimitLocal = toLimitGrobal;
      }

      var labelList = getLabelObjectList(productLabelNames);
      var customerLabelObject = getLabelObject(customerLabelName);
      labelList.push(customerLabelObject);
      
      var addressRegexs = [];
      if (addressStr && addressStr.trim().length != 0) {
        var addressStrs = addressStr.split(',');
        addressStrs.forEach(function(simgleAddressStr) {
          var addressRegex = new RegExp(simgleAddressStr.trim());
          addressRegexs.push(addressRegex);
        });
      }
      
      var subjectRegexs = [];
      if (subjectStr && subjectStr.trim().length != 0) {
        var subjectStrs = subjectStr.split(',');
        subjectStrs.forEach(function(simgleSubjectStr) {
          var subjectRegex = new RegExp(simgleSubjectStr.trim().replace(/[\\^$.*+?()[\]{}|]/g, '\\$&'));
          subjectRegexs.push(subjectRegex);
        });
      }
      
      customers.push({
        addressConditions: addressRegexs,
        subjectConditions: subjectRegexs,
        labels:labelList,
        toLimit:toLimitLocal
      });
    });
    
    var isExecuted = false;
    
    threads.forEach(function(thread) {
      var lastMessage = thread.getMessages()[thread.getMessageCount() - 1];
      var fromAddress = lastMessage.getFrom();
      var to = lastMessage.getTo() ? lastMessage.getTo().split(',') : [];
      var cc = lastMessage.getCc() ? lastMessage.getCc().split(',') : [];
      var date = thread.getLastMessageDate();
      var messageSubject = thread.getFirstMessageSubject();
      
      customers.forEach(function(customer) {
        var addressConditions = customer.addressConditions;
        var subjectConditions = customer.subjectConditions;
        var labels = customer.labels;
        var toLimit = customer.toLimit;
        
        if (toLimit < to.length) {
          return;
        }
        
        var match = false;
        addressConditions.forEach(function(condition) {
          if (fromAddress && fromAddress.match(condition)) {
            match = true;
          }
          to.forEach(function(toAddress) {
            if (toAddress && toAddress.match(condition)) {
              match = true;
            }
          });
          cc.forEach(function(ccAddress) {
            if (ccAddress && ccAddress.match(condition)) {
              match = true;
            }
          });
        });
        
        subjectConditions.forEach(function(condition) {
          if (messageSubject && messageSubject.match(condition)) {
            match = true;
          }
        });
        
        if (match) {
          isExecuted = true;
          Logger.log('Subject: %s, From: %s, Date: %s, Labels: %s<br/>', messageSubject, fromAddress, date, getLabelNames(labels));
          labels.forEach(function(label) {
            thread.addLabel(label);
          });
        }
      });
      
      var now = Date.now();
      var pastTime = (now - start.getTime())/1000;
      if (280 < pastTime) {
        throw 'TimeOutException';
      }
    });
    
    if (isExecuted) {
      var body = Logger.getLog();
      MailApp.sendEmail(email, 'GAS-Log:  Customer Label', body,
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
    MailApp.sendEmail(email, 'GAS-Log: Customer Label: ' + errorTitle, body,
                      { htmlBody: body, noReply: true });
  }
}
