function automaticLabelTest() {
  console.log(`automaticLabelTest`);
}

var toLimitGrobal = 10;

function getLabelObject(labelName) {
  var labelObject = null;
  if (labelName && labelName.trim().length != 0) {
    labelName = labelName.trim();
    labelObject = GmailApp.getUserLabelByName(labelName) ? GmailApp.getUserLabelByName(labelName) : GmailApp.createLabel(labelName);
  }
  return labelObject;
}

function getLabelObjectList(labelNames) {
  var labelNameList = labelNames ? labelNames.split(',') : [];
  var labelList = [];
  labelNameList.forEach(function(labelName) {
    var labelObject = getLabelObject(labelName);
    labelList.push(labelObject);
  });
  return labelList;
}

function getLabelNames(labels) {
  var labelNames = '';
  labels.forEach(function(label) {
    labelNames = labelNames + label.getName() + ', ';
  });
  return labelNames;
}

function automaticLabel() {
  var start = new Date();
  var email = Session.getActiveUser().getEmail();
  var executes = new Array();
  try {
    var spreadsheet = SpreadsheetApp.openById('1EYnNthMez3zFZlkkk9xt3-BxPv3y9oW4y5l8qfnwXDM');
    
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
      var customerLabelName = rowSetting[0];
      var productLabelNames = rowSetting[1];
      var addressStr = rowSetting[2];
      var subjectStr = rowSetting[3];
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
      
      var now = new Date();
      var pastTime = (now - start)/1000;
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
      Logger.severe('%s: %s (line: %s, file: "%s") Stack: "%s"<br/>',
                    e.name||'', e.message||'', e.lineNumber||'', e.fileName||'', e.stack||'');
    }
    var body = Logger.getLog();
    MailApp.sendEmail(email, 'GAS-Log: Customer Label: ' + errorTitle, body,
                      { htmlBody: body, noReply: true });
  }
}
