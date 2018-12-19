import Utils from "./utils";

// Automatic Label Sample
// https://docs.google.com/spreadsheets/d/1-LVfu4oSbUHPUmX55Q-3bXMW-10tJPVRaVcnOhJpsa8

function automaticLabel() {
  const start = new Date();

  try {
    // const labelSpreadsheetId = "1-LVfu4oSbUHPUmX55Q-3bXMW-10tJPVRaVcnOhJpsa8";
    const labelSpreadsheetId = Utils.getProperyValue("LabelSpreadsheetId");
    // const toLimitGrobal = 10;
    const toLimitGrobal = parseInt(Utils.getProperyValue("ToLimitGrobal"), 10);

    const spreadsheet = SpreadsheetApp.openById(labelSpreadsheetId);

    const sheetSettings = spreadsheet.getSheetByName("Settings");
    const rangeSettings = sheetSettings.getRange(2, 1, sheetSettings.getLastRow() - 1, sheetSettings.getLastColumn());
    const rowSettings = rangeSettings.getValues();

    const generalCondition = "is:inbox";
    const hour = start.getHours();
    const minute = start.getMinutes();
    let threads: GoogleAppsScript.Gmail.GmailThread[];

    if (hour % 24 === 2 && minute < 15) {
      threads = GmailApp.search(generalCondition, 0, 200);
      Logger.log('<span style="font-weight: bold;">Running night batch for (%s)...</span><br/>', threads.length);
    } else {
      threads = GmailApp.search(generalCondition, 0, 50);
    }

    const customers: Array<{
      addressConditions: RegExp[];
      labels: GoogleAppsScript.Gmail.GmailLabel[];
      subjectConditions: RegExp[];
      toLimit: number}> = [];
    rowSettings.forEach((rowSetting) => {
      const customerLabelName = rowSetting[0] as string;
      const productLabelNames = rowSetting[1] as string;
      const addressStr = rowSetting[2] as string;
      const subjectStr = rowSetting[3] as string;
      const toLimitLocal = rowSetting[4] ? rowSetting[4] as number : toLimitGrobal;

      const labelList = Utils.getLabelObjectList(productLabelNames);
      const customerLabelObject = Utils.getLabelObject(customerLabelName);
      if (customerLabelObject) {
        labelList.push(customerLabelObject);
      }

      const addressRegexs: RegExp[] = [];
      if (addressStr && addressStr.trim().length !== 0) {
        const addressStrs = addressStr.split(",");
        addressStrs.forEach((simgleAddressStr) => {
          const addressRegex = new RegExp(simgleAddressStr.trim());
          addressRegexs.push(addressRegex);
        });
      }

      const subjectRegexs: RegExp[] = [];
      if (subjectStr && subjectStr.trim().length !== 0) {
        const subjectStrs = subjectStr.split(",");
        subjectStrs.forEach((simgleSubjectStr) => {
          const subjectRegex = new RegExp(simgleSubjectStr.trim().replace(/[\\^$.*+?()[\]{}|]/g, "\\$&"));
          subjectRegexs.push(subjectRegex);
        });
      }

      customers.push({
        addressConditions: addressRegexs,
        labels: labelList,
        subjectConditions: subjectRegexs,
        toLimit: toLimitLocal,
      });
    });

    let isExecuted = false;

    threads.forEach((thread) => {
      const lastMessage = thread.getMessages()[thread.getMessageCount() - 1];
      const fromAddress = lastMessage.getFrom();
      const to = lastMessage.getTo() ? lastMessage.getTo().split(",") : [];
      const cc = lastMessage.getCc() ? lastMessage.getCc().split(",") : [];
      const date = thread.getLastMessageDate();
      const messageSubject = thread.getFirstMessageSubject();

      customers.forEach((customer) => {
        const addressConditions = customer.addressConditions;
        const subjectConditions = customer.subjectConditions;
        const labels = customer.labels;
        const toLimit = customer.toLimit;

        if (toLimit < to.length) {
          return;
        }

        let match = false;
        addressConditions.forEach((condition) => {
          if (fromAddress && fromAddress.match(condition)) {
            match = true;
          }
          to.forEach((toAddress) => {
            if (toAddress && toAddress.match(condition)) {
              match = true;
            }
          });
          cc.forEach((ccAddress) => {
            if (ccAddress && ccAddress.match(condition)) {
              match = true;
            }
          });
        });

        subjectConditions.forEach((condition) => {
          if (messageSubject && messageSubject.match(condition)) {
            match = true;
          }
        });

        if (match) {
          isExecuted = true;
          Logger.log("Subject: %s, From: %s, Date: %s, Labels: %s<br/>",
            messageSubject, fromAddress, date, Utils.getLabelNames(labels));
          labels.forEach((label) => {
            thread.addLabel(label);
          });
        }
      });

      const now = Date.now();
      const pastTime = (now - start.getTime()) / 1000;
      if (280 < pastTime) {
        throw new Error("TimeOutException");
      }
    });

    if (isExecuted) {
      Utils.handleExecuteLog("Customer Label");
    }
  } catch (e) {
    Utils.handleException(e, "Automatic Label");
  }
}
