import { LogWrapper } from "./log-wrapper";
import { Utils } from "./utils";

// Automatic Archive Sample
// https://docs.google.com/spreadsheets/d/1QVVpwG5GuvKh7eZCGjZY5sQbh-7mn-CpvuLjqjFy9mY

function automaticArchive() {
  const start = new Date();
  const executes = new Array();
  try {
    // const archiveSpreadsheetId = "1QVVpwG5GuvKh7eZCGjZY5sQbh-7mn-CpvuLjqjFy9mY";
    const archiveSpreadsheetId = Utils.getProperyValue("ArchiveSpreadsheetId");
    // const searchMaxGrobal = 10;
    const searchMaxGrobal = parseInt(Utils.getProperyValue("SearchMaxGrobal"), 10);

    const spreadsheet = SpreadsheetApp.openById(archiveSpreadsheetId);

    const rowSettings = Utils.getSheetSettings(spreadsheet, "Settings");

    rowSettings.forEach((rowSetting) => {
      const settingName = rowSetting[0] as string;
      const generalCondition = rowSetting[1] as string;
      const delayDays = rowSetting[2] as number;
      const labelNames = rowSetting[3] as string;
      const detailConditionSheetName = rowSetting[4] as string;
      const labelsToBeRemovedSheetName = rowSetting[5] as string;
      const ignoreImportance = rowSetting[6] as boolean;
      const searchMax = rowSetting[7] ? rowSetting[7] as number : searchMaxGrobal;

      const maxDate = new Date(Date.now() - 86400000 * delayDays);

      const labels = Utils.getLabelObjectList(labelNames);

      const hour = start.getHours();
      const minute = start.getMinutes();
      let threads: GoogleAppsScript.Gmail.GmailThread[];

      if (hour % 24 === 0 && minute < 5) {
        threads = GmailApp.search(generalCondition, 0, 100);
        LogWrapper.log('<span style="font-weight: bold;">Running night batch for %s(%s)...</span><br/>',
          settingName, threads.length);
      } else {
        threads = GmailApp.search(generalCondition, 0, searchMax);
      }

      const rowConditions = Utils.getSheetSettings(spreadsheet, detailConditionSheetName);

      const conditions = new Array();

      rowConditions.forEach((rowCondition) => {
        const subjectStr = rowCondition[0] as string;
        let subjectRegex = null;
        if (subjectStr && subjectStr.trim().length !== 0) {
          subjectRegex = new RegExp(subjectStr);
        }
        const fromStr = rowCondition[1] as string;
        let fromRegex = null;
        if (fromStr && fromStr.trim().length !== 0) {
          fromRegex = new RegExp(fromStr);
        }
        conditions.push([subjectRegex, fromRegex]);
      });

      const removeLabels = new Array();

      const sheetLabels = spreadsheet.getSheetByName(labelsToBeRemovedSheetName);
      if (sheetLabels != null) {
        const rangeLabels = sheetLabels.getRange(2, 1, sheetLabels.getLastRow() - 1, sheetLabels.getLastColumn());
        const rowLabels = rangeLabels.getValues();

        rowLabels.forEach((rowLabel) => {
          const labelNameToBeRemoved = rowLabel[0] as string;
          const labelObjectToBeRemoved = GmailApp.getUserLabelByName(labelNameToBeRemoved);
          if (labelObjectToBeRemoved !== null) {
            removeLabels.push(labelObjectToBeRemoved);
          }
        });
      }

      let isExecuted = false;

      threads.forEach((thread) => {
        let toBeArchived = false;
        const messageSubject = thread.getFirstMessageSubject();
        const from = thread.getMessages()[0].getFrom();
        const date = thread.getLastMessageDate();

        conditions.forEach((condition) => {
          const subjectRegex = condition[0];
          const fromRegex = condition[1];
          if (subjectRegex !== null) {
            if (messageSubject !== null && messageSubject.match(subjectRegex) !== null) {
              if (fromRegex === null || from.match(fromRegex) !== null) {
                toBeArchived = true;
                labels.forEach((label) => {
                  thread.addLabel(label);
                });
              }
            }
          } else {
            if (fromRegex === null || from.match(fromRegex) !== null) {
              toBeArchived = true;
              labels.forEach((label) => {
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
          removeLabels.forEach((label) => {
            thread.removeLabel(label);
          });
          if (!isExecuted) {
            isExecuted = true;
            LogWrapper.log('<span style="font-weight: bold;">%s &gt;----</span><br/>', settingName);
          }
          LogWrapper.log("Subject: %s, From: %s, Date: %s<br/>", messageSubject, from, date);
        }

        const now = Date.now();
        const pastTime = (now - start.getTime()) / 1000;
        if (280 < pastTime) {
          throw new Error("TimeOutException");
        }
      });

      if (isExecuted) {
        executes.push(settingName);
        LogWrapper.log('<span style="font-weight: bold;">----&gt; %s</span><br/>', settingName);
      }
    });
    if (executes.length !== 0) {
      Utils.handleExecuteLog(`Automatic Archive: ${executes.join(", ")}`);
    }
  } catch (e) {
    Utils.handleException(e, "Automatic Archive");
  }
}
