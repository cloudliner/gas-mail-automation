import Utils from "./utils";

const deleteSpreadsheetId = "1VneCqoMD28HW93COYPxmV3FrjSL5IlVvwz4MngtSqHE";

function automaticDelete() {
  const start = new Date();
  const executes = new Array();

  try {
    const spreadsheet = SpreadsheetApp.openById(deleteSpreadsheetId);

    const sheetSettings = spreadsheet.getSheetByName("Settings");
    const rangeSettings = sheetSettings.getRange(2, 1, sheetSettings.getLastRow() - 1, sheetSettings.getLastColumn());
    const rowSettings = rangeSettings.getValues();

    rowSettings.forEach((rowSetting) => {
      const settingName = rowSetting[0] as string;
      const generalCondition = rowSetting[1] as string;
      const delayDays = rowSetting[2] as number;

      const maxDate = new Date();
      maxDate.setDate(maxDate.getDate() - delayDays);

      const y = maxDate.getFullYear();
      const m = maxDate.getMonth() + 1;
      const d = maxDate.getDate() + 1;

      // 1回で最大100件削除
      const threads = GmailApp.search(`${generalCondition} before: ${y}/${m}/${d}`, 0, 100);
      let isExecuted = false;

      threads.forEach((thread) => {
        let toBeDeleted = true;
        const messageSubject = thread.getFirstMessageSubject();
        const from = thread.getMessages()[0].getFrom();
        const date = thread.getLastMessageDate();

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
          Logger.log("Subject: %s, From: %s, Date: %s<br/>", messageSubject, from, date);
        }

        const now = Date.now();
        const pastTime = (now - start.getTime()) / 1000;
        if (280 < pastTime) {
          throw new Error("TimeOutException");
        }
      });

      if (isExecuted) {
        executes.push(settingName);
        Logger.log('<span style="font-weight: bold;">----&gt; %s</span><br/>', settingName);
      }
    });
    if (executes.length !== 0) {
      Utils.handleExecuteLog(`Automatic Delete: ${executes.join(", ")}`);
    }
  } catch (e) {
    Utils.handleException(e, "Automatic Delete");
  }
}
