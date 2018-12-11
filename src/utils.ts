export default class Utils {
  public static getLabelObject(labelName: string): GoogleAppsScript.Gmail.GmailLabel | null {
    let labelObject = null;
    if (labelName && labelName.trim().length !== 0) {
      labelName = labelName.trim();
      labelObject = GmailApp.getUserLabelByName(labelName) ?
        GmailApp.getUserLabelByName(labelName) : GmailApp.createLabel(labelName);
    }
    return labelObject;
  }

  public static getLabelObjectList(labelNames: string): GoogleAppsScript.Gmail.GmailLabel[] {
    const labelNameList = labelNames ? labelNames.split(",") : [];
    const labelList = [] as GoogleAppsScript.Gmail.GmailLabel[];
    labelNameList.forEach((labelName) => {
      const labelObject = Utils.getLabelObject(labelName);
      if (labelObject) {
        labelList.push(labelObject);
      }
    });
    return labelList;
  }

  public static getLabelNames(labels: GoogleAppsScript.Gmail.GmailLabel[]) {
    let labelNames = "";
    labels.forEach((label) => {
      labelNames = labelNames + label.getName() + ", ";
    });
    return labelNames;
  }

  public static handleExecuteLog(subTitle: string) {
    const email = Session.getActiveUser().getEmail();
    const htmlBody = Logger.getLog();
    MailApp.sendEmail(email, `GAS-Log: ${subTitle}`, htmlBody,
                      { htmlBody, noReply: true });
  }

  public static handleException(e: any, subTitle: string) {
    const email = Session.getActiveUser().getEmail();
    let errorTitle = "Error";
    if (e.message === "TimeOutException") {
      errorTitle = "TimeOut";
      Logger.log(e);
    } else {
      Logger.log('%s: %s (line: %s, file: "%s") Stack: "%s"<br/>',
                    e.name || "", e.message || "", e.lineNumber || "", e.fileName || "", e.stack || "");
    }
    const htmlBody = Logger.getLog();
    MailApp.sendEmail(email, `GAS-Log: ${subTitle}: ${errorTitle}`, htmlBody,
                      { htmlBody, noReply: true });
  }
}
