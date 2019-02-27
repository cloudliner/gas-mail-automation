import { LogLevel, LogWrapper } from "./log-wrapper";

export class Utils {
  public static getProperyValue(property: string): string {
    if (property) {
      const value = PropertiesService.getScriptProperties().getProperty(property);
      if (value) {
        return value;
      }
    }
    const error = new Error("NoProperyException");
    error.message = `No property value for: ${property}`;
    throw error;
  }

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
    const title = `GAS-Log: ${subTitle}`;
    const htmlBody = LogWrapper.getLog(title, LogLevel.INFO);
    MailApp.sendEmail(email, title, htmlBody,
                      { htmlBody, noReply: true });
  }

  public static handleException(e: any, subTitle: string) {
    const email = Session.getActiveUser().getEmail();
    let errorTitle = "Error";
    if (e.message === "TimeOutException") {
      errorTitle = "TimeOut";
      LogWrapper.log(e);
    } else {
      LogWrapper.log('%s: %s (line: %s, file: "%s") Stack: "%s"<br/>',
                    e.name || "", e.message || "", e.lineNumber || "", e.fileName || "", e.stack || "");
    }
    const title = `GAS-Log: ${subTitle}: ${errorTitle}`;
    const htmlBody = LogWrapper.getLog(title, LogLevel.ERROR);
    MailApp.sendEmail(email, title, htmlBody,
                      { htmlBody, noReply: true });
  }
}
