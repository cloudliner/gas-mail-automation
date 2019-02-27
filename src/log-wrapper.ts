export default class LogWrapper {
  public static getLog(): string {
    const log = Logger.getLog();
    // tslint:disable-next-line:no-console
    console.log(log);
    return log;
  }

  public static log(format: string, ...values: any[]): GoogleAppsScript.Base.Logger {
    return Logger.log(format, ...values);
  }
}
