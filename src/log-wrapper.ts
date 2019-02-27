export enum LogLevel {
  DEBUG = "DEBUG",
  INFO = "INFO",
  WARN = "WARN",
  ERROR = "ERROR",
}

export class LogWrapper {
  public static getLog(level: LogLevel): string {
    const log = Logger.getLog();
    // tslint:disable-next-line:no-console
    console.log(log.replace("<br/>", "\n"), level);
    return log;
  }

  public static log(format: string, ...values: any[]): GoogleAppsScript.Base.Logger {
    return Logger.log(format, ...values);
  }
}
