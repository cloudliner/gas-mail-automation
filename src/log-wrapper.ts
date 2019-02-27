export enum LogLevel {
  ERROR,
  WARN,
  INFO,
  DEBUG,
}

export class LogWrapper {
  public static getLog(level: LogLevel): string {
    const log = Logger.getLog();
    switch (level) {
      case LogLevel.ERROR:
        // tslint:disable-next-line:no-console
        console.error(log);
        break;
      case LogLevel.WARN:
        // tslint:disable-next-line:no-console
        console.warn(log);
        break;
      case LogLevel.INFO:
        // tslint:disable-next-line:no-console
        console.info(log);
        break;
      case LogLevel.DEBUG:
        // tslint:disable-next-line:no-console
        console.debug(log);
        break;
      default:
        // tslint:disable-next-line:no-console
        console.log(log);
    }
    return log;
  }

  public static log(format: string, ...values: any[]): GoogleAppsScript.Base.Logger {
    return Logger.log(format, ...values);
  }
}
