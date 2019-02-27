export enum LogLevel {
  ERROR,
  WARN,
  INFO,
  DEBUG,
}

export class LogWrapper {
  public static getLog(title: string, level: LogLevel): string {
    const log = Logger.getLog();
    const content = log.split("<br/>");
    const message = `${title}: [${content.length}]`;
    const logObject = {
      content,
      message,
    };
    switch (level) {
      case LogLevel.ERROR:
        // tslint:disable-next-line:no-console
        console.error(logObject);
        break;
      case LogLevel.WARN:
        // tslint:disable-next-line:no-console
        console.warn(logObject);
        break;
      case LogLevel.INFO:
        // tslint:disable-next-line:no-console
        console.info(logObject);
        break;
      case LogLevel.DEBUG:
        // tslint:disable-next-line:no-console
        console.debug(logObject);
        break;
      default:
        // tslint:disable-next-line:no-console
        console.log(logObject);
    }
    return log;
  }

  public static log(format: string, ...values: any[]): GoogleAppsScript.Base.Logger {
    return Logger.log(format, ...values);
  }
}
