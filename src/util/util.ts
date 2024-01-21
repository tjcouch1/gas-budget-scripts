namespace Util {
  /**
   * Determine whether the object is a string
   * @param o object to determine if it is a string
   * @returns true if the object is a string; false otherwise
   *
   * Thanks to DRAX at https://stackoverflow.com/a/9436948
   */
  export function isString(o: unknown): o is string {
    return typeof o === "string" || o instanceof String;
  }

  export function isDate(o: unknown): o is Date {
    return o instanceof Date;
  }

  /**
   * Determine whether the two "date"-ish objects are equal
   * @param date1 Value from a range that we expect is a date. It could be a string, though, if it is unexpectedly a blank cell or something
   * @param date2 Value from a range that we expect is a date. It could be a string, though, if it is unexpectedly a blank cell or something
   * @returns
   */
  export function areDatesEqual(
    date1: GoogleAppsScript.Base.Date | string | undefined,
    date2: GoogleAppsScript.Base.Date | string | undefined
  ) {
    return isDate(date1) && isDate(date2)
      ? date1.getTime() === date2.getTime()
      : date1?.toString() === date2?.toString();
  }
}
