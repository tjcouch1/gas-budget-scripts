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
}
