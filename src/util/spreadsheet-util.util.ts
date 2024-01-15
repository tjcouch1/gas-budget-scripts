namespace SpreadsheetUtil {
  /** Thanks to oldpedro at https://stackoverflow.com/a/67916559 */
  export function doRangesIntersect(
    r1: GoogleAppsScript.Spreadsheet.Range,
    r2: GoogleAppsScript.Spreadsheet.Range
  ) {
    if (r1.getSheet().getIndex() != r2.getSheet().getIndex()) return false;
    if (r1.getLastRow() < r2.getRow()) return false;
    if (r2.getLastRow() < r1.getRow()) return false;
    if (r1.getLastColumn() < r2.getColumn()) return false;
    if (r2.getLastColumn() < r1.getColumn()) return false;
    return true;
  }
}
