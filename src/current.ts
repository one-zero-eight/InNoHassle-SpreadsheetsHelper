namespace Current {
  let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | undefined;
  let targetSheet: GoogleAppsScript.Spreadsheet.Sheet | undefined;
  let targetSheetId: number | undefined;
  let targetRange: GoogleAppsScript.Spreadsheet.Range | undefined;
  let targetValues: any[][] | undefined;

  export function getSpreadsheet() {
    if (spreadsheet) {
      return spreadsheet;
    }
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet;
  }

  export function getTargetSheet() {
    if (targetSheet) {
      return targetSheet;
    }
    targetSheet = getSpreadsheet().getActiveSheet();
    return targetSheet;
  }

  export function getTargetSheetId() {
    if (targetSheetId) {
      return targetSheetId;
    }
    targetSheetId = getTargetSheet().getSheetId();
    return targetSheetId;
  }

  export function getTargetRange() {
    if (targetRange) {
      return targetRange;
    }
    targetRange = getTargetSheet().getDataRange();
    return targetRange;
  }

  export function getTargetValues() {
    if (targetValues) {
      return targetValues;
    }
    targetValues = getTargetRange().getValues();
    return targetValues;
  }

  export function setTargetRange(range: GoogleAppsScript.Spreadsheet.Range) {
    targetRange = range;
    targetValues = range.getValues();
  }
}
