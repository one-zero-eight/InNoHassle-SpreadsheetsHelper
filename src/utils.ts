/** BEGIN src/utils.ts */
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let currentTargetSheet = spreadsheet.getActiveSheet();
let currentTargetRange = currentTargetSheet.getDataRange();
let currentTargetValues = currentTargetRange.getValues();
let mergedRangeRegistry: { [key: number]: { [key: number]: GoogleAppsScript.Spreadsheet.Range } } = {};

let _createMergedRangeRegistry = (mergedRanges: GoogleAppsScript.Spreadsheet.Range[]) => {
    // accessible by (row, column) in O(1)
    const mergedRangeRegistry: { [key: number]: { [key: number]: GoogleAppsScript.Spreadsheet.Range } } = {};

    for (let i = 0; i < mergedRanges.length; i++) {
        const mergedRange = mergedRanges[i];
        const notation = mergedRange.getA1Notation()
        const range = A1.rangeFromA1Notation(notation);
        for (let x = range.startRow; x <= range.endRow; x++) {
            if (!mergedRangeRegistry[x]) {
                mergedRangeRegistry[x] = {};
            }
            for (let y = range.startColumn; y <= range.endColumn; y++) {
                mergedRangeRegistry[x][y] = mergedRange;
            }
        }
    }
    return mergedRangeRegistry;
}

_createMergedRangeRegistry = Profiler.wrap(_createMergedRangeRegistry);


let _fastCheckMergedRows = (column: number, row: number) => {
    // using registry to check if cell is merged
    if (mergedRangeRegistry[column] && mergedRangeRegistry[column][row]) {
        return mergedRangeRegistry[column][row];
    }
    return null;
}

_fastCheckMergedRows = Profiler.wrap(_fastCheckMergedRows);

/** END src/utils.ts */
