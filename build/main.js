/** END src/interfaces.ts */
/** BEGIN src/settings.ts */
var Settings;
(function (Settings) {
    let _settings = null;
    let _settingsRange;
    function createSettingsRange() {
        // new sheet with the name "InNoHassle"
        const sheet = Current.getSpreadsheet().insertSheet("InNoHassle");
        // Add innohassle logo
        const blob = Utilities.newBlob(Utilities.base64Decode(innohassleLogoEncoded), "image/png", "innohassle.png");
        // merge cells and set "Settings" as the title
        sheet
            .getRange("A1:F1")
            .merge()
            .setValue("Settings")
            .setFontWeight("bold")
            .setFontStyle("italic")
            .setHorizontalAlignment("center")
            .setFontSize(14);
        // new named range with the name "Settings": "InNoHassle!A2:F1002"
        const header = sheet
            .getRange("A2:E2")
            .setValues([["subjects", "groups", "courses", "locations", "teachers"]]);
        // Bold italics
        header.setFontWeight("bold");
        header.setFontStyle("italic");
        header.setHorizontalAlignment("center");
        const range = sheet.getRange("A2:F1002");
        // Set type to text
        range.setNumberFormat('"text"');
        // Make cell borders
        range.setBorder(true, true, true, true, true, true);
        // Freeze the first row
        sheet.setFrozenRows(1);
        // Name the range
        Current.getSpreadsheet().setNamedRange("Settings", range);
        const logo = sheet.insertImage(blob, 10, 10);
        // logo.setWidth(512);
        // logo.setHeight(512);
        // Add a link to the website
        // sheet.getRange("")
        range.activate();
    }
    Settings.createSettingsRange = createSettingsRange;
    function getSettingsRange() {
        if (_settingsRange) {
            return _settingsRange;
        }
        _settingsRange = Current.getSpreadsheet().getRangeByName("Settings");
        return _settingsRange;
    }
    Settings.getSettingsRange = getSettingsRange;
    function getSettings() {
        if (_settings) {
            return _settings;
        }
        // get named range "Settings" from the spreadsheet
        const settingsRange = getSettingsRange();
        if (!settingsRange) {
            return null;
        }
        const settings = {
            subjects: new Set(),
            groups: new Set(),
            courses: new Set(),
            locations: new Set(),
            teachers: new Set(),
        };
        // iterate over columns
        for (let column = 0; column < settingsRange.getWidth(); column++) {
            const columnRange = settingsRange.offset(0, column, settingsRange.getHeight(), 1);
            const columnValues = columnRange.getValues();
            const columnName = columnValues[0][0];
            if (!columnName || !settings[columnName]) {
                continue;
            }
            // iterate over rows
            for (let row = 1; row < columnValues.length; row++) {
                const value = columnValues[row][0];
                if (value) {
                    settings[columnName].add(value.toString());
                }
            }
        }
        _settings = settings;
        return settings;
    }
    Settings.getSettings = getSettings;
    Settings.goToSettings = () => {
        const settingsRange = getSettingsRange();
        if (settingsRange) {
            settingsRange.activate();
        }
    };
})(Settings || (Settings = {}));
/** END src/settings.ts */
/** BEGIN src/profiling.ts */
var Profiler;
(function (Profiler) {
    const _executionTimes = {};
    function wrap(func, name = func.name) {
        if (!_executionTimes[name]) {
            _executionTimes[name] = {
                executionTime: null,
                calls: 0,
                cumulativeExecutionTime: 0,
            };
        }
        return (...args) => {
            const start = Date.now();
            const output = func(...args);
            const end = Date.now();
            const executionTime = end - start;
            _executionTimes[name].executionTime = executionTime;
            _executionTimes[name].calls += 1;
            _executionTimes[name].cumulativeExecutionTime += executionTime;
            return output;
        };
    }
    Profiler.wrap = wrap;
    function format() {
        const keys = Object.keys(_executionTimes);
        const sortedKeys = keys.sort((a, b) => {
            const aTime = _executionTimes[a].cumulativeExecutionTime;
            const bTime = _executionTimes[b].cumulativeExecutionTime;
            return bTime - aTime;
        });
        let output = "";
        for (const key of sortedKeys) {
            const { executionTime, calls, cumulativeExecutionTime } = _executionTimes[key];
            output += `${key}: ${executionTime}ms (cumulative: ${cumulativeExecutionTime}ms, calls: ${calls})\n`;
        }
        return output;
    }
    Profiler.format = format;
})(Profiler || (Profiler = {}));
/** END src/profiling.ts */
/** BEGIN src/a1notation.ts */
var A1;
(function (A1) {
    function offsetFromA1Notation(A1Notation) {
        // start:end
        const match = A1Notation.match(/(^[A-Z]+[0-9]+)|([A-Z]+[0-9]+$)/gm);
        if (match.length !== 2) {
            throw new Error("The given value was invalid. Cannot convert Google Sheet A1 notation to indexes");
        }
        const start_notation = match[0];
        const end_notation = match[1];
        const start = cellFromA1Notation(start_notation);
        const end = cellFromA1Notation(end_notation);
        return {
            startRow: start.row,
            startColumn: start.column,
            endRow: end.row,
            endColumn: end.column,
        };
    }
    A1.offsetFromA1Notation = offsetFromA1Notation;
    function cellFromA1Notation(A1Notation) {
        const match = A1Notation.match(/(^[A-Z]+)|([0-9]+$)/gm);
        if (match.length !== 2) {
            throw new Error("The given value was invalid. Cannot convert Google Sheet A1 notation to indexes");
        }
        const column_notation = match[0];
        const row_notation = match[1];
        const column = fromColumnA1Notation(column_notation);
        const row = fromRowA1Notation(row_notation);
        return { row, column };
    }
    function fromRowA1Notation(A1Row) {
        const num = parseInt(A1Row, 10);
        if (Number.isNaN(num)) {
            throw new Error("The given value was not a valid number. Cannot convert Google Sheet row notation to index");
        }
        return num;
    }
    function fromColumnA1Notation(A1Column) {
        const A = "A".charCodeAt(0);
        let output = 0;
        for (let i = 0; i < A1Column.length; i++) {
            const next_char = A1Column.charAt(i);
            const column_shift = 26 * i;
            output += column_shift + (next_char.charCodeAt(0) - A);
        }
        return output + 1;
    }
    function offsetToA1Notation(offset) {
        return rangeToA1Notation(offset.row, offset.column, offset.numRows, offset.numColumns);
    }
    A1.offsetToA1Notation = offsetToA1Notation;
    function rangeToA1Notation(row, column, numRows, numColumns) {
        Logger.log("row: " + row);
        Logger.log("column: " + column);
        const start = toA1Notation(row, column);
        const end = toA1Notation(row + numRows - 1, column + numColumns - 1);
        return `${start}:${end}`;
    }
    function toA1Notation(row, column) {
        const row_notation = row.toString();
        const column_notation = toColumnA1Notation(column);
        return column_notation + row_notation;
    }
    function toColumnA1Notation(column) {
        const A = "A".charCodeAt(0);
        let output = "";
        while (column > 0) {
            const remainder = column % 26;
            output = String.fromCharCode(A + remainder - 1) + output;
            column = Math.floor(column / 26);
        }
        return output;
    }
})(A1 || (A1 = {}));
/** END src/a1notation.ts */
/** BEGIN src/utils.ts */
let mergedRangeRegistry = {};
let _createMergedRangeRegistry = (mergedRanges) => {
    // accessible by (row, column) in O(1)
    const mergedRangeRegistry = {};
    for (let i = 0; i < mergedRanges.length; i++) {
        const mergedRange = mergedRanges[i];
        const notation = mergedRange.getA1Notation();
        const range = A1.offsetFromA1Notation(notation);
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
};
_createMergedRangeRegistry = Profiler.wrap(_createMergedRangeRegistry);
let _fastCheckMergedRows = (column, row) => {
    // using registry to check if cell is merged
    if (mergedRangeRegistry[column] && mergedRangeRegistry[column][row]) {
        return mergedRangeRegistry[column][row];
    }
    return null;
};
_fastCheckMergedRows = Profiler.wrap(_fastCheckMergedRows);
/**
 * @param {string} s1 Source string
 * @param {string} s2 Target string
 * @param {object} [costs] Costs for operations { [replace], [replaceCase], [insert], [remove] }
 * @return {number} Levenshtein distance
 */
function levenshtein(s1, s2, costs = {
    replace: 1,
    replaceCase: 1,
    insert: 1,
    remove: 1,
}) {
    let i, j, l1, l2, flip, ch, chl, ii, ii2, cost, cutHalf;
    l1 = s1.length;
    l2 = s2.length;
    const cr = costs.replace;
    const cri = costs.replaceCase;
    const ci = costs.insert;
    const cd = costs.remove;
    cutHalf = flip = Math.max(l1, l2);
    const minCost = Math.min(cd, ci, cr);
    const minD = Math.max(minCost, (l1 - l2) * cd);
    const minI = Math.max(minCost, (l2 - l1) * ci);
    const buf = new Array(cutHalf * 2 - 1);
    for (i = 0; i <= l2; ++i) {
        buf[i] = i * minD;
    }
    for (i = 0; i < l1; ++i, flip = cutHalf - flip) {
        ch = s1[i];
        chl = ch.toLowerCase();
        buf[flip] = (i + 1) * minI;
        ii = flip;
        ii2 = cutHalf - flip;
        for (j = 0; j < l2; ++j, ++ii, ++ii2) {
            cost = ch === s2[j] ? 0 : chl === s2[j].toLowerCase() ? cri : cr;
            buf[ii + 1] = Math.min(buf[ii2 + 1] + cd, buf[ii] + ci, buf[ii2] + cost);
        }
    }
    return buf[l2 + cutHalf - flip];
}
/** END src/utils.ts */
/** BEGIN src/index.ts */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("InNoHassle").addItem("Open linter", "openLinter").addToUi();
}
function openLinter() {
    var _a, _b;
    const ui = SpreadsheetApp.getUi();
    const template = HtmlService.createTemplateFromFile("src/dialog");
    template.templateData = {
        spreadsheetId: Current.getSpreadsheet().getId(),
        settingsGid: (_a = Settings.getSettingsRange()) === null || _a === void 0 ? void 0 : _a.getSheet().getSheetId(),
        settingsRange: (_b = Settings.getSettingsRange()) === null || _b === void 0 ? void 0 : _b.getA1Notation(),
    };
    ui.showSidebar(template.evaluate().setTitle("InNoHassle").setWidth(500));
}
function lintHeader() {
    return ScheduleLinter.lintHeader();
}
function lintSchedule() {
    return ScheduleLinter.lintSchedule();
}
function lintCommon() {
    return CommonLinter.lintCommon();
}
function fixSpaces() {
    return CommonLinter.fixSpaces();
}
function addUnknownSubjectsToSettings() {
    return ScheduleLinter.addUnknownSubjectsToSettings();
}
function addUnknownLocationsToSettings() {
    return ScheduleLinter.addUnknownLocationsToSettings();
}
function selectScheduleGrids() {
    return ScheduleLinter.selectScheduleGrids();
}
function focusOnRange(range) {
    Logger.log(`Focusing on range ${range}`);
    const rangeObj = Current.getTargetSheet().getRange(range);
    rangeObj.activate();
}
function goToSettings() {
    return Settings.goToSettings();
}
function createSettings() {
    if (Settings.getSettingsRange()) {
        return "Settings already exist";
    }
    return Settings.createSettingsRange();
}
/** END src/index.ts */
/** BEGIN index.ts */
/**
 * Here the triple slash directives allow to specify order
 * in which files get added to the output
 */
/// <reference path="src/interfaces.ts" />
/// <reference path="src/settings.ts" />
/// <reference path="src/profiling.ts" />
/// <reference path="src/a1notation.ts" />
/// <reference path="src/utils.ts" />
/// <reference path="src/index.ts" />
// other files in tsconfig scope (`files` and `include`) will be added past this point
/** END index.ts */
/** BEGIN src/commonLinter.ts */
var CommonLinter;
(function (CommonLinter) {
    const multipleSpacePattern = /\s{2,}/;
    const trailingSpacePattern = /\s$/;
    const leadingSpacePattern = /^\s/;
    const noSpaceBeforeBracketPattern = /(\S)([({<\[])/;
    const spaceAfterBracketPattern = /([({<\[])\s/;
    const cyrillicPattern = /[а-яА-Я]/;
    const brackets = {
        "(": ")",
        "[": "]",
        "{": "}",
        "<": ">",
    };
    CommonLinter.lintCommon = () => {
        const range = Current.getTargetRange();
        const values = Current.getTargetValues();
        Logger.log(`Linting common ${range.getA1Notation()}`);
        const warnings = [];
        for (let row = 0; row < values.length; row++) {
            for (let column = 0; column < values[row].length; column++) {
                const value = values[row][column];
                if (!value || typeof value !== "string") {
                    continue;
                }
                const cell = range.getCell(row + 1, column + 1);
                _lintCyrillicSymbols(value, cell, warnings);
                _lintTrailingSpaces(value, cell, warnings);
                _lintMultipleSpaces(value, cell, warnings);
                _lintUnclosedBrackets(value, cell, warnings);
                _lintSpacesNearBrackets(value, cell, warnings);
            }
        }
        for (const warning of warnings) {
            warning.gid = Current.getTargetSheetId();
        }
        return warnings;
    };
    CommonLinter.lintCommon = Profiler.wrap(CommonLinter.lintCommon);
    CommonLinter.fixSpaces = () => {
        return _fixSpaces(Current.getTargetRange(), Current.getTargetValues());
    };
    CommonLinter.fixSpaces = Profiler.wrap(CommonLinter.fixSpaces);
    let _lintCyrillicSymbols = (value, cell, warnings) => {
        if (value.match(cyrillicPattern)) {
            warnings.push({
                content: "Cyrillic symbols found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
    };
    let _lintMultipleSpaces = (value, cell, warnings) => {
        if (value.match(multipleSpacePattern)) {
            warnings.push({
                content: "Multiple spaces found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
    };
    let _lintTrailingSpaces = (value, cell, warnings) => {
        if (value.match(trailingSpacePattern)) {
            warnings.push({
                content: "Trailing space found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
        else if (value.match(leadingSpacePattern)) {
            warnings.push({
                content: "Leading space found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
    };
    let _lintUnclosedBrackets = (chars, cell, warnings) => {
        //   using stack to check if brackets are closed
        const stack = [];
        const bracketsStack = [];
        for (let i = 0; i < chars.length; i++) {
            const char = chars[i];
            if (brackets[char]) {
                stack.push(char);
                bracketsStack.push({ char, index: i });
            }
            else if (Object.values(brackets).includes(char)) {
                const lastBracket = stack.pop();
                if (brackets[lastBracket] !== char) {
                    warnings.push({
                        content: "Unclosed bracket in cell " + cell.getA1Notation(),
                        range: cell.getA1Notation(),
                    });
                    return;
                }
            }
        }
        if (stack.length > 0) {
            warnings.push({
                content: "Unclosed bracket in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
    };
    let _lintSpacesNearBrackets = (value, cell, warnings) => {
        if (value.match(noSpaceBeforeBracketPattern)) {
            warnings.push({
                content: "Space before bracket not found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
        if (value.match(spaceAfterBracketPattern)) {
            warnings.push({
                content: "Space after bracket found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation(),
            });
        }
    };
    let _fixSpaces = (range, values) => {
        for (let row = 0; row < values.length; row++) {
            for (let column = 0; column < values[row].length; column++) {
                const value = values[row][column];
                if (!value || typeof value !== "string") {
                    continue;
                }
                const cell = range.getCell(row + 1, column + 1);
                if (value.match(leadingSpacePattern) ||
                    value.match(trailingSpacePattern)) {
                    cell.setValue(value.trim());
                }
                if (value.match(multipleSpacePattern)) {
                    cell.setValue(value.replace(multipleSpacePattern, " "));
                }
                if (value.match(noSpaceBeforeBracketPattern)) {
                    cell.setValue(value.replace(noSpaceBeforeBracketPattern, "$1 $2"));
                }
                if (value.match(spaceAfterBracketPattern)) {
                    cell.setValue(value.replace(spaceAfterBracketPattern, "$1"));
                }
            }
        }
    };
})(CommonLinter || (CommonLinter = {}));
/** END src/commonLinter.ts */
var Current;
(function (Current) {
    let spreadsheet;
    let targetSheet;
    let targetSheetId;
    let targetRange;
    let targetValues;
    function getSpreadsheet() {
        if (spreadsheet) {
            return spreadsheet;
        }
        spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        return spreadsheet;
    }
    Current.getSpreadsheet = getSpreadsheet;
    function getTargetSheet() {
        if (targetSheet) {
            return targetSheet;
        }
        targetSheet = getSpreadsheet().getActiveSheet();
        return targetSheet;
    }
    Current.getTargetSheet = getTargetSheet;
    function getTargetSheetId() {
        if (targetSheetId) {
            return targetSheetId;
        }
        targetSheetId = getTargetSheet().getSheetId();
        return targetSheetId;
    }
    Current.getTargetSheetId = getTargetSheetId;
    function getTargetRange() {
        if (targetRange) {
            return targetRange;
        }
        targetRange = getTargetSheet().getDataRange();
        return targetRange;
    }
    Current.getTargetRange = getTargetRange;
    function getTargetValues() {
        if (targetValues) {
            return targetValues;
        }
        targetValues = getTargetRange().getValues();
        return targetValues;
    }
    Current.getTargetValues = getTargetValues;
    function setTargetRange(range) {
        targetRange = range;
        targetValues = range.getValues();
    }
    Current.setTargetRange = setTargetRange;
})(Current || (Current = {}));
/** BEGIN src/innohassleLogo.ts */
const innohassleLogoEncoded = "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";
/** END src/innohassleLogo.ts */
/** BEGIN src/scheduleLinter.ts */
var ScheduleLinter;
(function (ScheduleLinter) {
    const unknownSubjects = new Set();
    const unknownLocations = new Set();
    const subjectTypePattern = /^([a-zA-Zа-яА-Я 0-9-:,.&?]+\S)\s*(?:\((.+)\))?$/;
    const allowedSubjectTypes = new Set(["lec", "tut", "lab"]);
    const groupPattern = /^([a-zA-Zа-яА-Я0-9\-]+)\s*(\(\d+\))?$/;
    const timeslotPattern = /^[0-9]{1,2}:[0-9]{2}-[0-9]{1,2}:[0-9]{2}$/;
    // ONLY ON as event modifier
    // [ONLY ON **/**, **/**, ... ] - event will be held only on specified dates
    const ONLY_ON_PATTERN = /^ONLY ON (?:[0-9]{2}\/[0-9]{2},? ?)+$/;
    // ON as additional modifier
    // [ <event_modifier> : ON **/**, **/**, ... ] - event modifiers will be applied only on specified dates
    const ON_PATTERN = /^ON ([0-9]{2}\/[0-9]{2},? ?)+$/;
    // FROM as event modifier
    // [ FROM **/** ] - event will be held only starting from the specified date
    // FROM as additional modifier
    // [ <event modifier> : FROM **/** ] - event modifiers will be applied only starting from the specified date
    const FROM_PATTERN = /^FROM ([0-9]{2}\/[0-9]{2})$/;
    // UNTIL as event modifier
    // [ UNTIL **/** ] - event will be held only until specified date
    // UNTIL as additional modifier
    // [ <event modifier> : UNTIL **/**] - event modifiers will be applied only until the specified date
    const UNTIL_PATTERN = /^UNTIL ([0-9]{2}\/[0-9]{2})$/;
    // STARTS AT as event modifier
    // [ STARTS AT **:** ] - starting time will be changed to specified time
    const STARTS_AT_PATTERN = /^STARTS AT ([0-9]{1,2}:[0-9]{2})$/;
    // ENDS AT as event modifier
    // [ ENDS AT **:** ] - ending time will be changed to specified time
    const ENDS_AT_PATTERN = /^ENDS AT ([0-9]{1,2}:[0-9]{2})$/;
    // LOCATION event modifier
    // [ <location> ] - event location will be changed to specified location (Including "ONLINE")
    const LOCATION_PATTERN = /^[a-zA-Zа-яА-Я0-9\-]+$/;
    // ON TBA as event modifier (To be announced)
    const ON_TBA_PATTERN = /^ON TBA$/;
    let ModifierType;
    (function (ModifierType) {
        ModifierType[ModifierType["ON_TBA_EVENT"] = 0] = "ON_TBA_EVENT";
        ModifierType[ModifierType["ONLY_ON_EVENT"] = 1] = "ONLY_ON_EVENT";
        ModifierType[ModifierType["ON_ADDITIONAL"] = 2] = "ON_ADDITIONAL";
        ModifierType[ModifierType["FROM_EVENT"] = 3] = "FROM_EVENT";
        ModifierType[ModifierType["FROM_ADDITIONAL"] = 4] = "FROM_ADDITIONAL";
        ModifierType[ModifierType["UNTIL_EVENT"] = 5] = "UNTIL_EVENT";
        ModifierType[ModifierType["UNTIL_ADDITIONAL"] = 6] = "UNTIL_ADDITIONAL";
        ModifierType[ModifierType["STARTS_AT_EVENT"] = 7] = "STARTS_AT_EVENT";
        ModifierType[ModifierType["ENDS_AT_EVENT"] = 8] = "ENDS_AT_EVENT";
        ModifierType[ModifierType["LOCATION_EVENT"] = 9] = "LOCATION_EVENT";
    })(ModifierType || (ModifierType = {}));
    const days = [
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "SATURDAY",
    ];
    ScheduleLinter.lintHeader = () => {
        // first row is courses
        // second row is groups
        const startDate = new Date();
        const warnings = [];
        const range = Current.getTargetRange();
        const values = Current.getTargetValues();
        // every cell in the first row should be a course
        for (let column = 0; column < values[0].length; column++) {
            const value = values[0][column];
            if (!value) {
                continue;
            }
            if (!Settings.getSettings().courses.has(value) && value != "-") {
                warnings.push({
                    content: "Unknown course " + value,
                    range: range.getCell(1, column + 1).getA1Notation(),
                });
            }
        }
        // every cell in the second row should be a group, may be with a "(<count of students>)" suffix
        for (let column = 0; column < values[1].length; column++) {
            const value = values[1][column];
            if (!value || typeof value !== "string") {
                continue;
            }
            // count of students is optional
            const match = value.match(groupPattern);
            if (!match) {
                warnings.push({
                    content: "Group with wrong format: `" +
                        value +
                        "`, should be `group (count of students)`",
                    range: range.getCell(2, column + 1).getA1Notation(),
                });
            }
            else {
                const group = match[1];
                // const count = match[2] ? parseInt(match[2].slice(1, -1)) : null;
                if (!Settings.getSettings().groups.has(group)) {
                    warnings.push({
                        content: "Unknown group " + group,
                        range: range.getCell(2, column + 1).getA1Notation(),
                    });
                }
            }
        }
        Logger.log("Header Linting took " +
            (new Date().getTime() - startDate.getTime()) +
            " ms");
        for (const warning of warnings) {
            warning.gid = Current.getTargetSheetId();
        }
        return warnings;
    };
    ScheduleLinter.lintHeader = Profiler.wrap(ScheduleLinter.lintHeader);
    ScheduleLinter.lintSchedule = () => {
        const startDate = new Date();
        unknownSubjects.clear();
        const warnings = [];
        const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
        const endOfScheduleRow = Current.getTargetSheet().getMaxRows();
        Current.setTargetRange(Current.getTargetSheet().getRange(1, 1, endOfScheduleRow, endOfScheduleColumn));
        const grids = _getScheduleGrids(warnings);
        const mergedRanges = Current.getTargetRange().getMergedRanges();
        mergedRangeRegistry = _createMergedRangeRegistry(mergedRanges);
        if (grids) {
            for (let i = 0; i < grids.length; i++) {
                const grid = grids[i];
                warnings.push(..._lintScheduleGrid(grid));
            }
        }
        if (unknownSubjects.size) {
            Logger.log("Unknown subjects to cache: " +
                Array.from(unknownSubjects).join("\r\n"));
            CacheService.getDocumentCache().put("unknownSubjects", Array.from(unknownSubjects).join("\r\n"), 100);
        }
        else {
            Logger.log("No unknown subjects to cache");
        }
        if (unknownLocations.size) {
            Logger.log("Unknown locations to cache: " +
                Array.from(unknownLocations).join("\r\n"));
            CacheService.getDocumentCache().put("unknownLocations", Array.from(unknownLocations).join("\r\n"), 100);
        }
        else {
            Logger.log("No unknown locations to cache");
        }
        Logger.log("Whole schedule Linting took " +
            (new Date().getTime() - startDate.getTime()) +
            " ms");
        for (const warning of warnings) {
            warning.gid = Current.getTargetSheetId();
        }
        Logger.log("Profiler: \n" + Profiler.format());
        return warnings;
    };
    ScheduleLinter.lintSchedule = Profiler.wrap(ScheduleLinter.lintSchedule);
    ScheduleLinter.selectScheduleGrids = () => {
        const warnings = [];
        const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
        const endOfScheduleRow = Current.getTargetSheet().getMaxRows();
        Current.setTargetRange(Current.getTargetSheet().getRange(1, 1, endOfScheduleRow, endOfScheduleColumn));
        const grids = _getScheduleGrids(warnings);
        if (grids) {
            const gridA1Notations = grids.map((grid) => grid.getA1Notation());
            const gridRanges = Current.getSpreadsheet().getRangeList(gridA1Notations);
            Current.getSpreadsheet().setActiveRangeList(gridRanges);
        }
        else {
            warnings.push({
                content: "No schedule grids found",
            });
        }
        for (const warning of warnings) {
            warning.gid = Current.getTargetSheetId();
        }
        return warnings;
    };
    ScheduleLinter.selectScheduleGrids = Profiler.wrap(ScheduleLinter.selectScheduleGrids);
    ScheduleLinter.addUnknownSubjectsToSettings = () => {
        const unknownSubjects = CacheService.getDocumentCache()
            .get("unknownSubjects")
            .split("\r\n");
        Logger.log("Unknown subjects to add: " + unknownSubjects);
        const settingsRange = Current.getSpreadsheet().getRangeByName("Settings");
        const subjectsColumn = settingsRange.getValues()[0].indexOf("subjects");
        // add to the end of the column
        let rowOffset = 0;
        // get last element index
        const values = settingsRange.getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][subjectsColumn]) {
                rowOffset = i + 1;
                break;
            }
        }
        const toFillRange = settingsRange.offset(rowOffset, subjectsColumn, unknownSubjects.length, 1);
        toFillRange.setValues(unknownSubjects.map((subject) => [subject]));
        toFillRange.activate();
    };
    ScheduleLinter.addUnknownLocationsToSettings = () => {
        const unknownLocations = CacheService.getDocumentCache()
            .get("unknownLocations")
            .split("\r\n");
        Logger.log("Unknown locations to add: " + unknownLocations);
        const settingsRange = Current.getSpreadsheet().getRangeByName("Settings");
        const locationsColumn = settingsRange.getValues()[0].indexOf("locations");
        // add to the end of the column
        let rowOffset = 0;
        // get last element index
        const values = settingsRange.getValues();
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i][locationsColumn]) {
                rowOffset = i + 1;
                break;
            }
        }
        const toFillRange = settingsRange.offset(rowOffset, locationsColumn, unknownLocations.length, 1);
        toFillRange.setValues(unknownLocations.map((location) => [location]));
        toFillRange.activate();
    };
    let _getEndOfScheduleColumn = (warnings) => {
        // values in header row equal to "-"
        const headerRow = Current.getSpreadsheet()
            .getActiveSheet()
            .getRange(1, 1, 1, Current.getSpreadsheet().getActiveSheet().getLastColumn());
        const values = headerRow.getValues()[0];
        for (let i = 0; i < values.length; i++) {
            if (values[i] === "-") {
                return i;
            }
        }
        warnings.push({
            content: "No end of schedule column found",
            range: headerRow.getA1Notation(),
        });
    };
    _getEndOfScheduleColumn = Profiler.wrap(_getEndOfScheduleColumn);
    let _splitIntoScheduleEntries = (grid) => {
        // schedule entries are cells grouped by three rows (and merged columns)
        const values = grid.getValues();
        const gridHeight = grid.getHeight();
        const gridWidth = grid.getWidth();
        const gridStartRow = grid.getRow();
        const gridStartColumn = grid.getColumn();
        const scheduleEntries = [];
        for (let startRow = 0; startRow < gridHeight; startRow += 3) {
            const columnsToSkip = new Set();
            for (let targetColumn = 0; targetColumn < gridWidth; targetColumn++) {
                if (columnsToSkip.has(targetColumn)) {
                    continue;
                }
                // check if it has data
                const scheduleEntryValues = [
                    values[startRow][targetColumn],
                    values[startRow + 1][targetColumn],
                    values[startRow + 2][targetColumn],
                ];
                if (scheduleEntryValues.every((value) => !value)) {
                    continue;
                }
                let scheduleEntryOffset = {
                    row: gridStartRow + startRow,
                    column: gridStartColumn + targetColumn,
                    numRows: 3,
                    numColumns: 1,
                };
                // check if it is merged
                const mergedRange = _fastCheckMergedRows(scheduleEntryOffset.row, scheduleEntryOffset.column);
                if (mergedRange) {
                    const numColumns = mergedRange.getNumColumns();
                    scheduleEntryOffset = {
                        row: mergedRange.getRow(),
                        column: mergedRange.getColumn(),
                        numRows: 3,
                        numColumns: numColumns,
                    };
                    Logger.log("Merged range: " + mergedRange.getA1Notation());
                    for (let i = 1; i < numColumns; i++) {
                        columnsToSkip.add(targetColumn + i);
                    }
                }
                scheduleEntries.push({
                    offset: scheduleEntryOffset,
                    values: scheduleEntryValues,
                });
            }
        }
        return scheduleEntries;
    };
    _splitIntoScheduleEntries = Profiler.wrap(_splitIntoScheduleEntries);
    let _lintScheduleGrid = (grid) => {
        const warnings = [];
        const scheduleEntries = _splitIntoScheduleEntries(grid);
        const _lintScheduleEntriesStart = new Date();
        for (let i = 0; i < scheduleEntries.length; i++) {
            const scheduleEntry = scheduleEntries[i];
            Logger.log("Linting schedule entry " + A1.offsetToA1Notation(scheduleEntry.offset));
            if (scheduleEntry)
                warnings.push(..._lintScheduleEntry(scheduleEntry.offset, scheduleEntry.values));
        }
        Logger.log("Linting schedule entries took " +
            (new Date().getTime() - _lintScheduleEntriesStart.getTime()) +
            " ms");
        return warnings;
    };
    _lintScheduleGrid = Profiler.wrap(_lintScheduleGrid);
    let _lintScheduleEntry = (scheduleEntryOffset, rowsValues) => {
        const warnings = [];
        // schedule entry with three rows:
        // subject row
        // teacher row
        // location row
        const subjectValue = rowsValues[0];
        const teacherValue = rowsValues[1];
        const locationValue = rowsValues[2];
        if (subjectValue)
            _lintSubject(subjectValue, scheduleEntryOffset, warnings);
        // if (teacherValue)
        //     lintTeacher(teacherValue, scheduleEntryRange, warnings);
        if (locationValue)
            _lintLocation(locationValue.toString(), scheduleEntryOffset, warnings);
        return warnings;
    };
    _lintScheduleEntry = Profiler.wrap(_lintScheduleEntry);
    let _getScheduleGrids = (warnings) => {
        const gridColumns = _getTimeColumns(warnings);
        const range = Current.getTargetRange();
        gridColumns.push(range.getWidth());
        const gridRows = _getDayRows(warnings);
        gridRows.push(range.getHeight());
        const grids = [];
        // grid is not selected: pairwise iterate over columns and rows
        for (let i = 0; i < gridColumns.length - 1; i++) {
            const startColumn = gridColumns[i] + 1;
            const endColumn = gridColumns[i + 1];
            for (let j = 0; j < gridRows.length - 1; j++) {
                const startRow = gridRows[j] + 1;
                const endRow = gridRows[j + 1];
                // only between start and end columns and rows
                grids.push(range.offset(startRow, startColumn, endRow - startRow, endColumn - startColumn));
            }
        }
        const gridA1Notations = grids.map((grid) => grid.getA1Notation());
        Logger.log("Grids: " + gridA1Notations);
        return grids;
    };
    _getScheduleGrids = Profiler.wrap(_getScheduleGrids);
    let _lintSubject = (subjectValue, scheduleEntryOffset, warnings) => {
        // regex for subject value:
        // <subject>
        // <subject> (<type>)
        const match = subjectValue.match(subjectTypePattern);
        if (!match) {
            warnings.push({
                content: "Subject with wrong format: `" +
                    subjectValue +
                    "`, should be `subject (type)`",
                range: A1.offsetToA1Notation(scheduleEntryOffset),
            });
        }
        else {
            const subject = match[1];
            // maybe without type
            const type = match[2] ? match[2] : null;
            if (!Settings.getSettings().subjects.has(subject)) {
                let nearestSubject = "";
                let distance = 100;
                for (const knownSubject of Settings.getSettings().subjects) {
                    const currentDistance = levenshtein(subject, knownSubject);
                    if (currentDistance < distance) {
                        distance = currentDistance;
                        nearestSubject = knownSubject;
                    }
                }
                warnings.push({
                    content: `Unknown subject '${subject}'` +
                        (nearestSubject ? ` (did you mean '${nearestSubject}'?)` : ""),
                    range: A1.offsetToA1Notation(scheduleEntryOffset),
                });
                unknownSubjects.add(subject);
            }
            if (type && !allowedSubjectTypes.has(type)) {
                warnings.push({
                    content: "Unknown subject type `" +
                        type +
                        "`. Allowed types: `" +
                        Array.from(allowedSubjectTypes).join("`, `") +
                        "`",
                    range: A1.offsetToA1Notation(scheduleEntryOffset),
                });
            }
        }
    };
    _lintSubject = Profiler.wrap(_lintSubject);
    let _lintLocation = (locationValue, scheduleEntryOffset, warnings) => {
        // location string examples:
        // 1. STARTS AT 9:20 - For instance, "106 [STARTS AT 9:20]"
        // 2. ONLY ON [specific dates] - E.g., "107 [ONLY ON 22/01,29/01,25/03,01/04]"
        // 3. FROM [a specific date] - Such as "320 [FROM 05.02]"
        // 4. STARTS AT 11:00 - As in "421 [STARTS AT 11:00]" and "303 [STARTS AT 11:00]"
        // 5. ONLY ON [a specific date] AT [a specific time] - For example, "104 [ONLY ON 29/01 AT 12:00]"
        // 6. Location options - "106/108/301/312/321/421"
        // 7. only on [multiple specific dates]; starts at [a specific time] - E.g., "105 [only on 23/01, 30/01, 26/03, 02/04; starts at 9:10]"
        // 8. ONLINE [STARTS FROM a specific date] - Such as "ONLINE [STARTS FROM 07/02]"
        // 9. [a number] ON [a specific date] - For instance, "308 [313 ON 24.04]"
        // 10. ONLINE [ON TBA] - E.g., "ONLINE [ON TBA]"
        // 11. ONLINE [FROM a specific date] - As in "106 [ONLY ON 25/01] / ONLINE [FROM 22/02]"
        // 12. STARTS AT [a specific time] - For example, "ONLINE [STARTS AT 11:00]" and "ONLINE [STARTS AT 13:00]"
        //
        // Suggestions
        // Event modifier -- applied to event itself; Each event modifier in separate [ block ]
        // Additional modifier -- applied to event modifier. If no additional modifier is specified then event modifier applied to all event instances
        // Syntax for combining: [ <event modifiers delimited by semicolon> : <additional modifier>]
        // If no additional modifier is specified then event modifier applied to all event instances
        // divide modifiers and location
        const modifierPattern = /\[([^\[\]()]+)\]/g;
        const matches = locationValue.match(modifierPattern);
        const location = locationValue.replace(modifierPattern, "").trim();
        if (!location) {
            warnings.push({
                content: "No location found",
                range: A1.offsetToA1Notation(scheduleEntryOffset),
            });
        }
        else {
            // check if location contains many options (separated by /)
            const locationOptions = location.split("/");
            for (const option of locationOptions) {
                if (!Settings.getSettings().locations.has(option)) {
                    warnings.push({
                        content: `Unknown location '${option}'`,
                        range: A1.offsetToA1Notation(scheduleEntryOffset),
                    });
                    unknownLocations.add(option);
                }
            }
        }
        if (matches) {
            // convert to array of strings
            const modifiersStrings = matches.map((match) => match.slice(1, -1));
            const modifiers = _lintModifiers(modifiersStrings, scheduleEntryOffset, warnings);
        }
    };
    let _lintModifiers = (modifiersStrings, scheduleEntryOffset, warnings) => {
        let modifiers = [];
        for (let i = 0; i < modifiersStrings.length; i++) {
            let eventModifierString = modifiersStrings[i];
            let additionalModifierMatch = eventModifierString.match(/_(.+)/);
            let additionalModifier = "";
            if (additionalModifierMatch) {
                additionalModifier = additionalModifierMatch[1];
                // check if modifier is STARTS AT or ENDS AT`
                eventModifierString = eventModifierString.replace(/_(.+)/, "").trim();
            }
            for (const eventModifier of eventModifierString
                .split("AND")
                .map((s) => s.trim())) {
                let type = _getModifierType(eventModifier, false);
                if (type !== null) {
                    // check if additional modifier is present
                    if (additionalModifier) {
                        let additionalType = _getModifierType(additionalModifier, true);
                        if (additionalType !== null) {
                            modifiers.push({
                                origin: modifiersStrings[i],
                                eventModifier: eventModifier,
                                eventModifierType: type,
                                additionalModifier: additionalModifier,
                                additionalModifierType: additionalType,
                            });
                        }
                        else {
                            warnings.push({
                                content: `Unknown additional modifier '${additionalModifier}'`,
                                range: A1.offsetToA1Notation(scheduleEntryOffset),
                            });
                        }
                    }
                    else {
                        modifiers.push({
                            origin: modifiersStrings[i],
                            eventModifier: eventModifier,
                            eventModifierType: type,
                        });
                    }
                }
                else {
                    warnings.push({
                        content: `Unknown event modifier '${eventModifier}'`,
                        range: A1.offsetToA1Notation(scheduleEntryOffset),
                    });
                }
            }
        }
        return modifiers;
    };
    let _getModifierType = (modifier, is_additional = false) => {
        if (is_additional) {
            if (modifier.match(ON_PATTERN)) {
                return ModifierType.ON_ADDITIONAL;
            }
            if (modifier.match(FROM_PATTERN)) {
                return ModifierType.FROM_ADDITIONAL;
            }
            if (modifier.match(UNTIL_PATTERN)) {
                return ModifierType.UNTIL_ADDITIONAL;
            }
        }
        else {
            if (modifier.match(ON_TBA_PATTERN)) {
                return ModifierType.ON_TBA_EVENT;
            }
            if (modifier.match(ONLY_ON_PATTERN)) {
                return ModifierType.ONLY_ON_EVENT;
            }
            if (modifier.match(FROM_PATTERN)) {
                return ModifierType.FROM_EVENT;
            }
            if (modifier.match(UNTIL_PATTERN)) {
                return ModifierType.UNTIL_EVENT;
            }
            if (modifier.match(STARTS_AT_PATTERN)) {
                return ModifierType.STARTS_AT_EVENT;
            }
            if (modifier.match(ENDS_AT_PATTERN)) {
                return ModifierType.ENDS_AT_EVENT;
            }
            if (modifier.match(LOCATION_PATTERN)) {
                return ModifierType.LOCATION_EVENT;
            }
        }
        return null;
    };
    let _getTimeColumns = (warnings) => {
        const values = Current.getTargetValues();
        const timeColumns = [];
        for (let column = 0; column < values[0].length; column++) {
            // check if column contains time slots
            for (let row = 0; row < values.length; row++) {
                const value = values[row][column];
                if (!value || typeof value !== "string") {
                    continue;
                }
                if (value.match(timeslotPattern)) {
                    timeColumns.push(column);
                    break;
                }
            }
        }
        if (timeColumns.length === 0) {
            warnings.push({
                content: "No time columns found",
            });
        }
        Logger.log("Time columns: " + timeColumns);
        return timeColumns;
    };
    _getTimeColumns = Profiler.wrap(_getTimeColumns);
    let _getDayRows = (warnings) => {
        const values = Current.getTargetValues();
        const dayRows = [];
        for (let row = 0; row < values.length; row++) {
            const value = values[row][0];
            if (days.includes(value)) {
                dayRows.push(row);
            }
        }
        if (dayRows.length === 0) {
            warnings.push({
                content: "No day rows found",
            });
        }
        Logger.log("Day rows: " + dayRows);
        return dayRows;
    };
    _getDayRows = Profiler.wrap(_getDayRows);
})(ScheduleLinter || (ScheduleLinter = {}));
/** END src/scheduleLinter.ts */
