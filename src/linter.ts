// noinspection JSUnusedGlobalSymbols

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const fileId = spreadsheet.getId();
let ui: GoogleAppsScript.Base.Ui | null = null;
try {
    ui = SpreadsheetApp.getUi();
} catch (e) {
    Logger.log("No UI" + e.toString());
}

function onOpen() {
    ui.createMenu('InNoHassle').addItem("Open linter", "openLinter").addToUi();
    openLinter();
}

function openLinter() {
    ui.showModelessDialog(
        HtmlService.createHtmlOutputFromFile("src/index").setTitle('InNoHassle').setWidth(500),
        'InNoHassle'
    )
}

function lint() {
    Logger.log("Linting...");
    return lintAllCells();
}


function lintAllCells() {
    const range = selectionOrActiveSheet();
    const values = range.getValues();
    const warnings = [];
    let startDate = new Date();
    for (let row = 0; row < values.length; row++) {
        for (let column = 0; column < values[row].length; column++) {
            const value = values[row][column];

            if (!value || typeof value !== "string") {
                continue;
            }
            const cell = range.getCell(row + 1, column + 1);
            lintCyrillicSymbols(value, cell, warnings);
            lintTrailingSpaces(value, cell, warnings);
            lintMultipleSpaces(value, cell, warnings);
            lintUnclosedBrackets(value, cell, warnings);
        }
    }
    Logger.log("Cyrillic Linting took " + (new Date().getTime() - startDate.getTime()) + " ms");
    return warnings;
}

function lintCyrillicSymbols(value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) {
    if (value.match(/[а-яА-Я]/)) {
        warnings.push({
            content: "Cyrillic symbols found in cell " + cell.getA1Notation(),
            range: cell.getA1Notation()
        });
    }
}

function lintMultipleSpaces(value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) {
    if (value.match(/\s{2,}/)) {
        warnings.push({
            content: "Multiple spaces found in cell " + cell.getA1Notation(),
            range: cell.getA1Notation()
        });
    }
}


function lintTrailingSpaces(value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) {
    if (value.match(/\s$/)) {
        warnings.push({
            content: "Trailing space found in cell " + cell.getA1Notation(),
            range: cell.getA1Notation()
        });
    } else if (value.match(/^\s/)) {
        warnings.push({
            content: "Leading space found in cell " + cell.getA1Notation(),
            range: cell.getA1Notation()
        });
    }
}

const brackets = {
    "(": ")",
    "[": "]",
    "{": "}",
    "<": ">"
}

function lintUnclosedBrackets(chars: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) {
    //   using stack to check if brackets are closed
    const stack = [];
    const bracketsStack = [];
    for (let i = 0; i < chars.length; i++) {
        const char = chars[i];
        if (brackets[char]) {
            stack.push(char);
            bracketsStack.push({char, index: i});
        } else if (Object.values(brackets).includes(char)) {
            const lastBracket = stack.pop();
            if (brackets[lastBracket] !== char) {
                warnings.push({
                    content: "Unclosed bracket in cell " + cell.getA1Notation(),
                    range: cell.getA1Notation()
                });
                return;
            }
        }
    }

    if (stack.length > 0) {
        warnings.push({
            content: "Unclosed bracket in cell " + cell.getA1Notation(),
            range: cell.getA1Notation()
        });
    }
}

function trimSpaces() {
    const startDate = new Date();
    const range = selectionOrActiveSheet();
    range.trimWhitespace();
    //
    // const values = range.getValues();
    // let startDate = new Date();
    // for (let row = 0; row < values.length; row++) {
    //     for (let column = 0; column < values[row].length; column++) {
    //         const value = values[row][column];
    //
    //         if (!value || typeof value !== "string") {
    //             continue;
    //         }
    //         const cell = range.getCell(row + 1, column + 1);
    //         if (value.match(/\s$/) || value.match(/^\s/)) {
    //             cell.setValue(value.trim());
    //         }
    //     }
    // }
    Logger.log("Trailing spaces fix took " + (new Date().getTime() - startDate.getTime()) + " ms");
}


function selectionOrActiveSheet() {
    const sheet = spreadsheet.getActiveSheet();
    const selection = spreadsheet.getActiveRangeList();
    if (selection && selection.getRanges().length > 0 && (
        selection.getRanges()[0].getWidth() > 1 ||
        selection.getRanges()[0].getHeight() > 1
    )) {
        return selection.getRanges()[0];
    }
    return sheet.getDataRange();
}

function focusOnRange(range: string) {
    const sheet = spreadsheet.getActiveSheet();
    const rangeObj = sheet.getRange(range);
    spreadsheet.setActiveRange(rangeObj);
}

function getSettings() {
    // get named range "Settings" from the spreadsheet
    const settingsRange = spreadsheet.getRangeByName("Settings");
    if (!settingsRange) {
        return null;
    }
    const settingsSchema = getSettingsSchema(settingsRange);
    
}

function getSettingsSchema(settingsRange: GoogleAppsScript.Spreadsheet.Range) {
    // subjects	  groups
    // name	    name	count
    // Data Structures and Algorithms	B23-ISE-01	32
    // Software Systems Analysis and Design	B23-ISE-02	32
    // Mathematical Analysis II	B23-ISE-03	32
    // Analytical Geometry and Linear Algebra II	B23-ISE-01	32

    // bold italic values are the names of the columns
    const settingsInheritance: string[][] = [];
    // iterate over columns
    for (let column = 0; column <= settingsRange.getNumColumns(); column++) {
        const columnRange = settingsRange.offset(0, column, settingsRange.getNumRows(), 1);
        const columnSettings = [];
        // iterate over rows
        for (let row = 0; row <= settingsRange.getNumRows(); row++) {
            const cell = columnRange.getCell(row, 1);
            const value = cell.getRichTextValue();
            const runs = value.getRuns();
            if (runs.length == 1 && runs[0].getTextStyle().isBold() && runs[0].getTextStyle().isItalic()) {
                columnSettings.push(value.getText());
            } else {
                break
            }
        }
        settingsInheritance.push(columnSettings);
    }

    // {"subjects": {"name": {}}, "groups": {"name": {}, "count": {}}}
    const settingsSchema = {}

    // setup schema
    for (let i = 0; i < settingsInheritance.length; i++) {
        // ["subjects", "name"] or ["groups", "name"] or ["groups", "count"]
        const columnSettings = settingsInheritance[i];
        if (columnSettings.length == 0) {
            continue;
        }
        let parent = settingsSchema;
        for (let j = 0; j < columnSettings.length; j++) {
            const setting = columnSettings[j];
            if (!parent[setting]) {
                parent[setting] = {};
            }
            parent = parent[setting];
        }
    }

    return settingsSchema;
}
