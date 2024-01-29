/** BEGIN src/index.ts */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('InNoHassle').addItem("Open linter", "openLinter").addToUi();
}

function openLinter() {
    const ui = SpreadsheetApp.getUi();
    ui.showModelessDialog(
        HtmlService.createHtmlOutputFromFile("src/dialog").setTitle('InNoHassle').setWidth(500),
        'InNoHassle'
    )
}

function lintHeader() {
    return ScheduleLinter.lintHeader()
}

function lintSchedule() {
    return ScheduleLinter.lintSchedule()
}

function lintCommon() {
    return CommonLinter.lintCommon()
}

function fixSpaces() {
    return CommonLinter.fixSpaces()
}

function addUnknownSubjectsToSettings() {
    return ScheduleLinter.addUnknownSubjectsToSettings()
}

function selectScheduleGrids() {
    return ScheduleLinter.selectScheduleGrids()
}

function focusOnRange(range: string) {
    const rangeObj = spreadsheet.getRange(range);
    rangeObj.activate();
}

function goToSettings() {
    return Settings.goToSettings()
}
/** END src/index.ts */
