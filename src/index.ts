/** BEGIN src/index.ts */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("InNoHassle").addItem("Open linter", "openLinter").addToUi();
}

function openLinter() {
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile("src/dialog");
  template.templateData = {
    spreadsheetId: Current.getSpreadsheet().getId(),
    settingsGid: Settings.getSettingsRange()?.getSheet().getSheetId(),
    settingsRange: Settings.getSettingsRange()?.getA1Notation(),
  };
  ui.showSidebar(
    template.evaluate().setTitle("InNoHassle").setWidth(500),
    // "InNoHassle",
  );
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

function focusOnRange(range: string) {
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
