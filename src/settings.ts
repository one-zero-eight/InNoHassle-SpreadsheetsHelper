/** BEGIN src/settings.ts */

namespace Settings {
  let _settings: {
    subjects: Set<string>;
    groups: Set<string>;
    courses: Set<string>;
    locations: Set<string>;
    teachers: Set<string>;
  } | null = null;

  let _settingsRange: GoogleAppsScript.Spreadsheet.Range | undefined;

  export function createSettingsRange() {
    // new sheet with the name "InNoHassle"
    const sheet = Current.getSpreadsheet().insertSheet("InNoHassle");

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

    // Add innohassle logo
    const blob = Utilities.newBlob(
      Utilities.base64Decode(innohassleLogoEncoded),
      "image/png",
      "innohassle.png",
    );

    const logo = sheet.insertImage(blob, 10, 10);
    logo; // logo.setWidth(512);
    // logo.setHeight(512);
    // Add a link to the website
    // sheet.getRange("")

    range.activate();
  }

  export function getSettingsRange() {
    if (_settingsRange) {
      return _settingsRange;
    }
    _settingsRange = Current.getSpreadsheet().getRangeByName("Settings");
    return _settingsRange;
  }

  export function getSettings() {
    if (_settings) {
      return _settings;
    }

    // get named range "Settings" from the spreadsheet
    const settingsRange = getSettingsRange();
    if (!settingsRange) {
      return null;
    }
    const settings = {
      subjects: new Set<string>(),
      groups: new Set<string>(),
      courses: new Set<string>(),
      locations: new Set<string>(),
      teachers: new Set<string>(),
    };
    // iterate over columns
    for (let column = 0; column < settingsRange.getWidth(); column++) {
      const columnRange = settingsRange.offset(
        0,
        column,
        settingsRange.getHeight(),
        1,
      );
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

  export let goToSettings = () => {
    const settingsRange = getSettingsRange();
    if (settingsRange) {
      settingsRange.activate();
    }
  };
}
/** END src/settings.ts */
