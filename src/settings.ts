/** BEGIN src/settings.ts */

namespace Settings {

    let _settings: {
        subjects: Set<string>,
        groups: Set<string>,
        courses: Set<string>,
        locations: Set<string>,
        teachers: Set<string>
    } | null = null;

    export function getSettings() {
        if (_settings) {
            return _settings;
        }

        // get named range "Settings" from the spreadsheet
        const settingsRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("Settings");
        if (!settingsRange) {
            return null;
        }
        const settings = {
            "subjects": new Set<string>(),
            "groups": new Set<string>(),
            "courses": new Set<string>(),
            "locations": new Set<string>(),
            "teachers": new Set<string>()
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
                    settings[columnName].add(value);
                }
            }
        }
        _settings = settings;
        return settings;
    }

    export let goToSettings = () => {
        const settingsRange = spreadsheet.getRangeByName("Settings");
        if (settingsRange) {
            settingsRange.activate();
        }
    }
}
/** END src/settings.ts */