# InNoHassle-SpreadsheetsHelper

Plugin for Google Spreadsheets to lint and format Innopolis University's schedule

## Installation (just using)

1. Open your Google Sheets document
2. Go to `Extensions` -> `Apps Script`
3. Paste the code from [main.js](build/main.js) into the editor and save
4. Also copy [dialog.html](src/dialog.html) to `src` folder and [appsscript.json](appsscript.json) to the root
5. Reload page with the document, and you should see a new menu item `InNoHassle` with a few options

## Installation (for developers)

1. Install [node](https://nodejs.org/en), [clasp](https://github.com/google/clasp), [prettier](https://prettier.io)
2. Clone the repository
3. Run `clasp login` and `clasp clone <scriptId>` where `<scriptId>` is the id of the script you want to work with (you
   can find it in the url of the script's page in the Apps Script editor or in Apps Script settings)
4. Run `npm install` to install dependencies
5. Run `npm run <command>` to lint, format, or build the project (see [package.json](package.json) for available
   commands)
6. Done!