# InNoHassle-SpreadsheetsHelper

Plugin for Google Spreadsheets to lint and format Innopolis University's schedule

## Installation

1. Install [node](https://nodejs.org/en), [clasp](https://github.com/google/clasp), [prettier](https://prettier.io)
2. Clone the repository
3. Get <scriptId> from the Apps Script editor:
    - Open your Google Sheets document
    - Go to `Extensions` -> `Apps Script` -> `Project Settings`
    - Find `Script ID` and copy it
4. Run `clasp login` and `clasp clone <scriptId>` where `<scriptId>` is the id of the script you want to work with (you
   can find it in the url of the script's page in the Apps Script editor or in Apps Script settings)
5. Run `npm install` to install dependencies
6. Run `npm run bpp` to build, pretty and push the project (see [package.json](package.json)
7. Done!

## Usage

1. Install the plugin as described in the [Installation](#installation) section
2. Open your Google Sheets document
3. Click `InNoHassle` in top menu and `Open linter`
4. Create named range for the setting by clicking `Create settings`
5. Use `Lint` for just checking rules; and `Format` or `add` for fixing