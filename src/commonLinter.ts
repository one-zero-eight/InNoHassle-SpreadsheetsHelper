/** BEGIN src/commonLinter.ts */
namespace CommonLinter {
    const multipleSpacePattern = /\s{2,}/;
    const trailingSpacePattern = /\s$/;
    const leadingSpacePattern = /^\s/;
    const noSpaceBeforeBracketPattern = /\S([({<\[])/;
    const spaceAfterBracketPattern = /([({<\[])\s/;
    const cyrillicPattern = /[а-яА-Я]/;
    const brackets = {
        "(": ")",
        "[": "]",
        "{": "}",
        "<": ">"
    }

    export let lintCommon = () => {
        const range = currentTargetRange;
        const values = currentTargetValues;
        const warnings: Warning[] = [];
        let startDate = new Date();
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
        Logger.log("All cells Linting took " + (new Date().getTime() - startDate.getTime()) + " ms");
        return warnings;
    }

    export let fixSpaces = () => {
        return _fixSpaces(currentTargetRange, currentTargetValues);
    }


    let _lintCyrillicSymbols = (value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) => {
        if (value.match(cyrillicPattern)) {
            warnings.push({
                content: "Cyrillic symbols found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        }
    }

    let _lintMultipleSpaces = (value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) => {
        if (value.match(multipleSpacePattern)) {
            warnings.push({
                content: "Multiple spaces found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        }
    }

    let _lintTrailingSpaces = (value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) => {
        if (value.match(trailingSpacePattern)) {
            warnings.push({
                content: "Trailing space found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        } else if (value.match(leadingSpacePattern)) {
            warnings.push({
                content: "Leading space found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        }
    }

    let _lintUnclosedBrackets = (chars: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) => {
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

    let _lintSpacesNearBrackets = (value: string, cell: GoogleAppsScript.Spreadsheet.Range, warnings: any[]) => {
        if (value.match(noSpaceBeforeBracketPattern)) {
            warnings.push({
                content: "Space before bracket not found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        }
        if (value.match(spaceAfterBracketPattern)) {
            warnings.push({
                content: "Space after bracket found in cell " + cell.getA1Notation(),
                range: cell.getA1Notation()
            });
        }
    }

    let _fixSpaces = (range: GoogleAppsScript.Spreadsheet.Range, values: any[][]) => {
        for (let row = 0; row < values.length; row++) {
            for (let column = 0; column < values[row].length; column++) {
                const value = values[row][column];

                if (!value || typeof value !== "string") {
                    continue;
                }
                const cell = range.getCell(row + 1, column + 1);
                if (value.match(leadingSpacePattern) || value.match(trailingSpacePattern)) {
                    cell.setValue(value.trim());
                }

                if (value.match(multipleSpacePattern)) {
                    cell.setValue(value.replace(multipleSpacePattern, " "));
                }

                if (value.match(noSpaceBeforeBracketPattern)) {
                    cell.setValue(value.replace(noSpaceBeforeBracketPattern, " $1"));
                }

                if (value.match(spaceAfterBracketPattern)) {
                    cell.setValue(value.replace(spaceAfterBracketPattern, "$1"));
                }
            }
        }
    }
}
/** END src/commonLinter.ts */
