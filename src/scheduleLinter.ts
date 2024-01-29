/** BEGIN src/scheduleLinter.ts */
namespace ScheduleLinter {
    const unknownSubjects = new Set();
    const subjectTypePattern = /^([a-zA-Z 0-9-:,.&?]+\S)\s*(?:\((.+)\))?$/;
    const allowedSubjectTypes = new Set(["lec", "tut", "lab"])
    const groupPattern = /^([a-zA-Z0-9\-]+)\s*(\(\d+\))?$/;
    const timeslotPattern = /^[0-9]{1,2}:[0-9]{2}-[0-9]{1,2}:[0-9]{2}$/;
    const days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"];


    export let lintHeader = () => {
        // first row is courses
        // second row is groups

        const startDate = new Date();
        const warnings: Warning[] = [];

        const range = currentTargetRange;
        const values = currentTargetValues;

        // every cell in the first row should be a course
        for (let column = 0; column < values[0].length; column++) {
            const value = values[0][column];
            if (!value) {
                continue;
            }
            if (!Settings.getSettings().courses.has(value) && value != '-') {
                warnings.push({
                    content: "Unknown course " + value,
                    range: range.getCell(1, column + 1).getA1Notation()
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
                    content: "Group with wrong format: `" + value + "`, should be `group (count of students)`",
                    range: range.getCell(2, column + 1).getA1Notation()
                });
            } else {
                const group = match[1];
                // const count = match[2] ? parseInt(match[2].slice(1, -1)) : null;
                if (!Settings.getSettings().groups.has(group)) {
                    warnings.push({
                        content: "Unknown group " + group,
                        range: range.getCell(2, column + 1).getA1Notation()
                    });
                }
            }
        }

        Logger.log("Header Linting took " + (new Date().getTime() - startDate.getTime()) + " ms");
        return warnings;
    }

    export let lintSchedule = () => {
        const startDate = new Date();

        unknownSubjects.clear()
        const warnings: Warning[] = [];
        const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
        const endOfScheduleRow = spreadsheet.getActiveSheet().getMaxRows();
        currentTargetRange = currentTargetSheet.getRange(1, 1, endOfScheduleRow, endOfScheduleColumn);
        currentTargetValues = currentTargetRange.getValues();

        const grids = _getScheduleGrids(warnings);
        const mergedRanges = currentTargetRange.getMergedRanges();
        mergedRangeRegistry = _createMergedRangeRegistry(mergedRanges);
        if (grids) {
            for (let i = 0; i < grids.length; i++) {
                const grid = grids[i];
                warnings.push(..._lintScheduleGrid(grid));
            }
        }
        if (unknownSubjects.size) {
            Logger.log("Unknown subjects to cache: " + Array.from(unknownSubjects).join(","))
            CacheService.getDocumentCache().put("unknownSubjects", Array.from(unknownSubjects).join(","), 100)
        } else {
            Logger.log("No unknown subjects to cache")
        }
        Logger.log("Whole schedule Linting took " + (new Date().getTime() - startDate.getTime()) + " ms");

        Logger.log("Profiler: \n" + Profiler.format());
        return warnings;
    }


    export let selectScheduleGrids = () => {
        const warnings: Warning[] = [];
        const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
        const endOfScheduleRow = spreadsheet.getActiveSheet().getMaxRows();
        currentTargetRange = currentTargetSheet.getRange(1, 1, endOfScheduleRow, endOfScheduleColumn);
        currentTargetValues = currentTargetRange.getValues();
        const grids = _getScheduleGrids(warnings);

        if (grids) {
            const gridA1Notations = grids.map(grid => grid.getA1Notation());
            const gridRanges = spreadsheet.getRangeList(gridA1Notations);
            spreadsheet.setActiveRangeList(gridRanges);
        } else {
            warnings.push({
                content: "No schedule grids found",
            });
        }
        return warnings;
    }

    export let addUnknownSubjectsToSettings = () => {
        const unknownSubjects = CacheService.getDocumentCache().get("unknownSubjects").split(",");
        Logger.log("Unknown subjects to add: " + unknownSubjects)
        const settingsRange = spreadsheet.getRangeByName("Settings");
        const subjectsColumn = settingsRange.getValues()[0].indexOf("subjects");
        // add to the end of the column
        let rowOffset = 0;
        // last elem
        for (let i = settingsRange.getHeight() - 1; i >= 0; i--) {
            if (settingsRange.getCell(i + 1, subjectsColumn + 1).getValue()) {
                rowOffset = i + 1;
                break;
            }
        }
        const toFillRange = settingsRange.offset(rowOffset, subjectsColumn, unknownSubjects.length, 1);
        toFillRange.setValues(unknownSubjects.map(subject => [subject]));
        toFillRange.activate()
    }

    let _getEndOfScheduleColumn = (warnings: Warning[]) => {
        // values in header row equal to "-"
        const headerRow = spreadsheet.getActiveSheet().getRange(1, 1, 1, spreadsheet.getActiveSheet().getLastColumn());
        const values = headerRow.getValues()[0];
        for (let i = 0; i < values.length; i++) {
            if (values[i] === "-") {
                return i;
            }
        }
        warnings.push({
            content: "No end of schedule column found",
            range: headerRow.getA1Notation()
        });
    }

    let _splitIntoScheduleEntries = (grid: GoogleAppsScript.Spreadsheet.Range): {
        offset: Offset,
        values: string[]
    }[] => {
        // schedule entries are cells grouped by three rows (and merged columns)
        const values = grid.getValues();
        const gridHeight = grid.getHeight();
        const gridWidth = grid.getWidth();
        const gridStartRow = grid.getRow();
        const gridStartColumn = grid.getColumn();

        const scheduleEntries: { offset: Offset, values: string[] }[] = [];
        for (let startRow = 0; startRow < gridHeight; startRow += 3) {
            const columnsToSkip = new Set();
            for (let targetColumn = 0; targetColumn < gridWidth; targetColumn++) {
                if (columnsToSkip.has(targetColumn)) {
                    continue;
                }
                // check if it has data
                const scheduleEntryValues = [values[startRow][targetColumn], values[startRow + 1][targetColumn], values[startRow + 2][targetColumn]];
                if (scheduleEntryValues.every(value => !value)) {
                    continue;
                }

                let scheduleEntryOffset = {
                    row: gridStartRow + startRow + 1,
                    column: gridStartColumn + targetColumn + 1,
                    numRows: 3,
                    numColumns: 1
                }

                // check if it is merged
                const mergedRange = _fastCheckMergedRows(startRow, targetColumn);
                if (mergedRange) {
                    const numColumns = Profiler.wrap(mergedRange.getNumColumns, "getNumColumns")();
                    scheduleEntryOffset = {
                        row: mergedRange.getRow(),
                        column: mergedRange.getColumn(),
                        numRows: 3,
                        numColumns: numColumns
                    }
                    for (let i = 1; i < numColumns; i++) {
                        columnsToSkip.add(targetColumn + i);
                    }
                }

                scheduleEntries.push({
                    offset: scheduleEntryOffset,
                    values: scheduleEntryValues
                });
            }
        }
        return scheduleEntries;
    }

    _splitIntoScheduleEntries = Profiler.wrap(_splitIntoScheduleEntries);

    let _lintScheduleGrid = (grid: GoogleAppsScript.Spreadsheet.Range) => {
        const warnings = [];
        const scheduleEntries = _splitIntoScheduleEntries(grid);
        const _lintScheduleEntriesStart = new Date();
        for (let i = 0; i < scheduleEntries.length; i++) {
            const scheduleEntry = scheduleEntries[i];
            if (scheduleEntry)
                warnings.push(..._lintScheduleEntry(scheduleEntry.offset, scheduleEntry.values));
        }
        Logger.log("Linting schedule entries took " + (new Date().getTime() - _lintScheduleEntriesStart.getTime()) + " ms");
        return warnings;
    }

    _lintScheduleGrid = Profiler.wrap(_lintScheduleGrid);

    let _lintScheduleEntry = (scheduleEntryOffset: Offset, rowsValues: string[]) => {
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
        // if (locationValue)
        //     lintLocation(locationValue, scheduleEntryRange, warnings);

        return warnings;
    }

    let _getScheduleGrids = (warnings: Warning[]) => {
        const gridColumns = _getTimeColumns(warnings);
        gridColumns.push(currentTargetRange.getWidth());
        const gridRows = _getDayRows(warnings);
        gridRows.push(currentTargetRange.getHeight());
        const grids = [];

        // grid is not selected: pairwise iterate over columns and rows
        for (let i = 0; i < gridColumns.length - 1; i++) {
            const startColumn = gridColumns[i] + 1;
            const endColumn = gridColumns[i + 1];
            for (let j = 0; j < gridRows.length - 1; j++) {
                const startRow = gridRows[j] + 1;
                const endRow = gridRows[j + 1];
                // only between start and end columns and rows
                grids.push(currentTargetRange.offset(startRow, startColumn, endRow - startRow, endColumn - startColumn));
            }
        }

        const gridA1Notations = grids.map(grid => grid.getA1Notation());
        Logger.log("Grids: " + gridA1Notations);

        return grids;
    }

    let _lintSubject = (subjectValue: string, scheduleEntryOffset: Offset, warnings: Warning[]) => {
        // regex for subject value:
        // <subject>
        // <subject> (<type>)

        const match = subjectValue.match(subjectTypePattern);
        if (!match) {
            warnings.push({
                content: "Subject with wrong format: `" + subjectValue + "`, should be `subject (type)`",
                range: A1.offsetToA1Notation(scheduleEntryOffset)
            });
        } else {
            const subject = match[1];
            // maybe without type
            const type = match[2] ? match[2] : null;
            if (!Settings.getSettings().subjects.has(subject)) {
                warnings.push({
                    content: "Unknown subject " + subject,
                    range: A1.offsetToA1Notation(scheduleEntryOffset)
                });
                unknownSubjects.add(subject);
            }
            if (type && !allowedSubjectTypes.has(type)) {
                warnings.push({
                    content: "Unknown subject type `" + type + "`. Allowed types: `" + Array.from(allowedSubjectTypes).join("`, `") + "`",
                    range: A1.offsetToA1Notation(scheduleEntryOffset)
                });
            }
        }
    }

    _getScheduleGrids = Profiler.wrap(_getScheduleGrids);

    let _getTimeColumns = (warnings: Warning[]) => {
        const values = currentTargetValues;
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
    }

    let _getDayRows = (warnings: Warning[]) => {
        const values = currentTargetValues;
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
    }
}
/** END src/scheduleLinter.ts */
