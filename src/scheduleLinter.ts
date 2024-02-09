/** BEGIN src/scheduleLinter.ts */
namespace ScheduleLinter {
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

  enum ModifierType {
    ON_TBA_EVENT,
    ONLY_ON_EVENT,
    ON_ADDITIONAL,
    FROM_EVENT,
    FROM_ADDITIONAL,
    UNTIL_EVENT,
    UNTIL_ADDITIONAL,
    STARTS_AT_EVENT,
    ENDS_AT_EVENT,
    LOCATION_EVENT,
  }

  const days = [
    "MONDAY",
    "TUESDAY",
    "WEDNESDAY",
    "THURSDAY",
    "FRIDAY",
    "SATURDAY",
  ];

  export let lintHeader = () => {
    // first row is courses
    // second row is groups

    const startDate = new Date();
    const warnings: Warning[] = [];

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
          content:
            "Group with wrong format: `" +
            value +
            "`, should be `group (count of students)`",
          range: range.getCell(2, column + 1).getA1Notation(),
        });
      } else {
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

    Logger.log(
      "Header Linting took " +
        (new Date().getTime() - startDate.getTime()) +
        " ms",
    );

    for (const warning of warnings) {
      warning.gid = Current.getTargetSheetId();
    }
    return warnings;
  };

  lintHeader = Profiler.wrap(lintHeader);

  export let lintSchedule = () => {
    const startDate = new Date();

    unknownSubjects.clear();
    const warnings: Warning[] = [];
    const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
    const endOfScheduleRow = Current.getTargetSheet().getMaxRows();
    Current.setTargetRange(
      Current.getTargetSheet().getRange(
        1,
        1,
        endOfScheduleRow,
        endOfScheduleColumn,
      ),
    );

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
      Logger.log(
        "Unknown subjects to cache: " +
          Array.from(unknownSubjects).join("\r\n"),
      );
      CacheService.getDocumentCache().put(
        "unknownSubjects",
        Array.from(unknownSubjects).join("\r\n"),
        100,
      );
    } else {
      Logger.log("No unknown subjects to cache");
    }

    if (unknownLocations.size) {
      Logger.log(
        "Unknown locations to cache: " +
          Array.from(unknownLocations).join("\r\n"),
      );
      CacheService.getDocumentCache().put(
        "unknownLocations",
        Array.from(unknownLocations).join("\r\n"),
        100,
      );
    } else {
      Logger.log("No unknown locations to cache");
    }

    Logger.log(
      "Whole schedule Linting took " +
        (new Date().getTime() - startDate.getTime()) +
        " ms",
    );

    for (const warning of warnings) {
      warning.gid = Current.getTargetSheetId();
    }

    Logger.log("Profiler: \n" + Profiler.format());
    return warnings;
  };

  lintSchedule = Profiler.wrap(lintSchedule);

  export let selectScheduleGrids = () => {
    const warnings: Warning[] = [];
    const endOfScheduleColumn = _getEndOfScheduleColumn(warnings);
    const endOfScheduleRow = Current.getTargetSheet().getMaxRows();
    Current.setTargetRange(
      Current.getTargetSheet().getRange(
        1,
        1,
        endOfScheduleRow,
        endOfScheduleColumn,
      ),
    );

    const grids = _getScheduleGrids(warnings);

    if (grids) {
      const gridA1Notations = grids.map((grid) => grid.getA1Notation());
      const gridRanges = Current.getSpreadsheet().getRangeList(gridA1Notations);
      Current.getSpreadsheet().setActiveRangeList(gridRanges);
    } else {
      warnings.push({
        content: "No schedule grids found",
      });
    }

    for (const warning of warnings) {
      warning.gid = Current.getTargetSheetId();
    }

    return warnings;
  };

  selectScheduleGrids = Profiler.wrap(selectScheduleGrids);

  export let addUnknownSubjectsToSettings = () => {
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
    const toFillRange = settingsRange.offset(
      rowOffset,
      subjectsColumn,
      unknownSubjects.length,
      1,
    );
    toFillRange.setValues(unknownSubjects.map((subject) => [subject]));
    toFillRange.activate();
  };

  export let addUnknownLocationsToSettings = () => {
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
    const toFillRange = settingsRange.offset(
      rowOffset,
      locationsColumn,
      unknownLocations.length,
      1,
    );
    toFillRange.setValues(unknownLocations.map((location) => [location]));
    toFillRange.activate();
  };

  let _getEndOfScheduleColumn = (warnings: Warning[]) => {
    // values in header row equal to "-"
    const headerRow = Current.getSpreadsheet()
      .getActiveSheet()
      .getRange(
        1,
        1,
        1,
        Current.getSpreadsheet().getActiveSheet().getLastColumn(),
      );
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

  let _splitIntoScheduleEntries = (
    grid: GoogleAppsScript.Spreadsheet.Range,
  ): {
    offset: Offset;
    values: string[];
  }[] => {
    // schedule entries are cells grouped by three rows (and merged columns)
    const values = grid.getValues();
    const gridHeight = grid.getHeight();
    const gridWidth = grid.getWidth();
    const gridStartRow = grid.getRow();
    const gridStartColumn = grid.getColumn();

    const scheduleEntries: { offset: Offset; values: string[] }[] = [];
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
        const mergedRange = _fastCheckMergedRows(
          scheduleEntryOffset.row,
          scheduleEntryOffset.column,
        );
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

  let _lintScheduleGrid = (grid: GoogleAppsScript.Spreadsheet.Range) => {
    const warnings = [];
    const scheduleEntries = _splitIntoScheduleEntries(grid);
    const _lintScheduleEntriesStart = new Date();
    for (let i = 0; i < scheduleEntries.length; i++) {
      const scheduleEntry = scheduleEntries[i];
      Logger.log(
        "Linting schedule entry " + A1.offsetToA1Notation(scheduleEntry.offset),
      );

      if (scheduleEntry)
        warnings.push(
          ..._lintScheduleEntry(scheduleEntry.offset, scheduleEntry.values),
        );
    }
    Logger.log(
      "Linting schedule entries took " +
        (new Date().getTime() - _lintScheduleEntriesStart.getTime()) +
        " ms",
    );
    return warnings;
  };

  _lintScheduleGrid = Profiler.wrap(_lintScheduleGrid);

  let _lintScheduleEntry = (
    scheduleEntryOffset: Offset,
    rowsValues: string[],
  ) => {
    const warnings = [];
    // schedule entry with three rows:
    // subject row
    // teacher row
    // location row
    const subjectValue = rowsValues[0];
    const teacherValue = rowsValues[1];
    const locationValue = rowsValues[2];

    if (subjectValue) _lintSubject(subjectValue, scheduleEntryOffset, warnings);
    // if (teacherValue)
    //     lintTeacher(teacherValue, scheduleEntryRange, warnings);
    if (locationValue)
      _lintLocation(locationValue.toString(), scheduleEntryOffset, warnings);

    return warnings;
  };

  _lintScheduleEntry = Profiler.wrap(_lintScheduleEntry);

  let _getScheduleGrids = (warnings: Warning[]) => {
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
        grids.push(
          range.offset(
            startRow,
            startColumn,
            endRow - startRow,
            endColumn - startColumn,
          ),
        );
      }
    }

    const gridA1Notations = grids.map((grid) => grid.getA1Notation());
    Logger.log("Grids: " + gridA1Notations);

    return grids;
  };

  _getScheduleGrids = Profiler.wrap(_getScheduleGrids);

  let _lintSubject = (
    subjectValue: string,
    scheduleEntryOffset: Offset,
    warnings: Warning[],
  ) => {
    // regex for subject value:
    // <subject>
    // <subject> (<type>)

    const match = subjectValue.match(subjectTypePattern);
    if (!match) {
      warnings.push({
        content:
          "Subject with wrong format: `" +
          subjectValue +
          "`, should be `subject (type)`",
        range: A1.offsetToA1Notation(scheduleEntryOffset),
      });
    } else {
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
          content:
            `Unknown subject '${subject}'` +
            (nearestSubject ? ` (did you mean '${nearestSubject}'?)` : ""),
          range: A1.offsetToA1Notation(scheduleEntryOffset),
        });
        unknownSubjects.add(subject);
      }
      if (type && !allowedSubjectTypes.has(type)) {
        warnings.push({
          content:
            "Unknown subject type `" +
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

  let _lintLocation = (
    locationValue: string,
    scheduleEntryOffset: Offset,
    warnings: Warning[],
  ) => {
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
    } else {
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
      const modifiers = _lintModifiers(
        modifiersStrings,
        scheduleEntryOffset,
        warnings,
      );
    }
  };
  let _lintModifiers = (
    modifiersStrings: string[],
    scheduleEntryOffset: Offset,
    warnings: Warning[],
  ) => {
    let modifiers: {
      origin: string;
      eventModifier: string;
      eventModifierType: ModifierType;
      additionalModifier?: string;
      additionalModifierType?: ModifierType;
    }[] = [];

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
            } else {
              warnings.push({
                content: `Unknown additional modifier '${additionalModifier}'`,
                range: A1.offsetToA1Notation(scheduleEntryOffset),
              });
            }
          } else {
            modifiers.push({
              origin: modifiersStrings[i],
              eventModifier: eventModifier,
              eventModifierType: type,
            });
          }
        } else {
          warnings.push({
            content: `Unknown event modifier '${eventModifier}'`,
            range: A1.offsetToA1Notation(scheduleEntryOffset),
          });
        }
      }
    }
    return modifiers;
  };

  let _getModifierType = (modifier: string, is_additional: boolean = false) => {
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
    } else {
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

  let _getTimeColumns = (warnings: Warning[]) => {
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

  let _getDayRows = (warnings: Warning[]) => {
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
}
/** END src/scheduleLinter.ts */
