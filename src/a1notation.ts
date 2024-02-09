/** BEGIN src/a1notation.ts */
namespace A1 {
  export function offsetFromA1Notation(A1Notation: string): {
    startRow: number;
    startColumn: number;
    endRow: number;
    endColumn: number;
  } {
    // start:end
    const match = A1Notation.match(/(^[A-Z]+[0-9]+)|([A-Z]+[0-9]+$)/gm);

    if (match.length !== 2) {
      throw new Error(
        "The given value was invalid. Cannot convert Google Sheet A1 notation to indexes",
      );
    }

    const start_notation = match[0];
    const end_notation = match[1];

    const start = cellFromA1Notation(start_notation);
    const end = cellFromA1Notation(end_notation);

    return {
      startRow: start.row,
      startColumn: start.column,
      endRow: end.row,
      endColumn: end.column,
    };
  }

  function cellFromA1Notation(A1Notation: string) {
    const match = A1Notation.match(/(^[A-Z]+)|([0-9]+$)/gm);

    if (match.length !== 2) {
      throw new Error(
        "The given value was invalid. Cannot convert Google Sheet A1 notation to indexes",
      );
    }

    const column_notation = match[0];
    const row_notation = match[1];

    const column = fromColumnA1Notation(column_notation);
    const row = fromRowA1Notation(row_notation);

    return { row, column };
  }

  function fromRowA1Notation(A1Row: string) {
    const num = parseInt(A1Row, 10);
    if (Number.isNaN(num)) {
      throw new Error(
        "The given value was not a valid number. Cannot convert Google Sheet row notation to index",
      );
    }

    return num;
  }

  function fromColumnA1Notation(A1Column: string) {
    const A = "A".charCodeAt(0);

    let output = 0;
    for (let i = 0; i < A1Column.length; i++) {
      const next_char = A1Column.charAt(i);
      const column_shift = 26 * i;

      output += column_shift + (next_char.charCodeAt(0) - A);
    }

    return output + 1;
  }

  export function offsetToA1Notation(offset: Offset) {
    return rangeToA1Notation(
      offset.row,
      offset.column,
      offset.numRows,
      offset.numColumns,
    );
  }

  function rangeToA1Notation(
    row: number,
    column: number,
    numRows: number,
    numColumns: number,
  ) {
    Logger.log("row: " + row);
    Logger.log("column: " + column);
    const start = toA1Notation(row, column);
    const end = toA1Notation(row + numRows - 1, column + numColumns - 1);

    return `${start}:${end}`;
  }

  function toA1Notation(row: number, column: number) {
    const row_notation = row.toString();
    const column_notation = toColumnA1Notation(column);

    return column_notation + row_notation;
  }

  function toColumnA1Notation(column: number) {
    const A = "A".charCodeAt(0);

    let output = "";
    while (column > 0) {
      const remainder = column % 26;
      output = String.fromCharCode(A + remainder - 1) + output;
      column = Math.floor(column / 26);
    }

    return output;
  }
}
/** END src/a1notation.ts */
