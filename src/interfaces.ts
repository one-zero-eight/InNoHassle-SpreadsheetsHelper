/** BEGIN src/interfaces.ts */
interface Offset {
  row: number;
  column: number;
  numRows: number;
  numColumns: number;
}

interface Warning {
  content: string;
  gid?: number;
  range?: string; // A1 notation
}

/** END src/interfaces.ts */
