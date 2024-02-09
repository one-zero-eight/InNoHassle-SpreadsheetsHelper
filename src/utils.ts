/** BEGIN src/utils.ts */
let mergedRangeRegistry: {
  [key: number]: { [key: number]: GoogleAppsScript.Spreadsheet.Range };
} = {};

let _createMergedRangeRegistry = (
  mergedRanges: GoogleAppsScript.Spreadsheet.Range[],
) => {
  // accessible by (row, column) in O(1)
  const mergedRangeRegistry: {
    [key: number]: { [key: number]: GoogleAppsScript.Spreadsheet.Range };
  } = {};

  for (let i = 0; i < mergedRanges.length; i++) {
    const mergedRange = mergedRanges[i];
    const notation = mergedRange.getA1Notation();
    const range = A1.offsetFromA1Notation(notation);
    for (let x = range.startRow; x <= range.endRow; x++) {
      if (!mergedRangeRegistry[x]) {
        mergedRangeRegistry[x] = {};
      }
      for (let y = range.startColumn; y <= range.endColumn; y++) {
        mergedRangeRegistry[x][y] = mergedRange;
      }
    }
  }
  return mergedRangeRegistry;
};

_createMergedRangeRegistry = Profiler.wrap(_createMergedRangeRegistry);

let _fastCheckMergedRows = (column: number, row: number) => {
  // using registry to check if cell is merged
  if (mergedRangeRegistry[column] && mergedRangeRegistry[column][row]) {
    return mergedRangeRegistry[column][row];
  }
  return null;
};

_fastCheckMergedRows = Profiler.wrap(_fastCheckMergedRows);

/**
 * @param {string} s1 Source string
 * @param {string} s2 Target string
 * @param {object} [costs] Costs for operations { [replace], [replaceCase], [insert], [remove] }
 * @return {number} Levenshtein distance
 */
function levenshtein(
  s1: string,
  s2: string,
  costs: any = {
    replace: 1,
    replaceCase: 1,
    insert: 1,
    remove: 1,
  },
): number {
  let i, j, l1, l2, flip, ch, chl, ii, ii2, cost, cutHalf;
  l1 = s1.length;
  l2 = s2.length;

  const cr = costs.replace;
  const cri = costs.replaceCase;
  const ci = costs.insert;
  const cd = costs.remove;

  cutHalf = flip = Math.max(l1, l2);

  const minCost = Math.min(cd, ci, cr);
  const minD = Math.max(minCost, (l1 - l2) * cd);
  const minI = Math.max(minCost, (l2 - l1) * ci);
  const buf = new Array(cutHalf * 2 - 1);

  for (i = 0; i <= l2; ++i) {
    buf[i] = i * minD;
  }

  for (i = 0; i < l1; ++i, flip = cutHalf - flip) {
    ch = s1[i];
    chl = ch.toLowerCase();

    buf[flip] = (i + 1) * minI;

    ii = flip;
    ii2 = cutHalf - flip;

    for (j = 0; j < l2; ++j, ++ii, ++ii2) {
      cost = ch === s2[j] ? 0 : chl === s2[j].toLowerCase() ? cri : cr;
      buf[ii + 1] = Math.min(buf[ii2 + 1] + cd, buf[ii] + ci, buf[ii2] + cost);
    }
  }
  return buf[l2 + cutHalf - flip];
}

/** END src/utils.ts */
