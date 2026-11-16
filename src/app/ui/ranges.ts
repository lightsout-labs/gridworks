import { parse, references } from '../../engine/formula';
import {
  cellName,
  columnName,
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Range,
} from '../../engine/workbook';

/**
 * The ranges of a text such as `A1:B2,D:E,3:4`, or `null` when any part of
 * it is not a range of the sheet shown.
 */
export function parseRanges(text: string): Range[] | null {
  const ranges: Range[] = [];
  for (const part of text.split(',')) {
    const ast = parse(part.trim());
    if (ast?.type !== 'range' || ast.sheet !== undefined) return null;
    ranges.push(references(ast)[0]!.area);
  }
  return ranges;
}

/** The text of ranges, as `parseRanges` reads it. */
export function rangesText(ranges: readonly Range[]): string {
  return ranges
    .map(({ top, left, bottom, right }) => {
      if (top === 0 && bottom === SHEET_ROWS - 1) {
        return `${columnName(left)}:${columnName(right)}`;
      }
      if (left === 0 && right === SHEET_COLUMNS - 1) {
        return `${top + 1}:${bottom + 1}`;
      }
      const start = cellName(top, left);
      return top === bottom && left === right
        ? start
        : `${start}:${cellName(bottom, right)}`;
    })
    .join(',');
}
