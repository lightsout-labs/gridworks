import { isError, type Value } from '../value';

/** A column to sort by, ascending unless `descending`. */
export interface SortKey {
  readonly col: number;
  readonly descending?: boolean;
}

/**
 * Two values that are not empty in Excel's ascending order: numbers, then
 * text in any case, then `FALSE` and `TRUE`, then errors, which are equal.
 */
export function compare(a: Value, b: Value): number {
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') {
    const [x, y] = [a.toLowerCase(), (b as string).toLowerCase()];
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (isError(a)) return 0;
  return Number(a) - Number(b);
}

function rank(value: Value): number {
  return typeof value === 'number'
    ? 0
    : typeof value === 'string'
      ? 1
      : typeof value === 'boolean'
        ? 2
        : 3;
}

/**
 * The indexes of rows in the order the keys sort them, `values(row, key)`
 * giving the value of a row for each key. Empty values go last in either
 * direction, as in Excel, and rows equal on every key keep their order.
 */
export function sortOrder(
  count: number,
  keys: readonly SortKey[],
  values: (row: number, key: number) => Value,
): number[] {
  const read = keys.map((_, key) =>
    Array.from({ length: count }, (_, row) => values(row, key)),
  );
  const rows = Array.from({ length: count }, (_, row) => row);
  // `Array.prototype.sort` is stable.
  return rows.sort((a, b) => {
    for (const [key, { descending }] of keys.entries()) {
      const x = read[key]![a]!;
      const y = read[key]![b]!;
      if (x === null || y === null) {
        if (x !== y) return x === null ? 1 : -1;
        continue;
      }
      const order = compare(x, y);
      if (order !== 0) return descending ? -order : order;
    }
    return 0;
  });
}
