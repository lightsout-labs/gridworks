import type { Area } from '../formula';
import type { Value } from '../value';

/** How a filter compares text, in any case. */
export type TextTest = 'equals' | 'contains' | 'begins' | 'ends';

/** How a filter compares numbers; `between` includes both ends. */
export type NumberTest = 'equals' | 'greater' | 'less' | 'between';

/** What the rows of a filter must have in one column to be shown (R24). */
export type ColumnFilter =
  | {
      readonly col: number;
      /** Display texts checked in the list, `''` for empty cells. */
      readonly type: 'values';
      readonly values: readonly string[];
    }
  | {
      readonly col: number;
      readonly type: 'text';
      readonly test: TextTest;
      readonly text: string;
    }
  | {
      readonly col: number;
      readonly type: 'number';
      readonly test: NumberTest;
      readonly value: number;
      /** The upper end of `between`. */
      readonly to?: number;
    };

/**
 * A sheet's filter: a range whose first row holds the headers, and what the
 * rows below it must have in some of its columns to be shown.
 */
export interface Filter {
  readonly range: Area;
  readonly columns: readonly ColumnFilter[];
}

/** Whether a cell, of this value and display text, meets a column filter. */
export function meets(
  filter: ColumnFilter,
  value: Value,
  display: string,
): boolean {
  if (filter.type === 'values') return filter.values.includes(display);
  if (filter.type === 'text') {
    const text = display.toLowerCase();
    const wanted = filter.text.toLowerCase();
    return filter.test === 'equals'
      ? text === wanted
      : filter.test === 'contains'
        ? text.includes(wanted)
        : filter.test === 'begins'
          ? text.startsWith(wanted)
          : text.endsWith(wanted);
  }
  if (typeof value !== 'number') return false;
  return filter.test === 'equals'
    ? value === filter.value
    : filter.test === 'greater'
      ? value > filter.value
      : filter.test === 'less'
        ? value < filter.value
        : value >= filter.value && value <= (filter.to ?? filter.value);
}
