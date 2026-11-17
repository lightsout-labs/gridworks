import { formatValue, GENERAL } from '../format';
import type { Area } from '../formula';
import { isError, type Value } from '../value';

/** What a cell's value must be (R26). */
export type Criterion =
  | {
      /**
       * One of a list: items separated by commas, or, after `=`, a
       * reference to a range whose shown texts are the items.
       */
      readonly type: 'list';
      readonly source: string;
    }
  | {
      /**
       * A number, a date (a number too), or text of a length, between
       * bounds that are both included; an absent bound sets no limit.
       */
      readonly type: 'number' | 'date' | 'length';
      readonly min?: number;
      readonly max?: number;
    };

/** A rule of data validation: the cells it applies to, and what it allows. */
export interface Validation {
  readonly ranges: readonly Area[];
  readonly criterion: Criterion;
}

/**
 * Whether a value, shown as `display` in its cell, meets a criterion,
 * given the items of a list. Errors meet none.
 */
export function meetsCriterion(
  criterion: Criterion,
  value: Value,
  display: string,
  items: readonly string[],
): boolean {
  if (isError(value)) return false;
  if (criterion.type === 'list') {
    const folded = display.toLowerCase();
    return items.some((item) => item.toLowerCase() === folded);
  }
  let measure: number;
  if (criterion.type === 'length') {
    measure = (typeof value === 'string' ? value : formatValue(value, GENERAL))
      .length;
  } else if (typeof value === 'number') {
    measure = value;
  } else {
    return false;
  }
  const { min, max } = criterion;
  return (
    (min === undefined || measure >= min) &&
    (max === undefined || measure <= max)
  );
}

/** The items typed in a list's source, or `null` for a reference. */
export function typedItems(source: string): string[] | null {
  if (source.startsWith('=')) return null;
  return source
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '');
}

/**
 * Validations with the cells of `ranges` taken out of each, and those left
 * with no cells dropped.
 */
export function without(
  validations: readonly Validation[],
  ranges: readonly Area[],
): Validation[] {
  return validations.flatMap((validation) => {
    let left = [...validation.ranges];
    for (const cut of ranges) left = left.flatMap((area) => minus(area, cut));
    return left.length === 0 ? [] : [{ ...validation, ranges: left }];
  });
}

/** The cells of `area` outside `cut`, as up to four areas. */
function minus(area: Area, cut: Area): Area[] {
  if (
    cut.top > area.bottom ||
    cut.bottom < area.top ||
    cut.left > area.right ||
    cut.right < area.left
  ) {
    return [area];
  }
  const parts: Area[] = [];
  if (area.top < cut.top) parts.push({ ...area, bottom: cut.top - 1 });
  if (area.bottom > cut.bottom) parts.push({ ...area, top: cut.bottom + 1 });
  const top = Math.max(area.top, cut.top);
  const bottom = Math.min(area.bottom, cut.bottom);
  if (area.left < cut.left) {
    parts.push({ top, bottom, left: area.left, right: cut.left - 1 });
  }
  if (area.right > cut.right) {
    parts.push({ top, bottom, left: cut.right + 1, right: area.right });
  }
  return parts;
}
