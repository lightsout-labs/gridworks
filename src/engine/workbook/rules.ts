import { formatValue, GENERAL, parseInput } from '../format';
import type { Area } from '../formula';
import { isError, type Value } from '../value';
import type { Style } from './style';

/** What a cell's value must be for a rule to format it (R20). */
export type Condition =
  | {
      /** Compared with a value typed as in a cell: a number, text or date. */
      readonly type: 'greater' | 'less' | 'equal';
      readonly value: string;
    }
  | {
      /** Between two values, both included, in either order. */
      readonly type: 'between';
      readonly value: string;
      readonly other: string;
    }
  | {
      /** The text of the value contains `value`, in any case. */
      readonly type: 'contains';
      readonly value: string;
    }
  | {
      /** The value appears more than once in the rule's ranges. */
      readonly type: 'duplicate';
    }
  | {
      /** The number is among the `count` largest or smallest, ties included. */
      readonly type: 'top' | 'bottom';
      readonly count: number;
    };

/** What a rule changes in the cells it formats. */
export type RuleFormat = Pick<
  Style,
  'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fill'
>;

/** A conditional format: the cells it applies to, when, and how they look. */
export interface Rule {
  readonly ranges: readonly Area[];
  readonly condition: Condition;
  readonly format: RuleFormat;
}

/**
 * What a rule needs to know of all the values of its ranges: how often
 * each value appears, or the number a value must reach to be at the top or
 * bottom.
 */
type Totals = ReadonlyMap<string, number> | number | null;

/** Visits the values of the cells of an area that are not empty. */
export type VisitCells = (area: Area, visit: (value: Value) => void) => void;

/**
 * Works out the formats that rules give cells, keeping what a rule needs
 * of all its values until `forget` is called.
 */
export class Rules {
  private totals = new Map<Rule, Totals>();

  /** Forgets what was worked out of the values, after a change. */
  forget(): void {
    if (this.totals.size > 0) this.totals = new Map();
  }

  /**
   * The format the rules give a cell, the first rule's first where they
   * differ, as Excel's order of rules is their priority; `undefined` when
   * none applies. `visit` reads the values of the rules' sheet.
   */
  format(
    rules: readonly Rule[],
    row: number,
    col: number,
    value: Value,
    visit: VisitCells,
  ): RuleFormat | undefined {
    let result: Record<string, unknown> | undefined;
    for (const rule of rules) {
      if (!rule.ranges.some((area) => inside(area, row, col))) continue;
      if (!this.matches(rule, value, visit)) continue;
      result ??= {};
      for (const [property, setting] of Object.entries(rule.format)) {
        if (!(property in result)) result[property] = setting;
      }
    }
    return result as RuleFormat | undefined;
  }

  private matches(rule: Rule, value: Value, visit: VisitCells): boolean {
    if (value === null || isError(value)) return false;
    const { condition } = rule;
    switch (condition.type) {
      case 'greater':
        return compare(value, typed(condition.value)) > 0;
      case 'less':
        return compare(value, typed(condition.value)) < 0;
      case 'equal':
        return compare(value, typed(condition.value)) === 0;
      case 'between': {
        const one = typed(condition.value);
        const two = typed(condition.other);
        const [low, high] = compare(one, two) <= 0 ? [one, two] : [two, one];
        return compare(value, low) >= 0 && compare(value, high) <= 0;
      }
      case 'contains':
        return textOf(value)
          .toLowerCase()
          .includes(condition.value.toLowerCase());
      case 'duplicate': {
        const counts = this.totalsOf(rule, visit) as ReadonlyMap<
          string,
          number
        >;
        return (counts.get(identity(value)) ?? 0) > 1;
      }
      case 'top':
      case 'bottom': {
        const limit = this.totalsOf(rule, visit) as number | null;
        if (typeof value !== 'number' || limit === null) return false;
        return condition.type === 'top' ? value >= limit : value <= limit;
      }
    }
  }

  private totalsOf(rule: Rule, visit: VisitCells): Totals {
    if (this.totals.has(rule)) return this.totals.get(rule)!;
    const counts = new Map<string, number>();
    const numbers: number[] = [];
    for (const area of rule.ranges) {
      visit(area, (value) => {
        if (value === null || isError(value)) return;
        if (typeof value === 'number') numbers.push(value);
        const each = identity(value);
        counts.set(each, (counts.get(each) ?? 0) + 1);
      });
    }
    let totals: Totals = counts;
    const { condition } = rule;
    if (condition.type === 'top' || condition.type === 'bottom') {
      numbers.sort((a, b) => (condition.type === 'top' ? b - a : a - b));
      const count = Math.max(1, Math.floor(condition.count));
      totals =
        numbers.length === 0
          ? null
          : numbers[Math.min(count, numbers.length) - 1]!;
    }
    this.totals.set(rule, totals);
    return totals;
  }
}

function inside(area: Area, row: number, col: number): boolean {
  return (
    row >= area.top &&
    row <= area.bottom &&
    col >= area.left &&
    col <= area.right
  );
}

type Constant = number | string | boolean;

/** A value typed in a rule, read as it would be in a cell. */
function typed(text: string): Constant {
  return parseInput(text).value ?? '';
}

/** Numbers before text before booleans, text in any case, as Excel sorts. */
function compare(a: Constant, b: Constant): number {
  const rank = (value: Constant) =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') {
    const [x, y] = [a.toLowerCase(), (b as string).toLowerCase()];
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(a) - Number(b);
}

/** The text of a value, as a formula's `SEARCH` reads it. */
function textOf(value: Constant): string {
  return typeof value === 'string' ? value : formatValue(value, GENERAL);
}

/** A key that values equal for duplicates share: text in any case. */
function identity(value: Constant): string {
  return typeof value === 'string'
    ? `s${value.toLowerCase()}`
    : `${typeof value}${String(value)}`;
}
