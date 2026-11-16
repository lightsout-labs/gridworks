import { formatValue, GENERAL, parseInput } from '../format';
import type { Area } from '../formula';
import { isError, type Value } from '../value';
import { compare } from './sort';
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
    }
  | {
      /**
       * Fills each number with a color between those of the lowest number,
       * the median when there are three colors, and the highest.
       */
      readonly type: 'scale';
      readonly colors: readonly string[];
    }
  | {
      /** Draws in each number a bar as long as its share of the highest. */
      readonly type: 'bar';
      readonly color: string;
    };

/** What a rule changes in the cells it formats. */
export type RuleFormat = Pick<
  Style,
  'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fill'
>;

/**
 * A conditional format: the cells it applies to, when, and how they look.
 * The format of a scale or bar is empty, as its condition says how.
 */
export interface Rule {
  readonly ranges: readonly Area[];
  readonly condition: Condition;
  readonly format: RuleFormat;
}

/** A data bar: its color, and its length as a share of the cell's width. */
export interface DataBar {
  readonly color: string;
  readonly length: number;
}

/** What rules give a cell: a format over its style, and a data bar. */
export interface Conditional {
  readonly conditional?: RuleFormat;
  readonly bar?: DataBar;
}

/** The lowest, median and highest of the numbers of a rule's ranges. */
interface Spread {
  readonly low: number;
  readonly middle: number;
  readonly high: number;
}

/**
 * What a rule needs to know of all the values of its ranges: how often
 * each value appears, the number a value must reach to be at the top or
 * bottom, or the spread of the numbers.
 */
type Totals = ReadonlyMap<string, number> | number | Spread | null;

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
   * What the rules give a cell: the format, the first rule's first where
   * they differ, as Excel's order of rules is their priority, and the
   * first rule's data bar. `visit` reads the values of the rules' sheet.
   */
  format(
    rules: readonly Rule[],
    row: number,
    col: number,
    value: Value,
    visit: VisitCells,
  ): Conditional {
    let format: Record<string, unknown> | undefined;
    let bar: DataBar | undefined;
    for (const rule of rules) {
      if (!rule.ranges.some((area) => inside(area, row, col))) continue;
      const { condition } = rule;
      if (condition.type === 'scale' || condition.type === 'bar') {
        const spread = this.totalsOf(rule, visit) as Spread | null;
        if (typeof value !== 'number' || spread === null) continue;
        if (condition.type === 'bar') {
          bar ??= { color: condition.color, length: share(value, spread) };
          continue;
        }
        format ??= {};
        format.fill ??= scaleColor(value, spread, condition.colors);
        continue;
      }
      if (!this.matches(condition, value, rule, visit)) continue;
      format ??= {};
      for (const [property, setting] of Object.entries(rule.format)) {
        if (!(property in format)) format[property] = setting;
      }
    }
    return {
      ...(format && { conditional: format as RuleFormat }),
      ...(bar && { bar }),
    };
  }

  private matches(
    condition: Exclude<Condition, { type: 'scale' | 'bar' }>,
    value: Value,
    rule: Rule,
    visit: VisitCells,
  ): boolean {
    if (value === null || isError(value)) return false;
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
    if (condition.type === 'scale' || condition.type === 'bar') {
      numbers.sort((a, b) => a - b);
      totals =
        numbers.length === 0
          ? null
          : {
              low: numbers[0]!,
              middle: median(numbers),
              high: numbers[numbers.length - 1]!,
            };
    } else if (condition.type === 'top' || condition.type === 'bottom') {
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

/** The median of numbers in order, as Excel's 50th percentile. */
function median(sorted: readonly number[]): number {
  const at = (sorted.length - 1) / 2;
  return (sorted[Math.floor(at)]! + sorted[Math.ceil(at)]!) / 2;
}

/**
 * The length of a number's bar: its share of the way from zero, or the
 * lowest number below it, to the highest.
 */
function share(value: number, { low, high }: Spread): number {
  const from = Math.min(0, low);
  const to = Math.max(0, high);
  return to === from ? 0 : (value - from) / (to - from);
}

/**
 * The color of a number in a scale of two or three colors: the lowest
 * number has the first, the highest the last, and the median the middle.
 */
function scaleColor(
  value: number,
  { low, middle, high }: Spread,
  colors: readonly string[],
): string {
  const part = (from: number, to: number) =>
    to === from ? 0 : (value - from) / (to - from);
  if (colors.length < 3) return mix(colors[0]!, colors[1]!, part(low, high));
  return value <= middle
    ? mix(colors[0]!, colors[1]!, part(low, middle))
    : mix(colors[1]!, colors[2]!, part(middle, high));
}

/** The color a share of the way from one `#rrggbb` color to another. */
function mix(from: string, to: string, share: number): string {
  let result = '#';
  for (const at of [1, 3, 5]) {
    const a = parseInt(from.slice(at, at + 2), 16);
    const b = parseInt(to.slice(at, at + 2), 16);
    const channel = Math.round(a + (b - a) * share);
    result += channel.toString(16).padStart(2, '0');
  }
  return result;
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
