import {
  dateFromSerial,
  dateToSerial,
  format,
  formatColor as numfmtColor,
  getFormatDateInfo,
  getFormatFromId,
  getFormatInfo,
  isDateFormat,
  isPercentFormat,
  isTextFormat,
  isValidFormat,
  parseValue,
  tokenize,
} from 'numfmt';
import { isError, type Value } from '../value';

export { isDateFormat, isTextFormat };

/** The format of a cell that was given none. */
export const GENERAL = 'General';

/** What a typed entry stands for. */
export interface Input {
  readonly value: number | string | boolean | null;
  /** The format code that the way it was typed implies, as `0%` for `12%`. */
  readonly format?: string;
  /** The year given to a date typed without one. */
  readonly year?: number;
}

/**
 * Recognizes a typed entry as Excel does: a number, a date, a time, a
 * percentage or a boolean, with the format it implies; an empty cell when
 * nothing was typed; and otherwise the text itself. Formulas are not
 * recognized here. A date typed without a year is in `year`, by default
 * the current one.
 */
export function parseInput(
  text: string,
  year: number = new Date().getFullYear(),
): Input {
  if (text === '') return { value: null };
  const parsed = parseValue(text);
  if (!parsed) return { value: text };
  if (typeof parsed.value === 'boolean') return { value: parsed.value };
  if (parsed.formatId === undefined) return { value: parsed.value };
  const code = getFormatFromId(parsed.formatId);
  const info = getFormatDateInfo(code);
  if (!isDateFormat(code) || info.type === 'time' || info.year) {
    return { value: parsed.value, format: code };
  }
  // The parser gave it the year of the clock; it gets the one asked for.
  const [, month, day] = dateFromSerial(parsed.value);
  const time = parsed.value - Math.floor(parsed.value);
  const value = dateToSerial([year, month!, day!])! + time;
  return { value, format: code, year };
}

/**
 * The serial number of a text that reads as a date, a time or both, as
 * `DATEVALUE` and `TIMEVALUE` read it, if it does.
 */
export function parseDate(text: string): number | undefined {
  const { value, format } = parseInput(text.trim());
  return typeof value === 'number' && format && isDateFormat(format)
    ? value
    : undefined;
}

/**
 * The format code of one of Excel's built-in number formats, which files
 * name by number, or `null` for a number that names none.
 */
export function builtinFormat(id: number): string | null {
  return BUILT_IN.has(id) ? getFormatFromId(id) : null;
}

// The built-in formats that every language of Excel shares.
const BUILT_IN = new Set([
  0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38,
  39, 40, 45, 46, 47, 48, 49,
]);

/** The text that a cell shows for a value, in a format code. */
export function formatValue(value: Value, code: string = GENERAL): string {
  if (value === null) return '';
  if (isError(value)) return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  // Text goes through the text section of a format, as `"Text: "@`.
  if (typeof value === 'string') {
    return code === GENERAL ? value : format(code, value);
  }
  // The common case, which the grid meets in every cell it draws.
  if (code === GENERAL && Number.isInteger(value) && Math.abs(value) < 1e11) {
    return String(value);
  }
  return format(code, value);
}

/** Whether a format code is one Gridworks can show values in. */
export function validFormat(code: string): boolean {
  return code.trim() !== '' && isValidFormat(code);
}

/** The colors that format codes name, as Excel draws them. */
const COLORS: Readonly<Record<string, string>> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
};

/**
 * The color, as `#rrggbb`, that a format code gives a value, as `[Red]`
 * does, or `undefined` when it gives none.
 */
export function formatColor(value: Value, code: string): string | undefined {
  if (value === null || typeof value === 'boolean' || isError(value)) {
    return undefined;
  }
  const color = numfmtColor(code, value);
  if (typeof color !== 'string') return undefined;
  return COLORS[color.toLowerCase()] ?? color.toLowerCase();
}

/** The changes that the ribbon makes to the format of numbers. */
export type FormatChange = 'moreDecimals' | 'fewerDecimals' | 'thousands';

/**
 * A format code with one more or one fewer decimal place, or with the
 * thousands separator turned on or off, in each section that shows a
 * number. A General format starts from the decimals its value shows, as
 * in Excel; dates, times and text do not change.
 */
export function changedFormat(
  code: string,
  value: Value,
  change: FormatChange,
): string {
  if (code === GENERAL) code = shownDecimals(value);
  if (isDateFormat(code) || isTextFormat(code)) return code;
  const grouped = getFormatInfo(code).grouped === 1;
  const sections: string[][] = [[]];
  for (const token of tokenize(code)) {
    if (token.type === 'BREAK') sections.push([]);
    else sections.at(-1)!.push(token.raw);
  }
  return sections
    .map((raws) => changedSection(raws, change, grouped).join(''))
    .join(';');
}

/** The code that shows a value's number with the decimals General shows. */
function shownDecimals(value: Value): string {
  if (typeof value !== 'number') return '0';
  const [, decimals = ''] = formatValue(value).split('E')[0]!.split('.');
  return decimals === '' ? '0' : `0.${'0'.repeat(decimals.length)}`;
}

const DIGIT = /^[0#?]$/;

/** A section of a format code, as its tokens' text, changed. */
function changedSection(
  raws: readonly string[],
  change: FormatChange,
  grouped: boolean,
): string[] {
  // The number is the digits up to an exponent; a section with no `0` or
  // `#`, as one of literal text, is left as it is.
  const end = raws.findIndex((raw) => /^E[+-]$/i.test(raw));
  const number = end === -1 ? raws.length : end;
  const digits = raws
    .slice(0, number)
    .flatMap((raw, index) => (DIGIT.test(raw) ? [index] : []));
  if (!digits.some((index) => raws[index] !== '?')) return [...raws];
  const first = digits[0]!;
  const last = digits.at(-1)!;
  const point = raws.indexOf('.');
  const result = [...raws];
  if (change === 'moreDecimals') {
    result.splice(
      last + 1,
      0,
      ...(point === -1 || point > last ? ['.', '0'] : ['0']),
    );
  } else if (change === 'fewerDecimals') {
    if (point !== -1 && point < last) result.splice(last, 1);
    else return result;
    if (point === last - 1) result.splice(point, 1);
  } else {
    const integerEnd = point === -1 || point > last ? last + 1 : point;
    const integer = result
      .slice(first, integerEnd)
      .filter((raw) => raw !== ',');
    const zeros = integer.filter((raw) => raw !== '#');
    const plain = zeros.length > 0 ? zeros : ['0'];
    result.splice(
      first,
      integerEnd - first,
      ...(grouped ? plain : ['#', ',', '#', '#', ...plain]),
    );
  }
  return result;
}

/**
 * The text that edits a value, as the formula bar shows it: numbers in
 * full, percentages with their sign, and dates and times in a form that is
 * read back as the same value.
 */
export function editText(value: Value, code: string = GENERAL): string {
  if (typeof value !== 'number') return formatValue(value);
  if (isPercentFormat(code)) return `${plain(value * 100)}%`;
  if (isDateFormat(code)) {
    const { type } = getFormatDateInfo(code);
    const parts = [];
    if (type !== 'time') parts.push('m/d/yyyy');
    if (type !== 'date') parts.push('h:mm:ss AM/PM');
    return format(parts.join(' '), value);
  }
  return plain(value);
}

/** A number with the 15 significant digits that Excel keeps. */
function plain(value: number): string {
  return String(Number(value.toPrecision(15)));
}

// Dates are Excel's serial numbers: 1 is 1 January 1900, and 60 the
// 29 February 1900 that Excel keeps from Lotus 1-2-3, so that every date
// from 1 March 1900 is one day later than a count of real days.
const DAY = 86_400_000;
const START = Date.UTC(1899, 11, 31);

/** The serial number of a date, with months and days past their ends carried. */
export function dateSerial(year: number, month: number, day: number): number {
  const months = year * 12 + month - 1;
  const first =
    (Date.UTC(Math.floor(months / 12), months % 12, 1) - START) / DAY;
  return (first >= 60 ? first + 1 : first) + day - 1;
}

/** The year, month and day of a serial number, without its time. */
export function dateParts(
  date: number,
): [year: number, month: number, day: number] {
  if (date === 0) return [1900, 1, 0];
  if (date === 60) return [1900, 2, 29];
  const real = new Date(START + (date < 60 ? date : date - 1) * DAY);
  return [real.getUTCFullYear(), real.getUTCMonth() + 1, real.getUTCDate()];
}
