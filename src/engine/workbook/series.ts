import { dateParts, dateSerial, isDateFormat } from '../format';
import type { Constant } from './calc';

/** A cell of a row or column that the fill handle extends. */
export interface Seed {
  readonly value: Constant;
  readonly formula: boolean;
  /** The cell's number format code. */
  readonly format: string;
}

/** The value at a place of a series, the seeds being at 0, 1, 2 and on. */
export type Series = (place: number) => Constant;

// The names that a fill counts on, as Excel does, in English.
const NAMES: readonly (readonly string[])[] = [
  ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'],
  [
    'sunday',
    'monday',
    'tuesday',
    'wednesday',
    'thursday',
    'friday',
    'saturday',
  ],
  [
    'jan',
    'feb',
    'mar',
    'apr',
    'may',
    'jun',
    'jul',
    'aug',
    'sep',
    'oct',
    'nov',
    'dec',
  ],
  [
    'january',
    'february',
    'march',
    'april',
    'may',
    'june',
    'july',
    'august',
    'september',
    'october',
    'november',
    'december',
  ],
];

const modulo = (n: number, m: number) => ((n % m) + m) % m;

/**
 * The series that a row or column of cells starts, as the fill handle
 * extends it, or `null` when it starts none and is repeated instead:
 * numbers by their step (a trend line for more than two), dates by days or
 * by months when they fall on the same day, and weekday and month names.
 * One date or name counts on by one; one number is repeated.
 */
export function series(seeds: readonly Seed[]): Series | null {
  if (seeds.some((seed) => seed.formula || seed.value === null)) return null;
  const values = seeds.map((seed) => seed.value);
  if (values.every((value) => typeof value === 'string')) {
    return names(values as string[]);
  }
  if (!values.every((value) => typeof value === 'number')) return null;
  const numbers = values as number[];
  const dates = seeds.every((seed) => isDateFormat(seed.format));
  if (numbers.length === 1) {
    return dates ? (place) => numbers[0]! + place : null;
  }
  return (dates && months(numbers)) || trend(numbers);
}

/** Weekday or month names, each step the same, in the case of the first. */
function names(texts: readonly string[]): Series | null {
  for (const list of NAMES) {
    const places = texts.map((text) => list.indexOf(text.toLowerCase()));
    if (places.includes(-1)) continue;
    const size = list.length;
    const step =
      places.length === 1 ? 1 : modulo(places[1]! - places[0]!, size);
    const even = places.every(
      (place, index) => place === modulo(places[0]! + step * index, size),
    );
    if (!even) return null;
    const first = texts[0]!;
    const cased = (name: string) =>
      first === first.toUpperCase()
        ? name.toUpperCase()
        : first === first.toLowerCase()
          ? name
          : name[0]!.toUpperCase() + name.slice(1);
    return (place) => cased(list[modulo(places[0]! + step * place, size)]!);
  }
  return null;
}

/**
 * Dates on the same day of the month, the same number of months apart,
 * continued on that day, or on the last of a month without it.
 */
function months(dates: readonly number[]): Series | null {
  const parts = dates.map(dateParts);
  const [year, month, day] = parts[0]!;
  const index = ([y, m]: readonly number[]) => y! * 12 + m!;
  const step = index(parts[1]!) - index(parts[0]!);
  const even = parts.every(
    (each, place) =>
      each[2] === day && index(each) === index(parts[0]!) + step * place,
  );
  if (step === 0 || !even) return null;
  return (place) => {
    const target = month + step * place;
    return dateSerial(year, target, Math.min(day, lastDay(year, target)));
  };
}

const lastDay = (year: number, month: number) =>
  dateSerial(year, month + 1, 0) - dateSerial(year, month, 0);

/**
 * The least-squares line through the numbers, which is their step when
 * they have one, kept to the 15 significant digits of Excel.
 */
function trend(numbers: readonly number[]): Series {
  const count = numbers.length;
  const middle = (count - 1) / 2;
  const mean = numbers.reduce((sum, n) => sum + n, 0) / count;
  let covariance = 0;
  let variance = 0;
  numbers.forEach((n, place) => {
    covariance += (place - middle) * (n - mean);
    variance += (place - middle) ** 2;
  });
  const slope = covariance / variance;
  return (place) => Number((mean + slope * (place - middle)).toPrecision(15));
}
