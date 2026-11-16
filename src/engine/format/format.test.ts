import { describe, expect, it } from 'vitest';
import {
  changedFormat,
  isGroupedFormat,
  editText,
  formatColor,
  formatValue,
  parseInput,
  validFormat,
} from './index';

describe('parseInput', () => {
  it.each([
    ['2', 2],
    [' 2.5 ', 2.5],
    ['-3', -3],
    ['.5', 0.5],
    ['(5)', -5],
  ])('recognizes %j as a plain number', (text, value) => {
    expect(parseInput(text)).toEqual({ value });
  });

  it.each([
    ['1,234.5', 1234.5, '#,##0.00'],
    ['12%', 0.12, '0%'],
    ['1e3', 1000, '0.00E+00'],
    ['1/2/2026', 46024, 'm/d/yyyy'],
    ['2026-01-02', 46024, 'm/d/yyyy'],
    ['2-Jan', 46024, 'd-mmm'],
    ['13:45', 0.5729166666666666, 'h:mm'],
    ['1:30 PM', 0.5625, 'h:mm AM/PM'],
  ])('recognizes %j with its format', (text, value, format) => {
    const parsed = parseInput(text);
    expect(parsed.format).toBe(format);
    // Dates take the current year when none is typed.
    if (text !== '2-Jan') expect(parsed.value).toBeCloseTo(value, 10);
  });

  it('gives a date typed without a year the year it is told, and says which', () => {
    expect(parseInput('2-Jan', 2030)).toEqual({
      value: 47485, // 1/2/2030
      format: 'd-mmm',
      year: 2030,
    });
    expect(parseInput('2-Jan').year).toBe(new Date().getFullYear());
    expect(parseInput('1/2/2026', 2030)).toEqual({
      value: 46024,
      format: 'm/d/yyyy',
    });
    expect(parseInput('13:45', 2030).year).toBeUndefined();
  });

  it.each([
    ['TRUE', true],
    ['false', false],
  ])('recognizes %j as a boolean', (text, value) => {
    expect(parseInput(text)).toEqual({ value });
  });

  it.each(['abc', '2 apples', '1.2.3', '-', 'TRUE!'])(
    'keeps %j as text',
    (text) => {
      expect(parseInput(text)).toEqual({ value: text });
    },
  );

  it('reads an empty entry as an empty cell', () => {
    expect(parseInput('')).toEqual({ value: null });
  });
});

describe('formatValue', () => {
  it('shows each kind of value in the General format', () => {
    expect(formatValue(3)).toBe('3');
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
    expect(formatValue('abc')).toBe('abc');
    expect(formatValue(true)).toBe('TRUE');
    expect(formatValue(false)).toBe('FALSE');
    expect(formatValue(null)).toBe('');
    expect(formatValue({ error: '#VALUE!' })).toBe('#VALUE!');
  });

  it('shows numbers in a format code', () => {
    expect(formatValue(46024, 'm/d/yyyy')).toBe('1/2/2026');
    expect(formatValue(0.12, '0%')).toBe('12%');
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(0.5625, 'h:mm AM/PM')).toBe('1:30 PM');
  });

  it('shows text and booleans as they are, whatever the format', () => {
    expect(formatValue('abc', '0%')).toBe('abc');
    expect(formatValue(true, 'm/d/yyyy')).toBe('TRUE');
  });
});

describe('editText', () => {
  it('gives the text that edits a value back to itself, as Excel does', () => {
    expect(editText(1234.5, '#,##0.00')).toBe('1234.5');
    expect(editText(0.1 + 0.2)).toBe('0.3');
    expect(editText(0.12, '0%')).toBe('12%');
    expect(editText(46024, 'd-mmm')).toBe('1/2/2026');
    expect(editText(0.5625, 'h:mm')).toBe('1:30:00 PM');
    expect(editText(46024.5625, 'm/d/yyyy h:mm')).toBe('1/2/2026 1:30:00 PM');
    expect(editText(true)).toBe('TRUE');
    expect(editText('abc', '0%')).toBe('abc');
    expect(editText(null)).toBe('');
  });

  it.each(['1,234.5', '12%', '1/2/2026', '13:45', '1/2/2026 13:45', 'TRUE'])(
    'reads %j back as the same value',
    (text) => {
      const { value, format } = parseInput(text);
      expect(parseInput(editText(value, format)).value).toBe(value);
    },
  );
});

describe('formatValue with custom codes', () => {
  const code = '[Red]0.0;[Blue](0.0);"zero";"Text: "@';
  it.each([
    [2, '2.0', '#ff0000'],
    [-3, '(3.0)', '#0000ff'],
    [0, 'zero', undefined],
    ['ab', 'Text: ab', undefined],
  ])('shows %j by its section, with its color', (value, shown, color) => {
    expect(formatValue(value, code)).toBe(shown);
    expect(formatColor(value, code)).toBe(color);
  });

  it('gives no color to a format without one', () => {
    expect(formatColor(5, '0.00')).toBeUndefined();
    expect(formatColor(null, '[Red]0')).toBeUndefined();
  });

  it('tells a valid code from one that is not', () => {
    expect(validFormat('#,##0.00;[Red]-#,##0.00')).toBe(true);
    expect(validFormat('0.0"')).toBe(false);
    expect(validFormat('')).toBe(false);
  });
});

describe('changedFormat', () => {
  it.each([
    ['0', 'moreDecimals', '0.0'],
    ['0.00', 'fewerDecimals', '0.0'],
    ['0.0', 'fewerDecimals', '0'],
    ['0', 'fewerDecimals', '0'],
    ['0%', 'moreDecimals', '0.0%'],
    ['0.00E+00', 'moreDecimals', '0.000E+00'],
    [
      '[Red]#,##0.00_);[Blue](#,##0.00);"zero"',
      'moreDecimals',
      '[Red]#,##0.000_);[Blue](#,##0.000);"zero"',
    ],
    ['"$"#,##0.00', 'thousands', '"$"0.00'],
    ['0.00', 'thousands', '#,##0.00'],
    ['0.00;(0.00)', 'thousands', '#,##0.00;(#,##0.00)'],
    ['m/d/yyyy', 'moreDecimals', 'm/d/yyyy'],
    ['@', 'thousands', '@'],
    ['# ?/?', 'moreDecimals', '# ?/?'],
  ] as const)('changes %j by %s into %j', (code, change, changed) => {
    expect(changedFormat(code, 1, change)).toBe(changed);
  });

  it('tells whether a code separates thousands', () => {
    expect(isGroupedFormat('"$"#,##0.00')).toBe(true);
    expect(isGroupedFormat('0.00;(#,##0.00)')).toBe(false);
    expect(isGroupedFormat('0\\,000')).toBe(false);
  });

  it('starts a General cell from the decimals its value shows, as Excel does', () => {
    expect(changedFormat('General', 1.5, 'moreDecimals')).toBe('0.00');
    expect(changedFormat('General', 1.5, 'fewerDecimals')).toBe('0');
    expect(changedFormat('General', 1234.5, 'thousands')).toBe('#,##0.0');
    expect(changedFormat('General', 'text', 'moreDecimals')).toBe('0.0');
  });
});
