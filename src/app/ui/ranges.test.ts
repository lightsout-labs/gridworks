import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../../engine/workbook';
import { parseRanges, rangesText } from './ranges';

describe('the ranges a rule applies to', () => {
  it('reads cells, ranges, whole columns and rows, separated by commas', () => {
    expect(parseRanges(' b2:A1, C3 ,D:E,2:3')).toEqual([
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 2, left: 2, bottom: 2, right: 2 },
      { top: 0, left: 3, bottom: SHEET_ROWS - 1, right: 4 },
      { top: 1, left: 0, bottom: 2, right: SHEET_COLUMNS - 1 },
    ]);
  });

  it('refuses anything else', () => {
    for (const text of ['', 'A1,', 'Sheet2!A1', 'A1+1', 'SUM(A1)', 'x']) {
      expect(parseRanges(text)).toBeNull();
    }
  });

  it('writes ranges back as they are read', () => {
    const text = 'A1:B2,C3,D:E,2:3';
    expect(rangesText(parseRanges(text)!)).toBe(text);
  });
});
