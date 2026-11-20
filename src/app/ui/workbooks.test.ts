import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../../engine/workbook';
import {
  hashOf,
  modifiedText,
  numberedName,
  renamedSnapshot,
  workbookIdOf,
} from './workbooks';

describe('the address of a workbook', () => {
  it('names the workbook shown after #, and the workbook dialog alone without one', () => {
    expect(workbookIdOf('')).toBe(null);
    expect(workbookIdOf('#')).toBe(null);
    expect(workbookIdOf(hashOf('a b/c'))).toBe('a b/c');
    expect(workbookIdOf('#default')).toBe('default');
    // An address that does not decode names no workbook.
    expect(workbookIdOf('#%')).toBe(null);
  });
});

describe('numberedName', () => {
  it('gives the first number not taken, in any case', () => {
    expect(numberedName('Book', [])).toBe('Book1');
    expect(numberedName('Book', ['book1', 'Book3', 'Budget'])).toBe('Book2');
  });
});

describe('modifiedText', () => {
  const now = new Date(2026, 10, 18, 15, 30).getTime();

  it('gives the time alone today, and the date and time before', () => {
    expect(
      modifiedText(new Date(2026, 10, 18, 9, 5).getTime(), now, 'en'),
    ).toBe('9:05 AM');
    expect(modifiedText(new Date(2025, 0, 2, 14, 0).getTime(), now, 'en')).toBe(
      'Jan 2, 2025, 2:00 PM',
    );
    expect(modifiedText(new Date(2025, 0, 2, 14, 0).getTime(), now, 'zh')).toBe(
      '2025年1月2日 14:00',
    );
  });
});

describe('renamedSnapshot', () => {
  it('gives the stored workbook, its changes replayed, under a new name', () => {
    const workbook = createWorkbook();
    const sheet = workbook.sheets[0]!.id;
    const snapshot = renamedSnapshot(
      {
        snapshot: workbook.toSnapshot(),
        changes: [{ type: 'setCell', sheet, row: 0, col: 0, input: '5' }],
      },
      'Copy',
    );
    expect(snapshot.name).toBe('Copy');
    expect(snapshot.sheets[0]!.columns[0]).toEqual([5]);
  });

  it('refuses a workbook no longer stored', () => {
    expect(() => renamedSnapshot(null, 'Copy')).toThrow();
  });
});
