import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../formula';
import { isError } from '../value';
import { Workbook, createWorkbook } from '../workbook';
import { readCsv, writeCsv } from '.';

/** The workbook a CSV text opens as, and what was dropped. */
function open(text: string, name = 'Prices') {
  const { snapshot, dropped } = readCsv(text, name);
  const workbook = Workbook.fromSnapshot(snapshot);
  return { workbook, sheet: workbook.sheets[0]!.id, dropped };
}

/**
 * The values of the cells of a sheet in use, row by row, an error as its
 * text, as Gridworks has no error constants and reads an error as text.
 */
function values(workbook: Workbook, sheet: string) {
  const { rows, cols } = workbook.extent(sheet);
  return Array.from({ length: rows }, (_, row) =>
    Array.from({ length: cols }, (_, col) => {
      const { value } = workbook.cell(sheet, row, col);
      return isError(value) ? value.error : value;
    }),
  );
}

describe('readCsv', () => {
  it('opens as a workbook of one sheet named after the file', () => {
    const { workbook } = open('a,b\r\n', 'Prices 2026');
    expect(workbook.toSnapshot().name).toBe('Prices 2026');
    expect(workbook.sheets.map(({ name }) => name)).toEqual(['Prices 2026']);
  });

  it('names the sheet Sheet1 when the file name cannot name a sheet', () => {
    const { workbook } = open('a\n', 'Q1 [draft]');
    expect(workbook.sheets.map(({ name }) => name)).toEqual(['Sheet1']);
  });

  it('recognizes numbers, dates, times, percentages and booleans as typed, and keeps errors as text', () => {
    const { workbook, sheet } = open(
      'Item,Price,Sold,Share,Paid,Time,Check\n' +
        'Tea,"1,234.5",2026-01-15,50%,TRUE,13:45:30,#N/A\n',
    );
    expect(values(workbook, sheet)[1]).toEqual([
      'Tea',
      1234.5,
      46037,
      0.5,
      true,
      0.5732638888888889,
      '#N/A',
    ]);
    const display = (col: number) => workbook.cell(sheet, 1, col).display;
    expect(display(1)).toBe('1,234.50');
    expect(display(2)).toBe('1/15/2026');
    expect(display(3)).toBe('50%');
    expect(display(5)).toBe('13:45:30');
  });

  it('keeps a field that starts with = as text, not a formula', () => {
    const { workbook, sheet } = open('=1+2\n');
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: '=1+2',
      formula: null,
    });
  });

  it('reads quoted fields with commas, line breaks and doubled quotes', () => {
    const { workbook, sheet } = open(
      'name,note\r\n' +
        '"Smith, John","said ""hi""\r\nthen left"\r\n' +
        '"",plain "quote" inside\r\n',
    );
    expect(values(workbook, sheet)).toEqual([
      ['name', 'note'],
      ['Smith, John', 'said "hi"\r\nthen left'],
      [null, 'plain "quote" inside'],
    ]);
  });

  it('reads lines ended by LF, CRLF or CR, with empty lines and fields kept in place', () => {
    const { workbook, sheet } = open('a,,c\n\nd\re,f');
    expect(values(workbook, sheet)).toEqual([
      ['a', null, 'c'],
      [null, null, null],
      ['d', null, null],
      ['e', 'f', null],
    ]);
  });

  it('reads UTF-8 text with or without a byte order mark', () => {
    const text = 'Café,数据,😀\n';
    for (const csv of [text, `\uFEFF${text}`]) {
      const { workbook, sheet } = open(csv);
      expect(values(workbook, sheet)).toEqual([['Café', '数据', '😀']]);
    }
  });

  it('leaves out cells beyond the rows and columns of a sheet and counts them', () => {
    const wide = `${'x,'.repeat(SHEET_COLUMNS)}x\n`;
    const { workbook, sheet, dropped } = open(wide + '1\n'.repeat(SHEET_ROWS));
    expect(workbook.extent(sheet)).toEqual({
      rows: SHEET_ROWS,
      cols: SHEET_COLUMNS,
    });
    expect(dropped).toEqual([{ feature: 'cell', count: 2 }]);
  });

  it('opens an empty file as an empty sheet', () => {
    const { workbook, sheet, dropped } = open('');
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
    expect(dropped).toEqual([]);
  });
});

describe('writeCsv', () => {
  /** A workbook whose first sheet has the given inputs, by row. */
  function sheetOf(rows: string[][]) {
    const workbook = createWorkbook();
    const sheet = workbook.sheets[0]!.id;
    rows.forEach((cells, row) =>
      cells.forEach((input, col) =>
        workbook.execute({ type: 'setCell', sheet, row, col, input }),
      ),
    );
    return { workbook, sheet };
  }

  it('writes the cells in use with a byte order mark and CRLF line ends', () => {
    const { workbook, sheet } = sheetOf([
      ['a', '', 'c'],
      ['', '2'],
    ]);
    expect(writeCsv(workbook, sheet)).toBe('\uFEFFa,,c\r\n,2,\r\n');
  });

  it('quotes fields with commas, quotes or line breaks, doubling the quotes', () => {
    const { workbook, sheet } = sheetOf([
      ['Smith, John', 'say "hi"', 'two\nlines', 'plain'],
    ]);
    expect(writeCsv(workbook, sheet)).toBe(
      '\uFEFF"Smith, John","say ""hi""","two\nlines",plain\r\n',
    );
  });

  it('writes formulas as their results, and values in full whatever their format', () => {
    const { workbook, sheet } = sheetOf([
      ['1234.5678', '=A1*2', '=1/0', '=A1>1', '12%', '$1,234.50'],
    ]);
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
      style: { numberFormat: '#,##0' },
    });
    expect(writeCsv(workbook, sheet)).toBe(
      '\uFEFF1234.5678,2469.1356,#DIV/0!,TRUE,0.12,1234.5\r\n',
    );
  });

  it('writes dates and times as text that reads back as them', () => {
    const { workbook, sheet } = sheetOf([
      ['2026-01-15', '13:45:30', '1/15/2026 13:45'],
    ]);
    expect(writeCsv(workbook, sheet)).toBe(
      '\uFEFF2026-01-15,13:45:30,2026-01-15 13:45:00\r\n',
    );
  });

  it('gives a sheet that reopens with the same values', () => {
    const { workbook, sheet } = sheetOf([
      ['Item', 'Price', 'Sold', 'Note'],
      ['Tea, green', '3.75', '2026-03-01 08:30', 'said "fresh"\nand cheap'],
      ['', '=B2*4', 'FALSE', '0.1'],
      ['Café 数据', '1e-7', '-42', '=NA()'],
    ]);
    const reopened = open(writeCsv(workbook, sheet));
    expect(values(reopened.workbook, reopened.sheet)).toEqual(
      values(workbook, sheet),
    );
  });
});
