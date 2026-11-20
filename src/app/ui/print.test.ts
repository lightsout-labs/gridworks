import { describe, expect, it } from 'vitest';
import { createWorkbook, type Range } from '../../engine/workbook';
import {
  DEFAULT_PRINT,
  printDocument,
  printedRanges,
  type PrintSettings,
} from './print';

const measure = (text: string) => text.length * 7;

function setup(...inputs: [number, number, string][]) {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  for (const [row, col, input] of inputs) {
    workbook.execute({ type: 'setCell', sheet, row, col, input });
  }
  const print = (
    settings: Partial<PrintSettings> = {},
    selection: readonly Range[] = [],
  ) => {
    const chosen = { ...DEFAULT_PRINT, ...settings };
    const ranges = printedRanges(workbook, sheet, selection, chosen.what);
    return printDocument(workbook, sheet, ranges, chosen, measure);
  };
  return { workbook, sheet, print };
}

/** The cells of the tables of a document: their attributes and text. */
function cells(html: string): { attributes: string; text: string }[] {
  return [...html.matchAll(/<td([^>]*)>([^<]*)<\/td>/g)].map((match) => ({
    attributes: match[1]!,
    text: match[2]!,
  }));
}

const texts = (html: string) => cells(html).map((cell) => cell.text);

/** The attributes of the cell whose text is `text`. */
function cellWith(html: string, text: string): string {
  const cell = cells(html).find((each) => each.text === text);
  if (!cell) throw new Error(`No cell shows ${text}`);
  return cell.attributes;
}

describe('the printed range', () => {
  it('is the cells in use of the sheet, from A1, grown over merges', () => {
    const { workbook, sheet } = setup([1, 2, 'x']);
    workbook.execute({
      type: 'merge',
      sheet,
      ranges: [{ top: 1, left: 2, bottom: 3, right: 3 }],
    });
    expect(printedRanges(workbook, sheet, [], 'sheet')).toEqual([
      { top: 0, left: 0, bottom: 3, right: 3 },
    ]);
  });

  it('is each range selected, within the cells in use', () => {
    const { workbook, sheet } = setup([0, 0, 'a'], [9, 3, 'b']);
    const column = { top: 0, left: 1, bottom: 1_048_575, right: 1 };
    const block = { top: 2, left: 2, bottom: 3, right: 3 };
    expect(
      printedRanges(workbook, sheet, [column, block], 'selection'),
    ).toEqual([{ top: 0, left: 1, bottom: 9, right: 1 }, block]);
  });

  it('is nothing on an empty sheet or outside the cells in use', () => {
    const { workbook, sheet } = setup();
    expect(printedRanges(workbook, sheet, [], 'sheet')).toEqual([]);
    workbook.execute({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: 'a',
    });
    const outside = { top: 5, left: 5, bottom: 6, right: 6 };
    expect(printedRanges(workbook, sheet, [outside], 'selection')).toEqual([]);
  });
});

describe('the printable document', () => {
  it('shows the cells as the grid does, with their formats and widths', () => {
    const { workbook, sheet, print } = setup(
      [0, 0, 'Item'],
      [0, 1, '1234.5'],
      [1, 0, '<b>'],
      [1, 1, '=B1*2'],
    );
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [{ top: 0, left: 1, bottom: 1, right: 1 }],
      style: { numberFormat: '#,##0.00', bold: true, fill: '#ffff00' },
    });
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 0, end: 0, size: 150 }],
    });
    const html = print();
    expect(texts(html)).toEqual(['Item', '1,234.50', '&lt;b&gt;', '2,469.00']);
    expect(html).toContain('<title>Book1</title>');
    expect(html).toMatch(
      /<col style="width:150px">\s*<col style="width:96px">/,
    );
    const number = cellWith(html, '1,234.50');
    expect(number).toContain('text-align:right');
    expect(number).toContain('background-color:#ffff00');
    expect(number).toMatch(/font:bold /);
    expect(cellWith(html, 'Item')).toContain('text-align:left');
  });

  it('prints only the selection when it is chosen', () => {
    const { print } = setup([0, 0, 'a'], [0, 1, 'b'], [1, 0, 'c'], [1, 1, 'd']);
    const selection = [{ top: 1, left: 1, bottom: 1, right: 1 }];
    expect(texts(print({ what: 'selection' }, selection))).toEqual(['d']);
  });

  it('starts each range selected on a page of its own', () => {
    const { print } = setup([0, 0, 'a'], [5, 5, 'b']);
    const html = print({ what: 'selection' }, [
      { top: 0, left: 0, bottom: 0, right: 0 },
      { top: 5, left: 5, bottom: 5, right: 5 },
    ]);
    expect(html.match(/<table/g)).toHaveLength(2);
    expect(html).toContain('table + table { break-before: page; }');
  });

  it('gives the page its paper size, orientation and margins', () => {
    const { print } = setup([0, 0, 'a']);
    expect(print()).toContain(
      '@page { size: letter portrait; margin: 0.75in 0.7in 0.75in 0.7in; }',
    );
    expect(
      print({ paper: 'a4', orientation: 'landscape', margins: 'narrow' }),
    ).toContain(
      '@page { size: a4 landscape; margin: 0.75in 0.25in 0.75in 0.25in; }',
    );
    expect(print({ paper: 'legal', margins: 'wide' })).toContain(
      '@page { size: legal portrait; margin: 1in 1in 1in 1in; }',
    );
  });

  it('never breaks a page through a row, and repeats the headers', () => {
    const { print } = setup([0, 0, 'a']);
    expect(print()).toContain('tr { break-inside: avoid; }');
    expect(print({ headers: true })).toMatch(
      /<thead>.*<th>A<\/th>.*<\/thead>/s,
    );
  });

  it('shrinks a table wider than the page to its width', () => {
    // Ten columns of 96 pixels on a Letter page in portrait, whose width
    // inside the normal margins is 7.1 inches of 96 pixels.
    const { print } = setup([0, 9, 'a']);
    expect(print()).not.toContain('zoom');
    expect(print({ fitWidth: true })).toContain('zoom:0.71;');
    expect(print({ fitWidth: true, orientation: 'landscape' })).toContain(
      'zoom:0.96;',
    );
    // A narrow table is not enlarged.
    expect(setup([0, 0, 'a']).print({ fitWidth: true })).not.toContain('zoom');
  });

  it('prints gridlines and the row and column headers when chosen', () => {
    const { print } = setup([0, 1, 'a'], [1, 0, 'b']);
    const plain = print();
    expect(plain).not.toContain('<th');
    expect(cellWith(plain, 'a')).not.toContain('border');
    const full = print({ gridlines: true, headers: true });
    expect([...full.matchAll(/<th>([^<]*)<\/th>/g)].map((m) => m[1])).toEqual([
      '',
      'A',
      'B',
      '1',
      '2',
    ]);
    expect(cellWith(full, 'a')).toContain('border-top:1px solid #bfbfbf');
  });

  it('leaves out hidden rows and columns, as those a filter hides', () => {
    const { workbook, sheet, print } = setup(
      [0, 0, 'a'],
      [0, 1, 'b'],
      [1, 0, 'c'],
      [1, 1, 'd'],
    );
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 0, end: 0, hidden: true }],
    });
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 1, end: 1, hidden: true }],
    });
    const html = print({ headers: true });
    expect(texts(html)).toEqual(['c']);
    expect(html.match(/<col /g)).toHaveLength(2);
    expect(html).toContain('<th>2</th>');
    expect(html).not.toContain('<th>1</th>');
  });

  it('spans merged cells over their rows and columns', () => {
    const { workbook, sheet, print } = setup([0, 0, 'Title'], [2, 0, 'x']);
    workbook.execute({
      type: 'merge',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 1, right: 2 }],
    });
    const html = print();
    expect(cellWith(html, 'Title')).toContain('rowspan="2" colspan="3"');
    // The merged cell's row, the empty row it covers, and the third row.
    expect(cells(html)).toHaveLength(1 + 3);
  });

  it('draws borders on both cells that share them, over the gridlines', () => {
    const { workbook, sheet, print } = setup([0, 0, 'a'], [0, 1, 'b']);
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
      sides: 'right',
      border: { line: 'double', color: '#ff0000' },
    });
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [{ top: 0, left: 1, bottom: 0, right: 1 }],
      sides: 'bottom',
      border: { line: 'thick' },
    });
    const html = print({ gridlines: true });
    expect(cellWith(html, 'a')).toContain('border-right:3px double #ff0000');
    expect(cellWith(html, 'b')).toContain('border-left:3px double #ff0000');
    expect(cellWith(html, 'b')).toContain('border-bottom:3px solid #000000');
  });

  it('runs long text over empty cells only, and shows #### for a wide number', () => {
    const { workbook, sheet, print } = setup(
      [0, 0, 'a text longer than its cell'],
      [1, 0, 'another text longer than a cell'],
      [1, 1, 'x'],
      [2, 0, '123456789012345'],
    );
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [{ top: 2, left: 0, bottom: 2, right: 0 }],
      style: { numberFormat: '0' },
    });
    const html = print();
    expect(cellWith(html, 'a text longer than its cell')).toContain(
      'overflow:visible',
    );
    expect(cellWith(html, 'another text longer than a cell')).not.toContain(
      'overflow:visible',
    );
    expect(texts(html)).toContain('############');
  });
});
