import { describe, expect, it } from 'vitest';
import type { Style, Value } from '../../engine/workbook';
import { COLUMN_WIDTH } from './geometry';
import {
  cssFont,
  DEFAULT_FAMILY,
  rowTexts,
  shownCell,
  textColorOn,
} from './text';

/** A row of cells from values, each shown as its text, and their styles. */
function row(
  values: Record<number, Value>,
  styles: Record<number, Style> = {},
) {
  return (col: number) => {
    const value = values[col] ?? null;
    const display =
      value === null
        ? ''
        : typeof value === 'object'
          ? value.error
          : String(value).toUpperCase();
    return { value, display, style: styles[col] ?? {} };
  };
}

// Each character is a tenth of a column wide.
const measure = (text: string) => (text.length * COLUMN_WIDTH) / 10;
const standard = () => COLUMN_WIDTH;

describe('rowTexts', () => {
  it('leaves out the cells of merged cells, and runs no text over them', () => {
    const long = 'x'.repeat(25);
    const merged = (col: number) => col === 3 || col === 4;
    expect(
      rowTexts(row({ 1: long, 3: 'a' }), 2, 9, measure, standard, merged),
    ).toEqual([
      {
        col: 1,
        lines: [long.toUpperCase()],
        style: {},
        align: 'left',
        from: 1,
        to: 2,
      },
    ]);
  });

  it('aligns numbers right, text left, and booleans and errors in the middle', () => {
    const texts = rowTexts(
      row({ 0: 1, 1: 'a', 2: true, 3: { error: '#VALUE!' } }),
      0,
      5,
      measure,
      standard,
    );
    expect(texts).toEqual([
      { col: 0, lines: ['1'], style: {}, align: 'right', from: 0, to: 0 },
      { col: 1, lines: ['A'], style: {}, align: 'left', from: 1, to: 1 },
      { col: 2, lines: ['TRUE'], style: {}, align: 'center', from: 2, to: 2 },
      {
        col: 3,
        lines: ['#VALUE!'],
        style: {},
        align: 'center',
        from: 3,
        to: 3,
      },
    ]);
  });

  it('lets long text run over the empty cells to its right, as far as it needs', () => {
    const long = 'x'.repeat(15); // One and a half columns.
    expect(rowTexts(row({ 1: long }), 0, 9, measure, standard)).toEqual([
      {
        col: 1,
        lines: [long.toUpperCase()],
        style: {},
        align: 'left',
        from: 1,
        to: 2,
      },
    ]);
  });

  it('stops long text at the next cell that is not empty', () => {
    const long = 'x'.repeat(35);
    const texts = rowTexts(row({ 0: long, 2: 5 }), 0, 9, measure, standard);
    expect(texts[0]).toMatchObject({ col: 0, from: 0, to: 1 });
  });

  it('does not let numbers run over', () => {
    const texts = rowTexts(
      row({ 0: 1234567890123456 }),
      0,
      9,
      measure,
      standard,
    );
    expect(texts[0]).toMatchObject({ from: 0, to: 0 });
  });

  it('shows a number too wide for its column as #### across it, as Excel does', () => {
    // A column holds 9 characters inside its padding.
    const texts = rowTexts(
      row(
        { 0: 1234567890, 1: 123456789, 2: 'x'.repeat(10) },
        { 0: { wrap: true } },
      ),
      0,
      9,
      measure,
      standard,
    );
    expect(texts.map(({ lines }) => lines)).toEqual([
      ['#########'],
      ['123456789'],
      ['XXXXXXXXXX'],
    ]);
  });

  it('gives text the color its number format gives it', () => {
    const cell = (col: number) => ({
      value: col === 0 ? -1 : null,
      display: col === 0 ? '(1)' : '',
      style: {},
      color: '#ff0000',
    });
    expect(rowTexts(cell, 0, 3, measure, standard)[0]).toMatchObject({
      lines: ['(1)'],
      color: '#ff0000',
    });
  });

  it('includes long text from the left of the visible columns that reaches them', () => {
    const long = 'x'.repeat(35); // Three and a half columns.
    expect(rowTexts(row({ 1: long }), 3, 9, measure, standard)).toEqual([
      {
        col: 1,
        lines: [long.toUpperCase()],
        style: {},
        align: 'left',
        from: 1,
        to: 4,
      },
    ]);
    expect(rowTexts(row({ 1: long }), 5, 9, measure, standard)).toEqual([]);
    expect(rowTexts(row({ 0: long, 1: 'b' }), 3, 9, measure, standard)).toEqual(
      [],
    );
  });

  it('runs text over columns by their widths, and hides the text of hidden ones', () => {
    const long = 'x'.repeat(15);
    const wide = (col: number) => (col === 1 ? 2 * COLUMN_WIDTH : COLUMN_WIDTH);
    expect(rowTexts(row({ 1: long }), 0, 9, measure, wide)).toEqual([
      {
        col: 1,
        lines: [long.toUpperCase()],
        style: {},
        align: 'left',
        from: 1,
        to: 1,
      },
    ]);
    const hidden = (col: number) => (col === 2 ? 0 : COLUMN_WIDTH);
    expect(rowTexts(row({ 1: long, 2: 'b' }), 0, 9, measure, hidden)).toEqual([
      {
        col: 1,
        lines: [long.toUpperCase()],
        style: {},
        align: 'left',
        from: 1,
        to: 1,
      },
    ]);
  });

  it('aligns text as its style says, and lets only text aligned left run over', () => {
    const long = 'x'.repeat(15);
    const texts = rowTexts(
      row(
        { 0: 1, 1: long, 3: long },
        { 0: { horizontal: 'left' }, 1: { horizontal: 'center' } },
      ),
      0,
      9,
      measure,
      standard,
    );
    expect(texts.map(({ align, from, to }) => [align, from, to])).toEqual([
      ['left', 0, 0],
      ['center', 1, 1],
      ['left', 3, 4],
    ]);
  });

  it('wraps text onto lines that fit its column, which then does not run over', () => {
    const texts = rowTexts(
      row({ 0: 'aaaaa bbbbb cccc' }, { 0: { wrap: true } }),
      0,
      9,
      measure,
      standard,
    );
    expect(texts).toEqual([
      {
        col: 0,
        lines: ['AAAAA', 'BBBBB', 'CCCC'],
        style: { wrap: true },
        align: 'left',
        from: 0,
        to: 0,
      },
    ]);
  });

  it('measures text in the font of its cell', () => {
    const bold = (text: string, style: Style) =>
      measure(text) * (style.bold ? 2 : 1);
    const texts = rowTexts(
      row({ 0: 'x'.repeat(8) }, { 0: { bold: true } }),
      0,
      9,
      bold,
      standard,
    );
    expect(texts[0]).toMatchObject({ from: 0, to: 1 });
  });
});

describe('cssFont', () => {
  it('gives the font of a style at a zoom, with fallbacks', () => {
    expect(cssFont({}, 1)).toBe(`13.333px ${DEFAULT_FAMILY}`);
    expect(
      cssFont({ font: 'Courier New', size: 12, bold: true, italic: true }, 2),
    ).toBe('italic bold 32px "Courier New", "Liberation Mono", monospace');
  });
});

describe('textColorOn', () => {
  it('gives black text on light fills and white text on dark ones', () => {
    expect(textColorOn('#ffff00')).toBe('#000000');
    expect(textColorOn('#ffffff')).toBe('#000000');
    expect(textColorOn('#002060')).toBe('#ffffff');
    expect(textColorOn('#c00000')).toBe('#ffffff');
  });
});

describe('shownCell', () => {
  it('draws a conditional format over the style and the color of the number format', () => {
    const cell = {
      value: -1,
      display: '(1)',
      style: { bold: true, fill: '#ffffff', color: '#000000' } as const,
      color: '#ff0000',
    };
    expect(shownCell(cell, '#0563c1')).toBe(cell);
    expect(
      shownCell(
        { ...cell, conditional: { fill: '#ffc7ce', italic: true } },
        '#0563c1',
      ),
    ).toEqual({
      value: -1,
      display: '(1)',
      style: { bold: true, italic: true, fill: '#ffc7ce', color: '#000000' },
      color: '#ff0000',
    });
    expect(
      shownCell({ ...cell, conditional: { color: '#9c0006' } }, '#0563c1'),
    ).toMatchObject({ style: { color: '#9c0006' }, color: '#9c0006' });
  });
});

describe('shownCell with a link', () => {
  it('underlines a link, in the link color unless the cell has a color of its own', () => {
    const cell = { value: 'Site', display: 'Site', style: {} };
    expect(shownCell({ ...cell, link: '#A1' }, '#0563c1')).toEqual({
      ...cell,
      style: { underline: true, color: '#0563c1' },
    });
    expect(
      shownCell(
        { ...cell, style: { color: '#ff0000' }, link: '#A1' },
        '#0563c1',
      ).style,
    ).toEqual({ color: '#ff0000', underline: true });
    expect(shownCell(cell, '#0563c1')).toBe(cell);
  });
});
