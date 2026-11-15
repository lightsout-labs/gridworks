import { describe, expect, it } from 'vitest';
import { fromClipboard, toClipboard } from './clipboard';

describe('toClipboard', () => {
  it('gives tab-separated text and an HTML table of the cells', () => {
    expect(
      toClipboard([
        ['12%', '<b> & co'],
        ['', '3'],
      ]),
    ).toEqual({
      text: '12%\t<b> & co\n\t3\n',
      html:
        '<table><tr><td>12%</td><td>&lt;b&gt; &amp; co</td></tr>' +
        '<tr><td></td><td>3</td></tr></table>',
    });
  });

  it('quotes text with tabs, line breaks or a leading quote, as Excel does', () => {
    expect(toClipboard([['a\tb', 'two\nlines', '"q"', 'x"y']]).text).toBe(
      '"a\tb"\t"two\nlines"\t"""q"""\tx"y\n',
    );
  });
});

describe('fromClipboard', () => {
  it('reads tab-separated text as rows of cells, as Excel and Google Sheets write it', () => {
    expect(fromClipboard('12%\t1,234.5\r\n\t3\r\n')).toEqual([
      ['12%', '1,234.5'],
      ['', '3'],
    ]);
  });

  it('reads plain text as one line per row', () => {
    expect(fromClipboard('first\nsecond\rthird\n\nfifth')).toEqual([
      ['first'],
      ['second'],
      ['third'],
      [''],
      ['fifth'],
    ]);
  });

  it('reads quoted cells, with their tabs, line breaks and quotes', () => {
    expect(fromClipboard('"a\tb"\t"two\r\nlines"\n"""q"""\tx"y')).toEqual([
      ['a\tb', 'two\r\nlines'],
      ['"q"', 'x"y'],
    ]);
  });

  it('reads a quote that is never closed as plain text', () => {
    expect(fromClipboard('"abc\ndef')).toEqual([['"abc'], ['def']]);
  });

  it('reads back what it writes', () => {
    const cells = [['a\tb', '"', ''], ['two\nlines', 'x', 'y'], ['']];
    expect(fromClipboard(toClipboard(cells).text)).toEqual(cells);
  });
});
