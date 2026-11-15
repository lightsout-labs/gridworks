import type { ClipboardContents } from '../platform';

/**
 * The texts of a block of cells, by row, as other applications paste them:
 * tab-separated text, and an HTML table, which Excel and Google Sheets
 * paste as a grid.
 */
export function toClipboard(
  rows: readonly (readonly string[])[],
): ClipboardContents {
  // Every row ends with a line break, as in Excel, so that a last row of
  // one empty cell is kept.
  const text = rows.map((row) => row.map(quoted).join('\t') + '\n').join('');
  const html =
    '<table>' +
    rows
      .map(
        (row) =>
          `<tr>${row.map((cell) => `<td>${escaped(cell)}</td>`).join('')}</tr>`,
      )
      .join('') +
    '</table>';
  return { text, html };
}

/**
 * The texts of the cells in text from the clipboard, by row: tab-separated
 * text as Excel and Google Sheets write it, with cells in quotes where
 * they hold tabs, line breaks or quotes, and plain text one line per row.
 * A line break at the end ends the last row.
 */
export function fromClipboard(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let at = 0;
  for (;;) {
    let cell = '';
    // A quote that is never closed starts plain text.
    if (text[at] === '"' && text.indexOf('"', at + 1) >= 0) {
      // A quoted cell, where "" is a quote, then anything up to the next
      // tab or line break.
      at++;
      for (;;) {
        const quote = text.indexOf('"', at);
        if (quote < 0) {
          cell += text.slice(at);
          at = text.length;
          break;
        }
        cell += text.slice(at, quote);
        at = quote + 1;
        if (text[at] !== '"') break;
        cell += '"';
        at++;
      }
    }
    const end = at + text.slice(at).search(/[\t\r\n]|$/);
    cell += text.slice(at, end);
    row.push(cell);
    at = end;
    if (text[at] === '\t') {
      at++;
      continue;
    }
    rows.push(row);
    row = [];
    at += text.startsWith('\r\n', at) ? 2 : 1;
    if (at >= text.length) return rows;
  }
}

function quoted(cell: string): string {
  return /[\t\r\n]|^"/.test(cell) ? `"${cell.replaceAll('"', '""')}"` : cell;
}

function escaped(cell: string): string {
  return cell
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}
