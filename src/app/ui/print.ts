import {
  columnName,
  withMerges,
  type Border,
  type Range,
  type Style,
  type TextWidth,
  type Workbook,
} from '../../engine/workbook';
import {
  COLUMN_HEADER_HEIGHT,
  layoutOf,
  ROW_HEADER_WIDTH,
} from '../grid/geometry';
import {
  cellText,
  cssFont,
  DEFAULT_FAMILY,
  rowTexts,
  shownCell,
  textColorOn,
  type ShownCell,
} from '../grid/text';

export const PAPERS = ['letter', 'legal', 'a4', 'a3'] as const;
export const MARGINS = ['normal', 'wide', 'narrow'] as const;

/** How the user chose to print (R36). */
export interface PrintSettings {
  /** The cells in use of the active sheet, or the ranges selected. */
  readonly what: 'sheet' | 'selection';
  readonly orientation: 'portrait' | 'landscape';
  readonly paper: (typeof PAPERS)[number];
  readonly margins: (typeof MARGINS)[number];
  /** Whether a table wider than the page shrinks to its width. */
  readonly fitWidth: boolean;
  readonly gridlines: boolean;
  /** Whether the row numbers and column letters print. */
  readonly headers: boolean;
}

/** Excel's defaults, with Letter paper whatever the language. */
export const DEFAULT_PRINT: PrintSettings = {
  what: 'sheet',
  orientation: 'portrait',
  paper: 'letter',
  margins: 'normal',
  fitWidth: false,
  gridlines: false,
  headers: false,
};

/** The width and height of each paper in portrait, in inches. */
const PAPER_SIZES: Record<PrintSettings['paper'], [number, number]> = {
  letter: [8.5, 11],
  legal: [8.5, 14],
  a4: [8.27, 11.69],
  a3: [11.69, 16.54],
};

/** Excel's margins, top and bottom then left and right, in inches. */
const MARGIN_SIZES: Record<PrintSettings['margins'], [number, number]> = {
  normal: [0.75, 0.7],
  wide: [1, 1],
  narrow: [0.75, 0.25],
};

const CSS_INCH = 96;
const GRIDLINE = '1px solid #bfbfbf';
/** Excel's color of a link, as paper is white whatever the theme. */
const LINK_COLOR = '#0563c1';

/**
 * The ranges that print: the cells in use of the sheet from A1, or each
 * range selected within them; none when there are no cells to print.
 */
export function printedRanges(
  workbook: Workbook,
  sheet: string,
  selection: readonly Range[],
  what: PrintSettings['what'],
): Range[] {
  const { rows, cols } = workbook.extent(sheet);
  if (rows === 0) return [];
  const used = { top: 0, left: 0, bottom: rows - 1, right: cols - 1 };
  if (what === 'sheet') return [withMerges(used, workbook.merges(sheet))];
  return selection
    .map((range) => ({
      top: range.top,
      left: range.left,
      bottom: Math.min(range.bottom, used.bottom),
      right: Math.min(range.right, used.right),
    }))
    .filter((range) => range.top <= range.bottom && range.left <= range.right);
}

/**
 * An HTML document that prints the ranges of a sheet as tables, a page
 * each, in the page settings chosen: the cells as the grid shows them,
 * with their formats, borders, merges and sizes, leaving out hidden rows
 * and columns, and rows never broken across pages. `measure` gives the
 * width of a text at 100%.
 */
export function printDocument(
  workbook: Workbook,
  sheet: string,
  ranges: readonly Range[],
  settings: PrintSettings,
  measure: TextWidth,
): string {
  const [paperWidth, paperHeight] = PAPER_SIZES[settings.paper];
  const [vertical, horizontal] = MARGIN_SIZES[settings.margins];
  const width = settings.orientation === 'portrait' ? paperWidth : paperHeight;
  const printable = (width - 2 * horizontal) * CSS_INCH;
  const margin = `${vertical}in ${horizontal}in`;
  const style = [
    `@page { size: ${settings.paper} ${settings.orientation}; margin: ${margin} ${margin}; }`,
    'html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
    'body { margin: 0; color: #000000; background: #ffffff; }',
    `table { border-collapse: collapse; table-layout: fixed; font: 13.333px ${DEFAULT_FAMILY}; }`,
    'table + table { break-before: page; }',
    'tr { break-inside: avoid; }',
    'td { padding: 0 4px; white-space: pre; overflow: hidden; }',
    `th { border: ${GRIDLINE}; background: #f2f2f2; font-weight: normal; text-align: center; padding: 0; }`,
  ].join('\n');
  const tables = ranges
    .map((range) => table(workbook, sheet, range, settings, measure, printable))
    .join('\n');
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>${escaped(workbook.name)}</title>
<style>
${style}
</style></head>
<body>
${tables}
</body></html>`;
}

function table(
  workbook: Workbook,
  sheet: string,
  range: Range,
  settings: PrintSettings,
  measure: TextWidth,
  printable: number,
): string {
  const layout = layoutOf(workbook.layout(sheet));
  const lines = (from: number, to: number, size: (index: number) => number) =>
    Array.from({ length: to - from + 1 }, (_, at) => from + at).filter(
      (index) => size(index) > 0,
    );
  const rows = lines(range.top, range.bottom, (row) => layout.rows.size(row));
  const cols = lines(range.left, range.right, (col) => layout.cols.size(col));
  const colWidth = (col: number) => layout.cols.size(col);
  const total =
    cols.reduce((sum, col) => sum + colWidth(col), 0) +
    (settings.headers ? ROW_HEADER_WIDTH : 0);
  const zoom =
    settings.fitWidth && total > printable
      ? Math.round((printable / total) * 10000) / 10000
      : 1;

  // The merged cells: each shows as its first row and column printed,
  // spanning those printed, and covers the others.
  const spans = new Map<string, { rows: number; cols: number; merge: Range }>();
  const covered = new Set<string>();
  for (const merge of workbook.merges(sheet)) {
    const mergeRows = rows.filter((r) => r >= merge.top && r <= merge.bottom);
    const mergeCols = cols.filter((c) => c >= merge.left && c <= merge.right);
    if (mergeRows.length === 0 || mergeCols.length === 0) continue;
    for (const row of mergeRows) {
      for (const col of mergeCols) covered.add(`${row},${col}`);
    }
    const anchor = `${mergeRows[0]},${mergeCols[0]}`;
    covered.delete(anchor);
    spans.set(anchor, {
      rows: mergeRows.length,
      cols: mergeCols.length,
      merge,
    });
  }

  const shown = (row: number, col: number): ShownCell =>
    shownCell(workbook.cell(sheet, row, col), LINK_COLOR);
  const styleAt = (row: number, col: number): Style =>
    workbook.cell(sheet, row, col).style;
  // A side's border is the cell's own, or else that of the cell beside it
  // on the side, so that both cells draw it alike. The cell printed spans
  // the rows and columns printed from `rowAt` and `colAt` to `rowEnd` and
  // `colEnd`.
  const borders = (
    rowAt: number,
    colAt: number,
    rowEnd: number,
    colEnd: number,
  ): string => {
    const [top, bottom] = [rows[rowAt]!, rows[rowEnd]!];
    const [left, right] = [cols[colAt]!, cols[colEnd]!];
    const above = rows[rowAt - 1];
    const below = rows[rowEnd + 1];
    const leftOf = cols[colAt - 1];
    const rightOf = cols[colEnd + 1];
    const sides: [string, Border | undefined][] = [
      [
        'top',
        styleAt(top, left).borderTop ??
          (above === undefined ? undefined : styleAt(above, left).borderBottom),
      ],
      [
        'right',
        styleAt(top, right).borderRight ??
          (rightOf === undefined
            ? undefined
            : styleAt(top, rightOf).borderLeft),
      ],
      [
        'bottom',
        styleAt(bottom, left).borderBottom ??
          (below === undefined ? undefined : styleAt(below, left).borderTop),
      ],
      [
        'left',
        styleAt(top, left).borderLeft ??
          (leftOf === undefined ? undefined : styleAt(top, leftOf).borderRight),
      ],
    ];
    return sides
      .map(([side, border]) =>
        border
          ? `border-${side}:${borderCss(border)};`
          : settings.gridlines
            ? `border-${side}:${GRIDLINE};`
            : '',
      )
      .join('');
  };

  const body = rows.map((row, rowAt) => {
    const texts = new Map(
      rowTexts(
        (col) => shown(row, col),
        range.left,
        range.right,
        measure,
        colWidth,
        (col) => covered.has(`${row},${col}`) || spans.has(`${row},${col}`),
      ).map((text) => [text.col, text]),
    );
    const tds = cols.flatMap((col, colAt) => {
      const key = `${row},${col}`;
      if (covered.has(key)) return [];
      const span = spans.get(key);
      const cell = shown(span?.merge.top ?? row, span?.merge.left ?? col);
      let text = texts.get(col);
      if (span) {
        const spanned = cols
          .slice(colAt, colAt + span.cols)
          .reduce((sum, each) => sum + colWidth(each), 0);
        text = {
          ...cellText(cell, spanned, measure),
          col,
          style: cell.style,
          from: col,
          to: col,
        };
      }
      const { style } = cell;
      const color =
        cell.color ??
        style.color ??
        (style.fill ? textColorOn(style.fill) : '#000000');
      const decoration = [
        style.underline && 'underline',
        style.strike && 'line-through',
      ].filter(Boolean);
      const bar = workbook.cell(
        sheet,
        span?.merge.top ?? row,
        span?.merge.left ?? col,
      ).bar;
      const css =
        `font:${cssFont(style, 1)};color:${color};` +
        (style.fill ? `background-color:${style.fill};` : '') +
        (bar
          ? `background-image:linear-gradient(to right, ${bar.color} ${Math.round(bar.length * 100)}%, transparent 0);`
          : '') +
        (decoration.length > 0
          ? `text-decoration:${decoration.join(' ')};`
          : '') +
        `text-align:${text?.align ?? 'left'};` +
        `vertical-align:${style.vertical ?? 'bottom'};` +
        borders(
          rowAt,
          colAt,
          rowAt + (span?.rows ?? 1) - 1,
          colAt + (span?.cols ?? 1) - 1,
        ) +
        (text && Math.min(text.to, range.right) > col
          ? 'overflow:visible;'
          : '');
      const spanning = span
        ? ` rowspan="${span.rows}" colspan="${span.cols}"`
        : '';
      return [
        `<td${spanning} style="${attribute(css)}">${escaped(text?.lines.join('\n') ?? '')}</td>`,
      ];
    });
    const header = settings.headers ? `<th>${row + 1}</th>` : '';
    return `<tr style="height:${layout.rows.size(row)}px">${header}${tds.join('')}</tr>`;
  });

  const colgroup =
    (settings.headers ? `<col style="width:${ROW_HEADER_WIDTH}px">` : '') +
    cols.map((col) => `<col style="width:${colWidth(col)}px">`).join('\n');
  const head = settings.headers
    ? `<thead><tr style="height:${COLUMN_HEADER_HEIGHT}px"><th></th>${cols
        .map((col) => `<th>${columnName(col)}</th>`)
        .join('')}</tr></thead>`
    : '';
  const tableStyle = `width:${total}px;` + (zoom < 1 ? `zoom:${zoom};` : '');
  return `<table style="${tableStyle}"><colgroup>${colgroup}</colgroup>${head}<tbody>
${body.join('\n')}
</tbody></table>`;
}

const LINES: Record<Border['line'], string> = {
  thin: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  dashed: '1px dashed',
  dotted: '1px dotted',
  double: '3px double',
};

function borderCss({ line, color }: Border): string {
  return `${LINES[line]} ${color ?? '#000000'}`;
}

function escaped(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function attribute(text: string): string {
  return escaped(text).replaceAll('"', '&quot;');
}
