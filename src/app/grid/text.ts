import {
  fontPixels,
  SHEET_COLUMNS,
  TEXT_PADDING,
  wrapText,
  type Style,
  type TextWidth,
  type Value,
} from '../../engine/workbook';

export { TEXT_PADDING } from '../../engine/workbook';

/** The interface font, which cells have unless their style names another. */
export const DEFAULT_FAMILY =
  "system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif";

/** The fallbacks of each font cells can have, for systems without it. */
const FAMILIES: Readonly<Record<string, string>> = {
  Arial: 'Arial, Helvetica, "Liberation Sans", sans-serif',
  'Courier New': '"Courier New", "Liberation Mono", monospace',
  Georgia: 'Georgia, "DejaVu Serif", serif',
  'Times New Roman': '"Times New Roman", "Liberation Serif", serif',
  'Trebuchet MS': '"Trebuchet MS", "DejaVu Sans", sans-serif',
  Verdana: 'Verdana, "DejaVu Sans", sans-serif',
};

/** The CSS font of a style at a zoom, as a canvas takes it. */
export function cssFont(style: Style, zoom: number): string {
  const family =
    (style.font !== undefined && FAMILIES[style.font]) || DEFAULT_FAMILY;
  const size = Math.round(fontPixels(style) * zoom * 1000) / 1000;
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${size}px ${family}`;
}

/**
 * The color of text that has none of its own on a fill: black on a light
 * fill and white on a dark one, whatever the theme, as the theme's text
 * color may not show on the fill.
 */
export function textColorOn(fill: string): string {
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(fill.slice(at, at + 2), 16));
  const luminance = 0.299 * r! + 0.587 * g! + 0.114 * b!;
  return luminance > 140 ? '#000000' : '#ffffff';
}

/** The text of a cell as a row of the grid draws it. */
export interface RowText {
  readonly col: number;
  /** One line, or those that wrapped text takes. */
  readonly lines: readonly string[];
  readonly style: Style;
  /** The color the number format gives the text, over the style's. */
  readonly color?: string;
  readonly align: 'left' | 'right' | 'center';
  /** The first and last columns the text may cover. */
  readonly from: number;
  readonly to: number;
}

/** What a cell holds, as the grid draws it. */
export interface ShownCell {
  readonly value: Value;
  readonly display: string;
  readonly style: Style;
  readonly color?: string;
}

/**
 * The lines of a cell's text in a width, and how they align: as its style
 * says, or else numbers right, text left, and booleans and errors in the
 * middle. Text that does not wrap is one line, which may be too long; a
 * number too wide for the width is as many `#` as fit, as in Excel.
 */
export function cellText(
  { value, display, style }: ShownCell,
  width: number,
  measure: TextWidth,
): Pick<RowText, 'lines' | 'align'> {
  const align =
    style.horizontal ??
    (typeof value === 'number'
      ? 'right'
      : typeof value === 'string'
        ? 'left'
        : 'center');
  const inside = width - 2 * TEXT_PADDING;
  if (typeof value === 'number' && measure(display, style) > inside) {
    const count = Math.floor(inside / measure('#', style));
    return { lines: ['#'.repeat(Math.max(1, count))], align };
  }
  const lines = style.wrap
    ? wrapText(display, inside, (text) => measure(text, style))
    : [display];
  return { lines, align };
}

/**
 * The texts that the columns `left` to `right` of a row show, as in Excel:
 * aligned as `cellText` says; text aligned left that is too long for its
 * cell runs over the empty cells to its right, unless it wraps.
 * It includes text from a cell to the left of `left` that runs into them,
 * and leaves out that of hidden columns. `measure` gives the width of a
 * text at 100%, and `width` that of a column. The cells of merged cells,
 * which `merged` names, are drawn apart: they show nothing here, and no
 * text runs over them.
 */
export function rowTexts(
  cell: (col: number) => ShownCell,
  left: number,
  right: number,
  measure: TextWidth,
  width: (col: number) => number,
  merged: (col: number) => boolean = () => false,
): RowText[] {
  const texts: RowText[] = [];
  const empty = (col: number) => !merged(col) && cell(col).display === '';
  const add = (col: number) => {
    if (merged(col)) return;
    const shown = cell(col);
    const { display, style } = shown;
    if (display === '' || width(col) === 0) return;
    const { lines, align } = cellText(shown, width(col), measure);
    let to = col;
    if (align === 'left' && !style.wrap) {
      const needed = measure(display, style) + 2 * TEXT_PADDING;
      let covered = width(col);
      while (covered < needed && to + 1 < SHEET_COLUMNS && empty(to + 1)) {
        to++;
        covered += width(to);
      }
    }
    if (to >= left) {
      const { color } = shown;
      texts.push({
        col,
        lines,
        style,
        ...(color !== undefined && { color }),
        align,
        from: col,
        to,
      });
    }
  };
  // The nearest text to the left may run into the columns.
  let before = left - 1;
  while (before >= 0 && empty(before)) before--;
  if (before >= 0) add(before);
  for (let col = left; col <= right; col++) add(col);
  return texts;
}
