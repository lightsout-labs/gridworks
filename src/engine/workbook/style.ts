/** How cells look. Cells that look the same share one. */
export interface Style {
  /** An Excel format code; `General` when absent. */
  readonly numberFormat?: string;
  /** One of `FONTS`; the system's interface font when absent. */
  readonly font?: string;
  /** In points; `DEFAULT_FONT_SIZE` when absent. */
  readonly size?: number;
  readonly bold?: true;
  readonly italic?: true;
  readonly underline?: true;
  readonly strike?: true;
  /** The text color, as `#rrggbb`; the theme's when absent. */
  readonly color?: string;
  /** The fill color, as `#rrggbb`; none when absent. */
  readonly fill?: string;
  /** By the type of the value when absent, as Excel's General alignment. */
  readonly horizontal?: 'left' | 'center' | 'right';
  /** At the bottom when absent, as in Excel. */
  readonly vertical?: 'top' | 'middle';
  /** Whether long text wraps onto more lines instead of running over. */
  readonly wrap?: true;
}

/** Properties to give cells, `null` taking one back to the default. */
export type StyleChange = {
  readonly [K in keyof Style]?: Style[K] | null;
};

/** The order of a style's properties, so that equal styles print the same. */
const ORDER: readonly (keyof Style)[] = [
  'numberFormat',
  'font',
  'size',
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'fill',
  'horizontal',
  'vertical',
  'wrap',
];

/** A style with a change made to it. */
export function changedStyle(style: Style, change: StyleChange): Style {
  const result: Record<string, unknown> = {};
  for (const name of ORDER) {
    const value = name in change ? change[name] : style[name];
    if (value !== null && value !== undefined) result[name] = value;
  }
  return result as Style;
}

/**
 * The fonts cells can have, present on Windows and macOS, each with
 * fallbacks for the other systems.
 */
export const FONTS: readonly string[] = [
  'Arial',
  'Courier New',
  'Georgia',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];

/** The font sizes offered, in points, as Excel offers them. */
export const FONT_SIZES: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72,
];

export const DEFAULT_FONT_SIZE = 10;

/** The standard sizes of rows and columns, in CSS pixels at 100%. */
export const STANDARD_ROW_HEIGHT = 24;
export const STANDARD_COLUMN_WIDTH = 96;

/** The space between the text of a cell and its borders, at 100%. */
export const TEXT_PADDING = 4;

/** The width of a text in a style, in CSS pixels at 100%. */
export type TextWidth = (text: string, style: Style) => number;

/** The size of a style's font, in CSS pixels at 100%. */
export function fontPixels(style: Style): number {
  return ((style.size ?? DEFAULT_FONT_SIZE) * 4) / 3;
}

/** The height of a line of text in a style, in CSS pixels at 100%. */
export function lineHeight(style: Style): number {
  return Math.ceil(fontPixels(style) * 1.2);
}

/**
 * The lines a text takes in a width, broken between words; a word longer
 * than the width has a line of its own.
 */
export function wrapText(
  text: string,
  width: number,
  measure: (text: string) => number,
): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split('\n')) {
    let line = '';
    for (const word of paragraph.split(' ')) {
      const longer = line === '' ? word : `${line} ${word}`;
      if (line === '' || measure(longer) <= width) {
        line = longer;
      } else {
        lines.push(line);
        line = word;
      }
    }
    lines.push(line);
  }
  return lines;
}

/** Whether cells of a style may need a row higher than the standard one. */
export function mayGrowRow(style: Style): boolean {
  return style.wrap === true || (style.size ?? 0) > DEFAULT_FONT_SIZE;
}

/** The height a cell's text needs in a column, in CSS pixels at 100%. */
export function cellHeight(
  text: string,
  style: Style,
  width: number,
  measure: TextWidth,
): number {
  const lines = style.wrap
    ? wrapText(text, width - 2 * TEXT_PADDING, (each) => measure(each, style))
        .length
    : 1;
  return lines * lineHeight(style) + 2 * TEXT_PADDING;
}

/** A guess at the width of a text, where no fonts can measure it. */
export const estimatedWidth: TextWidth = (text, style) =>
  text.length * fontPixels(style) * 0.55;
