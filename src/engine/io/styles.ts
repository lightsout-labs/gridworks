import { builtinFormat, GENERAL } from '../format';
import {
  BORDER_LINES,
  FONTS,
  type Border,
  type RuleFormat,
  type Style,
} from '../workbook';
import { attributes, elements, encode, inside } from './xml';

/** A color of a file, as its element's attributes give it. */
export type ColorOf = (
  color: Record<string, string> | undefined,
) => string | undefined;

/** The styles of a file, as Gridworks has them. */
export interface FileStyles {
  /** The table of styles, the default one first. */
  readonly styles: Style[];
  /** The index in `styles` of each cell format of the file. */
  readonly styleOf: (cellFormat: number) => number;
  /** What each differential format, of conditional formats, changes. */
  readonly ruleFormat: (dxf: number) => RuleFormat;
  readonly colorOf: ColorOf;
}

/** Excel's palette of indexed colors, 64 and 65 being automatic. */
const INDEXED = (
  '000000FFFFFFFF000000FF000000FFFFFF00FF00FF00FFFF000000FFFFFFFF000000FF00' +
  '0000FFFFFF00FF00FF00FFFF800000008000000080808000800080008080C0C0C0808080' +
  '9999FF993366FFFFCCCCFFFF660066FF80800066CCCCCCFF000080FF00FFFFFF0000FFFF' +
  '8000808000000080800000FF00CCFFCCFFFFCCFFCCFFFF9999CCFFFF99CCCC99FFFFCC99' +
  '3366FF33CCCC99CC00FFCC00FF9900FF6600666699969696003366339966003300333300' +
  '993300993366333399333333'
).match(/.{6}/g)!;

/** The names of a theme's colors, in the order of Excel's theme indexes. */
const THEME_COLORS = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/** Excel's lines that Gridworks draws as the nearest of its own. */
const NEAREST_LINE: Record<string, Border['line']> = {
  hair: 'dotted',
  mediumDashed: 'dashed',
  dashDot: 'dashed',
  mediumDashDot: 'dashed',
  dashDotDot: 'dashed',
  mediumDashDotDot: 'dashed',
  slantDashDot: 'dashed',
};

const SIDES = [
  ['borderLeft', 'left'],
  ['borderRight', 'right'],
  ['borderTop', 'top'],
  ['borderBottom', 'bottom'],
] as const;

/** The font of a file's style, as far as Gridworks reads it. */
interface Font {
  readonly name?: string;
  readonly size?: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly color?: Record<string, string>;
}

/** Whether a font has a property such as `<b/>`, not turned off. */
function flag(xml: string, name: string): boolean {
  const match = new RegExp(`<${name}(?:\\s+val="([^"]*)")?\\s*/>`).exec(xml);
  return !!match && !['0', 'false', 'none'].includes(match[1] ?? '');
}

function fontOf(xml: string): Font {
  const value = (name: string) =>
    new RegExp(`<${name}\\s+val="([^"]*)"`).exec(xml)?.[1];
  const size = value('sz');
  return {
    name: value('name'),
    ...(size !== undefined && { size: Number(size) }),
    bold: flag(xml, 'b'),
    italic: flag(xml, 'i'),
    underline: flag(xml, 'u'),
    strike: flag(xml, 'strike'),
    color: elements(xml, 'color')[0]?.attributes,
  };
}

/**
 * Reads the styles part of a file, with the colors of its theme: fonts
 * that Gridworks has, sizes other than the file's default, colors, solid
 * fills, alignment, wrapping, borders and number formats. Excel's default
 * font and automatic colors are left to Gridworks' own, and the Hyperlink
 * style, as the grid draws links itself.
 */
export function readStyles(xml: string, theme: string): FileStyles {
  const scheme = inside(theme, 'a:clrScheme');
  const palette = THEME_COLORS.map(
    (name) =>
      /(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"/.exec(
        inside(scheme, `a:${name}`),
      )?.[1],
  );
  const colorOf: ColorOf = (color) => {
    if (!color) return undefined;
    let hex: string | undefined;
    if (color.rgb) hex = color.rgb.slice(-6);
    else if (color.theme) hex = palette[Number(color.theme)];
    else if (color.indexed) hex = INDEXED[Number(color.indexed)];
    if (!hex) return undefined;
    const tint = Number(color.tint ?? 0);
    return `#${(tint ? tinted(hex, tint) : hex).toLowerCase()}`;
  };
  // Automatic, and the theme's text color, are the text color of Gridworks.
  const textColor: ColorOf = (color) =>
    !color || color.auto === '1' || (color.theme === '1' && !color.tint)
      ? undefined
      : colorOf(color);

  const custom = new Map<number, string>();
  for (const tag of xml.match(/<numFmt\b[^>]*>/g) ?? []) {
    const { numFmtId, formatCode } = attributes(tag);
    if (numFmtId && formatCode) custom.set(Number(numFmtId), formatCode);
  }
  const fonts = elements(inside(xml, 'fonts'), 'font').map(({ inner }) =>
    fontOf(inner),
  );
  const fills = elements(inside(xml, 'fills'), 'fill').map(({ inner }) => {
    const pattern = elements(inner, 'patternFill')[0];
    return pattern?.attributes.patternType === 'solid'
      ? colorOf(elements(pattern.inner, 'fgColor')[0]?.attributes)
      : undefined;
  });
  const borders = elements(inside(xml, 'borders'), 'border').map(
    ({ inner }) => {
      const border: Partial<Record<(typeof SIDES)[number][0], Border>> = {};
      for (const [property, side] of SIDES) {
        const element = elements(inner, side)[0];
        const style = element?.attributes.style;
        if (!element || !style || style === 'none') continue;
        const line = (BORDER_LINES as readonly string[]).includes(style)
          ? (style as Border['line'])
          : (NEAREST_LINE[style] ?? 'thin');
        const color = textColor(
          elements(element.inner, 'color')[0]?.attributes,
        );
        border[property] = { line, ...(color && { color }) };
      }
      return border;
    },
  );
  const linkStyles = new Set(
    elements(inside(xml, 'cellStyles'), 'cellStyle')
      .filter(({ attributes }) => attributes.builtinId === '8')
      .map(({ attributes }) => attributes.xfId),
  );
  const defaultFont = fonts[0];

  const styles: Style[] = [{}];
  const index = new Map<string, number>([['{}', 0]]);
  const ids = elements(inside(xml, 'cellXfs'), 'xf').map(
    ({ attributes: xf, inner }) => {
      const code =
        custom.get(Number(xf.numFmtId ?? 0)) ??
        builtinFormat(Number(xf.numFmtId ?? 0)) ??
        GENERAL;
      const fontId = Number(xf.fontId ?? 0);
      const font = fontId === 0 ? undefined : fonts[fontId];
      const link = linkStyles.has(xf.xfId ?? '0');
      const fill = fills[Number(xf.fillId ?? 0)];
      const alignment = elements(inner, 'alignment')[0]?.attributes ?? {};
      const horizontal = (
        {
          left: 'left',
          center: 'center',
          centerContinuous: 'center',
          right: 'right',
          fill: 'left',
          justify: 'left',
          distributed: 'center',
        } as const
      )[alignment.horizontal ?? ''];
      const vertical = (
        {
          top: 'top',
          center: 'middle',
          justify: 'top',
          distributed: 'middle',
        } as const
      )[alignment.vertical ?? ''];
      const color = font && !link ? textColor(font.color) : undefined;
      const style: Style = {
        ...(code !== GENERAL && { numberFormat: code }),
        ...(font?.name && FONTS.includes(font.name) && { font: font.name }),
        ...(font?.size !== undefined &&
          font.size !== defaultFont?.size && { size: font.size }),
        ...(font?.bold && { bold: true }),
        ...(font?.italic && { italic: true }),
        ...(font?.underline && !link && { underline: true }),
        ...(font?.strike && { strike: true }),
        ...(color && { color }),
        ...(fill && { fill }),
        ...(horizontal && { horizontal }),
        ...(vertical && { vertical }),
        ...(['1', 'true'].includes(alignment.wrapText ?? '') && { wrap: true }),
        ...(borders[Number(xf.borderId ?? 0)] ?? {}),
      };
      const key = JSON.stringify(style);
      let styleId = index.get(key);
      if (styleId === undefined) {
        styleId = styles.push(style) - 1;
        index.set(key, styleId);
      }
      return styleId;
    },
  );
  const formats = elements(inside(xml, 'dxfs'), 'dxf').map(({ inner }) => {
    const font = fontOf(inside(inner, 'font'));
    const pattern = inside(inner, 'fill');
    const color = colorOf(font.color);
    const fill = colorOf(
      (elements(pattern, 'bgColor')[0] ?? elements(pattern, 'fgColor')[0])
        ?.attributes,
    );
    const format: RuleFormat = {
      ...(font.bold && { bold: true }),
      ...(font.italic && { italic: true }),
      ...(font.underline && { underline: true }),
      ...(font.strike && { strike: true }),
      ...(color && { color }),
      ...(fill && { fill }),
    };
    return format;
  });
  return {
    styles,
    styleOf: (format) => ids[format] ?? 0,
    ruleFormat: (dxf) => formats[dxf] ?? {},
    colorOf,
  };
}

/** A color lightened or darkened by a tint, as Excel does to theme colors. */
function tinted(hex: string, tint: number): string {
  const [r, g, b] = [0, 2, 4].map(
    (at) => parseInt(hex.slice(at, at + 2), 16) / 255,
  ) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const chroma = max - min;
  let hue = 0;
  if (chroma > 0) {
    hue =
      max === r
        ? ((g - b) / chroma + 6) % 6
        : max === g
          ? (b - r) / chroma + 2
          : (r - g) / chroma + 4;
  }
  let lightness = (max + min) / 2;
  const saturation =
    chroma === 0 ? 0 : chroma / (1 - Math.abs(2 * lightness - 1));
  lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const c = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = c * (1 - Math.abs((hue % 2) - 1));
  const m = lightness - c / 2;
  const [red, green, blue] = [
    [c, x, 0],
    [x, c, 0],
    [0, c, x],
    [0, x, c],
    [x, 0, c],
    [c, 0, x],
  ][Math.floor(hue) % 6]!;
  return [red!, green!, blue!]
    .map((channel) =>
      Math.round((channel + m) * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('');
}

/** Excel's default font, which a file's styles start with. */
const DEFAULT_FONT = '<sz val="11"/><name val="Calibri"/><family val="2"/>';

/** A color as a file holds it, from `#rrggbb`. */
const argb = (color: string) => `rgb="FF${color.slice(1).toUpperCase()}"`;

/** The properties of a font, in the order of the file's schema. */
function fontProperties(style: RuleFormat): string {
  return (
    (style.bold ? '<b/>' : '') +
    (style.italic ? '<i/>' : '') +
    (style.strike ? '<strike/>' : '') +
    (style.underline ? '<u/>' : '')
  );
}

/**
 * The styles part of a file holding the table of styles, each style at its
 * own index among the cell formats, and the formats of conditional formats
 * as differential formats, in their order.
 */
export function writeStyles(
  styles: readonly Style[],
  formats: readonly RuleFormat[],
): string {
  const builtin = new Map<string, number>();
  for (let id = 49; id > 0; id--) {
    const code = builtinFormat(id);
    if (code !== null) builtin.set(code, id);
  }
  const numberFormats: string[] = [];
  const fonts = [`<font>${DEFAULT_FONT}</font>`];
  const fills = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];
  const borders = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
  ];
  const indexIn = (list: string[], xml: string) => {
    const found = list.indexOf(xml);
    return found >= 0 ? found : list.push(xml) - 1;
  };
  const cellFormats = styles.map((style) => {
    let format = 0;
    const code = style.numberFormat;
    if (code !== undefined && code !== GENERAL) {
      format = builtin.get(code) ?? 164 + indexIn(numberFormats, code);
    }
    const font =
      style.font !== undefined ||
      style.size !== undefined ||
      style.bold ||
      style.italic ||
      style.underline ||
      style.strike ||
      style.color !== undefined
        ? indexIn(
            fonts,
            `<font>${fontProperties(style)}<sz val="${style.size ?? 11}"/>${
              style.color ? `<color ${argb(style.color)}/>` : ''
            }<name val="${encode(style.font ?? 'Calibri')}"/><family val="2"/></font>`,
          )
        : 0;
    const fill = style.fill
      ? indexIn(
          fills,
          `<fill><patternFill patternType="solid"><fgColor ${argb(style.fill)}/><bgColor indexed="64"/></patternFill></fill>`,
        )
      : 0;
    const sides = SIDES.map(([property, side]) => {
      const border = style[property];
      return border
        ? `<${side} style="${border.line}"><color ${border.color ? argb(border.color) : 'indexed="64"'}/></${side}>`
        : `<${side}/>`;
    }).join('');
    const border = SIDES.some(([property]) => style[property])
      ? indexIn(borders, `<border>${sides}<diagonal/></border>`)
      : 0;
    const alignment =
      style.horizontal || style.vertical || style.wrap
        ? `<alignment${style.horizontal ? ` horizontal="${style.horizontal}"` : ''}${
            style.vertical
              ? ` vertical="${style.vertical === 'middle' ? 'center' : 'top'}"`
              : ''
          }${style.wrap ? ' wrapText="1"' : ''}/>`
        : '';
    const applied =
      (format ? ' applyNumberFormat="1"' : '') +
      (font ? ' applyFont="1"' : '') +
      (fill ? ' applyFill="1"' : '') +
      (border ? ' applyBorder="1"' : '') +
      (alignment ? ' applyAlignment="1"' : '');
    const tag = `<xf numFmtId="${format}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${applied}`;
    return alignment ? `${tag}>${alignment}</xf>` : `${tag}/>`;
  });
  const dxfs = formats.map((format) => {
    const font =
      fontProperties(format) +
      (format.color ? `<color ${argb(format.color)}/>` : '');
    return `<dxf>${font ? `<font>${font}</font>` : ''}${
      format.fill
        ? `<fill><patternFill><bgColor ${argb(format.fill)}/></patternFill></fill>`
        : ''
    }</dxf>`;
  });
  const list = (name: string, items: string[]) =>
    `<${name} count="${items.length}">${items.join('')}</${name}>`;
  return (
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    (numberFormats.length > 0
      ? list(
          'numFmts',
          numberFormats.map(
            (code, at) =>
              `<numFmt numFmtId="${164 + at}" formatCode="${encode(code)}"/>`,
          ),
        )
      : '') +
    list('fonts', fonts) +
    list('fills', fills) +
    list('borders', borders) +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    list('cellXfs', cellFormats) +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    list('dxfs', dxfs) +
    '</styleSheet>'
  );
}
