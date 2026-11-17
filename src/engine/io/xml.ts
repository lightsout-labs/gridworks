import { columnName, type Area } from '../formula';

/** The attributes of an element's start tag, decoded. */
export function attributes(tag: string): Record<string, string> {
  const found: Record<string, string> = {};
  for (const [, name, double, single] of tag.matchAll(
    /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g,
  )) {
    found[name!] = decode(double ?? single ?? '');
  }
  return found;
}

/**
 * Each element of a name in some XML, with its attributes and the XML
 * inside it. The name must not have a prefix that another name ends with.
 */
export function elements(
  xml: string,
  name: string,
): { attributes: Record<string, string>; inner: string }[] {
  const pattern = new RegExp(
    `<${name}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${name}>)`,
    'g',
  );
  return Array.from(xml.matchAll(pattern), ([, tag = '', inner = '']) => ({
    attributes: attributes(tag),
    inner,
  }));
}

/** The XML inside the first element of a name, or `''`. */
export function inside(xml: string, name: string): string {
  return elements(xml, name)[0]?.inner ?? '';
}

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Text with XML's entities and Excel's `_xHHHH_` escapes replaced. */
export function decode(text: string): string {
  if (!text.includes('&') && !text.includes('_x')) return text;
  return text
    .replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (entity, name: string) =>
      name.startsWith('#')
        ? String.fromCodePoint(
            name[1] === 'x' || name[1] === 'X'
              ? parseInt(name.slice(2), 16)
              : Number(name.slice(1)),
          )
        : (ENTITIES[name] ?? entity),
    )
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_, code: string) =>
      String.fromCharCode(parseInt(code, 16)),
    );
}

/**
 * Text written for XML, in an element or an attribute, with Excel's
 * escapes for the control characters XML cannot hold and for text that
 * would read as one.
 */
export function encode(text: string): string {
  return text
    .replace(/_x[0-9A-Fa-f]{4}_/g, (escape) => `_x005F${escape}`)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replace(
      // eslint-disable-next-line no-control-regex
      /[\u0000-\u0008\u000B-\u001F]/g,
      (char) => `_x${char.charCodeAt(0).toString(16).padStart(4, '0')}_`,
    );
}

/** The index of a column from its letters, `A` being 0. */
export function columnIndex(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + letters.charCodeAt(i) - 64;
  }
  return col - 1;
}

const CELL = /^\$?([A-Z]{1,3})\$?(\d+)$/;

/** The area of a reference such as `A1` or `B2:C3`, or `null`. */
export function areaOf(reference: string): Area | null {
  const [first = '', last = first] = reference.split(':');
  const start = CELL.exec(first);
  const end = CELL.exec(last);
  if (!start || !end) return null;
  const rows = [Number(start[2]) - 1, Number(end[2]) - 1];
  const cols = [columnIndex(start[1]!), columnIndex(end[1]!)];
  return {
    top: Math.min(...rows),
    left: Math.min(...cols),
    bottom: Math.max(...rows),
    right: Math.max(...cols),
  };
}

/** The areas of a list of references separated by spaces, as `sqref`. */
export function areasOf(references: string): Area[] {
  return references
    .split(/\s+/)
    .map(areaOf)
    .filter((area) => area !== null);
}

/** An area as Excel writes it: `A1` for one cell, `B2:C3` for more. */
export function areaName({ top, left, bottom, right }: Area): string {
  const first = columnName(left) + (top + 1);
  return top === bottom && left === right
    ? first
    : `${first}:${columnName(right)}${bottom + 1}`;
}
