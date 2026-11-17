import JSZip from 'jszip';
import {
  calculable,
  parse,
  print,
  shift,
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Area,
  type Ast,
} from '../formula';
import type {
  Condition,
  Criterion,
  NamedRange,
  Rule,
  SheetSnapshot,
  Validation,
} from '../workbook';
import type { Dropped, DroppedFeature, Opened } from '.';
import { readStyles, type FileStyles } from './styles';
import {
  areaOf,
  areasOf,
  attributes,
  columnIndex,
  decode,
  elements,
  inside,
} from './xml';

type Constant = SheetSnapshot['columns'][number][number];

/** How many of each feature a file had that Gridworks leaves out. */
type Counts = Record<DroppedFeature, number>;

/** The parts of a file that hold content Gridworks does not support. */
const UNSUPPORTED: [DroppedFeature, RegExp][] = [
  ['chart', /^xl\/charts\/chart\d+\.xml$/],
  ['pivotTable', /^xl\/pivotTables\/pivotTable\d+\.xml$/],
  ['picture', /^xl\/media\//],
  ['table', /^xl\/tables\/table\d+\.xml$/],
];

/** The order in which what was dropped is listed. */
const FEATURES: readonly DroppedFeature[] = [
  'chart',
  'pivotTable',
  'picture',
  'table',
  'formula',
  'cell',
  'conditionalFormat',
  'validation',
  'name',
  'link',
];

const RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** The schemes of the web addresses Gridworks keeps as links (R29). */
const WEB_LINK = /^(?:https?|mailto):/i;

/**
 * Reads an `.xlsx` file as a workbook named `name`. Rejects a file that is
 * not an `.xlsx` file.
 *
 * Our own reader, not a library, so that a 10 MB file opens within 5
 * seconds (R47; ADR 0006): it scans the parts it needs and skips
 * everything else.
 */
export async function readXlsx(
  bytes: Uint8Array,
  name: string,
): Promise<Opened> {
  const zip = await JSZip.loadAsync(bytes);
  const read = async (path: string): Promise<string> => {
    const file = zip.file(path);
    if (!file) throw new Error(`Not an .xlsx file: ${path} is missing`);
    return file.async('string');
  };
  const optional = (path: string | undefined) =>
    path && zip.file(path) ? read(path) : Promise.resolve('');

  const [workbookPath] = relationships(await optional('_rels/.rels'), '')
    .filter(({ type }) => type === 'officeDocument')
    .map(({ path }) => path);
  if (!workbookPath) throw new Error('Not an .xlsx file: no workbook');
  const workbookXml = await read(workbookPath);
  // Another kind of document, such as a `.docx` file, has another root; a
  // file whose elements have a namespace prefix is not read either.
  if (!/<workbook\b[^>]*\bxmlns="[^"]*spreadsheetml/.test(workbookXml)) {
    throw new Error('Not an .xlsx file: no spreadsheet workbook');
  }
  const parts = relationships(
    await optional(relationshipsPath(workbookPath)),
    directory(workbookPath),
  );
  const partOf = (type: string) => parts.find((part) => part.type === type);
  const [strings, styleXml, theme] = await Promise.all([
    optional(partOf('sharedStrings')?.path).then(sharedStrings),
    optional(partOf('styles')?.path),
    optional(partOf('theme')?.path),
  ]);
  const styles = readStyles(styleXml, theme);
  const counts = Object.fromEntries(
    FEATURES.map((feature) => [feature, 0]),
  ) as Counts;
  for (const [feature, path] of UNSUPPORTED) {
    counts[feature] = zip.file(path).length;
  }
  const names = definedNames(workbookXml, counts);
  const known = new Set(names.map(({ name }) => name.toLowerCase()));

  const sheets: SheetSnapshot[] = [];
  for (const tag of workbookXml.match(/<sheet\b[^>]*>/g) ?? []) {
    const { name: sheetName = '', 'r:id': id } = attributes(tag);
    const part = parts.find((each) => each.id === id);
    // A chart sheet is a chart, which the chart parts count.
    if (part?.type !== 'worksheet') continue;
    const links = relationships(
      await optional(relationshipsPath(part.path)),
      directory(part.path),
    );
    const comments = links.find(({ type }) => type === 'comments')?.path;
    sheets.push(
      readSheet(`sheet${sheets.length + 1}`, sheetName, await read(part.path), {
        strings,
        styles,
        counts,
        isName: (each) => known.has(each.toLowerCase()),
        links,
        comments: await optional(comments),
      }),
    );
  }
  if (sheets.length === 0) {
    sheets.push({ id: 'sheet1', name: 'Sheet1', columns: [], formulas: [] });
  }

  const dropped: Dropped[] = FEATURES.filter(
    (feature) => counts[feature] > 0,
  ).map((feature) => ({ feature, count: counts[feature] }));
  return {
    snapshot: {
      version: 1,
      name,
      sheets,
      styles: styles.styles,
      ...(names.length > 0 && { names }),
    },
    dropped,
  };
}

/**
 * The defined names of a workbook that Gridworks supports: names of a
 * range, or of `#REF!`, in the whole workbook. Excel's own names, such as
 * a print area, are skipped; the others are counted as dropped.
 */
function definedNames(xml: string, counts: Counts): NamedRange[] {
  const names: NamedRange[] = [];
  for (const { attributes: defined, inner } of elements(
    inside(xml, 'definedNames'),
    'definedName',
  )) {
    const name = defined.name ?? '';
    if (name.startsWith('_xlnm.')) continue;
    const refersTo = decode(inner);
    const target = parse(refersTo);
    const supported =
      defined.localSheetId === undefined &&
      defined.hidden !== '1' &&
      ((target?.type === 'range' && target.sheet !== undefined) ||
        refersTo === '#REF!');
    if (supported) names.push({ name, refersTo });
    else counts.name++;
  }
  return names;
}

/** What reading a worksheet needs from the rest of its file. */
interface SheetParts {
  readonly strings: string[];
  readonly styles: FileStyles;
  readonly counts: Counts;
  /** Whether a name is one of the workbook's defined names. */
  readonly isName: (name: string) => boolean;
  /** The worksheet's relationships, to hyperlinks and its comments. */
  readonly links: Relationship[];
  /** The part of its notes, or `''`. */
  readonly comments: string;
}

// One element in the cells of a worksheet: a row, or a cell with its contents.
const ROW_OR_CELL = /<row\b([^>]*)>|<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
const REFERENCE = /\br="([A-Z]+)(\d+)"/;
const ROW_NUMBER = /\br="(\d+)"/;
const TYPE = /\bt="(\w+)"/;
const STYLE = /\bs="(\d+)"/;
const VALUE = /<v>([\s\S]*?)<\/v>/;
const FORMULA = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/;
const INLINE = /<is>([\s\S]*?)<\/is>/;
const SHARED_INDEX = /\bsi="(\d+)"/;

/** Excel's widths, in characters of 7 pixels, as Gridworks' pixels. */
const columnPixels = (width: number) =>
  Math.round((width < 1 ? width * 12 : width * 7 + 5) * 1.5);

/** Excel's heights, in points, as Gridworks' pixels. */
const rowPixels = (points: number) => Math.round(points * 1.6);

function readSheet(
  id: string,
  name: string,
  xml: string,
  { strings, styles, counts, isName, links, comments }: SheetParts,
): SheetSnapshot {
  const columns: Constant[][] = [];
  const styleIds: number[][] = [];
  const formulas: [number, number, string][] = [];
  const rowHeights: [number, number][] = [];
  const hiddenRows: number[] = [];
  // The first formula of each group shared down or across, where it is.
  const shared = new Map<
    string,
    { ast: Ast | null; row: number; col: number }
  >();
  if (!xml.includes('<worksheet')) {
    throw new Error(`Not an .xlsx file: ${name} is not a worksheet`);
  }
  const start = xml.indexOf('<sheetData');
  const end = start < 0 ? xml.length : xml.indexOf('</sheetData>', start);
  const before = xml.slice(0, Math.max(start, 0));
  const after = end < 0 ? '' : xml.slice(end);

  let row = -1;
  let col = -1;
  ROW_OR_CELL.lastIndex = start < 0 ? xml.length : start;
  for (let match; (match = ROW_OR_CELL.exec(xml));) {
    if (end >= 0 && match.index > end) break;
    const [, rowAttributes, cellAttributes = '', contents = ''] = match;
    if (rowAttributes !== undefined) {
      const number = ROW_NUMBER.exec(rowAttributes)?.[1];
      row = number ? Number(number) - 1 : row + 1;
      col = -1;
      if (row < SHEET_ROWS) {
        if (rowAttributes.includes('hidden="1"')) hiddenRows.push(row);
        if (rowAttributes.includes('customHeight="1"')) {
          const height = Number(attributes(rowAttributes).ht);
          if (height > 0) rowHeights.push([row, rowPixels(height)]);
        }
      }
      continue;
    }
    const reference = REFERENCE.exec(cellAttributes);
    if (reference) {
      col = columnIndex(reference[1]!);
      row = Number(reference[2]) - 1;
    } else {
      col++;
    }
    if (row >= SHEET_ROWS || col >= SHEET_COLUMNS) {
      counts.cell++;
      continue;
    }

    let value = cellValue(TYPE.exec(cellAttributes)?.[1], contents, strings);
    const formula = contents.includes('<f') && FORMULA.exec(contents);
    if (formula) {
      const [, formulaAttributes = '', text] = formula;
      const group = SHARED_INDEX.exec(formulaAttributes)?.[1];
      let ast = text === undefined ? null : parse(decode(text));
      if (group !== undefined) {
        const first = shared.get(group);
        if (text !== undefined || !first) {
          shared.set(group, { ast, row, col });
        } else {
          ast =
            first.ast &&
            shift(first.ast, {
              type: 'offset',
              rows: row - first.row,
              cols: col - first.col,
            });
        }
      }
      if (ast && calculable(ast, isName)) {
        formulas.push([row, col, `=${print(ast)}`]);
        value = null;
      } else {
        counts.formula++;
      }
    }
    if (value !== null) (columns[col] ??= [])[row] = value;
    const style = STYLE.exec(cellAttributes);
    if (style) {
      const styleId = styles.styleOf(Number(style[1]));
      if (styleId !== 0) (styleIds[col] ??= [])[row] = styleId;
    }
  }

  const columnWidths: [number, number][] = [];
  const hiddenColumns: number[] = [];
  for (const { attributes: line } of elements(inside(before, 'cols'), 'col')) {
    const last = Math.min(Number(line.max ?? line.min), SHEET_COLUMNS);
    for (let each = Number(line.min) - 1; each < last; each++) {
      if (line.hidden === '1') hiddenColumns.push(each);
      if (line.customWidth === '1' && Number(line.width) > 0) {
        columnWidths.push([each, columnPixels(Number(line.width))]);
      }
    }
  }
  const pane = elements(before, 'pane')[0]?.attributes;
  const frozen =
    pane && pane.state?.startsWith('frozen')
      ? { rows: Number(pane.ySplit ?? 0), cols: Number(pane.xSplit ?? 0) }
      : undefined;
  const color = styles.colorOf(elements(before, 'tabColor')[0]?.attributes);
  const merges = elements(inside(after, 'mergeCells'), 'mergeCell')
    .map(({ attributes: merge }) => areaOf(merge.ref ?? ''))
    .filter((area) => area !== null);
  const rules = conditionalFormats(after, styles, counts);
  const validations = dataValidations(after, counts);
  const notes = notesOf(comments);
  const hyperlinks = hyperlinksOf(after, links, counts);

  return {
    id,
    name,
    ...(color && { color }),
    columns: Array.from(columns, (column) =>
      Array.from(column ?? [], (value) => value ?? null),
    ),
    formulas,
    styleIds: Array.from(styleIds, (column) =>
      Array.from(column ?? [], (styleId) => styleId ?? 0),
    ),
    ...(rowHeights.length > 0 && { rowHeights }),
    ...(columnWidths.length > 0 && { columnWidths }),
    ...(hiddenRows.length > 0 && { hiddenRows }),
    ...(hiddenColumns.length > 0 && { hiddenColumns }),
    ...(merges.length > 0 && { merges }),
    ...(frozen && (frozen.rows > 0 || frozen.cols > 0) && { frozen }),
    ...(rules.length > 0 && { conditionalFormats: rules }),
    ...(validations.length > 0 && { validations }),
    ...(notes.length > 0 && { notes }),
    ...(hyperlinks.length > 0 && { links: hyperlinks }),
  };
}

/**
 * The value of a condition, from a formula of a constant: a number, a
 * boolean or a quoted text. `null` for any other formula.
 */
function operand(formula: string): string | null {
  const text = decode(formula).trim();
  if (/^-?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?$/i.test(text)) {
    return String(Number(text));
  }
  if (/^(?:TRUE|FALSE)$/i.test(text)) return text.toUpperCase();
  const quoted = /^"((?:[^"]|"")*)"$/.exec(text);
  return quoted ? quoted[1]!.replaceAll('""', '"') : null;
}

const COMPARISONS: Record<string, 'greater' | 'less' | 'equal'> = {
  greaterThan: 'greater',
  lessThan: 'less',
  equal: 'equal',
};

/** The condition of a rule of a file, if Gridworks has it. */
function conditionOf(
  rule: Record<string, string>,
  inner: string,
  colorOf: FileStyles['colorOf'],
): Condition | null {
  const values = elements(inner, 'formula').map(({ inner: each }) =>
    operand(each),
  );
  const [value, other] = values;
  const points = elements(inner, 'cfvo').map(
    ({ attributes: point }) => `${point.type}${point.val ?? ''}`,
  );
  const colors = elements(inner, 'color').map(({ attributes: color }) =>
    colorOf(color),
  );
  switch (rule.type) {
    case 'cellIs': {
      if (values.length === 0 || values.includes(null)) return null;
      if (rule.operator === 'between' && other !== undefined) {
        return { type: 'between', value: value!, other: other! };
      }
      const type = COMPARISONS[rule.operator ?? ''];
      return type ? { type, value: value! } : null;
    }
    case 'containsText':
      return rule.operator === 'containsText' && rule.text !== undefined
        ? { type: 'contains', value: rule.text }
        : null;
    case 'duplicateValues':
      return { type: 'duplicate' };
    case 'top10':
      return rule.percent === '1'
        ? null
        : {
            type: rule.bottom === '1' ? 'bottom' : 'top',
            count: Number(rule.rank ?? 10),
          };
    case 'colorScale': {
      const scale =
        points.join() === 'min,max' || points.join() === 'min,percentile50,max';
      return scale && colors.every((color) => color !== undefined)
        ? { type: 'scale', colors: colors as string[] }
        : null;
    }
    case 'dataBar':
      return points.join() === 'min,max' && colors[0]
        ? { type: 'bar', color: colors[0] }
        : null;
    default:
      return null;
  }
}

/**
 * The conditional formats of a worksheet that Gridworks has, in the order
 * of their priority; the others are counted as dropped.
 */
function conditionalFormats(
  xml: string,
  { ruleFormat, colorOf }: FileStyles,
  counts: Counts,
): Rule[] {
  const found: [number, Rule][] = [];
  for (const { attributes: group, inner } of elements(
    xml,
    'conditionalFormatting',
  )) {
    const ranges = areasOf(group.sqref ?? '');
    for (const { attributes: rule, inner: body } of elements(inner, 'cfRule')) {
      const condition = conditionOf(rule, body, colorOf);
      if (!condition || ranges.length === 0) {
        counts.conditionalFormat++;
        continue;
      }
      const format =
        condition.type === 'scale' || condition.type === 'bar'
          ? {}
          : ruleFormat(Number(rule.dxfId ?? -1));
      found.push([Number(rule.priority ?? 0), { ranges, condition, format }]);
    }
  }
  return found.sort(([a], [b]) => a - b).map(([, rule]) => rule);
}

const BOUNDS: Record<string, (first: number, second: number) => object> = {
  between: (min, max) => ({ min, max }),
  greaterThanOrEqual: (min) => ({ min }),
  lessThanOrEqual: (max) => ({ max }),
};

const KINDS: Record<string, 'number' | 'date' | 'length'> = {
  decimal: 'number',
  date: 'date',
  textLength: 'length',
};

/** The criterion of a rule of data validation, if Gridworks has it. */
function criterionOf(
  rule: Record<string, string>,
  first: string,
  second: string,
): Criterion | null {
  if (rule.type === 'list') {
    if (first.startsWith('"')) {
      return { type: 'list', source: first.slice(1, -1).replaceAll('""', '"') };
    }
    const type = parse(first)?.type;
    return type === 'range' || type === 'name'
      ? { type: 'list', source: `=${first}` }
      : null;
  }
  const kind = KINDS[rule.type ?? ''];
  const bounds = BOUNDS[rule.operator ?? 'between'];
  const numbers = [first, second].map((text) =>
    /^-?\d+(?:\.\d+)?(?:E[+-]?\d+)?$/i.test(text) ? Number(text) : NaN,
  );
  if (
    !kind ||
    !bounds ||
    Number.isNaN(numbers[0]) ||
    (rule.operator === undefined || rule.operator === 'between'
      ? Number.isNaN(numbers[1])
      : false)
  ) {
    return null;
  }
  return { type: kind, ...bounds(numbers[0]!, numbers[1]!) } as Criterion;
}

/**
 * The rules of data validation of a worksheet that Gridworks has, those
 * Excel keeps in its extension for lists of another sheet last; the others
 * are counted as dropped.
 */
function dataValidations(xml: string, counts: Counts): Validation[] {
  const found: Validation[] = [];
  const add = (
    rule: Record<string, string>,
    first: string,
    second: string,
    sqref: string,
  ) => {
    if (rule.type === undefined || rule.type === 'none') return;
    const criterion = criterionOf(rule, decode(first), decode(second));
    const ranges = areasOf(sqref);
    if (criterion && ranges.length > 0) found.push({ ranges, criterion });
    else counts.validation++;
  };
  for (const { attributes: rule, inner } of elements(xml, 'dataValidation')) {
    add(
      rule,
      inside(inner, 'formula1'),
      inside(inner, 'formula2'),
      rule.sqref ?? '',
    );
  }
  for (const { attributes: rule, inner } of elements(
    xml,
    'x14:dataValidation',
  )) {
    add(
      rule,
      inside(inside(inner, 'x14:formula1'), 'xm:f'),
      inside(inside(inner, 'x14:formula2'), 'xm:f'),
      inside(inner, 'xm:sqref'),
    );
  }
  return found;
}

/** The notes of a comments part, without the name of their author. */
function notesOf(xml: string): [number, number, string][] {
  const authors = elements(inside(xml, 'authors'), 'author').map(({ inner }) =>
    decode(inner),
  );
  const notes: [number, number, string][] = [];
  for (const { attributes: comment, inner } of elements(xml, 'comment')) {
    const area = areaOf(comment.ref ?? '');
    if (!area || area.top >= SHEET_ROWS || area.left >= SHEET_COLUMNS) continue;
    let text = richText(inside(inner, 'text'));
    const author = authors[Number(comment.authorId ?? -1)];
    if (author && text.startsWith(`${author}:`)) {
      text = text.slice(author.length + 1).replace(/^\r?\n/, '');
    }
    notes.push([area.top, area.left, text]);
  }
  return notes;
}

/**
 * The hyperlinks of a worksheet that Gridworks has, on each of their
 * cells: web addresses and places in the workbook. The others are counted
 * as dropped.
 */
function hyperlinksOf(
  xml: string,
  links: Relationship[],
  counts: Counts,
): [number, number, string][] {
  const found: [number, number, string][] = [];
  for (const { attributes: hyperlink } of elements(
    inside(xml, 'hyperlinks'),
    'hyperlink',
  )) {
    const id = hyperlink['r:id'];
    const address = id && links.find((each) => each.id === id)?.target;
    const location = hyperlink.location;
    const link = address
      ? address + (location ? `#${location}` : '')
      : location && `#${location}`;
    const area = areaOf(hyperlink.ref ?? '');
    if (!link || !area || !(link.startsWith('#') || WEB_LINK.test(link))) {
      counts.link++;
      continue;
    }
    eachCell(area, (row, col) => found.push([row, col, link]));
  }
  return found;
}

/** Visits the cells of an area that are inside a sheet. */
function eachCell(area: Area, visit: (row: number, col: number) => void) {
  const bottom = Math.min(area.bottom, SHEET_ROWS - 1);
  const right = Math.min(area.right, SHEET_COLUMNS - 1);
  for (let row = area.top; row <= bottom; row++) {
    for (let col = area.left; col <= right; col++) visit(row, col);
  }
}

/** The value a cell holds, or the last result of its formula. */
function cellValue(
  type: string | undefined,
  contents: string,
  strings: string[],
): Constant {
  if (type === 'inlineStr') {
    const inline = INLINE.exec(contents)?.[1];
    return inline === undefined ? null : richText(inline);
  }
  const text = VALUE.exec(contents)?.[1];
  if (text === undefined) return null;
  switch (type) {
    case 's':
      return strings[Number(text)] ?? null;
    case 'b':
      return text === '1' || text === 'true';
    case 'str':
    case 'e':
      // An error is kept as its text, as Gridworks has no error constants.
      return decode(text);
    default: {
      const number = Number(text);
      return Number.isNaN(number) ? decode(text) : number;
    }
  }
}

/** The shared strings of a workbook, by index. */
function sharedStrings(xml: string): string[] {
  return Array.from(xml.matchAll(/<si>([\s\S]*?)<\/si>|<si\/>/g), (match) =>
    richText(match[1] ?? ''),
  );
}

/** The text of a string that may be in runs, without its phonetic guide. */
function richText(xml: string): string {
  const text = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  return Array.from(text.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (match) =>
    decode(match[1]!),
  ).join('');
}

/** A relationship of a part, its target as a path in the file or an address. */
interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly path: string;
  /** The target as written: the address of an external one. */
  readonly target: string;
}

/** The relationships of a part, with their targets as paths in the file. */
function relationships(xml: string, base: string): Relationship[] {
  return (xml.match(/<Relationship\b[^>]*>/g) ?? []).map((tag) => {
    const { Id = '', Type = '', Target = '' } = attributes(tag);
    return {
      id: Id,
      type: Type.startsWith(RELATIONSHIP)
        ? Type.slice(RELATIONSHIP.length)
        : Type,
      path: resolve(base, Target),
      target: Target,
    };
  });
}

/** The path of a target relative to a directory, or absolute in the file. */
function resolve(base: string, target: string): string {
  const parts: string[] = [];
  const full = target.startsWith('/') ? target.slice(1) : base + target;
  for (const part of full.split('/')) {
    if (part === '..') parts.pop();
    else if (part !== '.') parts.push(part);
  }
  return parts.join('/');
}

const directory = (path: string) => path.slice(0, path.lastIndexOf('/') + 1);

const relationshipsPath = (path: string) =>
  `${directory(path)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
