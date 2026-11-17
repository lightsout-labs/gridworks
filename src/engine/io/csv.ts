import {
  formatValue,
  GENERAL,
  isDateFormat,
  parseInput,
  type Input,
} from '../format';
import { SHEET_COLUMNS, SHEET_ROWS } from '../formula';
import {
  createWorkbook,
  type SheetSnapshot,
  type Style,
  type Workbook,
} from '../workbook';
import type { Opened } from '.';

type Constant = SheetSnapshot['columns'][number][number];

const BYTE_ORDER_MARK = '\uFEFF';

/**
 * Reads the text of a `.csv` file as a workbook named `name`, with one sheet
 * named after it too. Each field is recognized as if it were typed, a
 * number, a date, a percentage or a boolean taking the format it implies,
 * and an error such as `#N/A` kept as text; a field starting with
 * `=` stays text, so that a file opened runs no formula. Cells beyond the
 * rows and columns of a sheet are left out and counted.
 */
export function readCsv(text: string, name: string): Opened {
  const columns: Constant[][] = [];
  const styleIds: number[][] = [];
  const styles: Style[] = [{}];
  const formats = new Map<string, number>();
  let beyond = 0;
  const year = new Date().getFullYear();
  records(
    text.startsWith(BYTE_ORDER_MARK) ? text.slice(1) : text,
    (row, fields) =>
      fields.forEach((field, col) => {
        if (field === '') return;
        if (row >= SHEET_ROWS || col >= SHEET_COLUMNS) {
          beyond++;
          return;
        }
        const { value, format } = recognized(field, year);
        (columns[col] ??= [])[row] = value;
        if (format === undefined) return;
        let id = formats.get(format);
        if (id === undefined) {
          id = styles.push({ numberFormat: format }) - 1;
          formats.set(format, id);
        }
        (styleIds[col] ??= [])[row] = id;
      }),
  );
  // Columns and rows skipped are holes that a snapshot fills.
  const dense = <T>(list: T[][], empty: T): T[][] =>
    Array.from(list, (column) =>
      Array.from(column ?? [], (each) => each ?? empty),
    );
  const sheetName = createWorkbook().sheetNameProblem(name, '')
    ? 'Sheet1'
    : name;
  return {
    snapshot: {
      version: 1,
      name,
      sheets: [
        {
          id: 'sheet1',
          name: sheetName,
          columns: dense<Constant>(columns, null),
          formulas: [],
          styleIds: dense(styleIds, 0),
        },
      ],
      styles,
    },
    dropped: beyond > 0 ? [{ feature: 'cell', count: beyond }] : [],
  };
}

/**
 * A field as a value, with the format its text implies. An error such as
 * `#N/A` stays text, as in an `.xlsx` file, since Gridworks has no error
 * constants.
 */
function recognized(field: string, year: number): Input {
  if (field.startsWith('=')) return { value: field };
  return parseInput(field, year);
}

/**
 * Calls `record` with the fields of each line of a CSV text, as RFC 4180
 * has them and Excel writes them: quoted fields may hold commas, line
 * breaks and doubled quotes, and lines end with CRLF, LF or CR. A quote
 * inside an unquoted field is kept as it is.
 */
function records(
  text: string,
  record: (row: number, fields: string[]) => void,
): void {
  let row = 0;
  let fields: string[] = [];
  let index = 0;
  const end = text.length;
  while (index < end) {
    let field = '';
    if (text[index] === '"') {
      index++;
      for (;;) {
        const quote = text.indexOf('"', index);
        if (quote < 0) {
          field += text.slice(index);
          index = end;
          break;
        }
        field += text.slice(index, quote);
        index = quote + 1;
        if (text[index] !== '"') break;
        field += '"';
        index++;
      }
    }
    // What follows a closing quote, or an unquoted field, up to its end.
    let stop = index;
    while (stop < end && !',\r\n'.includes(text[stop]!)) stop++;
    field += text.slice(index, stop);
    fields.push(field);
    index = stop;
    if (text[index] === ',') {
      index++;
      if (index === end) fields.push('');
      continue;
    }
    if (text[index] === '\r') index++;
    if (text[index] === '\n') index++;
    record(row++, fields);
    fields = [];
  }
}

/**
 * The text of a `.csv` file of a sheet's cells in use, starting with a
 * byte order mark so that Excel reads it as UTF-8, with lines ended by
 * CRLF. A formula is written as its result, a number in full whatever its
 * format, and a date or time as text that reads back as the same value.
 */
export function writeCsv(workbook: Workbook, sheet: string): string {
  const { rows, cols } = workbook.extent(sheet);
  let text = BYTE_ORDER_MARK;
  for (let row = 0; row < rows; row++) {
    const fields: string[] = [];
    for (let col = 0; col < cols; col++) {
      const { value, style } = workbook.cell(sheet, row, col);
      fields.push(quoted(written(value, style.numberFormat ?? GENERAL)));
    }
    text += `${fields.join(',')}\r\n`;
  }
  return text;
}

/** A value as the text of a field. */
function written(value: ReturnType<Workbook['cell']>['value'], format: string) {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'object') return value.error;
  if (typeof value === 'string') return value;
  if (!isDateFormat(format) || value < 0) return String(value);
  const time = value % 1 !== 0;
  if (value < 1) return formatValue(value, 'hh:mm:ss');
  return formatValue(value, time ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd');
}

/** A field quoted when it holds a comma, a quote or a line break. */
function quoted(field: string): string {
  return /[",\r\n]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field;
}
