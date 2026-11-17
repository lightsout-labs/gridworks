import JSZip from 'jszip';
import { parseInput } from '../format';
import { columnName, parse, print, type Area } from '../formula';
import type {
  Rule,
  RuleFormat,
  SheetSnapshot,
  Snapshot,
  Validation,
} from '../workbook';
import { writeStyles } from './styles';
import { areaName, encode } from './xml';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const CONTENT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
/** The extension of a rule that has more in Excel 2010's extension list. */
const RULE_EXTENSION = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';
const SHEET_EXTENSION = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';

const xml = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;

/** Gridworks' pixels as Excel's widths, in characters of 7 pixels. */
function columnWidth(pixels: number): number {
  const excel = pixels / 1.5;
  return Number((excel < 12 ? excel / 12 : (excel - 5) / 7).toFixed(6));
}

/** Gridworks' pixels as Excel's heights, in points. */
const rowHeight = (pixels: number) => Number((pixels / 1.6).toFixed(4));

/** The relationships of a part, as `[id, type, target, external]`. */
function relationshipsPart(
  relationships: [string, string, string, boolean?][],
): string {
  return xml(
    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">${relationships
      .map(
        ([id, type, target, external]) =>
          `<Relationship Id="${id}" Type="${RELATIONSHIPS}/${type}" Target="${encode(target)}"${
            external ? ' TargetMode="External"' : ''
          }/>`,
      )
      .join('')}</Relationships>`,
  );
}

/**
 * Writes a workbook as an `.xlsx` file. Formulas are written without their
 * results, and the file asks to be calculated when it is opened.
 */
export async function writeXlsx(snapshot: Snapshot): Promise<Uint8Array> {
  const zip = new JSZip();
  const strings = new Map<string, number>();
  const formats: RuleFormat[] = [];
  const formatIndex = new Map<string, number>();
  const shared: Shared = {
    string: (text) => {
      let index = strings.get(text);
      if (index === undefined) {
        index = strings.size;
        strings.set(text, index);
      }
      return index;
    },
    format: (format) => {
      const key = JSON.stringify(format);
      let index = formatIndex.get(key);
      if (index === undefined) {
        index = formats.push(format) - 1;
        formatIndex.set(key, index);
      }
      return index;
    },
  };
  const overrides: [string, string][] = [
    ['/xl/workbook.xml', `${CONTENT}.sheet.main+xml`],
    ['/xl/styles.xml', `${CONTENT}.styles+xml`],
    ['/xl/sharedStrings.xml', `${CONTENT}.sharedStrings+xml`],
  ];
  let notes = 0;
  snapshot.sheets.forEach((sheet, index) => {
    const number = index + 1;
    const path = `xl/worksheets/sheet${number}.xml`;
    overrides.push([`/${path}`, `${CONTENT}.worksheet+xml`]);
    const relationships: [string, string, string, boolean?][] = [];
    const hasNotes = (sheet.notes?.length ?? 0) > 0;
    if (hasNotes) {
      notes++;
      relationships.push(
        ['rId1', 'comments', `../comments${notes}.xml`],
        ['rId2', 'vmlDrawing', `../drawings/vmlDrawing${notes}.vml`],
      );
      zip.file(`xl/comments${notes}.xml`, commentsPart(sheet));
      zip.file(`xl/drawings/vmlDrawing${notes}.vml`, vmlDrawing(sheet, notes));
      overrides.push([`/xl/comments${notes}.xml`, `${CONTENT}.comments+xml`]);
    }
    zip.file(path, worksheet(sheet, index, shared, relationships, hasNotes));
    if (relationships.length > 0) {
      zip.file(
        `xl/worksheets/_rels/sheet${number}.xml.rels`,
        relationshipsPart(relationships),
      );
    }
  });

  zip.file(
    '[Content_Types].xml',
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>${overrides
        .map(
          ([part, type]) =>
            `<Override PartName="${part}" ContentType="${type}"/>`,
        )
        .join('')}</Types>`,
    ),
  );
  zip.file(
    '_rels/.rels',
    relationshipsPart([['rId1', 'officeDocument', 'xl/workbook.xml']]),
  );
  const names = snapshot.names ?? [];
  zip.file(
    'xl/workbook.xml',
    xml(
      `<workbook xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}"><workbookPr/><bookViews><workbookView activeTab="0"/></bookViews><sheets>${snapshot.sheets
        .map(
          (sheet, index) =>
            `<sheet name="${encode(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 3}"/>`,
        )
        .join('')}</sheets>${
        names.length > 0
          ? `<definedNames>${names
              .map(
                ({ name, refersTo }) =>
                  `<definedName name="${encode(name)}">${encode(refersTo)}</definedName>`,
              )
              .join('')}</definedNames>`
          : ''
      }<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    ),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    relationshipsPart([
      ['rId1', 'styles', 'styles.xml'],
      ['rId2', 'sharedStrings', 'sharedStrings.xml'],
      ...snapshot.sheets.map((_, index): [string, string, string] => [
        `rId${index + 3}`,
        'worksheet',
        `worksheets/sheet${index + 1}.xml`,
      ]),
    ]),
  );
  zip.file('xl/styles.xml', xml(writeStyles(snapshot.styles ?? [{}], formats)));
  zip.file(
    'xl/sharedStrings.xml',
    xml(
      `<sst xmlns="${MAIN}" count="${strings.size}" uniqueCount="${strings.size}">${Array.from(
        strings.keys(),
        (text) => `<si><t xml:space="preserve">${encode(text)}</t></si>`,
      ).join('')}</sst>`,
    ),
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** The shared strings and the differential formats of the whole file. */
interface Shared {
  /** The index of a text among the shared strings. */
  readonly string: (text: string) => number;
  /** The index of a conditional format's format among the dxfs. */
  readonly format: (format: RuleFormat) => number;
}

/** What a cell of a worksheet holds. */
interface CellXml {
  value?: SheetSnapshot['columns'][number][number];
  formula?: string;
  style?: number;
}

/** The worksheet part of a sheet. */
function worksheet(
  sheet: SheetSnapshot,
  index: number,
  shared: Shared,
  relationships: [string, string, string, boolean?][],
  hasNotes: boolean,
): string {
  const rows = new Map<number, Map<number, CellXml>>();
  const cell = (row: number, col: number) => {
    let cells = rows.get(row);
    if (!cells) rows.set(row, (cells = new Map()));
    let found = cells.get(col);
    if (!found) cells.set(col, (found = {}));
    return found;
  };
  sheet.columns.forEach((column, col) =>
    column.forEach((value, row) => {
      if (value !== null) cell(row, col).value = value;
    }),
  );
  for (const [row, col, text] of sheet.formulas) {
    const ast = parse(text.slice(1));
    cell(row, col).formula = ast ? print(ast, { file: true }) : text.slice(1);
  }
  sheet.styleIds?.forEach((column, col) =>
    column.forEach((id, row) => {
      if (id) cell(row, col).style = id;
    }),
  );
  const heights = new Map(sheet.rowHeights ?? []);
  const hiddenRows = new Set(sheet.hiddenRows ?? []);
  for (const row of [...heights.keys(), ...hiddenRows]) {
    if (!rows.has(row)) rows.set(row, new Map());
  }

  const sheetData = [...rows.keys()]
    .sort((a, b) => a - b)
    .map((row) => {
      const height = heights.get(row);
      const cells = [...rows.get(row)!]
        .sort(([a], [b]) => a - b)
        .map(([col, { value, formula, style }]) => {
          const reference = `r="${columnName(col)}${row + 1}"${style ? ` s="${style}"` : ''}`;
          if (formula !== undefined) {
            return `<c ${reference}><f>${encode(formula)}</f></c>`;
          }
          if (typeof value === 'string') {
            return `<c ${reference} t="s"><v>${shared.string(value)}</v></c>`;
          }
          if (typeof value === 'boolean') {
            return `<c ${reference} t="b"><v>${value ? 1 : 0}</v></c>`;
          }
          return typeof value === 'number'
            ? `<c ${reference}><v>${value}</v></c>`
            : `<c ${reference}/>`;
        })
        .join('');
      return `<row r="${row + 1}"${
        height === undefined
          ? ''
          : ` ht="${rowHeight(height)}" customHeight="1"`
      }${hiddenRows.has(row) ? ' hidden="1"' : ''}>${cells}</row>`;
    })
    .join('');

  const widths = new Map(sheet.columnWidths ?? []);
  const hiddenColumns = new Set(sheet.hiddenColumns ?? []);
  const cols = [...new Set([...widths.keys(), ...hiddenColumns])]
    .sort((a, b) => a - b)
    .map((col) => {
      const width = widths.get(col);
      return `<col min="${col + 1}" max="${col + 1}" width="${
        width === undefined ? 9.140625 : columnWidth(width)
      }"${width === undefined ? '' : ' customWidth="1"'}${
        hiddenColumns.has(col) ? ' hidden="1"' : ''
      }/>`;
    })
    .join('');

  const { rows: frozenRows = 0, cols: frozenCols = 0 } = sheet.frozen ?? {};
  const pane =
    frozenRows > 0 || frozenCols > 0
      ? `<pane${frozenCols > 0 ? ` xSplit="${frozenCols}"` : ''}${
          frozenRows > 0 ? ` ySplit="${frozenRows}"` : ''
        } topLeftCell="${columnName(frozenCols)}${frozenRows + 1}" activePane="${
          frozenRows > 0 && frozenCols > 0
            ? 'bottomRight'
            : frozenRows > 0
              ? 'bottomLeft'
              : 'topRight'
        }" state="frozen"/>`
      : '';

  const links = (sheet.links ?? [])
    .map(([row, col, link]) => {
      const ref = `ref="${columnName(col)}${row + 1}"`;
      if (link.startsWith('#')) {
        return `<hyperlink ${ref} location="${encode(link.slice(1))}"/>`;
      }
      const id = `rId${relationships.length + 1}`;
      relationships.push([id, 'hyperlink', link, true]);
      return `<hyperlink ${ref} r:id="${id}"/>`;
    })
    .join('');

  const bars: string[] = [];
  const rules = (sheet.conditionalFormats ?? [])
    .map((rule, at) => {
      const sqref = rule.ranges.map(areaName).join(' ');
      const id = `{00000000-0000-4000-8000-${(index * 10_000 + at)
        .toString(16)
        .padStart(12, '0')}}`;
      if (rule.condition.type === 'bar') {
        bars.push(
          `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`,
        );
      }
      return `<conditionalFormatting sqref="${sqref}">${ruleXml(rule, at + 1, shared, id)}</conditionalFormatting>`;
    })
    .join('');

  const validations = sheet.validations ?? [];
  return xml(
    `<worksheet xmlns="${MAIN}" xmlns:r="${RELATIONSHIPS}">` +
      (sheet.color
        ? `<sheetPr><tabColor rgb="FF${sheet.color.slice(1).toUpperCase()}"/></sheetPr>`
        : '') +
      `<sheetViews><sheetView${index === 0 ? ' tabSelected="1"' : ''} workbookViewId="0">${pane}</sheetView></sheetViews>` +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      (cols ? `<cols>${cols}</cols>` : '') +
      `<sheetData>${sheetData}</sheetData>` +
      ((sheet.merges?.length ?? 0) > 0
        ? `<mergeCells count="${sheet.merges!.length}">${sheet
            .merges!.map((merge) => `<mergeCell ref="${areaName(merge)}"/>`)
            .join('')}</mergeCells>`
        : '') +
      rules +
      (validations.length > 0
        ? `<dataValidations count="${validations.length}">${validations
            .map(validationXml)
            .join('')}</dataValidations>`
        : '') +
      (links ? `<hyperlinks>${links}</hyperlinks>` : '') +
      (hasNotes ? '<legacyDrawing r:id="rId2"/>' : '') +
      (bars.length > 0
        ? `<extLst><ext uri="${SHEET_EXTENSION}" xmlns:x14="${X14}"><x14:conditionalFormattings>${bars.join('')}</x14:conditionalFormattings></ext></extLst>`
        : '') +
      '</worksheet>',
  );
}

/** A condition's value as a formula: a number, a boolean or a text. */
function constant(text: string): string {
  const { value } = parseInput(text);
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return `"${text.replaceAll('"', '""')}"`;
}

const OPERATORS = { greater: 'greaterThan', less: 'lessThan', equal: 'equal' };

/** A conditional format as a `cfRule` of priority `priority`. */
function ruleXml(
  rule: Rule,
  priority: number,
  shared: Shared,
  id: string,
): string {
  const { condition } = rule;
  const dxf = () => ` dxfId="${shared.format(rule.format)}"`;
  const formula = (text: string) => `<formula>${encode(text)}</formula>`;
  switch (condition.type) {
    case 'greater':
    case 'less':
    case 'equal':
      return `<cfRule type="cellIs"${dxf()} priority="${priority}" operator="${OPERATORS[condition.type]}">${formula(constant(condition.value))}</cfRule>`;
    case 'between':
      return `<cfRule type="cellIs"${dxf()} priority="${priority}" operator="between">${formula(constant(condition.value))}${formula(constant(condition.other))}</cfRule>`;
    case 'contains': {
      const first = rule.ranges[0]!;
      const text = `"${condition.value.replaceAll('"', '""')}"`;
      return `<cfRule type="containsText"${dxf()} priority="${priority}" operator="containsText" text="${encode(condition.value)}">${formula(
        `NOT(ISERROR(SEARCH(${text},${columnName(first.left)}${first.top + 1})))`,
      )}</cfRule>`;
    }
    case 'duplicate':
      return `<cfRule type="duplicateValues"${dxf()} priority="${priority}"/>`;
    case 'top':
    case 'bottom':
      return `<cfRule type="top10"${dxf()} priority="${priority}"${
        condition.type === 'bottom' ? ' bottom="1"' : ''
      } rank="${condition.count}"/>`;
    case 'scale':
      return `<cfRule type="colorScale" priority="${priority}"><colorScale><cfvo type="min"/>${
        condition.colors.length === 3
          ? '<cfvo type="percentile" val="50"/>'
          : ''
      }<cfvo type="max"/>${condition.colors
        .map((color) => `<color rgb="FF${color.slice(1).toUpperCase()}"/>`)
        .join('')}</colorScale></cfRule>`;
    case 'bar':
      return `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF${condition.color
        .slice(1)
        .toUpperCase()}"/></dataBar><extLst><ext uri="${RULE_EXTENSION}" xmlns:x14="${X14}"><x14:id>${id}</x14:id></ext></extLst></cfRule>`;
  }
}

const VALIDATION_TYPES = {
  number: 'decimal',
  date: 'date',
  length: 'textLength',
};

/** A rule of data validation as a `dataValidation`. */
function validationXml({ ranges, criterion }: Validation): string {
  const sqref = ranges.map((area: Area) => areaName(area)).join(' ');
  const start = `<dataValidation type="${
    criterion.type === 'list' ? 'list' : VALIDATION_TYPES[criterion.type]
  }"`;
  const end = `allowBlank="1" showErrorMessage="1" sqref="${sqref}">`;
  if (criterion.type === 'list') {
    const { source } = criterion;
    const first = source.startsWith('=')
      ? source.slice(1)
      : `"${source.replaceAll('"', '""')}"`;
    return `${start} ${end}<formula1>${encode(first)}</formula1></dataValidation>`;
  }
  const { min, max } = criterion;
  if (min !== undefined && max !== undefined) {
    return `${start} operator="between" ${end}<formula1>${min}</formula1><formula2>${max}</formula2></dataValidation>`;
  }
  return min !== undefined
    ? `${start} operator="greaterThanOrEqual" ${end}<formula1>${min}</formula1></dataValidation>`
    : `${start} operator="lessThanOrEqual" ${end}<formula1>${max}</formula1></dataValidation>`;
}

/** The comments part of a sheet's notes. */
function commentsPart(sheet: SheetSnapshot): string {
  return xml(
    `<comments xmlns="${MAIN}"><authors><author></author></authors><commentList>${(
      sheet.notes ?? []
    )
      .map(
        ([row, col, text]) =>
          `<comment ref="${columnName(col)}${row + 1}" authorId="0"><text><t xml:space="preserve">${encode(text)}</t></text></comment>`,
      )
      .join('')}</commentList></comments>`,
  );
}

/** The drawing that Excel needs to show a sheet's notes, hidden until hovered. */
function vmlDrawing(sheet: SheetSnapshot, number: number): string {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${number}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${(
    sheet.notes ?? []
  )
    .map(
      ([row, col], at) =>
        `<v:shape id="_x0000_s${number * 1024 + at + 1}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${at + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1}, 15, ${row}, 2, ${col + 3}, 15, ${row + 4}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`,
    )
    .join('')}</xml>`;
}
