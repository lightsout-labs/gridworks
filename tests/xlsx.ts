import JSZip from 'jszip';

/** A worksheet as Excel writes it: the XML inside `<sheetData>`. */
export interface SheetPart {
  readonly name: string;
  readonly rows: string;
  /** The XML before `<sheetData>`, such as `<sheetPr>`. */
  readonly before?: string;
  /** The XML after `<sheetData>`, such as `<mergeCells>`. */
  readonly after?: string;
  /** The relationships of the worksheet, to hyperlinks and other parts. */
  readonly relationships?: Relationship[];
}

/** A relationship of a part: its type as the last word of Excel's URI. */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: true;
}

/**
 * Builds an `.xlsx` file laid out as Excel saves one: a workbook part with
 * relationships, shared strings, styles and one part for each worksheet.
 * `parts` adds other files, such as charts.
 */
export async function xlsxFile({
  sheets,
  strings = [],
  numberFormats = [],
  cellFormats = [0],
  styles,
  theme,
  names = '',
  parts = {},
}: {
  sheets: SheetPart[];
  /** Shared strings, as the XML inside each `<si>`. */
  strings?: string[];
  /** Custom number formats, as `[id, code]`. */
  numberFormats?: [number, string][];
  /** The number format identifier of each cell format, by index. */
  cellFormats?: number[];
  /** The XML inside `<styleSheet>`, in place of the formats above. */
  styles?: string;
  /** The theme part, whose colors styles may refer to. */
  theme?: string;
  /** The `<definedNames>` element of the workbook. */
  names?: string;
  parts?: Record<string, string>;
}): Promise<Uint8Array> {
  const zip = new JSZip();
  const sheetEntries = sheets.map((_, index) => index + 1);
  zip.file(
    '[Content_Types].xml',
    xml(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheetEntries
        .map(
          (n) =>
            `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('')}</Types>`,
    ),
  );
  zip.file(
    '_rels/.rels',
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
  );
  zip.file(
    'xl/workbook.xml',
    xml(
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><sheets>${sheets
        .map(
          (sheet, index) =>
            `<sheet name="${escape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 3}"/>`,
        )
        .join('')}</sheets>${names}</workbook>`,
    ),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    xml(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>${sheetEntries
        .map(
          (n) =>
            `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`,
        )
        .join('')}${
        theme
          ? `<Relationship Id="rId${sheets.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`
          : ''
      }</Relationships>`,
    ),
  );
  if (theme) zip.file('xl/theme/theme1.xml', xml(theme));
  zip.file(
    'xl/sharedStrings.xml',
    xml(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings
        .map((string) => `<si>${string}</si>`)
        .join('')}</sst>`,
    ),
  );
  zip.file(
    'xl/styles.xml',
    xml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">${
        styles ??
        `<numFmts count="${numberFormats.length}">${numberFormats
          .map(
            ([id, code]) =>
              `<numFmt numFmtId="${id}" formatCode="${escape(code)}"/>`,
          )
          .join(
            '',
          )}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><cellXfs count="${cellFormats.length}">${cellFormats
          .map(
            (id) =>
              `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ''}/>`,
          )
          .join('')}</cellXfs>`
      }</styleSheet>`,
    ),
  );
  sheets.forEach((sheet, index) => {
    zip.file(
      `xl/worksheets/sheet${index + 1}.xml`,
      xml(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">${sheet.before ?? ''}<sheetData>${sheet.rows}</sheetData>${sheet.after ?? ''}</worksheet>`,
      ),
    );
    if (sheet.relationships) {
      zip.file(
        `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
        xml(
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheet.relationships
            .map(
              ({ id, type, target, external }) =>
                `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${escape(target)}"${external ? ' TargetMode="External"' : ''}/>`,
            )
            .join('')}</Relationships>`,
        ),
      );
    }
  });
  for (const [path, contents] of Object.entries(parts)) {
    zip.file(path, contents);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const xml = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;

const escape = (text: string) =>
  text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

/**
 * The parts that a chart, a pivot table and a picture add to a file saved
 * by Excel.
 */
export const UNSUPPORTED_PARTS: Record<string, string> = {
  'xl/drawings/drawing1.xml': xml(
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
  ),
  'xl/charts/chart1.xml': xml(
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
  ),
  'xl/pivotTables/pivotTable1.xml': xml(
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"/>',
  ),
  'xl/pivotCache/pivotCacheDefinition1.xml': xml(
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
  ),
  'xl/media/image1.png': 'not really a picture',
};

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** The theme of Office 2013–2022, as Excel saves it, cut to its colors. */
const OFFICE_THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

const EXCEL_STYLES =
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>' +
  '<fonts count="5" x14ac:knownFonts="1">' +
  '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
  '<font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>' +
  '<font><i/><strike/><u/><sz val="11"/><color theme="4" tint="-0.249977111117893"/><name val="Georgia"/><family val="1"/></font>' +
  '<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
  '<font><sz val="9"/><color indexed="10"/><name val="Courier New"/><family val="3"/></font>' +
  '</fonts>' +
  '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor theme="9" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
  '<border><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF0070C0"/></right><top style="dashed"><color auto="1"/></top><bottom style="double"><color theme="5"/></bottom><diagonal/></border></borders>' +
  '<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>' +
  '<cellXfs count="7">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
  '<xf numFmtId="164" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="1"/>' +
  '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>' +
  '</cellXfs>' +
  '<cellStyles count="2"><cellStyle name="Hyperlink" xfId="1" builtinId="8"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="3">' +
  '<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>' +
  '<dxf><font><b/><i/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>' +
  '<dxf><font><u/><strike/></font></dxf>' +
  '</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>';

const STRINGS = [
  'Total',
  'Wrapped',
  'Merged',
  'Small',
  'Example',
  'Jump',
  'Red',
  'Green',
  'Blue',
  'hidden',
  'Report',
];

const s = (ref: string, text: string, style = 0) =>
  `<c r="${ref}"${style ? ` s="${style}"` : ''} t="s"><v>${STRINGS.indexOf(text)}</v></c>`;
const n = (ref: string, value: number, style = 0) =>
  `<c r="${ref}"${style ? ` s="${style}"` : ''}><v>${value}</v></c>`;

const cfRule = (attributes: string, body = '') =>
  `<cfRule ${attributes}${body ? `>${body}</cfRule>` : '/>'}`;

/** The worksheet `Formats` of the Excel workbook. */
const FORMATS: SheetPart = {
  name: 'Formats',
  before:
    '<dimension ref="A1:H12"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="B3" sqref="B3"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>' +
    '<cols><col min="1" max="1" width="20.7109375" customWidth="1"/><col min="8" max="8" width="9.140625" hidden="1"/></cols>',
  rows:
    '<row r="1" spans="1:8" ht="18.75" x14ac:dyDescent="0.3">' +
    s('A1', 'Total', 1) +
    n('B1', 0.125, 2) +
    s('C1', 'Wrapped', 3) +
    n('D1', 46340, 4) +
    s('E1', 'Merged') +
    s('G1', 'Report', 6) +
    '</row>' +
    '<row r="2" spans="1:8" ht="30" customHeight="1" x14ac:dyDescent="0.25">' +
    s('A2', 'Example', 5) +
    s('B2', 'Jump', 5) +
    '</row>' +
    `<row r="3" spans="1:8" x14ac:dyDescent="0.25">${n('A3', 5)}<c r="B3"><f>SUM(Rates)</f><v>45</v></c>${n('C3', 1)}${n('D3', 10)}</row>` +
    `<row r="4" spans="1:8" x14ac:dyDescent="0.25">${n('A4', 20)}${n('C4', 2)}${n('D4', 50)}</row>` +
    `<row r="5" spans="1:8" x14ac:dyDescent="0.25">${n('A5', 20)}${n('C5', 3)}${n('D5', 100)}</row>` +
    `<row r="6" spans="1:8" hidden="1" x14ac:dyDescent="0.25">${s('A6', 'hidden')}</row>`,
  after:
    '<mergeCells count="1"><mergeCell ref="E1:F2"/></mergeCells>' +
    '<conditionalFormatting sqref="A3:A5">' +
    cfRule(
      'type="cellIs" dxfId="0" priority="3" operator="greaterThan"',
      '<formula>10</formula>',
    ) +
    cfRule('type="duplicateValues" dxfId="1" priority="2"') +
    cfRule(
      'type="cellIs" dxfId="0" priority="12" operator="greaterThan"',
      '<formula>$B$1</formula>',
    ) +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="A2:B2">' +
    cfRule(
      'type="containsText" dxfId="2" priority="4" operator="containsText" text="Ex"',
      '<formula>NOT(ISERROR(SEARCH("Ex",A2)))</formula>',
    ) +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="C3:C5">' +
    cfRule(
      'type="colorScale" priority="5"',
      '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
    ) +
    cfRule('type="top10" dxfId="1" priority="7" bottom="1" rank="2"') +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="D3:D5">' +
    cfRule(
      'type="dataBar" priority="6"',
      `<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14}"><x14:id>{6A5A1E4C-0F7B-4E4C-9B57-6D2E6F0E2C11}</x14:id></ext></extLst>`,
    ) +
    cfRule('type="top10" dxfId="0" priority="1" rank="1"') +
    cfRule(
      'type="iconSet" priority="11"',
      '<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>',
    ) +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="A3:A5 C3:C5">' +
    cfRule(
      'type="cellIs" dxfId="2" priority="8" operator="between"',
      '<formula>2</formula><formula>6</formula>',
    ) +
    cfRule(
      'type="expression" dxfId="0" priority="13"',
      '<formula>MOD(ROW(),2)=0</formula>',
    ) +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="A1:G1">' +
    cfRule(
      'type="cellIs" dxfId="0" priority="9" operator="lessThan"',
      '<formula>"M"</formula>',
    ) +
    cfRule(
      'type="cellIs" dxfId="1" priority="10" operator="equal"',
      '<formula>46340</formula>',
    ) +
    '</conditionalFormatting>' +
    '<dataValidations count="6">' +
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A10:A12"><formula1>"Yes,No,Maybe"</formula1></dataValidation>' +
    '<dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B10"><formula1>1</formula1><formula2>10</formula2></dataValidation>' +
    '<dataValidation type="date" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C10"><formula1>46023</formula1></dataValidation>' +
    '<dataValidation type="textLength" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D10 D12"><formula1>5</formula1></dataValidation>' +
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E10"><formula1>$A$3:$A$5</formula1></dataValidation>' +
    '<dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F10"><formula1>0</formula1></dataValidation>' +
    '</dataValidations>' +
    '<hyperlinks><hyperlink ref="A2" r:id="rId1" display="https://example.com/" xr:uid="{1B0F7D2A-3C55-4B8B-9E7E-2A1C3B4D5E6F}"/><hyperlink ref="B2" location="\'My lists\'!A1" display="Jump" xr:uid="{2C1E8E3B-4D66-4C9C-8F8F-3B2D4C5E6F70}"/><hyperlink ref="G1" r:id="rId2" xr:uid="{3D2F9F4C-5E77-4DAD-9090-4C3E5D6F7081}"/></hyperlinks>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<legacyDrawing r:id="rId4"/><tableParts count="1"><tablePart r:id="rId5"/></tableParts>' +
    `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{6A5A1E4C-0F7B-4E4C-9B57-6D2E6F0E2C11}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>D3:D5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>` +
    `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" xr:uid="{4E3A0A5D-6F88-4EBE-A1A1-5D4F6E708192}"><x14:formula1><xm:f>'My lists'!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>G10</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`,
  relationships: [
    {
      id: 'rId1',
      type: 'hyperlink',
      target: 'https://example.com/',
      external: true,
    },
    {
      id: 'rId2',
      type: 'hyperlink',
      target: 'file:///C:\\Data\\Report.xlsx',
      external: true,
    },
    { id: 'rId3', type: 'comments', target: '../comments1.xml' },
    { id: 'rId4', type: 'vmlDrawing', target: '../drawings/vmlDrawing1.vml' },
    { id: 'rId5', type: 'table', target: '../tables/table1.xml' },
  ],
};

const NOTE_RUN =
  '<rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>';

/**
 * A workbook as Excel 365 saves one, with each feature of R30 and some
 * that Gridworks does not support: three conditional formats, a rule of
 * data validation, two defined names, a link to a file and a table.
 */
export function excelWorkbook(): Promise<Uint8Array> {
  return xlsxFile({
    strings: STRINGS.map((text) => `<t>${text}</t>`),
    styles: EXCEL_STYLES,
    theme: OFFICE_THEME,
    names:
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Formats!$A$1:$H$12</definedName><definedName name="Colors">\'My lists\'!$A$1:$A$3</definedName><definedName name="Local" localSheetId="1">\'My lists\'!$B$1</definedName><definedName name="Rates">Formats!$A$3:$A$5</definedName><definedName name="Tax">0.2</definedName></definedNames>',
    sheets: [
      FORMATS,
      {
        name: 'My lists',
        before:
          '<sheetPr><tabColor theme="5"/></sheetPr><dimension ref="A1:A3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>',
        rows: ['Red', 'Green', 'Blue']
          .map(
            (text, index) =>
              `<row r="${index + 1}" spans="1:1" x14ac:dyDescent="0.25">${s(`A${index + 1}`, text)}</row>`,
          )
          .join(''),
        after:
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
      },
    ],
    parts: {
      'xl/comments1.xml': xml(
        `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ana Lima</author></authors><commentList><comment ref="A1" authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>Ana Lima:</t></r><r>${NOTE_RUN}<t xml:space="preserve">\nCheck the total</t></r></text></comment><comment ref="D1" authorId="0" shapeId="0"><text><r>${NOTE_RUN}<t>Due &amp; paid</t></r></text></comment></commentList></comments>`,
      ),
      'xl/drawings/vmlDrawing1.vml':
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t202" style="visibility:hidden"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>',
      'xl/tables/table1.xml': xml(
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:D5"/>',
      ),
    },
  });
}
