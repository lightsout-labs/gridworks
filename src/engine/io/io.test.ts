import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  excelWorkbook,
  UNSUPPORTED_PARTS,
  xlsxFile,
} from '../../../tests/xlsx';
import { Workbook, type Snapshot } from '../workbook';
import { readXlsx, writeXlsx } from '.';

/** The value, formula and display text of a cell given by its name. */
function cellOf(workbook: Workbook, sheet: string, name: string) {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(name)!;
  const col = letters!.charCodeAt(0) - 65;
  const { value, formula, display } = workbook.cell(
    sheet,
    Number(digits) - 1,
    col,
  );
  return { value, formula, display };
}

async function open(bytes: Uint8Array) {
  const { snapshot, dropped } = await readXlsx(bytes, 'Budget');
  return { workbook: Workbook.fromSnapshot(snapshot), snapshot, dropped };
}

describe('readXlsx', () => {
  it('reads the sheets of a file in order, with their names and tab colors', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        sheets: [
          { name: 'Income', rows: '' },
          {
            name: 'Q&A 2026',
            before: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
            rows: '',
          },
        ],
      }),
    );
    expect(snapshot.name).toBe('Budget');
    expect(workbook.sheets.map(({ name, color }) => ({ name, color }))).toEqual(
      [
        { name: 'Income', color: undefined },
        { name: 'Q&A 2026', color: '#ff0000' },
      ],
    );
  });

  it('reads numbers, shared and inline strings, booleans and errors', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        strings: [
          '<t>Rent &amp; bills</t>',
          '<r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> part</t></r>',
          '<t>Line_x000D_</t><rPh><t>ignored</t></rPh>',
        ],
        sheets: [
          {
            name: 'Sheet1',
            rows:
              '<row r="1"><c r="A1"><v>12.5</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
              '<row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3" t="inlineStr"><is><t>&lt;inline&gt;</t></is></c><c r="C3" t="e"><v>#N/A</v></c><c r="D3" t="s"><v>2</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').value).toBe(12.5);
    expect(cellOf(workbook, id, 'B1').value).toBe('Rent & bills');
    expect(cellOf(workbook, id, 'C1').value).toBe('Bold part');
    expect(cellOf(workbook, id, 'A3').value).toBe(true);
    expect(cellOf(workbook, id, 'B3').value).toBe('<inline>');
    expect(cellOf(workbook, id, 'C3').display).toBe('#N/A');
    expect(cellOf(workbook, id, 'D3').value).toBe('Line\r');
    expect(cellOf(workbook, id, 'A2').value).toBeNull();
  });

  it('reads number formats, built in and custom', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        numberFormats: [[164, '0.0%']],
        cellFormats: [0, 14, 164],
        sheets: [
          {
            name: 'Sheet1',
            rows: '<row r="1"><c r="A1" s="1"><v>46340</v></c><c r="B1" s="2"><v>0.125</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').display).toBe('11/14/2026');
    expect(cellOf(workbook, id, 'B1').display).toBe('12.5%');
  });

  it('reads formulas, filling in those shared down a column, and calculates them', async () => {
    const { workbook, snapshot, dropped } = await open(
      await xlsxFile({
        sheets: [
          { name: 'Rates', rows: '<row r="1"><c r="A1"><v>2</v></c></row>' },
          {
            name: 'Sheet2',
            rows:
              '<row r="1"><c r="A1"><v>10</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*Rates!$A$1</f><v>20</v></c></row>' +
              '<row r="2"><c r="A2"><v>20</v></c><c r="B2"><f t="shared" si="0"/><v>40</v></c></row>' +
              '<row r="3"><c r="A3"><v>30</v></c><c r="B3"><f t="shared" si="0"/><v>60</v></c><c r="C3"><f>SUM(B1:B3)</f><v>120</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[1]!.id;
    expect(cellOf(workbook, id, 'B3')).toEqual({
      value: 60,
      formula: '=A3*Rates!$A$1',
      display: '60',
    });
    expect(cellOf(workbook, id, 'C3')).toMatchObject({
      value: 120,
      formula: '=SUM(B1:B3)',
    });
    expect(dropped).toEqual([]);
  });

  it('fills in a formula shared across a row, with #REF! for a reference moved off the sheet', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        sheets: [
          {
            name: 'Sheet1',
            rows: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:C1" si="0">A1*2</f><v>2</v></c><c r="C1"><f t="shared" si="0"/><v>4</v></c><c r="CU1"><f t="shared" ref="CU1:CV1" si="1">CV1+1</f><v>1</v></c><c r="CV1"><f t="shared" si="1"/><v>0</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'C1')).toMatchObject({
      value: 4,
      formula: '=B1*2',
    });
    expect(workbook.cell(id, 0, 99)).toMatchObject({
      formula: '=#REF!+1',
      display: '#REF!',
    });
  });

  it('opens a file with charts, pivot tables, pictures and unknown formulas, and lists what it dropped', async () => {
    const { workbook, snapshot, dropped } = await open(
      await xlsxFile({
        strings: ['<t>old text</t>'],
        sheets: [
          {
            name: 'Sheet1',
            rows:
              '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>OFFSET(A1,1,1)</f><v>9</v></c><c r="C1" t="str"><f>[1]Other!A1</f><v>linked</v></c></row>' +
              '<row r="2"><c r="A2"><f>NOSUCHFUNCTION(A1)</f><v>#NAME?</v></c><c r="DA2"><v>1</v></c></row>' +
              '<row r="100001"><c r="A100001"><v>1</v></c></row>',
          },
        ],
        parts: UNSUPPORTED_PARTS,
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').value).toBe(3);
    // The last result is kept in place of a formula that cannot be calculated.
    expect(cellOf(workbook, id, 'B1')).toMatchObject({
      value: 9,
      formula: null,
    });
    expect(cellOf(workbook, id, 'C1').value).toBe('linked');
    expect(dropped).toEqual([
      { feature: 'chart', count: 1 },
      { feature: 'pivotTable', count: 1 },
      { feature: 'picture', count: 1 },
      { feature: 'formula', count: 3 },
      { feature: 'cell', count: 2 },
    ]);
  });

  it('rejects a file that is not an .xlsx file', async () => {
    await expect(
      readXlsx(new TextEncoder().encode('just text'), 'notes'),
    ).rejects.toThrow();
    const zip = new JSZip();
    zip.file('readme.txt', 'not a workbook');
    await expect(
      readXlsx(await zip.generateAsync({ type: 'uint8array' }), 'archive'),
    ).rejects.toThrow();
    // A Word document renamed to .xlsx.
    const word = new JSZip();
    word.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    word.file(
      'word/document.xml',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
    );
    await expect(
      readXlsx(await word.generateAsync({ type: 'uint8array' }), 'letter'),
    ).rejects.toThrow();
  });
});

describe('readXlsx, of a workbook saved by Excel', () => {
  const opened = async () => {
    const result = await open(await excelWorkbook());
    const [formats, lists] = result.snapshot.sheets.map(({ id }) => id);
    return { ...result, formats: formats!, lists: lists! };
  };

  it('reads fonts, colors, fills, alignment, wrap, borders and number formats', async () => {
    const { workbook, formats, lists } = await opened();
    const style = (name: string) => {
      const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(name)!;
      return workbook.cell(
        formats,
        Number(digits) - 1,
        letters!.charCodeAt(0) - 65,
      ).style;
    };
    expect(style('A1')).toEqual({
      font: 'Arial',
      size: 14,
      bold: true,
      color: '#ff0000',
      fill: '#ffff00',
    });
    // Theme colors, lightened or darkened by their tint.
    expect(style('B1')).toEqual({
      numberFormat: '0.0%',
      font: 'Georgia',
      italic: true,
      underline: true,
      strike: true,
      color: '#2f5597',
      fill: '#e2efda',
    });
    // Automatic colors are left to the theme.
    expect(style('C1')).toEqual({
      horizontal: 'center',
      vertical: 'middle',
      wrap: true,
      borderTop: { line: 'dashed' },
      borderRight: { line: 'medium', color: '#0070c0' },
      borderBottom: { line: 'double', color: '#ed7d31' },
      borderLeft: { line: 'thin' },
    });
    expect(cellOf(workbook, formats, 'D1').display).toBe('11/14/2026');
    expect(style('G1')).toEqual({
      font: 'Courier New',
      size: 9,
      color: '#ff0000',
      horizontal: 'right',
      vertical: 'top',
    });
    // Excel's default font, and its Hyperlink style, are not kept.
    expect(style('A2')).toEqual({});
    expect(style('E1')).toEqual({});
    expect(workbook.sheets[1]).toMatchObject({ id: lists, color: '#ed7d31' });
  });

  it('reads column widths, row heights, hidden rows and columns, merges and frozen panes', async () => {
    const { snapshot, workbook, formats } = await opened();
    // Excel's standard 64 by 20 pixels are Gridworks' 96 by 24.
    expect(snapshot.sheets[0]).toMatchObject({
      columnWidths: [[0, 225]],
      rowHeights: [[1, 48]],
      hiddenColumns: [7],
      hiddenRows: [5],
      merges: [{ top: 0, left: 4, bottom: 1, right: 5 }],
      frozen: { rows: 1, cols: 1 },
    });
    expect(workbook.layout(formats).rows.hidden).toEqual(new Set([5]));
  });

  it('reads conditional formats in their order, data validation and named ranges, and calculates formulas that use names', async () => {
    const { snapshot, workbook, formats } = await opened();
    const area = (top: number, left: number, bottom = top, right = left) => ({
      top,
      left,
      bottom,
      right,
    });
    const red = { color: '#9c0006', fill: '#ffc7ce' };
    const green = {
      bold: true,
      italic: true,
      color: '#006100',
      fill: '#c6efce',
    };
    expect(snapshot.sheets[0]!.conditionalFormats).toEqual([
      {
        ranges: [area(2, 3, 4)],
        condition: { type: 'top', count: 1 },
        format: red,
      },
      {
        ranges: [area(2, 0, 4)],
        condition: { type: 'duplicate' },
        format: green,
      },
      {
        ranges: [area(2, 0, 4)],
        condition: { type: 'greater', value: '10' },
        format: red,
      },
      {
        ranges: [area(1, 0, 1, 1)],
        condition: { type: 'contains', value: 'Ex' },
        format: { underline: true, strike: true },
      },
      {
        ranges: [area(2, 2, 4)],
        condition: { type: 'scale', colors: ['#f8696b', '#ffeb84', '#63be7b'] },
        format: {},
      },
      {
        ranges: [area(2, 3, 4)],
        condition: { type: 'bar', color: '#638ec6' },
        format: {},
      },
      {
        ranges: [area(2, 2, 4)],
        condition: { type: 'bottom', count: 2 },
        format: green,
      },
      {
        ranges: [area(2, 0, 4), area(2, 2, 4)],
        condition: { type: 'between', value: '2', other: '6' },
        format: { underline: true, strike: true },
      },
      {
        ranges: [area(0, 0, 0, 6)],
        condition: { type: 'less', value: 'M' },
        format: red,
      },
      {
        ranges: [area(0, 0, 0, 6)],
        condition: { type: 'equal', value: '46340' },
        format: green,
      },
    ]);
    expect(snapshot.sheets[0]!.validations).toEqual([
      {
        ranges: [area(9, 0, 11, 0)],
        criterion: { type: 'list', source: 'Yes,No,Maybe' },
      },
      { ranges: [area(9, 1)], criterion: { type: 'number', min: 1, max: 10 } },
      { ranges: [area(9, 2)], criterion: { type: 'date', min: 46023 } },
      {
        ranges: [area(9, 3), area(11, 3)],
        criterion: { type: 'length', max: 5 },
      },
      {
        ranges: [area(9, 4)],
        criterion: { type: 'list', source: '=$A$3:$A$5' },
      },
      {
        ranges: [area(9, 6)],
        criterion: { type: 'list', source: "='My lists'!$A$1:$A$3" },
      },
    ]);
    expect(workbook.names).toEqual([
      { name: 'Colors', refersTo: "'My lists'!$A$1:$A$3" },
      { name: 'Rates', refersTo: 'Formats!$A$3:$A$5' },
    ]);
    expect(cellOf(workbook, formats, 'B3')).toEqual({
      value: 45,
      formula: '=SUM(Rates)',
      display: '45',
    });
  });

  it('reads notes without their author, and hyperlinks to addresses and to places in the workbook', async () => {
    const { workbook, formats } = await opened();
    expect(workbook.cell(formats, 0, 0).note).toBe('Check the total');
    expect(workbook.cell(formats, 0, 3).note).toBe('Due & paid');
    expect(workbook.cell(formats, 1, 0).link).toBe('https://example.com/');
    expect(workbook.cell(formats, 1, 1).link).toBe("#'My lists'!A1");
    // A link to a file is dropped.
    expect(workbook.cell(formats, 0, 6).link).toBeUndefined();
  });

  it('lists what it dropped', async () => {
    const { dropped } = await opened();
    expect(dropped).toEqual([
      { feature: 'table', count: 1 },
      { feature: 'conditionalFormat', count: 3 },
      { feature: 'validation', count: 1 },
      { feature: 'name', count: 2 },
      { feature: 'link', count: 1 },
    ]);
  });
});

describe('writeXlsx', () => {
  it('writes a file that reads back as the same sheets, values, formulas and formats', async () => {
    const snapshot: Snapshot = {
      version: 1,
      name: 'Round trip',
      styles: [{}, { numberFormat: 'm/d/yyyy' }, { numberFormat: '0.0%' }],
      sheets: [
        {
          id: 'sheet1',
          name: 'Data',
          color: '#00aa55',
          columns: [
            [1, 'two', true, 46340, null],
            [null, 0.5, 'Q&A <b>'],
          ],
          formulas: [
            [0, 1, '=A1+Totals!A1'],
            [3, 1, '=A1&" items"'],
          ],
          styleIds: [
            [0, 0, 0, 1],
            [0, 2],
          ],
        },
        {
          id: 'sheet2',
          name: 'Totals',
          columns: [[5]],
          formulas: [],
          styleIds: [],
        },
      ],
    };
    const before = Workbook.fromSnapshot(snapshot);
    const { snapshot: read, dropped } = await readXlsx(
      await writeXlsx(snapshot),
      'Round trip',
    );
    const after = Workbook.fromSnapshot(read);
    expect(dropped).toEqual([]);
    expect(after.sheets).toEqual(before.sheets);
    for (const sheet of before.sheets) {
      for (let row = 0; row < 5; row++) {
        for (let col = 0; col < 2; col++) {
          expect(after.cell(sheet.id, row, col)).toEqual(
            before.cell(sheet.id, row, col),
          );
        }
      }
    }
  });

  it('writes newer functions with the prefix Excel expects, and reads them back', async () => {
    const snapshot: Snapshot = {
      version: 1,
      name: 'Prefixes',
      sheets: [
        {
          id: 'sheet1',
          name: 'Sheet1',
          columns: [[1, 3]],
          formulas: [[0, 1, '=IFS(A1>0,STDEV.S(A1:A2))+SUM(A1:A2)']],
        },
      ],
    };
    const bytes = await writeXlsx(snapshot);
    const zip = await JSZip.loadAsync(bytes);
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toContain(
      '_xlfn.IFS(A1&gt;0,_xlfn.STDEV.S(A1:A2))+SUM(A1:A2)',
    );
    const { workbook, snapshot: read } = await open(bytes);
    expect(cellOf(workbook, read.sheets[0]!.id, 'B1')).toMatchObject({
      value: Math.SQRT2 + 4,
      formula: '=IFS(A1>0,STDEV.S(A1:A2))+SUM(A1:A2)',
    });
  });

  it('writes every feature of a workbook opened from Excel so that it reads back the same', async () => {
    const { snapshot } = await readXlsx(await excelWorkbook(), 'Excel');
    const { snapshot: again, dropped } = await readXlsx(
      await writeXlsx(snapshot),
      'Excel',
    );
    expect(dropped).toEqual([]);
    expect(again).toEqual(snapshot);
  });

  it('writes a package whose every part has a content type and every relationship a target', async () => {
    const { snapshot } = await readXlsx(await excelWorkbook(), 'Excel');
    const zip = await JSZip.loadAsync(await writeXlsx(snapshot));
    const types = await zip.file('[Content_Types].xml')!.async('string');
    const extensions = [...types.matchAll(/Extension="(\w+)"/g)].map(
      ([, extension]) => extension,
    );
    for (const path of Object.keys(zip.files)) {
      if (zip.files[path]!.dir) continue;
      const declared =
        types.includes(`PartName="/${path}"`) ||
        extensions.includes(path.slice(path.lastIndexOf('.') + 1));
      expect(declared, path).toBe(true);
    }
    for (const path of Object.keys(zip.files).filter((each) =>
      each.endsWith('.rels'),
    )) {
      const base = path.replace(/_rels\/([^/]*)\.rels$/, '');
      const rels = await zip.file(path)!.async('string');
      for (const [tag] of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        if (tag.includes('TargetMode="External"')) continue;
        const target = /Target="([^"]*)"/.exec(tag)![1]!;
        const resolved = new URL(target, `file:///${base}`).pathname.slice(1);
        expect(zip.file(resolved), `${path}: ${target}`).not.toBeNull();
      }
    }
  });
});
