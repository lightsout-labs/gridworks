import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { excelWorkbook, UNSUPPORTED_PARTS, xlsxFile } from '../xlsx';
import {
  cellBackground,
  clickCell,
  enter,
  expectCell,
  openApp,
  openFile,
  XLSX_TYPE,
} from './grid';

const sheetTabs = (page: Page) =>
  page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab');

const budget = () =>
  xlsxFile({
    strings: ['<t>Rent</t>'],
    sheets: [
      {
        name: 'Income',
        rows:
          '<row r="1"><c r="A1"><v>5</v></c></row>' +
          '<row r="2"><c r="A2"><f>A1*2</f><v>10</v></c></row>',
      },
      { name: 'Costs', rows: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' },
    ],
  });

test('an .xlsx file opens, and saving it gives a file that opens the same', async ({
  page,
}) => {
  await openApp(page);
  await openFile(page, 'Budget.xlsx', await budget());

  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await expectCell(page, 0, 0, 'A1 5');
  await expectCell(page, 1, 0, 'A2 10');
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A1*2');
  await expect(page.getByRole('alert')).toHaveCount(0);

  // The workbook opened is the one kept in browser storage.
  await page.reload();
  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await clickCell(page, 0, 0);
  await enter(page, '7');
  await expectCell(page, 1, 0, 'A2 14');

  await page.getByRole('tab', { name: 'File' }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save' }).click();
  const download = await downloading;
  expect(download.suggestedFilename()).toBe('Budget.xlsx');
  const saved = await readFile(await download.path());

  // A change after the save is not in the file saved.
  await clickCell(page, 0, 1);
  await enter(page, 'later');
  await openFile(page, 'Budget.xlsx', saved);
  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await expectCell(page, 0, 1, 'B1');
  await expectCell(page, 0, 0, 'A1 7');
  await expectCell(page, 1, 0, 'A2 14');
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A1*2');
  await sheetTabs(page).filter({ hasText: 'Costs' }).click();
  await expectCell(page, 0, 0, 'A1 Rent');
});

test('a file with features Gridworks does not support opens and says what was dropped', async ({
  page,
}) => {
  await openApp(page);
  await openFile(
    page,
    'Report.xlsx',
    await xlsxFile({
      sheets: [
        {
          name: 'Report',
          rows: '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>OFFSET(A1,1,1)</f><v>9</v></c></row>',
        },
      ],
      parts: UNSUPPORTED_PARTS,
    }),
  );

  await expect(page.getByRole('alert')).toHaveText(
    'Some content of "Report.xlsx" is not supported and was not opened: ' +
      '1 chart, 1 pivot table, 1 picture, 1 formula that Gridworks cannot ' +
      'calculate (its last result was kept as a value).',
  );
  await expect(sheetTabs(page)).toHaveText(['Report']);
  await expectCell(page, 0, 1, 'B1 9');

  await page.getByRole('button', { name: 'Close' }).click();
  await expect(page.getByRole('alert')).toHaveCount(0);
});

/** Selects a cell or a named range by typing it in the name box. */
async function goTo(page: Page, text: string): Promise<void> {
  const nameBox = page.getByRole('combobox', { name: 'Name box' });
  await nameBox.fill(text);
  await nameBox.press('Enter');
}

/** Expects the features of the workbook saved by Excel to show. */
async function expectExcelFeatures(page: Page): Promise<void> {
  await expect(sheetTabs(page)).toHaveText(['Formats', 'My lists']);
  await page.getByRole('tab', { name: 'Home' }).click();
  await expectCell(page, 0, 0, 'A1 Total');
  await expect(
    page.getByRole('button', { name: 'Bold', exact: true }),
  ).toHaveAttribute('aria-pressed', 'true');
  await expect(
    page.getByRole('button', { name: 'Font', exact: true }),
  ).toHaveText('Arial');
  await clickCell(page, 8, 8);
  expect(await cellBackground(page, 0, 0)).toEqual([255, 255, 0]);
  await clickCell(page, 0, 0);
  await page.getByRole('tab', { name: 'Review' }).click();
  await expect(
    page.getByRole('button', { name: 'Edit note', exact: true }),
  ).toBeVisible();
  await goTo(page, 'B1');
  await expect(page.getByRole('status')).toHaveText('B1 12.5%');
  await goTo(page, 'B3');
  await expect(page.getByRole('status')).toHaveText('B3 45');
  await expect(page.getByLabel('Formula bar')).toHaveValue('=SUM(Rates)');
  await goTo(page, 'Colors');
  await expect(page.getByRole('status')).toHaveText('A1 Red');
}

test('a workbook saved by Excel opens with its features, and saving it keeps them', async ({
  page,
}) => {
  await openApp(page);
  await openFile(page, 'Features.xlsx', await excelWorkbook());

  await expect(page.getByRole('alert')).toHaveText(
    'Some content of "Features.xlsx" is not supported and was not opened: ' +
      '1 table, 3 conditional formats, 1 data validation rule, ' +
      '2 defined names, 1 hyperlink that is not a web or email address.',
  );
  await page.getByRole('button', { name: 'Close' }).click();
  await expectExcelFeatures(page);

  await page.getByRole('tab', { name: 'File' }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save' }).click();
  const saved = await readFile(await (await downloading).path());
  await openFile(page, 'Features.xlsx', saved);
  await expect(page.getByRole('alert')).toHaveCount(0);
  await expectExcelFeatures(page);
});

test('a .csv file opens with typed values, and exporting the sheet gives a file that reopens the same', async ({
  page,
}) => {
  await openApp(page);
  const csv =
    '\uFEFFItem,Price,Sold,Note\r\n' +
    '"Tea, green",3.75,2026-03-01,"said ""fresh""\nand cheap"\r\n' +
    'Café,"1,200",FALSE,\r\n';
  await openFile(page, 'Prices.csv', new TextEncoder().encode(csv), 'text/csv');

  await expect(sheetTabs(page)).toHaveText(['Prices']);
  await expect(page.getByRole('alert')).toHaveCount(0);
  await expectCell(page, 1, 0, 'A2 Tea, green');
  await expectCell(page, 1, 1, 'B2 3.75');
  await expectCell(page, 1, 2, 'C2 3/1/2026');
  await expectCell(page, 1, 3, 'D2 said "fresh"\nand cheap');
  await expectCell(page, 2, 0, 'A3 Café');
  await expectCell(page, 2, 1, 'B3 1,200');
  await expectCell(page, 2, 2, 'C3 FALSE');

  // A formula is exported as its result.
  await clickCell(page, 3, 1);
  await enter(page, '=B2+B3');
  await page.getByRole('tab', { name: 'File' }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Export CSV' }).click();
  const download = await downloading;
  expect(download.suggestedFilename()).toBe('Prices.csv');
  const saved = await readFile(await download.path());
  expect(saved.toString('utf8')).toBe(
    '\uFEFFItem,Price,Sold,Note\r\n' +
      '"Tea, green",3.75,2026-03-01,"said ""fresh""\nand cheap"\r\n' +
      'Café,1200,FALSE,\r\n' +
      ',1203.75,,\r\n',
  );

  await openFile(page, 'Prices.csv', saved, 'text/csv');
  await expectCell(page, 3, 1, 'B4 1203.75');
  await expect(page.getByLabel('Formula bar')).toHaveValue('1203.75');
  await expectCell(page, 1, 2, 'C2 3/1/2026');
});

test('a file that is not an .xlsx file leaves the workbook as it was', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'kept');
  await openFile(page, 'notes.xlsx', new TextEncoder().encode('plain text'));

  await expect(page.getByRole('alert')).toHaveText(
    '"notes.xlsx" could not be opened because it is not a valid Excel ' +
      'workbook (.xlsx).',
  );
  await expect(sheetTabs(page)).toHaveText(['Sheet1']);
  await expectCell(page, 0, 0, 'A1 kept');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the File tab and the message bar are translated and dark', async ({
    page,
  }) => {
    await openApp(page);
    await page.getByRole('tab', { name: '文件' }).click();
    await page.getByRole('button', { name: '打开' }).hover();
    await expect(
      page.getByRole('tooltip', {
        name: '从计算机中打开 Excel 工作簿 (.xlsx) 或 CSV 文件 (.csv)',
      }),
    ).toBeVisible();
    await page.getByRole('button', { name: '导出 CSV' }).hover();
    await expect(
      page.getByRole('tooltip', {
        name: '将当前工作表以 CSV 文件 (.csv) 格式保存到计算机，公式保存为其结果',
      }),
    ).toBeVisible();
    const chooser = page.waitForEvent('filechooser');
    await page.getByRole('button', { name: '打开' }).click();
    await (
      await chooser
    ).setFiles({
      name: 'notes.xlsx',
      mimeType: XLSX_TYPE,
      buffer: Buffer.from('plain text'),
    });

    const alert = page.getByRole('alert');
    await expect(alert).toHaveText(
      '无法打开“notes.xlsx”，因为它不是有效的 Excel 工作簿 (.xlsx)。',
    );
    const background = await alert.evaluate(
      (node) => getComputedStyle(node).backgroundColor,
    );
    const [red, green, blue] = background.match(/\d+/g)!.map(Number);
    expect((red! + green! + blue!) / 3).toBeLessThan(128);
    await page.getByRole('button', { name: '关闭' }).click();
    await expect(alert).toHaveCount(0);
  });
});
