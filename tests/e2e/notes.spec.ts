import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { clickCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const TOP = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/** The middle of a cell, on the page. */
async function middle(page: Page, row: number, col: number) {
  const box = (await page.locator('canvas').boundingBox())!;
  const rect = cellRect({ row, col }, TOP);
  return {
    x: box.x + rect.x + rect.width / 2,
    y: box.y + rect.y + rect.height / 2,
  };
}

/** Whether the canvas draws the red corner of a note in a cell. */
async function hasMarker(page: Page, row: number, col: number) {
  const rect = cellRect({ row, col }, TOP);
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const { data } = canvas
      .getContext('2d')!
      .getImageData(
        (rect.x + rect.width - 3) * scale,
        (rect.y + 3) * scale,
        1,
        1,
      );
    return data[0]! > 180 && data[1]! < 100 && data[2]! < 100;
  }, rect);
}

/** Right-clicks a cell and chooses an item of its menu. */
async function fromMenu(page: Page, row: number, col: number, item: string) {
  const { x, y } = await middle(page, row, col);
  await page.mouse.click(x, y, { button: 'right' });
  await page.getByRole('menuitem', { name: item, exact: true }).click();
}

test('adds, shows, edits and deletes a note, from the ribbon and the menu', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 1, 1);
  await page.getByRole('tab', { name: 'Review' }).click();
  await expect(button(page, 'Delete note')).toBeDisabled();
  await button(page, 'New note').click();
  let dialog = page.getByRole('dialog', { name: 'New note' });
  await dialog.getByLabel('Note', { exact: true }).fill('Check the total');
  await button(page, 'OK').click();
  await expect(dialog).toBeHidden();
  await expect.poll(() => hasMarker(page, 1, 1)).toBe(true);
  expect(await hasMarker(page, 1, 2)).toBe(false);

  // The note appears while the pointer rests on its cell.
  const note = page.getByRole('tooltip', { name: 'Check the total' });
  let { x, y } = await middle(page, 1, 1);
  await page.mouse.move(x, y);
  await expect(note).toBeVisible();
  ({ x, y } = await middle(page, 4, 4));
  await page.mouse.move(x, y);
  await expect(note).toBeHidden();

  // The cell's menu edits it.
  await fromMenu(page, 1, 1, 'Edit note');
  dialog = page.getByRole('dialog', { name: 'Edit note' });
  await expect(dialog.getByLabel('Note', { exact: true })).toHaveValue(
    'Check the total',
  );
  await dialog.getByLabel('Note', { exact: true }).fill('Checked');
  await button(page, 'OK').click();
  ({ x, y } = await middle(page, 1, 1));
  await page.mouse.move(x, y);
  await expect(page.getByRole('tooltip', { name: 'Checked' })).toBeVisible();

  // The ribbon deletes it, and undo brings it back.
  await clickCell(page, 1, 1);
  await button(page, 'Delete note').click();
  await expect.poll(() => hasMarker(page, 1, 1)).toBe(false);
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => hasMarker(page, 1, 1)).toBe(true);
  await page.keyboard.press('ControlOrMeta+y');
  await expect.poll(() => hasMarker(page, 1, 1)).toBe(false);
  await fromMenu(page, 1, 1, 'New note');
  await page.getByLabel('Note', { exact: true }).fill('Again');
  await button(page, 'OK').click();
  await fromMenu(page, 1, 1, 'Delete note');
  await expect.poll(() => hasMarker(page, 1, 1)).toBe(false);
});

test('inserts, follows, edits and removes links to a web address and to a cell', async ({
  page,
  context,
}) => {
  await context.route('https://example.com/**', (route) =>
    route.fulfill({ contentType: 'text/html', body: 'Example' }),
  );
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.getByRole('tab', { name: 'Insert' }).click();
  await button(page, 'Link').click();
  let dialog = page.getByRole('dialog', { name: 'Insert link' });
  await dialog.getByLabel('Text to display').fill('Example');
  await dialog.getByLabel('Address', { exact: true }).fill('example.com');
  await button(page, 'OK').click();
  await expect(dialog).toBeHidden();
  await expect(page.getByRole('status')).toHaveText('A1 Example');

  // Clicking a web link opens it in a new tab.
  const opened = page.waitForEvent('popup');
  await clickCell(page, 0, 0);
  const popup = await opened;
  await expect(popup).toHaveURL('https://example.com/');
  await popup.close();

  // A link to a cell, from the menu, selects that cell.
  await fromMenu(page, 0, 1, 'Link');
  dialog = page.getByRole('dialog', { name: 'Insert link' });
  await dialog.getByLabel('Place in this workbook').check();
  await dialog.getByLabel('Cell reference').fill('D10');
  await dialog.getByLabel('Text to display').fill('Totals');
  await button(page, 'OK').click();
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('D10');

  // A wrong reference is refused.
  await fromMenu(page, 0, 1, 'Edit link');
  dialog = page.getByRole('dialog', { name: 'Edit link' });
  await expect(dialog.getByLabel('Cell reference')).toHaveValue('D10');
  await dialog.getByLabel('Cell reference').fill('nowhere');
  await button(page, 'OK').click();
  await expect(dialog.getByRole('alert')).toHaveText(
    'Type a cell reference, such as A1.',
  );
  await dialog.getByLabel('Cell reference').fill('E5');
  await dialog.getByLabel('Text to display').fill('Sums');
  await button(page, 'OK').click();
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('E5');

  // Removing it keeps the text, and undo puts it back.
  await fromMenu(page, 0, 1, 'Remove link');
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('B1 Sums');
  await page.keyboard.press('ControlOrMeta+z');
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('E5');
  await page.keyboard.press('ControlOrMeta+z');
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('D10');
});

/** Whether the canvas draws any text of a cell in the dark theme's blue. */
async function hasLinkColor(page: Page, row: number, col: number) {
  const rect = cellRect({ row, col }, TOP);
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const { data } = canvas
      .getContext('2d')!
      .getImageData(
        (rect.x + 2) * scale,
        (rect.y + 2) * scale,
        (rect.width - 4) * scale,
        (rect.height - 4) * scale,
      );
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 2]! > 200 && data[i + 2]! - data[i]! > 80) return true;
    }
    return false;
  }, rect);
}

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the note and link controls are translated and drawn in the theme', async ({
    page,
  }) => {
    await openApp(page);
    await clickCell(page, 1, 1);
    await page.getByRole('tab', { name: '审阅' }).click();
    await button(page, '新建注释').click();
    const dialog = page.getByRole('dialog', { name: '新建注释' });
    await dialog.getByLabel('注释', { exact: true }).fill('核对');
    await button(page, '确定').click();
    await expect(dialog).toBeHidden();
    await expect(button(page, '编辑注释')).toBeVisible();
    await expect(button(page, '删除注释')).toBeEnabled();
    await expect.poll(() => hasMarker(page, 1, 1)).toBe(true);
    const { x, y } = await middle(page, 1, 1);
    await page.mouse.move(x, y);
    const note = page.getByRole('tooltip', { name: '核对' });
    await expect(note).toHaveCSS('background-color', 'rgb(37, 36, 35)');
    await expect(note).toHaveCSS('color', 'rgb(237, 235, 232)');

    await clickCell(page, 0, 0);
    await page.getByRole('tab', { name: '插入' }).click();
    await button(page, '链接').click();
    const link = page.getByRole('dialog', { name: '插入链接' });
    await link.getByLabel('要显示的文字').fill('Example');
    await link.getByLabel('地址', { exact: true }).fill('');
    await button(page, '确定').click();
    await expect(link.getByRole('alert')).toHaveText('请输入网址。');
    await link.getByLabel('地址', { exact: true }).fill('example.com');
    await button(page, '确定').click();
    await expect(link).toBeHidden();
    await expect.poll(() => hasLinkColor(page, 0, 0)).toBe(true);
    expect(await hasLinkColor(page, 1, 1)).toBe(false);

    await fromMenu(page, 0, 0, '编辑链接');
    await expect(page.getByRole('dialog', { name: '编辑链接' })).toBeVisible();
    await button(page, '取消').click();
    await fromMenu(page, 0, 0, '删除链接');
    await expect.poll(() => hasLinkColor(page, 0, 0)).toBe(false);
  });
});
