import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

/** Types rows of cells from A1, each row left to right. */
async function typeRows(page: Page, rows: string[][]): Promise<void> {
  for (const [row, cells] of rows.entries()) {
    for (const [col, text] of cells.entries()) {
      await clickCell(page, row, col);
      await enter(page, text);
    }
  }
}

/** Drags from one cell to another, the first staying active. */
async function selectRange(
  page: Page,
  from: [number, number],
  to: [number, number],
): Promise<void> {
  const top = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };
  const center = ([row, col]: [number, number]) => {
    const rect = cellRect({ row, col }, top);
    return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  };
  const box = (await page.locator('canvas').boundingBox())!;
  const start = center(from);
  const end = center(to);
  await page.mouse.move(box.x + start.x, box.y + start.y);
  await page.mouse.down();
  await page.mouse.move(box.x + end.x, box.y + end.y);
  await page.mouse.up();
}

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

test('sorts the selection by the active column, and undoes and redoes it', async ({
  page,
}) => {
  await openApp(page);
  await typeRows(page, [
    ['b', '2'],
    ['c', '1'],
    ['a', '3'],
  ]);
  await page.getByRole('tab', { name: 'Data' }).click();
  await expect(button(page, 'Sort A to Z')).toBeDisabled();
  // B1 stays active.
  await selectRange(page, [0, 1], [2, 0]);
  await button(page, 'Sort A to Z').click();
  await expectCell(page, 0, 0, 'A1 c');
  await expectCell(page, 2, 1, 'B3 3');

  await selectRange(page, [0, 0], [2, 1]);
  await button(page, 'Sort Z to A').click();
  await expectCell(page, 0, 0, 'A1 c');
  await expectCell(page, 1, 1, 'B2 2');
  await expectCell(page, 2, 0, 'A3 a');

  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 0, 'A1 b');
  await page.keyboard.press('ControlOrMeta+y');
  await expectCell(page, 0, 0, 'A1 c');
});

test('sorts by several columns with a header row in the dialog', async ({
  page,
}) => {
  await openApp(page);
  await typeRows(page, [
    ['Name', 'Team'],
    ['dan', 'b'],
    ['amy', 'a'],
    ['cat', 'b'],
  ]);
  await selectRange(page, [0, 0], [3, 1]);
  await page.getByRole('tab', { name: 'Data' }).click();
  await button(page, 'Custom sort').click();
  const dialog = page.getByRole('dialog', { name: 'Sort' });
  await dialog.getByLabel('My data has headers').check();
  await dialog.getByLabel('Sort by').selectOption('Team');
  await dialog.getByLabel('Order', { exact: true }).selectOption('Z to A');
  await expect(button(page, 'Delete level')).toBeDisabled();
  await button(page, 'Add level').click();
  await dialog.getByLabel('Then by').selectOption('Name');
  await button(page, 'OK').click();
  await expect(dialog).toBeHidden();
  await expectCell(page, 0, 0, 'A1 Name');
  await expectCell(page, 1, 0, 'A2 cat');
  await expectCell(page, 2, 0, 'A3 dan');
  await expectCell(page, 3, 0, 'A4 amy');
});

test('translates the sort controls', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '数据' }).click();
  for (const name of ['升序排序', '降序排序', '自定义排序']) {
    await expect(button(page, name)).toBeVisible();
  }
});

test.describe('in the dark theme', () => {
  test.use({ colorScheme: 'dark' });

  test('the sort dialog is dark', async ({ page }) => {
    await openApp(page);
    await typeRows(page, [['b'], ['a']]);
    await selectRange(page, [0, 0], [1, 0]);
    await page.getByRole('tab', { name: 'Data' }).click();
    await button(page, 'Custom sort').click();
    const dialog = page.getByRole('dialog', { name: 'Sort' });
    await expect(dialog).toBeVisible();
    for (const element of [dialog, dialog.getByLabel('Sort by')]) {
      const background = await element.evaluate(
        (each) => getComputedStyle(each).backgroundColor,
      );
      const [r, g, b] = background.match(/\d+/g)!.map(Number);
      expect(r! + g! + b!).toBeLessThan(3 * 96);
    }
  });
});
