import { expect, test, type Page } from '@playwright/test';
import {
  COLUMN_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from '../../src/app/grid/geometry';
import { cellBackground, clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/** Expects a toggle of the ribbon to be on or off. */
async function expectPressed(page: Page, name: string, pressed: boolean) {
  await expect(button(page, name)).toHaveAttribute(
    'aria-pressed',
    String(pressed),
  );
}

/** Opens a menu of the ribbon and chooses one of its items. */
async function choose(page: Page, menu: string, item: string) {
  await button(page, menu).click();
  await page.getByRole('menuitemradio', { name: item, exact: true }).click();
}

/** The cell the grid announces for a click at a height below the headers. */
async function rowAt(page: Page, y: number): Promise<string | null> {
  await page.locator('canvas').click({
    position: { x: ROW_HEADER_WIDTH + 10, y: COLUMN_HEADER_HEIGHT + y },
  });
  return page.getByRole('status').textContent();
}

test('formats cells from the ribbon and the keyboard, shows the format of the active cell, and undoes it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'Total');
  await clickCell(page, 0, 0);

  await button(page, 'Bold').click();
  await page.keyboard.press('ControlOrMeta+i');
  await page.keyboard.press('ControlOrMeta+u');
  await button(page, 'Strikethrough').click();
  await choose(page, 'Font', 'Georgia');
  await choose(page, 'Font size', '12');
  await button(page, 'Center').click();
  await button(page, 'Top align').click();
  for (const name of [
    'Bold',
    'Italic',
    'Underline',
    'Strikethrough',
    'Center',
    'Top align',
  ]) {
    await expectPressed(page, name, true);
  }
  await expectPressed(page, 'Bottom align', false);
  await expect(button(page, 'Font')).toHaveText('Georgia');
  await expect(button(page, 'Font size')).toHaveText('12');

  // Another cell shows its own format; Ctrl+B toggles bold off again.
  await clickCell(page, 1, 0);
  await expectPressed(page, 'Bold', false);
  await expectPressed(page, 'Bottom align', true);
  await expect(button(page, 'Font')).toHaveText('Default font');
  await expect(button(page, 'Font size')).toHaveText('10');
  await clickCell(page, 0, 0);
  await page.keyboard.press('ControlOrMeta+b');
  await expectPressed(page, 'Bold', false);

  // Each change is undone and redone on its own.
  await page.keyboard.press('ControlOrMeta+z');
  await expectPressed(page, 'Bold', true);
  await page.keyboard.press('ControlOrMeta+z');
  await expectPressed(page, 'Top align', false);
  await page.keyboard.press('ControlOrMeta+y');
  await expectPressed(page, 'Top align', true);
  await expectCell(page, 0, 0, 'A1 Total');
});

test('colors text and fills from the palette or a custom color', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 1, 1);
  await enter(page, 'x');
  const white = await cellBackground(page, 1, 1);
  await clickCell(page, 1, 1);

  await button(page, 'Fill color').click();
  await page.getByRole('menuitem', { name: 'Yellow', exact: true }).click();
  await clickCell(page, 4, 4);
  expect(await cellBackground(page, 1, 1)).toEqual([255, 255, 0]);

  // A custom color, chosen in the color picker.
  await clickCell(page, 1, 1);
  await button(page, 'Fill color').click();
  await page.getByRole('menuitem', { name: 'Custom color' }).click();
  await page.getByLabel('Fill color: Custom color').fill('#123456');
  await clickCell(page, 4, 4);
  expect(await cellBackground(page, 1, 1)).toEqual([18, 52, 86]);

  await clickCell(page, 1, 1);
  await button(page, 'Font color').click();
  await page.getByRole('menuitem', { name: 'Red', exact: true }).click();
  await expect(
    button(page, 'Font color').locator('span span').last(),
  ).toHaveCSS('background-color', 'rgb(255, 0, 0)');
  await button(page, 'Font color').click();
  await page.getByRole('menuitem', { name: 'Custom color' }).click();
  await page.getByLabel('Font color: Custom color').fill('#00ff00');
  await expect(
    button(page, 'Font color').locator('span span').last(),
  ).toHaveCSS('background-color', 'rgb(0, 255, 0)');

  await button(page, 'Fill color').click();
  await page.getByRole('menuitem', { name: 'No fill' }).click();
  await clickCell(page, 4, 4);
  expect(await cellBackground(page, 1, 1)).toEqual(white);
});

test('grows a row to fit wrapped text and a large font', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'A long line of text that needs many lines to show');
  await clickCell(page, 0, 0);
  // Just below the first row, before it grows, is the second.
  expect(await rowAt(page, ROW_HEIGHT + 4)).toBe('A2');

  await clickCell(page, 0, 0);
  await button(page, 'Wrap text').click();
  await expectPressed(page, 'Wrap text', true);
  expect(await rowAt(page, ROW_HEIGHT + 4)).toMatch(/^A1 /);

  await page.keyboard.press('ControlOrMeta+z');
  expect(await rowAt(page, ROW_HEIGHT + 4)).toBe('A2');

  await clickCell(page, 0, 1);
  await choose(page, 'Font size', '36');
  expect(await rowAt(page, ROW_HEIGHT + 4)).toMatch(/^A1 /);
});

test('translates the format controls', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '开始' }).click();
  for (const name of ['加粗', '倾斜', '下划线', '填充颜色', '自动换行']) {
    await expect(button(page, name)).toBeVisible();
  }
  await expect(button(page, '字体')).toHaveText('默认字体');
});
