import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { cellBackground, clickCell, enter, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const WHITE = [255, 255, 255];
const LIGHT_RED = [255, 199, 206];
const YELLOW = [255, 255, 0];
// The colors of Excel's scales and bars, and those between them.
const SCALE_RED = [248, 105, 107];
const SCALE_YELLOW = [255, 235, 132];
const SCALE_GREEN = [99, 190, 123];
const HALF_RED_GREEN = [174, 148, 115];
const QUARTER_RED_GREEN = [211, 126, 111];
const BAR_BLUE = [99, 142, 198];
const BAR_GREEN = [0, 176, 80];

/** Types values down column A from A1. */
async function fillColumn(page: Page, values: string[]): Promise<void> {
  await clickCell(page, 0, 0);
  for (const value of values) await enter(page, value);
}

/** Selects a range by dragging from one cell to another. */
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

/** Chooses an item of the Conditional formatting menu of the Home tab. */
async function openRuleMenu(page: Page, item: string): Promise<void> {
  await page.getByRole('button', { name: 'Conditional formatting' }).click();
  await page.getByRole('menuitem', { name: item }).click();
}

/** The backgrounds of A1 down to A{count}, with a cell elsewhere selected. */
async function columnBackgrounds(page: Page, count: number) {
  await clickCell(page, 10, 5);
  const result = [];
  for (let row = 0; row < count; row++) {
    result.push(await cellBackground(page, row, 0));
  }
  return result;
}

test('formats the cells greater than a value, and follows their values and formula results', async ({
  page,
}) => {
  await openApp(page);
  await fillColumn(page, ['5', '20', '=A1*3']);
  await selectRange(page, [0, 0], [2, 0]);

  await openRuleMenu(page, 'Greater than...');
  const dialog = page.getByRole('dialog');
  await expect(dialog.getByLabel('Applies to')).toHaveValue('A1:A3');
  await dialog.getByLabel('Value', { exact: true }).fill('10');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  expect(await columnBackgrounds(page, 3)).toEqual([
    WHITE,
    LIGHT_RED,
    LIGHT_RED,
  ]);

  // A1 below the value makes A3, its triple, fall below it too.
  await clickCell(page, 0, 0);
  await enter(page, '2');
  expect(await columnBackgrounds(page, 3)).toEqual([WHITE, LIGHT_RED, WHITE]);

  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  expect(await columnBackgrounds(page, 3)).toEqual([WHITE, WHITE, WHITE]);
  await page.keyboard.press('ControlOrMeta+y');
  expect(await columnBackgrounds(page, 3)).toEqual([
    WHITE,
    LIGHT_RED,
    LIGHT_RED,
  ]);
});

/** The color the canvas drew at a share of the width of a cell. */
async function pixelAt(
  page: Page,
  row: number,
  col: number,
  share: number,
): Promise<number[]> {
  const top = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };
  const rect = cellRect({ row, col }, top);
  const x = rect.x + rect.width * share;
  return page.locator('canvas').evaluate(
    (canvas: HTMLCanvasElement, { x, y }) => {
      const scale = canvas.width / canvas.clientWidth;
      const { data } = canvas
        .getContext('2d')!
        .getImageData(x * scale, y * scale, 1, 1);
      return [data[0]!, data[1]!, data[2]!];
    },
    { x, y: rect.y + 4 },
  );
}

test('shades numbers with a two-color or three-color scale, following their values, each undone', async ({
  page,
}) => {
  await openApp(page);
  await fillColumn(page, ['0', '5', '10']);
  await selectRange(page, [0, 0], [2, 0]);
  await openRuleMenu(page, '2-color scale...');
  const dialog = page.getByRole('dialog');
  await expect(
    dialog.getByRole('button', { name: 'Lowest value: Salmon' }),
  ).toHaveAttribute('aria-pressed', 'true');
  await dialog.getByRole('button', { name: 'OK' }).click();
  expect(await columnBackgrounds(page, 3)).toEqual([
    SCALE_RED,
    HALF_RED_GREEN,
    SCALE_GREEN,
  ]);

  await clickCell(page, 2, 0);
  await enter(page, '20');
  expect(await columnBackgrounds(page, 3)).toEqual([
    SCALE_RED,
    QUARTER_RED_GREEN,
    SCALE_GREEN,
  ]);
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  expect(await columnBackgrounds(page, 3)).toEqual([WHITE, WHITE, WHITE]);
  await page.keyboard.press('ControlOrMeta+y');
  expect(await columnBackgrounds(page, 3)).toEqual([
    SCALE_RED,
    HALF_RED_GREEN,
    SCALE_GREEN,
  ]);

  await selectRange(page, [0, 0], [2, 0]);
  await openRuleMenu(page, '3-color scale...');
  await dialog.getByRole('button', { name: 'OK' }).click();
  expect(await columnBackgrounds(page, 3)).toEqual([
    SCALE_RED,
    SCALE_YELLOW,
    SCALE_GREEN,
  ]);
});

test('draws data bars as long as their share of the highest, and edits and deletes them in the list', async ({
  page,
}) => {
  await openApp(page);
  await fillColumn(page, ['5', '20']);
  await selectRange(page, [0, 0], [1, 0]);
  await openRuleMenu(page, 'Data bars...');
  const dialog = page.getByRole('dialog');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await clickCell(page, 10, 5);
  expect(await pixelAt(page, 0, 0, 0.1)).toEqual(BAR_BLUE);
  expect(await pixelAt(page, 0, 0, 0.5)).toEqual(WHITE);
  expect(await pixelAt(page, 1, 0, 0.9)).toEqual(BAR_BLUE);

  // The bars follow the values.
  await clickCell(page, 1, 0);
  await enter(page, '10');
  await clickCell(page, 10, 5);
  expect(await pixelAt(page, 0, 0, 0.4)).toEqual(BAR_BLUE);

  await openRuleMenu(page, 'Manage rules...');
  const rows = dialog.getByRole('listitem');
  await expect(rows.nth(0)).toContainText('Data bars');
  await expect(rows.nth(0)).toContainText('A1:A2');
  await rows.nth(0).getByRole('button', { name: 'Edit' }).click();
  await dialog.getByRole('button', { name: 'Bar color: Green' }).click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await dialog.getByRole('button', { name: 'Close' }).click();
  expect(await pixelAt(page, 1, 0, 0.9)).toEqual(BAR_GREEN);

  await openRuleMenu(page, 'Manage rules...');
  await rows.nth(0).getByRole('button', { name: 'Edit' }).click();
  await dialog.getByLabel('Rule').selectOption('Two-color scale');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(rows.nth(0)).toContainText('Two-color scale');
  await rows.nth(0).getByRole('button', { name: 'Delete' }).click();
  await expect(rows).toHaveCount(0);
  await dialog.getByRole('button', { name: 'Close' }).click();
  expect(await pixelAt(page, 1, 0, 0.9)).toEqual(WHITE);

  await page.keyboard.press('ControlOrMeta+z');
  expect(await columnBackgrounds(page, 2)).toEqual([SCALE_RED, SCALE_GREEN]);
  await page.keyboard.press('ControlOrMeta+z');
  // The active cell of a selection shows its bar too.
  await expect(page.getByLabel('Name box')).toHaveText('A1');
  expect(await pixelAt(page, 0, 0, 0.1)).toEqual(BAR_GREEN);
  await clickCell(page, 10, 5);
  expect(await pixelAt(page, 1, 0, 0.9)).toEqual(BAR_GREEN);
});

test('refuses a rule without a value or with ranges it cannot read', async ({
  page,
}) => {
  await openApp(page);
  await openRuleMenu(page, 'Between...');
  const dialog = page.getByRole('dialog');
  const ok = dialog.getByRole('button', { name: 'OK' });
  await expect(ok).toBeDisabled();
  await dialog.getByLabel('Value', { exact: true }).fill('1');
  await dialog.getByLabel('And').fill('5');
  await expect(ok).toBeEnabled();
  await dialog.getByLabel('Applies to').fill('A1:');
  await expect(ok).toBeDisabled();
  await expect(dialog.getByText('Enter cells such as A1:B10')).toBeVisible();
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
});

test('lists the rules of the sheet, the first winning, to edit, reorder and delete them, each undone', async ({
  page,
}) => {
  await openApp(page);
  await fillColumn(page, ['3', '9', '3']);
  await selectRange(page, [0, 0], [2, 0]);
  await openRuleMenu(page, 'Duplicate values...');
  const dialog = page.getByRole('dialog');
  await dialog.getByRole('button', { name: 'OK' }).click();
  // A new rule comes first, as in Excel, so its yellow wins over the red.
  await openRuleMenu(page, 'Greater than...');
  await dialog.getByLabel('Value', { exact: true }).fill('2');
  await dialog.getByRole('button', { name: 'Fill color: Yellow' }).click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  expect(await columnBackgrounds(page, 3)).toEqual([YELLOW, YELLOW, YELLOW]);

  await openRuleMenu(page, 'Manage rules...');
  await expect(dialog).toHaveAccessibleName('Conditional formatting rules');
  const rows = dialog.getByRole('listitem');
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0)).toContainText('Cell value > 2');
  await expect(rows.nth(0)).toContainText('A1:A3');
  await expect(rows.nth(1)).toContainText('Duplicate values');
  await expect(
    rows.nth(0).getByRole('button', { name: 'Move up' }),
  ).toBeDisabled();
  await expect(
    rows.nth(1).getByRole('button', { name: 'Move down' }),
  ).toBeDisabled();

  await rows.nth(1).getByRole('button', { name: 'Move up' }).click();
  await expect(rows.nth(0)).toContainText('Duplicate values');
  await rows.nth(1).getByRole('button', { name: 'Edit' }).click();
  await dialog.getByLabel('Rule').selectOption('Equal to');
  await dialog.getByLabel('Value', { exact: true }).fill('9');
  await dialog.getByRole('button', { name: 'Bold' }).click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(rows.nth(1)).toContainText('Cell value = 9');
  await dialog.getByRole('button', { name: 'Close' }).click();
  expect(await columnBackgrounds(page, 3)).toEqual([
    LIGHT_RED,
    YELLOW,
    LIGHT_RED,
  ]);

  await openRuleMenu(page, 'Manage rules...');
  await rows.nth(0).getByRole('button', { name: 'Delete' }).click();
  await expect(rows).toHaveCount(1);
  await dialog.getByRole('button', { name: 'Close' }).click();
  expect(await columnBackgrounds(page, 3)).toEqual([WHITE, YELLOW, WHITE]);

  // Each change is undone on its own: the deletion, the edit, the move.
  await page.keyboard.press('ControlOrMeta+z');
  expect(await columnBackgrounds(page, 3)).toEqual([
    LIGHT_RED,
    YELLOW,
    LIGHT_RED,
  ]);
  await page.keyboard.press('ControlOrMeta+z');
  await openRuleMenu(page, 'Manage rules...');
  await expect(rows.nth(1)).toContainText('Cell value > 2');
  await dialog.getByRole('button', { name: 'Close' }).click();
  await page.keyboard.press('ControlOrMeta+z');
  expect(await columnBackgrounds(page, 3)).toEqual([YELLOW, YELLOW, YELLOW]);
  await page.keyboard.press('ControlOrMeta+y');
  expect(await columnBackgrounds(page, 3)).toEqual([
    LIGHT_RED,
    YELLOW,
    LIGHT_RED,
  ]);
});

test('offers the rules in Chinese, and stays usable in the dark theme', async ({
  page,
}) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: 'Dark' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '开始' }).click();
  await page.getByRole('button', { name: '条件格式' }).hover();
  await expect(page.getByRole('tooltip')).toBeVisible();
  await page.getByRole('button', { name: '条件格式' }).click();
  await expect(page.getByRole('menuitem', { name: '双色刻度…' })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: '三色刻度…' })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: '数据条…' })).toBeVisible();
  await page.getByRole('menuitem', { name: '文本包含…' }).click();
  const dialog = page.getByRole('dialog');
  await expect(dialog.getByLabel('应用于')).toHaveValue('A1');
  await expect(dialog).toHaveCSS('background-color', 'rgb(27, 26, 25)');
});
