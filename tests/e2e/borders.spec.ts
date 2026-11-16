import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { cellBackground, clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const TOP = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/** Selects the range between two cells with a click and a Shift+click. */
async function selectRange(
  page: Page,
  from: [number, number],
  to: [number, number],
) {
  await clickCell(page, ...from);
  const rect = cellRect({ row: to[0], col: to[1] }, TOP);
  await page.locator('canvas').click({
    position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
    modifiers: ['Shift'],
  });
}

/** The red, green and blue the canvas drew in the middle of a cell's top edge. */
async function topEdge(page: Page, row: number, col: number) {
  const rect = cellRect({ row, col }, TOP);
  return page.locator('canvas').evaluate(
    (canvas: HTMLCanvasElement, { x, y }) => {
      const scale = canvas.width / canvas.clientWidth;
      const { data } = canvas
        .getContext('2d')!
        .getImageData(Math.round(x * scale), Math.round(y * scale), 1, 1);
      return [data[0]!, data[1]!, data[2]!];
    },
    { x: rect.x + rect.width / 2, y: rect.y + 1 },
  );
}

/** Chooses an item of the Borders menu of the ribbon. */
async function borders(page: Page, item: string) {
  await button(page, 'Borders').click();
  await page.getByRole('menuitem', { name: item, exact: true }).click();
}

test('draws borders of a chosen line and color, and undoes them', async ({
  page,
}) => {
  await openApp(page);
  const plain = await topEdge(page, 2, 2);

  await selectRange(page, [2, 2], [3, 3]);
  await button(page, 'Borders').click();
  await page.getByRole('menuitemradio', { name: 'Thick line' }).click();
  await page.getByRole('menuitem', { name: 'Red', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Outside borders' }).click();
  await clickCell(page, 8, 8);
  expect(await topEdge(page, 2, 2)).toEqual([255, 0, 0]);
  expect(await topEdge(page, 2, 3)).toEqual([255, 0, 0]);
  // Not inside the range.
  expect(await topEdge(page, 3, 2)).toEqual(plain);

  await selectRange(page, [2, 2], [3, 3]);
  await borders(page, 'No border');
  await clickCell(page, 8, 8);
  expect(await topEdge(page, 2, 2)).toEqual(plain);
  await page.keyboard.press('ControlOrMeta+z');
  // Undo selects the cells it changed, which shades them.
  await clickCell(page, 8, 8);
  expect(await topEdge(page, 2, 2)).toEqual([255, 0, 0]);
});

test('merges cells into one that selection and the keyboard treat as one, and unmerges them', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'Title');
  await clickCell(page, 0, 1);
  await enter(page, 'gone');

  await clickCell(page, 0, 0);
  await expect(button(page, 'Merge cells')).toBeDisabled();
  await expect(button(page, 'Unmerge cells')).toBeDisabled();
  await selectRange(page, [0, 0], [1, 2]);
  await button(page, 'Merge cells').click();
  await expectCell(page, 1, 2, 'A1 Title');
  await expect(button(page, 'Merge cells')).toBeDisabled();

  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('status')).toHaveText('D1');
  await page.keyboard.press('ArrowLeft');
  await expect(page.getByRole('status')).toHaveText('A1 Title');
  await page.keyboard.press('ArrowDown');
  await expect(page.getByRole('status')).toHaveText('A3');

  await clickCell(page, 0, 1);
  await button(page, 'Unmerge cells').click();
  await expectCell(page, 0, 1, 'B1');
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 1, 1, 'A1 Title');
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 1, 'B1 gone');
});

test('paints the formats of a range onto the next one selected', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await button(page, 'Fill color').click();
  await page.getByRole('menuitem', { name: 'Yellow', exact: true }).click();
  const white = await cellBackground(page, 2, 2);

  await button(page, 'Format painter').click();
  await expect(button(page, 'Format painter')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await clickCell(page, 2, 2);
  await clickCell(page, 5, 5);
  expect(await cellBackground(page, 2, 2)).toEqual([255, 255, 0]);
  await expect(button(page, 'Format painter')).toHaveAttribute(
    'aria-pressed',
    'false',
  );

  // Escape puts the painter away; the next click paints nothing.
  await clickCell(page, 0, 0);
  await button(page, 'Format painter').click();
  await page.keyboard.press('Escape');
  await clickCell(page, 3, 3);
  await clickCell(page, 5, 5);
  expect(await cellBackground(page, 3, 3)).toEqual(white);

  await page.keyboard.press('ControlOrMeta+z');
  await clickCell(page, 5, 5);
  expect(await cellBackground(page, 2, 2)).toEqual(white);
});

test('translates the border, merge and painter controls', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '开始' }).click();
  for (const name of ['边框', '合并单元格', '取消合并单元格', '格式刷']) {
    await expect(button(page, name)).toBeVisible();
  }
});
