import { expect, test, type Page } from '@playwright/test';
import {
  cellRect,
  COLUMN_HEADER_HEIGHT,
  ROW_HEIGHT,
  UNIFORM,
} from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

const TOP = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };

/** Chooses an item of the Freeze panes menu on the View tab. */
async function freeze(page: Page, item: string): Promise<void> {
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: 'Freeze panes' }).click();
  await page.getByRole('menuitem', { name: item, exact: true }).click();
}

/** Scrolls the grid down by some rows, and waits for it to be drawn. */
async function scrollRows(page: Page, rows: number): Promise<void> {
  await page.locator('canvas').evaluate((canvas, top) => {
    canvas.parentElement!.parentElement!.scrollTop = top;
  }, rows * ROW_HEIGHT);
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(resolve)),
  );
}

/** The red, green and blue the canvas drew at a point. */
async function pixel(page: Page, x: number, y: number): Promise<number[]> {
  return page.locator('canvas').evaluate(
    (canvas: HTMLCanvasElement, [x, y]) => {
      const scale = canvas.width / canvas.clientWidth;
      const { data } = canvas
        .getContext('2d')!
        .getImageData(x! * scale, y! * scale, 1, 1);
      return [data[0]!, data[1]!, data[2]!];
    },
    [x, y],
  );
}

test('the top row stays in place while the rest scrolls, with a line under it, and works across it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'Name');
  await clickCell(page, 0, 1);
  await enter(page, '5');
  await clickCell(page, 11, 1);
  await enter(page, '7');

  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: 'Freeze panes' }).click();
  await expect(
    page.getByRole('menuitem', { name: 'Unfreeze panes' }),
  ).toHaveAttribute('data-disabled');
  await page.keyboard.press('Escape');
  await freeze(page, 'Freeze top row');
  await scrollRows(page, 10);

  // Row 1 is still at the top; the row under it is row 12.
  await expectCell(page, 0, 0, 'A1 Name');
  await expectCell(page, 1, 1, 'B12 7');
  // The boundary is marked by a line unlike the grid lines.
  const x = cellRect({ row: 0, col: 2 }, TOP).x + 20;
  const boundary = COLUMN_HEADER_HEIGHT + ROW_HEIGHT;
  expect(await pixel(page, x, boundary)).not.toEqual(
    await pixel(page, x, boundary + ROW_HEIGHT),
  );

  // A drag from the frozen row selects across the boundary.
  const box = (await page.locator('canvas').boundingBox())!;
  const b1 = cellRect({ row: 0, col: 1 }, TOP);
  const b2 = cellRect({ row: 1, col: 1 }, TOP);
  await page.mouse.move(box.x + b1.x + 10, box.y + b1.y + 10);
  await page.mouse.down();
  await page.mouse.move(box.x + b2.x + 10, box.y + b2.y + 10, { steps: 3 });
  await page.mouse.up();
  await expect(page.getByLabel('Summary of the selection')).toContainText(
    'Sum: 12',
  );

  // The fill handle drags from the frozen row into the scrolled ones.
  await clickCell(page, 0, 1);
  await page.mouse.move(box.x + b1.x + b1.width, box.y + b1.y + b1.height);
  await page.mouse.down();
  const b3 = cellRect({ row: 2, col: 1 }, TOP);
  await page.mouse.move(box.x + b3.x + 10, box.y + b3.y + 10, { steps: 5 });
  await page.mouse.up();
  await expectCell(page, 2, 1, 'B13 5');

  // A frozen cell is edited in place while the sheet is scrolled.
  await page
    .locator('canvas')
    .dblclick({ position: { x: b1.x + 10, y: b1.y + 10 } });
  await page.keyboard.press('End');
  await page.keyboard.type('0');
  await page.keyboard.press('Enter');
  await expectCell(page, 0, 1, 'B1 50');

  // The frozen row survives a reload.
  await page.waitForTimeout(1000);
  await page.reload();
  await expect(page.locator('canvas')).toBeVisible();
  await scrollRows(page, 10);
  await expectCell(page, 0, 0, 'A1 Name');

  // Undo takes the freeze back, and redo brings it again.
  await freeze(page, 'Unfreeze panes');
  await scrollRows(page, 10);
  await expectCell(page, 0, 0, 'A11');
  await page.keyboard.press('ControlOrMeta+z');
  await scrollRows(page, 10);
  await expectCell(page, 0, 0, 'A1 Name');
  await page.keyboard.press('ControlOrMeta+y');
  await scrollRows(page, 10);
  await expectCell(page, 0, 0, 'A11');
});

test('freezes the rows above and the columns left of the active cell, or the first column', async ({
  page,
}) => {
  await openApp(page);
  // At A1 there is nothing to freeze.
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: 'Freeze panes' }).click();
  await expect(
    page.getByRole('menuitem', { name: 'Freeze panes', exact: true }),
  ).toHaveAttribute('data-disabled');
  await page.keyboard.press('Escape');

  await clickCell(page, 2, 1);
  await freeze(page, 'Freeze panes');
  await page.locator('canvas').evaluate((canvas) => {
    canvas.parentElement!.parentElement!.scrollTo(960, 240);
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(resolve)),
  );
  await expectCell(page, 1, 0, 'A2');
  await expectCell(page, 2, 1, 'L13');
  await expectCell(page, 2, 0, 'A13');
  await expectCell(page, 1, 1, 'L2');

  await freeze(page, 'Freeze first column');
  await expectCell(page, 1, 0, 'A12');
  await expectCell(page, 1, 1, 'L12');
});

test('the Freeze panes menu is translated', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('button', { name: '冻结窗格' }).click();
  await expect(page.getByRole('menuitem', { name: '冻结首行' })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: '冻结首列' })).toBeVisible();
  await expect(
    page.getByRole('menuitem', { name: '取消冻结窗格' }),
  ).toBeVisible();
});

test.describe('in the dark theme', () => {
  test.use({ colorScheme: 'dark' });

  test('the Freeze panes menu is dark', async ({ page }) => {
    await openApp(page);
    await page.getByRole('tab', { name: 'View' }).click();
    await page.getByRole('button', { name: 'Freeze panes' }).click();
    const menu = page.getByRole('menu');
    await expect(menu).toBeVisible();
    const background = await menu.evaluate(
      (element) => getComputedStyle(element).backgroundColor,
    );
    const [r, g, b] = background.match(/\d+/g)!.map(Number);
    expect(r! + g! + b!).toBeLessThan(3 * 96);
  });
});
