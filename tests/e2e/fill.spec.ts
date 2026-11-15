import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

const TOP = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };

/** Where the fill handle of a range ending at a cell is, on the page. */
async function handle(page: Page, bottom: number, right: number) {
  const box = (await page.locator('canvas').boundingBox())!;
  const corner = cellRect({ row: bottom + 1, col: right + 1 }, TOP);
  return { x: box.x + corner.x, y: box.y + corner.y };
}

/** Drags the fill handle of the range ending at a cell to the middle of another. */
async function dragFill(
  page: Page,
  from: [row: number, col: number],
  to: [row: number, col: number],
): Promise<void> {
  const start = await handle(page, ...from);
  const box = (await page.locator('canvas').boundingBox())!;
  const end = cellRect({ row: to[0], col: to[1] }, TOP);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(
    box.x + end.x + end.width / 2,
    box.y + end.y + end.height / 2,
    { steps: 5 },
  );
  await page.mouse.up();
}

test('dragging the fill handle extends series and copies formulas, and undo takes it back', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await clickCell(page, 0, 1);
  await enter(page, '=A1*10');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  await dragFill(page, [1, 0], [4, 0]);
  await expectCell(page, 4, 0, 'A5 5');

  await clickCell(page, 0, 1);
  await dragFill(page, [0, 1], [3, 1]);
  await expectCell(page, 3, 1, 'B4 40');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(
    '=A4*10',
  );

  // Weekday names, to the right.
  await clickCell(page, 6, 0);
  await enter(page, 'Mon');
  await clickCell(page, 6, 0);
  await dragFill(page, [6, 0], [6, 2]);
  await expectCell(page, 6, 2, 'C7 Wed');

  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 6, 2, 'C7');
  await page.keyboard.press('ControlOrMeta+y');
  await expectCell(page, 6, 2, 'C7 Wed');
});

test('double-clicking the fill handle fills down beside the data', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  for (const n of ['3', '4', '5']) await enter(page, n);
  await clickCell(page, 0, 1);
  await enter(page, '=A1+1');
  await clickCell(page, 0, 1);
  const at = await handle(page, 0, 1);
  await page.mouse.dblclick(at.x, at.y);
  await expectCell(page, 2, 1, 'B3 6');
  await expectCell(page, 3, 1, 'B4');
});

test('Ctrl/Cmd+D and Ctrl/Cmd+R fill down and right', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '7');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('ControlOrMeta+d');
  await expectCell(page, 2, 0, 'A3 7');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ControlOrMeta+r');
  await expect(page.getByRole('status')).toHaveText('B3 7');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.getByRole('status')).toHaveText('B3');
});
