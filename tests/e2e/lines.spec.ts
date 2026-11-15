import { expect, test, type Page } from '@playwright/test';
import {
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

/** The middle of a row or column header, on the canvas, at the standard sizes. */
function header(kind: 'row' | 'column', index: number) {
  return kind === 'row'
    ? {
        x: ROW_HEADER_WIDTH / 2,
        y: COLUMN_HEADER_HEIGHT + (index + 0.5) * ROW_HEIGHT,
      }
    : {
        x: ROW_HEADER_WIDTH + (index + 0.5) * COLUMN_WIDTH,
        y: COLUMN_HEADER_HEIGHT / 2,
      };
}

/** Selects the rows or columns from one header to another by dragging. */
async function selectHeaders(
  page: Page,
  kind: 'row' | 'column',
  from: number,
  to = from,
): Promise<void> {
  const box = (await page.locator('canvas').boundingBox())!;
  const start = header(kind, from);
  const end = header(kind, to);
  await page.mouse.move(box.x + start.x, box.y + start.y);
  await page.mouse.down();
  await page.mouse.move(box.x + end.x, box.y + end.y, { steps: 3 });
  await page.mouse.up();
}

/** Right-clicks a header and chooses an action of its menu. */
async function headerMenu(
  page: Page,
  kind: 'row' | 'column',
  index: number,
  action: string,
): Promise<void> {
  await page
    .locator('canvas')
    .click({ button: 'right', position: header(kind, index) });
  await page.getByRole('menuitem', { name: action, exact: true }).click();
}

test('inserts and deletes rows and columns from the header menus, with the references to them, and undo', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  for (const value of ['1', '2', '3']) await enter(page, value);
  await clickCell(page, 0, 1);
  await enter(page, '=A3*10');

  // Two rows selected, two rows inserted above them.
  await selectHeaders(page, 'row', 1, 2);
  await headerMenu(page, 'row', 2, 'Insert rows');
  await expectCell(page, 4, 0, 'A5 3');
  await expectCell(page, 0, 1, 'B1 30');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(
    '=A5*10',
  );

  // A row that is not selected is selected by its menu.
  await headerMenu(page, 'row', 4, 'Delete rows');
  await expectCell(page, 0, 1, 'B1 #REF!');
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 1, 'B1 30');

  await headerMenu(page, 'column', 0, 'Insert columns');
  await expectCell(page, 0, 2, 'C1 30');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(
    '=B5*10',
  );
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 1, 'B1 30');
  await page.keyboard.press('ControlOrMeta+y');
  await expectCell(page, 0, 2, 'C1 30');
});

test('resizes a column by dragging its header border, fits it with a double-click, and undoes it', async ({
  page,
}) => {
  await openApp(page);
  const canvas = page.locator('canvas');
  const box = (await canvas.boundingBox())!;
  const border = {
    x: box.x + ROW_HEADER_WIDTH + COLUMN_WIDTH,
    y: box.y + COLUMN_HEADER_HEIGHT / 2,
  };
  // Column A twice as wide: its old right half and more are still A.
  await page.mouse.move(border.x, border.y);
  await page.mouse.down();
  await page.mouse.move(border.x + COLUMN_WIDTH, border.y, { steps: 5 });
  await page.mouse.up();
  const inA = {
    x: ROW_HEADER_WIDTH + COLUMN_WIDTH * 1.5,
    y: COLUMN_HEADER_HEIGHT + ROW_HEIGHT / 2,
  };
  await canvas.click({ position: inA });
  await expect(page.getByRole('status')).toHaveText('A1');

  await page.keyboard.press('ControlOrMeta+z');
  await canvas.click({ position: inA });
  await expect(page.getByRole('status')).toHaveText('B1');

  // Fitted to a long text, column A covers where B was.
  await clickCell(page, 0, 0);
  await enter(page, 'A rather long piece of text for one column');
  await page.mouse.dblclick(border.x, border.y);
  await canvas.click({ position: inA });
  await expect(page.getByRole('status')).toHaveText(
    'A1 A rather long piece of text for one column',
  );

  // Dragged across several selected rows, the border resizes them all.
  await selectHeaders(page, 'row', 1, 2);
  const rowBorder = {
    x: box.x + ROW_HEADER_WIDTH / 2,
    y: box.y + COLUMN_HEADER_HEIGHT + 3 * ROW_HEIGHT,
  };
  await page.mouse.move(rowBorder.x, rowBorder.y);
  await page.mouse.down();
  await page.mouse.move(rowBorder.x, rowBorder.y + ROW_HEIGHT, { steps: 5 });
  await page.mouse.up();
  // Rows 2 and 3 are each twice as tall, so row 4 starts after five rows.
  await canvas.click({
    position: {
      x: ROW_HEADER_WIDTH + 5,
      y: COLUMN_HEADER_HEIGHT + 5.5 * ROW_HEIGHT,
    },
  });
  await expect(page.getByRole('status')).toHaveText('A4');
});

test('hides columns and shows them again from the header menus', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 1);
  await enter(page, 'b');
  await selectHeaders(page, 'column', 1, 2);
  await headerMenu(page, 'column', 1, 'Hide columns');
  // Where B was, D is now.
  await clickCell(page, 0, 1);
  await expect(page.getByRole('status')).toHaveText('D1');

  // Selected across the hidden columns, they are shown again.
  await selectHeaders(page, 'column', 0, 1);
  await headerMenu(page, 'column', 0, 'Unhide columns');
  await expectCell(page, 0, 1, 'B1 b');
});

test.describe('in Chinese', () => {
  test.use({ locale: 'zh-CN' });

  test('names the header menu actions in Chinese', async ({ page }) => {
    await openApp(page);
    await page
      .locator('canvas')
      .click({ button: 'right', position: header('row', 3) });
    await expect(page.getByRole('menuitem', { name: '插入行' })).toBeVisible();
    await expect(
      page.getByRole('menuitem', { name: '取消隐藏行' }),
    ).toBeVisible();
  });
});
