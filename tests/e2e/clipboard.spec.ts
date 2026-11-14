import { expect, test, type Page } from '@playwright/test';
import { cellRect } from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

/** The element that scrolls the grid, which holds the canvas. */
const grid = (page: Page) => page.locator('canvas').locator('../..');
const formulaBar = (page: Page) =>
  page.getByRole('textbox', { name: 'Formula bar' });

/** Types entries into a row of cells, from the active one rightwards. */
async function typeRow(page: Page, ...entries: string[]) {
  for (const entry of entries) {
    await page.keyboard.type(entry);
    await page.keyboard.press('Tab');
  }
}

/** Right-clicks a cell, opening the cell menu. */
async function rightClickCell(page: Page, row: number, col: number) {
  const rect = cellRect({ row, col }, { scrollX: 0, scrollY: 0, zoom: 1 });
  await page.locator('canvas').click({
    button: 'right',
    position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
  });
}

test('copies and pastes with the keyboard, moving references and carrying formats, with undo', async ({
  page,
}) => {
  await openApp(page);
  const paste = page.getByRole('button', { name: 'Paste', exact: true });
  await expect(paste).toBeDisabled();
  await clickCell(page, 0, 0);
  await typeRow(page, '12%', '=A1*2', '=$A$1');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('ControlOrMeta+c');
  // A moving border shows the cells copied.
  await expect(grid(page)).toHaveAttribute('data-copied', 'A1:C1');
  const canvas = page.locator('canvas');
  const still = await canvas.screenshot();
  await expect
    .poll(async () => (await canvas.screenshot()).equals(still))
    .toBe(false);
  await expect(paste).toBeEnabled();

  await clickCell(page, 2, 0);
  await page.keyboard.press('ControlOrMeta+v');
  await expectCell(page, 2, 0, 'A3 12%');
  await expect(formulaBar(page)).toHaveValue('12%');
  await expectCell(page, 2, 1, 'B3 0.24');
  await expect(formulaBar(page)).toHaveValue('=A3*2');
  await clickCell(page, 2, 2);
  await expect(formulaBar(page)).toHaveValue('=$A$1');
  // What was copied can be pasted again until Escape.
  await expect(grid(page)).toHaveAttribute('data-copied', 'A1:C1');
  await page.keyboard.press('Escape');
  await expect(grid(page)).not.toHaveAttribute('data-copied');
  await expect(paste).toBeDisabled();

  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.getByRole('status')).toHaveText('A3');
  await expectCell(page, 2, 1, 'B3');
  await page.keyboard.press('ControlOrMeta+y');
  await expectCell(page, 2, 1, 'B3 0.24');
});

test('cuts with the ribbon and pastes with the cell menu, moving the cells and the references to them', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '5');
  await clickCell(page, 0, 2);
  await enter(page, '=A1+1');
  await clickCell(page, 0, 0);
  await page.getByRole('button', { name: 'Cut' }).click();
  await expect(grid(page)).toHaveAttribute('data-copied', 'A1');

  // Right-clicking a cell outside the selection selects it.
  await rightClickCell(page, 3, 1);
  await page.getByRole('menuitem', { name: 'Paste', exact: true }).click();
  await expect(page.getByRole('status')).toHaveText('B4 5');
  await expectCell(page, 0, 0, 'A1');
  await expectCell(page, 0, 2, 'C1 6');
  await expect(formulaBar(page)).toHaveValue('=B4+1');
  // Cut cells are pasted once.
  await expect(grid(page)).not.toHaveAttribute('data-copied');
  await expect(
    page.getByRole('button', { name: 'Paste', exact: true }),
  ).toBeDisabled();

  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 0, 'A1 5');
  await expectCell(page, 0, 2, 'C1 6');
  await expect(formulaBar(page)).toHaveValue('=A1+1');

  // The keyboard cuts too, and the ribbon pastes.
  await clickCell(page, 0, 0);
  await page.keyboard.press('ControlOrMeta+x');
  await clickCell(page, 1, 0);
  await page.getByRole('button', { name: 'Paste', exact: true }).click();
  await expectCell(page, 0, 2, 'C1 6');
  await expect(formulaBar(page)).toHaveValue('=A2+1');

  // The cell menu cuts, and copies.
  await rightClickCell(page, 1, 0);
  await page.getByRole('menuitem', { name: 'Cut' }).click();
  await expect(grid(page)).toHaveAttribute('data-copied', 'A2');
  await rightClickCell(page, 0, 2);
  await page.getByRole('menuitem', { name: 'Copy' }).click();
  await expect(grid(page)).toHaveAttribute('data-copied', 'C1');
});

test('fills a selection with one cell, and pastes only values, formulas or formatting', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await typeRow(page, '10%', '=A1*2');
  await clickCell(page, 8, 0);
  await enter(page, '0.5');

  // One cell pasted into a larger selection fills it.
  await clickCell(page, 0, 1);
  await page.getByRole('button', { name: 'Copy' }).click();
  await clickCell(page, 0, 3);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('ControlOrMeta+v');
  await clickCell(page, 2, 3);
  await expect(formulaBar(page)).toHaveValue('=C3*2');

  // Values only, from the ribbon.
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('ControlOrMeta+c');
  await clickCell(page, 4, 0);
  await page.getByRole('button', { name: 'Paste special' }).click();
  await page.getByRole('menuitem', { name: 'Values' }).click();
  await expectCell(page, 4, 1, 'B5 0.2');
  await expect(formulaBar(page)).toHaveValue('0.2');

  // Formulas only, from the cell menu, without the format of A1.
  await rightClickCell(page, 6, 0);
  await page.getByRole('menuitem', { name: 'Paste special' }).click();
  await page.getByRole('menuitem', { name: 'Formulas' }).click();
  await expectCell(page, 6, 0, 'A7 0.1');
  await expectCell(page, 6, 1, 'B7 0.2');
  await expect(formulaBar(page)).toHaveValue('=A7*2');

  // Formatting only: the cell keeps its contents in the format of A1.
  await rightClickCell(page, 8, 0);
  await page.getByRole('menuitem', { name: 'Paste special' }).click();
  await page.getByRole('menuitem', { name: 'Formatting' }).click();
  await expectCell(page, 8, 0, 'A9 50%');
});
