import { expect, test } from '@playwright/test';
import {
  clickCell,
  enter,
  expectCell,
  openApp,
  openStoredSheet,
} from '../e2e/grid';

test('a change that 100,000 formulas depend on shows within a second, and input goes on', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await openApp(page);
  const rows = 100_000;
  await openStoredSheet(page, {
    columns: [[1]],
    formulas: Array.from({ length: rows }, (_, row) => [
      row,
      1,
      row === 0 ? '=A1+1' : `=B${row}+1`,
    ]),
  });
  await expectCell(page, 0, 1, `B1 2`);

  await clickCell(page, 0, 0);
  await page.keyboard.type('2');
  const start = Date.now();
  await page.keyboard.press('Enter');
  // Typed at once, while the change is still taking effect.
  await enter(page, '7');
  await expectCell(page, 0, 1, 'B1 3');
  expect(Date.now() - start).toBeLessThan(1000);
  await expectCell(page, 1, 0, 'A2 7');

  // The end of the chain, which depends on every formula before it.
  await clickCell(page, 0, 1);
  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await expect(page.getByRole('status')).toHaveText(`B${rows} ${rows + 2}`);
});
