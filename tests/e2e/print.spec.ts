import { expect, test, type Page } from '@playwright/test';
import { clickCell, enter, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

// The browser's print dialog cannot be driven, so each frame counts its
// calls to print instead of opening it.
test.beforeEach(async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(window, 'print', {
      value: () => {
        const top = window.top as unknown as { printed?: number };
        top.printed = (top.printed ?? 0) + 1;
      },
    });
  });
});

const panel = (page: Page) => page.getByRole('dialog');
const printed = (page: Page) => page.frameLocator('iframe');
const prints = (page: Page) =>
  page.evaluate(() => (window as unknown as { printed?: number }).printed);

/** Types texts into the first column, from A1 down. */
async function typeColumn(page: Page, texts: string[]): Promise<void> {
  await clickCell(page, 0, 0);
  for (const text of texts) await enter(page, text);
}

test('Ctrl+P prints the sheet with the page chosen', async ({ page }) => {
  await openApp(page);
  await typeColumn(page, ['Fruit', 'apple', 'pear']);
  await page.keyboard.press('ControlOrMeta+p');
  await expect(panel(page)).toHaveAccessibleName('Print');
  await panel(page).getByLabel('Orientation').selectOption('Landscape');
  await panel(page).getByLabel('Paper size').selectOption('A4 (210 x 297 mm)');
  await panel(page).getByLabel('Margins').selectOption('Narrow');
  await panel(page).getByLabel('Gridlines').check();
  await panel(page).getByLabel('Row and column headings').check();
  await panel(page).getByRole('button', { name: 'Print', exact: true }).click();

  await expect(panel(page)).toBeHidden();
  await expect.poll(() => prints(page)).toBe(1);
  const table = printed(page).locator('table');
  await expect(table.locator('td')).toHaveText(['Fruit', 'apple', 'pear']);
  await expect(table.locator('thead th')).toHaveText(['', 'A']);
  expect(await printed(page).locator('style').textContent()).toContain(
    'size: a4 landscape; margin: 0.75in 0.25in 0.75in 0.25in;',
  );
});

test('the ribbon prints the selection only', async ({ page }) => {
  await openApp(page);
  await typeColumn(page, ['one', 'two', 'three']);
  await clickCell(page, 1, 0);
  await page.getByRole('tab', { name: 'File', exact: true }).click();
  await page.getByRole('button', { name: 'Print', exact: true }).click();
  await panel(page)
    .getByLabel('Print', { exact: true })
    .selectOption('Selection');
  await panel(page).getByRole('button', { name: 'Print', exact: true }).click();
  await expect.poll(() => prints(page)).toBe(1);
  await expect(printed(page).locator('td')).toHaveText(['two']);
});

test('an empty sheet has nothing to print', async ({ page }) => {
  await openApp(page);
  await page.keyboard.press('ControlOrMeta+p');
  await expect(panel(page)).toContainText('There is nothing to print.');
  await expect(
    panel(page).getByRole('button', { name: 'Print', exact: true }),
  ).toBeDisabled();
});

test('the print panel is in Chinese', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.keyboard.press('ControlOrMeta+p');
  await expect(panel(page)).toHaveAccessibleName('打印');
  await expect(panel(page).getByLabel('纸张大小')).toBeVisible();
});

test.describe('in the dark theme', () => {
  test.use({ colorScheme: 'dark' });

  test('the print panel is dark', async ({ page }) => {
    await openApp(page);
    await page.keyboard.press('ControlOrMeta+p');
    for (const element of [panel(page), panel(page).getByLabel('Margins')]) {
      const background = await element.evaluate(
        (each) => getComputedStyle(each).backgroundColor,
      );
      const [r, g, b] = background.match(/\d+/g)!.map(Number);
      expect(r! + g! + b!).toBeLessThan(3 * 96);
    }
  });
});
