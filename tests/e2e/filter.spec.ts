import { expect, test, type Page } from '@playwright/test';
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

const TABLE = [
  ['Fruit', 'Price'],
  ['apple', '3'],
  ['banana', '12'],
  ['cherry', '7'],
];

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/** The cell the arrow down from A1 goes to, over the rows hidden. */
async function belowHeader(page: Page): Promise<string> {
  await clickCell(page, 0, 0);
  await page.keyboard.press('ArrowDown');
  return (await page.getByRole('status').textContent()) ?? '';
}

test('filters rows by values and conditions, undoes it, and clears it', async ({
  page,
}) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'Data' }).click();
  await clickCell(page, 5, 5);
  await expect(button(page, 'Filter')).toBeDisabled();
  await typeRows(page, TABLE);
  await clickCell(page, 1, 0);
  await button(page, 'Filter').click();
  await expect(button(page, 'Filter')).toHaveAttribute('aria-pressed', 'true');
  await expect(button(page, 'Filter column A')).toBeVisible();
  await expect(button(page, 'Filter column B')).toBeVisible();

  // The checklist, narrowed by the search box.
  await button(page, 'Filter column A').click();
  const panel = page.getByRole('dialog', { name: 'Filter column A' });
  await expect(panel.getByRole('checkbox')).toHaveCount(4);
  await panel.getByLabel('Search').fill('an');
  await expect(panel.getByRole('checkbox', { name: 'banana' })).toBeVisible();
  await expect(panel.getByRole('checkbox', { name: 'apple' })).toBeHidden();
  await panel.getByLabel('Search').fill('');
  await panel.getByRole('checkbox', { name: 'apple' }).uncheck();
  await button(page, 'OK').click();
  await expect(panel).toBeHidden();
  expect(await belowHeader(page)).toBe('A3 banana');
  // Row 2 is gone from the grid: the second row shown is row 3.
  await expectCell(page, 1, 0, 'A3 banana');
  await expect(button(page, 'Filter column A (filtered)')).toBeVisible();

  // A number condition on another column combines with it.
  await button(page, 'Filter column B').click();
  const prices = page.getByRole('dialog', { name: 'Filter column B' });
  await prices.getByLabel('Show rows where').selectOption('Less than');
  await prices.getByLabel('Value', { exact: true }).fill('10');
  await button(page, 'OK').click();
  expect(await belowHeader(page)).toBe('A4 cherry');

  await page.keyboard.press('ControlOrMeta+z');
  expect(await belowHeader(page)).toBe('A3 banana');
  await page.keyboard.press('ControlOrMeta+y');
  expect(await belowHeader(page)).toBe('A4 cherry');

  await button(page, 'Clear').click();
  expect(await belowHeader(page)).toBe('A2 apple');
  await expect(button(page, 'Filter column A')).toBeVisible();
  await expect(button(page, 'Clear')).toBeDisabled();

  await button(page, 'Filter').click();
  await expect(button(page, 'Filter column A')).toBeHidden();
  await expectCell(page, 2, 0, 'A3 banana');
});

test('filters by a text condition, and clears one column from its dropdown', async ({
  page,
}) => {
  await openApp(page);
  await typeRows(page, TABLE);
  await clickCell(page, 0, 0);
  await page.getByRole('tab', { name: 'Data' }).click();
  await button(page, 'Filter').click();
  await button(page, 'Filter column A').click();
  const panel = page.getByRole('dialog', { name: 'Filter column A' });
  await panel.getByLabel('Show rows where').selectOption('Begins with');
  await panel.getByLabel('Value', { exact: true }).fill('CH');
  await button(page, 'OK').click();
  expect(await belowHeader(page)).toBe('A4 cherry');

  await button(page, 'Filter column A (filtered)').click();
  await button(page, 'Clear filter').click();
  expect(await belowHeader(page)).toBe('A2 apple');
});

test('translates the filter controls', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '数据' }).click();
  for (const name of ['筛选', '清除']) {
    await expect(button(page, name)).toBeVisible();
  }
});

test.describe('in the dark theme', () => {
  test.use({ colorScheme: 'dark' });

  test('the filter dropdown is dark', async ({ page }) => {
    await openApp(page);
    await typeRows(page, [['Fruit'], ['apple']]);
    await clickCell(page, 0, 0);
    await page.getByRole('tab', { name: 'Data' }).click();
    await button(page, 'Filter').click();
    await button(page, 'Filter column A').click();
    const panel = page.getByRole('dialog', { name: 'Filter column A' });
    await expect(panel).toBeVisible();
    for (const element of [panel, panel.getByLabel('Search')]) {
      const background = await element.evaluate(
        (each) => getComputedStyle(each).backgroundColor,
      );
      const [r, g, b] = background.match(/\d+/g)!.map(Number);
      expect(r! + g! + b!).toBeLessThan(3 * 96);
    }
  });
});
