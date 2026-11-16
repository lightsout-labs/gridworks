import { expect, test, type Page } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

/** Types texts into the cells of the first column, from A1 down. */
async function typeColumn(page: Page, texts: string[]): Promise<void> {
  await clickCell(page, 0, 0);
  for (const text of texts) await enter(page, text);
}

const panel = (page: Page) => page.getByRole('dialog');
const button = (page: Page, name: string) =>
  panel(page).getByRole('button', { name, exact: true });
const announced = (page: Page) => page.getByRole('status');

test('Ctrl+F moves through the matches and counts them', async ({ page }) => {
  await openApp(page);
  await typeColumn(page, ['apple', 'pear', 'Apple pie']);
  await clickCell(page, 0, 0);
  await page.keyboard.press('ControlOrMeta+f');
  await expect(panel(page)).toHaveAccessibleName('Find');
  await expect(panel(page).getByLabel('Replace with')).toHaveCount(0);
  await expect(button(page, 'Find next')).toBeDisabled();
  await panel(page).getByLabel('Find what').fill('apple');
  await page.keyboard.press('Enter');
  await expect(announced(page)).toHaveText('A3 Apple pie');
  await expect(panel(page)).toContainText('2 of 2');
  await page.keyboard.press('Enter');
  await expect(announced(page)).toHaveText('A1 apple');
  await expect(panel(page)).toContainText('1 of 2');

  await panel(page).getByLabel('Match case').check();
  await panel(page).getByLabel('Match entire cell contents').check();
  await button(page, 'Find next').click();
  await expect(panel(page)).toContainText('1 of 1');
  await panel(page).getByLabel('Find what').fill('plum');
  await button(page, 'Find next').click();
  await expect(panel(page)).toContainText('No matches found');

  await page.keyboard.press('Escape');
  await expect(panel(page)).toHaveCount(0);
});

test('searches the whole workbook, in values or in formulas', async ({
  page,
}) => {
  await openApp(page);
  await typeColumn(page, ['=1+1']);
  await page.getByRole('button', { name: 'New sheet' }).click();
  await typeColumn(page, ['x', '2']);
  await page.getByRole('tab', { name: 'Sheet1', exact: true }).click();
  await page.keyboard.press('ControlOrMeta+f');
  await panel(page).getByLabel('Find what').fill('2');
  await panel(page).getByLabel('Within').selectOption('Workbook');
  await panel(page).getByLabel('Look in').selectOption('Values');
  await page.keyboard.press('Enter');
  await expect(announced(page)).toHaveText('A2 2');
  await expect(
    page.getByRole('tab', { name: 'Sheet2', exact: true }),
  ).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Enter');
  await expect(announced(page)).toHaveText('A1 2');
  await expect(panel(page)).toContainText('1 of 2');
  await panel(page).getByLabel('Look in').selectOption('Formulas');
  await button(page, 'Find next').click();
  await expect(panel(page)).toContainText('1 of 1');
});

test('Ctrl+H replaces one match and all of them, and undo restores them', async ({
  page,
}) => {
  await openApp(page);
  await typeColumn(page, ['apple', 'apple tart', 'apple pie']);
  await clickCell(page, 0, 0);
  await page.keyboard.press('ControlOrMeta+h');
  await expect(panel(page)).toHaveAccessibleName('Replace');
  await expect(panel(page).getByLabel('Look in')).toBeDisabled();
  await panel(page).getByLabel('Find what').fill('apple');
  await panel(page).getByLabel('Replace with').fill('pear');
  await button(page, 'Replace').click();
  await expect(announced(page)).toHaveText('A2 apple tart');
  await button(page, 'Replace').click();
  await expect(announced(page)).toHaveText('A3 apple pie');
  await button(page, 'Replace all').click();
  await expect(panel(page)).toContainText('Replaced: 1');
  await page.keyboard.press('Escape');

  await expectCell(page, 0, 0, 'A1 pear');
  await expectCell(page, 1, 0, 'A2 pear tart');
  await expectCell(page, 2, 0, 'A3 pear pie');
  // Replace all is one change, and each single replace another.
  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 2, 0, 'A3 apple pie');
  await expectCell(page, 0, 0, 'A1 pear');
  await page.keyboard.press('ControlOrMeta+y');
  await expectCell(page, 2, 0, 'A3 pear pie');
});

test('the ribbon opens find and replace, in Chinese too', async ({ page }) => {
  await openApp(page);
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await expect(panel(page)).toHaveAccessibleName('Find');
  await page.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(panel(page)).toHaveAccessibleName('Replace');
  await expect(panel(page).getByLabel('Find what')).toBeFocused();
  await button(page, 'Close').click();

  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '开始' }).click();
  await page.getByRole('button', { name: '替换', exact: true }).click();
  await expect(panel(page)).toHaveAccessibleName('替换');
  await expect(button(page, '全部替换')).toBeVisible();
});

test.describe('in the dark theme', () => {
  test.use({ colorScheme: 'dark' });

  test('the find panel is dark', async ({ page }) => {
    await openApp(page);
    await page.keyboard.press('ControlOrMeta+f');
    for (const element of [panel(page), panel(page).getByLabel('Find what')]) {
      const background = await element.evaluate(
        (each) => getComputedStyle(each).backgroundColor,
      );
      const [r, g, b] = background.match(/\d+/g)!.map(Number);
      expect(r! + g! + b!).toBeLessThan(3 * 96);
    }
  });
});
