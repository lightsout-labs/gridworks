import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  clickCell,
  enter,
  expectCell,
  expectOpensInPlace,
  openApp,
} from './grid';

const stored = (page: Page) =>
  page.getByRole('list', { name: 'Workbooks in this browser' });

/** The names of the workbooks the start screen lists, in order. */
const listed = (page: Page) =>
  stored(page)
    .getByRole('listitem')
    .getByRole('button', { name: /^Open / });

const titleBox = (page: Page) => page.getByLabel('Workbook name');

async function goHome(page: Page): Promise<void> {
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Start screen' }).click();
  await expect(stored(page)).toBeVisible();
}

async function action(page: Page, workbook: string, item: string) {
  await page
    .getByRole('button', { name: `More actions for ${workbook}` })
    .click();
  await page.getByRole('menuitem', { name: item }).click();
}

test.use({ locale: 'en-US' });

test('the start screen creates, lists, opens, renames, duplicates and deletes workbooks', async ({
  page,
}) => {
  await page.goto('./');
  await expect(page).toHaveTitle('Gridworks');
  await expect(page.getByText('No workbooks yet')).toBeVisible();

  // A new workbook, named in the title area, which the page title follows.
  await page.getByRole('button', { name: 'New workbook' }).click();
  await expect(page.locator('canvas')).toBeVisible();
  await expect(titleBox(page)).toHaveValue('Book1');
  await expect(page).toHaveTitle('Book1 - Gridworks');
  await clickCell(page, 0, 0);
  await enter(page, 'one');
  await titleBox(page).fill('Budget');
  await titleBox(page).press('Enter');
  await expect(page).toHaveTitle('Budget - Gridworks');
  // An empty name keeps the one before.
  await titleBox(page).fill('');
  await titleBox(page).press('Enter');
  await expect(titleBox(page)).toHaveValue('Budget');

  await goHome(page);
  await expect(page).toHaveTitle('Gridworks');
  await page.getByRole('button', { name: 'New workbook' }).click();
  await expect(titleBox(page)).toHaveValue('Book1');
  await clickCell(page, 0, 0);
  await enter(page, 'two');

  // Most recent first, each with its modification time.
  await goHome(page);
  await expect(listed(page)).toHaveText([/^Book1/, /^Budget/]);
  await expect(stored(page).getByRole('listitem').first()).toContainText(
    /\d:\d\d\s[AP]M/,
  );

  await listed(page).filter({ hasText: 'Budget' }).click();
  await expect(titleBox(page)).toHaveValue('Budget');
  await expectCell(page, 0, 0, 'A1 one');

  // Duplicate and rename, in the menu of each workbook.
  await goHome(page);
  await action(page, 'Budget', 'Duplicate');
  await expect(listed(page)).toHaveText([
    /^Budget \(copy\)/,
    /^Book1/,
    /^Budget/,
  ]);
  await action(page, 'Budget (copy)', 'Rename');
  const dialog = page.getByRole('dialog', { name: 'Rename workbook' });
  await dialog.getByLabel('Name').fill('Plan');
  await dialog.getByRole('button', { name: 'Rename' }).click();
  await expect(listed(page)).toHaveText([/^Plan/, /^Book1/, /^Budget/]);
  await listed(page).filter({ hasText: 'Plan' }).click();
  await expect(page).toHaveTitle('Plan - Gridworks');
  await expectCell(page, 0, 0, 'A1 one');

  // Deleting asks first.
  await goHome(page);
  await action(page, 'Book1', 'Delete');
  const confirm = page.getByRole('alertdialog', { name: 'Delete workbook' });
  await expect(confirm).toContainText('Book1');
  await confirm.getByRole('button', { name: 'Cancel' }).click();
  await expect(listed(page)).toHaveCount(3);
  await action(page, 'Book1', 'Delete');
  await confirm.getByRole('button', { name: 'Delete' }).click();
  await expect(listed(page)).toHaveText([/^Plan/, /^Budget/]);

  // The list outlives a reload, and so does an open workbook.
  await page.reload();
  await expect(listed(page)).toHaveText([/^Plan/, /^Budget/]);
  await listed(page).filter({ hasText: 'Budget' }).click();
  await page.reload();
  await expectCell(page, 0, 0, 'A1 one');
  await expect(titleBox(page)).toHaveValue('Budget');
});

test('the start screen opens a file from disk as a new stored workbook', async ({
  page,
}) => {
  await page.goto('./');
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Open file' }).click();
  await (
    await chooser
  ).setFiles({
    name: 'Prices.csv',
    mimeType: 'text/csv',
    buffer: Buffer.from('Apple,3\r\n'),
  });
  await expect(titleBox(page)).toHaveValue('Prices');
  await expectCell(page, 0, 1, 'B1 3');
  await goHome(page);
  await expect(listed(page)).toHaveText([/^Prices/]);
});

test('a failure to save in the browser is shown', async ({ page }) => {
  await openApp(page);
  // Storage refuses every write, as when it is full.
  await page.evaluate(() => {
    IDBObjectStore.prototype.put = () => {
      throw new DOMException('Storage is full', 'QuotaExceededError');
    };
  });
  await clickCell(page, 0, 0);
  await enter(page, 'lost?');
  await expect(page.getByRole('alert')).toContainText(
    'could not be saved in this browser',
  );
});

test('the start screen is translated and follows the theme', async ({
  page,
}) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('button', { name: 'Dark' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '文件' }).click();
  await page.getByRole('button', { name: '开始屏幕' }).click();
  await expect(page.getByRole('button', { name: '新建工作簿' })).toBeVisible();
  const background = await page
    .getByRole('main')
    .evaluate((node) => getComputedStyle(node).backgroundColor);
  const [red, green, blue] = background.match(/\d+/g)!.map(Number);
  expect((red! + green! + blue!) / 3).toBeLessThan(128);
});

test('the start screen controls have tooltips and show hover, keyboard focus and disabled states', async ({
  page,
  browserName,
}) => {
  const style = (control: Locator, property: string) =>
    control.evaluate(
      (node, property) => getComputedStyle(node).getPropertyValue(property),
      property,
    );
  await openApp(page);
  await goHome(page);
  const controls = [
    [
      page.getByRole('button', { name: 'New workbook' }),
      'Start an empty workbook',
    ],
    [
      page.getByRole('button', { name: 'Open file' }),
      'Open an Excel workbook (.xlsx) or a CSV file (.csv) from your computer',
    ],
    [page.getByRole('button', { name: 'Open Book1' }), 'Open this workbook'],
    [
      page.getByRole('button', { name: 'More actions for Book1' }),
      'Rename, duplicate or delete this workbook',
    ],
  ] as const;
  const accent = await style(
    page.getByRole('heading', { name: 'Gridworks' }),
    'color',
  );
  for (const [control, tooltip] of controls) {
    await page.mouse.move(0, 700);
    // The primary button brightens; the others take a background.
    const look = async () =>
      `${await style(control, 'background-color')} ${await style(control, 'filter')}`;
    const resting = await look();
    await control.hover();
    await expect(page.getByRole('tooltip', { name: tooltip })).toBeVisible();
    await expect.poll(look).not.toBe(resting);
    expect(parseFloat(await style(control, 'border-radius'))).toBeGreaterThan(
      0,
    );
    const duration = parseFloat(await style(control, 'transition-duration'));
    expect(duration).toBeGreaterThan(0);
    expect(duration).toBeLessThanOrEqual(0.2);
  }
  await expect(page.getByRole('button')).toHaveCount(controls.length);

  // The keyboard reaches each control in turn, outlined in the accent color.
  const key = browserName === 'webkit' ? 'Alt+Tab' : 'Tab';
  await page.mouse.click(0, 700);
  await page.keyboard.press(key);
  for (const [control] of controls) {
    await expect(control).toBeFocused();
    expect(await style(control, 'outline-style')).toBe('solid');
    expect(await style(control, 'outline-color')).toBe(accent);
    await page.keyboard.press(key);
  }

  // Until the stored workbooks are listed, New workbook is disabled: faded,
  // with the plain cursor. Here storage never answers.
  await page.addInitScript(() => {
    IDBObjectStore.prototype.getAll = () => ({}) as IDBRequest;
  });
  await page.reload();
  const create = page.getByRole('button', { name: 'New workbook' });
  await expect(create).toBeDisabled();
  expect(await style(create, 'opacity')).toBe('0.5');
  expect(await style(create, 'cursor')).toBe('default');
});

test('dialogs open in place at the center with a short fade', async ({
  page,
}) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'Formulas' }).click();
  await page.getByRole('button', { name: 'Name manager' }).click();
  const names = page.getByRole('dialog', { name: 'Name manager' });
  await expectOpensInPlace(names);
  await page.keyboard.press('Escape');

  await goHome(page);
  await action(page, 'Book1', 'Delete');
  await expectOpensInPlace(
    page.getByRole('alertdialog', { name: 'Delete workbook' }),
  );
});
