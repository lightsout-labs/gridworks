import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  clickCell,
  enter,
  expectCell,
  expectOpensInPlace,
  openApp,
} from './grid';

const dialog = (page: Page) =>
  page.getByRole('dialog', { name: 'Workbooks', exact: true });

const stored = (page: Page) =>
  page.getByRole('list', { name: 'Workbooks in this browser' });

/** The names of the workbooks the dialog lists, in order. */
const listed = (page: Page) =>
  stored(page)
    .getByRole('listitem')
    .getByRole('button', { name: /^Open / });

const titleBox = (page: Page) => page.getByLabel('Workbook name');

/** Opens the workbook dialog from the File tab. */
async function openDialog(page: Page): Promise<void> {
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Workbooks', exact: true }).click();
  await expect(stored(page)).toBeVisible();
}

async function action(page: Page, workbook: string, item: string) {
  await page
    .getByRole('button', { name: `More actions for ${workbook}` })
    .click();
  await page.getByRole('menuitem', { name: item }).click();
}

/** The element a click at the middle of `control` would reach. */
const reached = (control: Locator) =>
  control.evaluate((node) => {
    const box = node.getBoundingClientRect();
    const top = document.elementFromPoint(
      box.x + box.width / 2,
      box.y + box.height / 2,
    );
    return {
      self: top === node || node.contains(top),
      background: top ? getComputedStyle(top).backgroundColor : '',
    };
  });

test.use({ locale: 'en-US' });

test('the dialog creates, lists, opens, renames, duplicates and deletes workbooks', async ({
  page,
}) => {
  await page.goto('./');
  await expect(page).toHaveTitle('Gridworks');
  await expect(dialog(page).getByText('No workbooks yet')).toBeVisible();

  // A new workbook, named in the title area, which the page title follows.
  await page.getByRole('button', { name: 'New workbook' }).click();
  await expect(dialog(page)).toHaveCount(0);
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

  await openDialog(page);
  await page.getByRole('button', { name: 'New workbook' }).click();
  await expect(titleBox(page)).toHaveValue('Book1');
  await clickCell(page, 0, 0);
  await enter(page, 'two');

  // Most recent first, each with its modification time.
  await openDialog(page);
  await expect(listed(page)).toHaveText([/^Book1/, /^Budget/]);
  await expect(stored(page).getByRole('listitem').first()).toContainText(
    /\d:\d\d\s[AP]M/,
  );

  await listed(page).filter({ hasText: 'Budget' }).click();
  await expect(dialog(page)).toHaveCount(0);
  await expect(titleBox(page)).toHaveValue('Budget');
  await expectCell(page, 0, 0, 'A1 one');

  // Duplicate and rename, in the menu of each workbook.
  await openDialog(page);
  await action(page, 'Budget', 'Duplicate');
  await expect(listed(page)).toHaveText([
    /^Budget \(copy\)/,
    /^Book1/,
    /^Budget/,
  ]);
  await action(page, 'Budget (copy)', 'Rename');
  const rename = page.getByRole('dialog', { name: 'Rename workbook' });
  await rename.getByLabel('Name').fill('Plan');
  await rename.getByRole('button', { name: 'Rename' }).click();
  await expect(listed(page)).toHaveText([/^Plan/, /^Book1/, /^Budget/]);
  await listed(page).filter({ hasText: 'Plan' }).click();
  await expect(page).toHaveTitle('Plan - Gridworks');
  await expectCell(page, 0, 0, 'A1 one');

  // Deleting asks first.
  await openDialog(page);
  await action(page, 'Book1', 'Delete');
  const confirm = page.getByRole('alertdialog', { name: 'Delete workbook' });
  await expect(confirm).toContainText('Book1');
  await confirm.getByRole('button', { name: 'Cancel' }).click();
  await expect(listed(page)).toHaveCount(3);
  await action(page, 'Book1', 'Delete');
  await confirm.getByRole('button', { name: 'Delete' }).click();
  await expect(listed(page)).toHaveText([/^Plan/, /^Budget/]);

  // The open workbook is renamed in place, and cannot be deleted.
  await action(page, 'Plan', 'Rename');
  await rename.getByLabel('Name').fill('Forecast');
  await rename.getByRole('button', { name: 'Rename' }).click();
  await expect(listed(page)).toHaveText([/^Forecast/, /^Budget/]);
  await expect(titleBox(page)).toHaveValue('Forecast');
  await page.getByRole('button', { name: 'More actions for Forecast' }).click();
  await expect(page.getByRole('menuitem', { name: 'Delete' })).toBeDisabled();
  await page.keyboard.press('Escape');

  // The list outlives a reload, and so does an open workbook.
  await page.goto('./');
  await expect(listed(page)).toHaveText([/^Forecast/, /^Budget/]);
  await listed(page).filter({ hasText: 'Budget' }).click();
  await page.reload();
  await expectCell(page, 0, 0, 'A1 one');
  await expect(titleBox(page)).toHaveValue('Budget');
});

test('with no workbook open, the dialog shows over the dimmed, inactive interface and stays open', async ({
  page,
}) => {
  await page.goto('./');
  await expect(dialog(page)).toBeVisible();
  // The ribbon, an empty grid and the status bar are behind it, covered by
  // a dimming layer that takes every click.
  const fileTab = page.getByRole('tab', { name: 'File', includeHidden: true });
  await expect(fileTab).toBeAttached();
  await expect(page.locator('canvas')).toBeAttached();
  await expect(
    page.getByRole('button', { name: 'Zoom in', includeHidden: true }),
  ).toBeAttached();
  const covered = await reached(fileTab);
  expect(covered.self).toBe(false);
  expect(covered.background).toMatch(/rgba\(0, 0, 0, 0\.[1-9]/);
  await page.mouse.click(5, 5);
  await expect(dialog(page)).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).getByRole('button', { name: 'Close' })).toHaveCount(
    0,
  );
});

test('the File tab opens the dialog over the workbook, which Esc, the close button or a click outside return to', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'kept');
  await clickCell(page, 1, 1);
  const closes = [
    () => page.keyboard.press('Escape'),
    () => dialog(page).getByRole('button', { name: 'Close' }).click(),
    () => page.mouse.click(5, 5),
  ];
  for (const close of closes) {
    await openDialog(page);
    await expect(listed(page)).toHaveText([/^Book1/]);
    await close();
    await expect(dialog(page)).toHaveCount(0);
    await expect(titleBox(page)).toHaveValue('Book1');
    await expect(page.getByLabel('Name box')).toHaveValue('B2');
  }
  await expectCell(page, 0, 0, 'A1 kept');
  // Opening the workbook shown returns to it too.
  await openDialog(page);
  await listed(page).click();
  await expect(dialog(page)).toHaveCount(0);
  await expectCell(page, 0, 0, 'A1 kept');
});

test('the dialog opens a file from disk as a new stored workbook', async ({
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
  await openDialog(page);
  await expect(listed(page)).toHaveText([/^Prices/]);
});

test('a file that cannot be read is reported in the dialog', async ({
  page,
}) => {
  await page.goto('./');
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Open file' }).click();
  await (
    await chooser
  ).setFiles({
    name: 'Broken.xlsx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from('not a zip'),
  });
  await expect(dialog(page).getByRole('alert')).toContainText('Broken.xlsx');
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

test('the dialog is translated and follows the theme', async ({ page }) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('button', { name: 'Dark' }).click();
  await page.getByRole('button', { name: '中文' }).click();
  await page.getByRole('tab', { name: '文件' }).click();
  await page.getByRole('button', { name: '工作簿', exact: true }).click();
  const translated = page.getByRole('dialog', { name: '工作簿' });
  await expect(
    translated.getByRole('button', { name: '新建工作簿' }),
  ).toBeVisible();
  await expect(translated.getByRole('button', { name: '关闭' })).toBeVisible();
  const background = await translated.evaluate(
    (node) => getComputedStyle(node).backgroundColor,
  );
  const [red, green, blue] = background.match(/\d+/g)!.map(Number);
  expect((red! + green! + blue!) / 3).toBeLessThan(128);
});

test('the dialog controls have tooltips and show hover, keyboard focus and disabled states', async ({
  page,
  browserName,
}) => {
  const style = (control: Locator, property: string) =>
    control.evaluate(
      (node, property) => getComputedStyle(node).getPropertyValue(property),
      property,
    );
  await openApp(page);
  await openDialog(page);
  const accent = await style(
    dialog(page).getByRole('heading', { name: 'Workbooks' }),
    'color',
  );
  const controls = [
    [
      page.getByRole('button', { name: 'New workbook' }),
      'Start an empty workbook',
    ],
    [
      page.getByRole('button', { name: 'Open file' }),
      'Open an Excel workbook (.xlsx) or a CSV file (.csv) from your computer',
    ],
    [
      page.getByRole('button', { name: 'Close' }),
      'Return to the open workbook',
    ],
    [page.getByRole('button', { name: 'Open Book1' }), 'Open this workbook'],
    [
      page.getByRole('button', { name: 'More actions for Book1' }),
      'Rename, duplicate or delete this workbook',
    ],
  ] as const;
  for (const [control, tooltip] of controls) {
    await page.mouse.move(0, 0);
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
  await expect(dialog(page).getByRole('button')).toHaveCount(controls.length);

  // From the dialog, focused when it opens, the keyboard reaches each
  // control in turn, outlined in the accent color.
  const key = browserName === 'webkit' ? 'Alt+Tab' : 'Tab';
  await expect(dialog(page)).toBeFocused();
  for (const [control] of controls) {
    await page.keyboard.press(key);
    await expect(control).toBeFocused();
    expect(await style(control, 'outline-style')).toBe('solid');
    expect(await style(control, 'outline-color')).toBe(accent);
  }

  // Until the stored workbooks are listed, New workbook is disabled: faded,
  // with the plain cursor. Here storage never answers.
  await page.addInitScript(() => {
    IDBObjectStore.prototype.getAll = () => ({}) as IDBRequest;
  });
  await page.goto('./');
  const create = page.getByRole('button', { name: 'New workbook' });
  await expect(create).toBeDisabled();
  expect(await style(create, 'opacity')).toBe('0.5');
  expect(await style(create, 'cursor')).toBe('default');
});

test('dialogs open in place at the center with a short fade', async ({
  page,
}) => {
  await page.goto('./');
  await expectOpensInPlace(dialog(page));
  await page.getByRole('button', { name: 'New workbook' }).click();
  await expect(page.locator('canvas')).toBeVisible();
  await clickCell(page, 0, 0);
  await enter(page, '3');
  await enter(page, '4');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  const button = (name: string) =>
    page.getByRole('button', { name, exact: true });
  const opens = async (name: string, open: () => Promise<void>) => {
    await open();
    await expectOpensInPlace(page.getByRole('dialog', { name }));
    await page.keyboard.press('Escape');
    await expect(page.getByRole('dialog')).toHaveCount(0);
  };

  await opens('Custom number format', async () => {
    await button('Number format').click();
    await page
      .getByRole('menuitem', { name: 'More number formats...' })
      .click();
  });
  await opens('Sort', async () => {
    await page.getByRole('tab', { name: 'Data' }).click();
    await button('Custom sort').click();
  });
  await opens('Name manager', async () => {
    await page.getByRole('tab', { name: 'Formulas' }).click();
    await button('Name manager').click();
  });
  await opens('Print', () => page.keyboard.press('ControlOrMeta+p'));
  await opens('Workbooks', () => openDialog(page));

  // With no workbook open, as only a workbook not open can be deleted.
  await page.goto('./');
  await action(page, 'Book1', 'Rename');
  await expectOpensInPlace(
    page.getByRole('dialog', { name: 'Rename workbook' }),
  );
  await page.keyboard.press('Escape');
  await action(page, 'Book1', 'Delete');
  await expectOpensInPlace(
    page.getByRole('alertdialog', { name: 'Delete workbook' }),
  );
});
