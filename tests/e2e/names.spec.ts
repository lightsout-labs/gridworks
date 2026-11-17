import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  COLUMN_HEADER_HEIGHT,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
} from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });
const nameBox = (page: Page) =>
  page.getByRole('combobox', { name: 'Name box' });

/** Types a text in the name box and presses Enter. */
async function goTo(page: Page, text: string) {
  await nameBox(page).fill(text);
  await nameBox(page).press('Enter');
}

test('names a range in the name box, uses it in a formula, and goes to it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2');
  await enter(page, '3');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  await goTo(page, 'Sales');
  await expect(nameBox(page)).toHaveValue('Sales');
  // The grid has the keyboard back.
  await page.keyboard.press('ArrowRight');
  await expect(nameBox(page)).toHaveValue('B1');

  // Formula autocomplete offers the name, which then counts as its range.
  await clickCell(page, 0, 2);
  await page.keyboard.type('=SUM(sa');
  await expect(
    page.getByRole('listbox').getByRole('option', { name: 'Sales' }),
  ).toBeVisible();
  await page.keyboard.press('Tab');
  await page.keyboard.type(')');
  await page.keyboard.press('Enter');
  await expectCell(page, 0, 2, 'C1 5');

  // Choosing the name from the name box's list selects its range.
  await clickCell(page, 5, 5);
  await button(page, 'Defined names').click();
  await page.getByRole('menuitem', { name: 'Sales' }).click();
  await expect(nameBox(page)).toHaveValue('Sales');
  await expect(page.getByRole('status')).toHaveText('A1 2');

  // A cell address goes to that cell.
  await goTo(page, 'd10');
  await expect(nameBox(page)).toHaveValue('D10');

  // Neither an address nor a name that can be defined.
  await goTo(page, '1st');
  await expect(page.getByRole('alert')).toContainText(
    '"1st" is not a valid reference or name.',
  );
  await page.keyboard.press('Escape');
  await expect(nameBox(page)).toHaveValue('D10');

  // A formula using a name not defined yet gives #NAME? until it is, and
  // naming can be undone and redone.
  await clickCell(page, 1, 2);
  await enter(page, '=Rate*2');
  await expectCell(page, 1, 2, 'C2 #NAME?');
  await clickCell(page, 1, 0);
  await goTo(page, 'Rate');
  await expectCell(page, 1, 2, 'C2 6');
  await button(page, 'Undo').click();
  await expectCell(page, 1, 2, 'C2 #NAME?');
  await button(page, 'Redo').click();
  await expectCell(page, 1, 2, 'C2 6');
});

test('edits and deletes names in the name manager, and names follow inserted rows', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '4');
  await enter(page, '6');
  await clickCell(page, 0, 0);
  await goTo(page, 'Cost');
  await clickCell(page, 1, 0);
  await goTo(page, 'Qty');
  await clickCell(page, 0, 1);
  await enter(page, '=Cost+Qty');
  await expectCell(page, 0, 1, 'B1 10');

  await page.getByRole('tab', { name: 'Formulas' }).click();
  await button(page, 'Name manager').click();
  const dialog = page.getByRole('dialog', { name: 'Name manager' });
  const row = (name: string) =>
    dialog.getByRole('listitem').filter({ hasText: name });
  await expect(dialog.getByRole('listitem')).toHaveText([
    /Cost.*Sheet1!\$A\$1/,
    /Qty.*Sheet1!\$A\$2/,
  ]);

  // A name taken or a range that is not one is refused.
  await row('Cost').getByRole('button', { name: 'Edit' }).click();
  await dialog.getByLabel('Name', { exact: true }).fill('qty');
  await button(page, 'OK').click();
  await expect(dialog).toContainText('A name with that name already exists.');
  await dialog.getByLabel('Name', { exact: true }).fill('Start');
  await dialog.getByLabel('Refers to').fill('=nowhere');
  await button(page, 'OK').click();
  await expect(dialog).toContainText('Type a range, as =Sheet1!$A$1:$B$2.');
  await dialog.getByLabel('Refers to').fill('=$A$2');
  await button(page, 'OK').click();
  await expect(row('Start')).toContainText('Sheet1!$A$2');

  // The formula follows the renamed name; a deleted name gives #NAME?.
  await row('Qty').getByRole('button', { name: 'Delete' }).click();
  await expect(dialog.getByRole('listitem')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expectCell(page, 0, 1, 'B1 #NAME?');
  await page.getByRole('tab', { name: 'Home' }).click();
  await button(page, 'Undo').click();
  await expectCell(page, 0, 1, 'B1 12');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(
    '=Start+Qty',
  );

  // Names follow their cells when rows are inserted.
  await clickCell(page, 0, 0);
  await page.locator('canvas').click({
    button: 'right',
    position: {
      x: ROW_HEADER_WIDTH / 2,
      y: COLUMN_HEADER_HEIGHT + ROW_HEIGHT / 2,
    },
  });
  await page
    .getByRole('menuitem', { name: 'Insert rows', exact: true })
    .click();
  await expectCell(page, 1, 1, 'B2 12');
  await goTo(page, 'Start');
  await expect(page.getByRole('status')).toHaveText('A3 6');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the name box and the name manager are translated and dark, with tooltips and hover states', async ({
    page,
  }) => {
    await openApp(page);
    const background = (locator: Locator) =>
      locator.evaluate((node) => getComputedStyle(node).backgroundColor);
    const dark = (color: string) => {
      const [red, green, blue] = color.match(/\d+/g)!.map(Number);
      return (red! + green! + blue!) / 3 < 128;
    };
    const box = page.getByRole('combobox', { name: '名称框' });
    expect(dark(await background(box))).toBe(true);
    await box.fill('Total');
    await box.press('Enter');

    const list = button(page, '已定义的名称');
    const resting = await background(list);
    await list.hover();
    await expect.poll(() => background(list)).not.toBe(resting);
    await expect(
      page.getByRole('tooltip', { name: '转到工作簿中的名称' }),
    ).toBeVisible();

    await page.getByRole('tab', { name: '公式' }).click();
    await button(page, '名称管理器').hover();
    await expect(
      page.getByRole('tooltip', { name: /^查看、编辑和删除工作簿中的名称/ }),
    ).toBeVisible();
    await button(page, '名称管理器').click();
    const dialog = page.getByRole('dialog', { name: '名称管理器' });
    expect(dark(await background(dialog))).toBe(true);
    await expect(dialog.getByRole('listitem')).toHaveText([/Total/]);
    await expect(dialog.getByRole('button', { name: '编辑' })).toBeVisible();
  });
});
