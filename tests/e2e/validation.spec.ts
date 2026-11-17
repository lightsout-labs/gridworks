import { expect, test, type Page } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/**
 * Answers the message that an entry was rejected, and waits for the
 * keyboard to come back: to the editor for Retry, to the grid for Cancel.
 */
async function answer(page: Page, name: 'Retry' | 'Cancel' | '重试' | '取消') {
  await button(page, name).click();
  await expect(page.getByRole('alertdialog')).toBeHidden();
  await expect(
    name === 'Retry' || name === '重试'
      ? page.getByRole('textbox', { name: /^[A-Z]+\d+$/ })
      : page.locator('canvas').locator('..').locator('..'),
  ).toBeFocused();
}

/** Opens the Data validation dialog for the selection. */
async function openDialog(page: Page) {
  await button(page, 'Data validation').click();
  return page.getByRole('dialog', { name: 'Data validation' });
}

test('restricts cells to a list chosen from a dropdown, and rejects other input', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 2);
  await enter(page, 'Yes');
  await enter(page, 'No');
  await page.getByRole('tab', { name: 'Data' }).click();
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  const dialog = await openDialog(page);
  await dialog.getByLabel('Allow').selectOption('List');
  await dialog.getByLabel('Source').fill('=$C$1:$C$2');
  await button(page, 'OK').click();
  await expect(dialog).toBeHidden();

  // The active cell of a list shows a dropdown of its items.
  await clickCell(page, 1, 0);
  await button(page, 'Choose from the list').click();
  await page.getByRole('menuitem', { name: 'No' }).click();
  await expect(page.getByRole('status')).toHaveText('A2 No');
  await clickCell(page, 1, 1);
  await expect(button(page, 'Choose from the list')).toBeHidden();

  // Other input is rejected with a message; Retry goes back to it.
  await clickCell(page, 0, 0);
  await enter(page, 'Maybe');
  const message = page.getByRole('alertdialog', { name: 'Invalid entry' });
  await expect(message).toContainText(
    "This value doesn't match the data validation restrictions defined for this cell.",
  );
  await answer(page, 'Retry');
  await expect(page.getByRole('textbox', { name: 'A1' })).toHaveValue('Maybe');
  await page.keyboard.press('ControlOrMeta+a');
  await enter(page, 'yes');
  await expectCell(page, 0, 0, 'A1 yes');

  // Cancel abandons it.
  await enter(page, 'Perhaps');
  await answer(page, 'Cancel');
  await expectCell(page, 0, 0, 'A1 yes');
});

test('restricts cells to a number range, a date range or a text length, and changes and removes the rule', async ({
  page,
}) => {
  await openApp(page);
  await page.getByRole('tab', { name: 'Data' }).click();
  await clickCell(page, 0, 0);
  let dialog = await openDialog(page);
  await dialog.getByLabel('Allow').selectOption('Number');
  await dialog.getByLabel('Minimum').fill('1');
  await dialog.getByLabel('Maximum').fill('10');
  await button(page, 'OK').click();
  await clickCell(page, 0, 0);
  await enter(page, '20');
  await answer(page, 'Cancel');
  await expectCell(page, 0, 0, 'A1');
  await enter(page, '5');
  await expectCell(page, 0, 0, 'A1 5');

  // The dialog shows the rule of the active cell, to change it.
  dialog = await openDialog(page);
  await expect(dialog.getByLabel('Allow')).toHaveValue('number');
  await expect(dialog.getByLabel('Minimum')).toHaveValue('1');
  await dialog.getByLabel('Allow').selectOption('Date');
  await dialog.getByLabel('Minimum').fill('1/1/2025');
  await dialog.getByLabel('Maximum').fill('');
  await button(page, 'OK').click();
  await clickCell(page, 0, 0);
  await enter(page, '12/31/2024');
  await answer(page, 'Cancel');
  await enter(page, '3/1/2025');
  await expectCell(page, 0, 0, 'A1 3/1/2025');

  dialog = await openDialog(page);
  await expect(dialog.getByLabel('Minimum')).toHaveValue('1/1/2025');
  await dialog.getByLabel('Allow').selectOption('Text length');
  await dialog.getByLabel('Minimum').fill('');
  await dialog.getByLabel('Maximum').fill('abc');
  await button(page, 'OK').click();
  await expect(dialog.getByText('Enter a whole number.')).toBeVisible();
  await dialog.getByLabel('Maximum').fill('3');
  await button(page, 'OK').click();
  await clickCell(page, 0, 0);
  await enter(page, 'abcd');
  await answer(page, 'Cancel');

  // Removed, any input goes; undone, the rule is back.
  dialog = await openDialog(page);
  await button(page, 'Clear all').click();
  await expect(dialog).toBeHidden();
  await clickCell(page, 0, 0);
  await enter(page, 'abcd');
  await expectCell(page, 0, 0, 'A1 abcd');
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('ControlOrMeta+z');
  await enter(page, 'abcd');
  await expect(
    page.getByRole('alertdialog', { name: 'Invalid entry' }),
  ).toBeVisible();
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the dialog, the list and the message are translated and dark, with tooltips and hover states', async ({
    page,
  }) => {
    await openApp(page);
    const background = (locator: ReturnType<Page['locator']>) =>
      locator.evaluate((node) => getComputedStyle(node).backgroundColor);
    const dark = (color: string) => {
      const [red, green, blue] = color.match(/\d+/g)!.map(Number);
      return (red! + green! + blue!) / 3 < 128;
    };
    const hovered = async (locator: ReturnType<Page['locator']>) => {
      const resting = await background(locator);
      await locator.hover();
      await expect.poll(() => background(locator)).not.toBe(resting);
    };

    await page.getByRole('tab', { name: '数据' }).click();
    await clickCell(page, 0, 0);
    await button(page, '数据验证').hover();
    await expect(
      page.getByRole('tooltip', { name: /^选择所选单元格中可以输入的内容/ }),
    ).toBeVisible();
    await button(page, '数据验证').click();
    const dialog = page.getByRole('dialog', { name: '数据验证' });
    expect(dark(await background(dialog))).toBe(true);
    await dialog.getByLabel('允许').selectOption('序列');
    await dialog.getByLabel('来源').fill('是,否');
    await button(page, '确定').click();

    const list = button(page, '从列表中选择');
    await hovered(list);
    await expect(
      page.getByRole('tooltip', { name: '从列表中选择' }),
    ).toBeVisible();
    await list.click();
    const menu = page.getByRole('menu');
    expect(dark(await background(menu))).toBe(true);
    await hovered(page.getByRole('menuitem', { name: '否' }));
    await page.keyboard.press('Escape');

    await clickCell(page, 0, 0);
    await enter(page, 'maybe');
    const message = page.getByRole('alertdialog', { name: '输入无效' });
    await expect(message).toContainText(
      '此值与此单元格定义的数据验证限制不匹配。',
    );
    expect(dark(await background(message))).toBe(true);
    await hovered(button(page, '取消'));
    await answer(page, '取消');
  });
});
