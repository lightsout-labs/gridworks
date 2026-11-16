import { expect, test, type Page } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US', colorScheme: 'light' });

const button = (page: Page, name: string) =>
  page.getByRole('button', { name, exact: true });

/** Chooses a format from the Number format menu of the ribbon. */
async function chooseFormat(page: Page, name: string) {
  await button(page, 'Number format').click();
  await page.getByRole('menuitemradio', { name, exact: true }).click();
}

/** Opens the dialog of custom format codes from the Number format menu. */
async function openCustom(
  page: Page,
  menu = 'Number format',
  item = 'More number formats...',
) {
  await button(page, menu).click();
  await page.getByRole('menuitem', { name: item }).click();
}

test('applies each number format of the ribbon, and undoes it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1234.5');
  for (const [name, shown] of [
    ['Number', '1234.50'],
    ['Currency', '$1,234.50'],
    ['Accounting', '$1,234.50'],
    ['Percentage', '123450.00%'],
    ['Scientific', '1.23E+03'],
    ['Date', '5/18/1903'],
    ['Time', '12:00:00 PM'],
    ['Text', '1234.5'],
    ['General', '1234.5'],
  ]) {
    await clickCell(page, 0, 0);
    await chooseFormat(page, name!);
    await expect(page.getByRole('status')).toHaveText(`A1 ${shown}`);
    // The menu shows the format of the active cell.
    await expect(button(page, 'Number format')).toHaveText(name!);
  }

  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 0, 'A1 1234.5');
  await expect(button(page, 'Number format')).toHaveText('Text');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(button(page, 'Number format')).toHaveText('Time');
  await page.keyboard.press('ControlOrMeta+y');
  await expect(button(page, 'Number format')).toHaveText('Text');
});

test('adds and removes decimals and the thousands separator', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1234.5');
  await clickCell(page, 0, 0);

  await button(page, 'Increase decimal').click();
  await expect(page.getByRole('status')).toHaveText('A1 1234.50');
  await button(page, 'Thousands separator').click();
  await expect(page.getByRole('status')).toHaveText('A1 1,234.50');
  await expect(button(page, 'Thousands separator')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await button(page, 'Decrease decimal').click();
  await button(page, 'Decrease decimal').click();
  await expect(page.getByRole('status')).toHaveText('A1 1,235');
  await button(page, 'Thousands separator').click();
  await expect(page.getByRole('status')).toHaveText('A1 1235');

  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 0, 'A1 1,235');
});

test('keeps what is typed into a cell with the text format as text', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await chooseFormat(page, 'Text');
  await clickCell(page, 0, 0);
  await enter(page, '007');
  await expectCell(page, 0, 0, 'A1 007');
  await expect(page.getByLabel('Formula bar')).toHaveValue('007');
  await clickCell(page, 0, 0);
  await enter(page, '=1+1');
  await expectCell(page, 0, 0, 'A1 =1+1');
});

test('applies a custom format code from a dialog that previews it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '-3');
  await clickCell(page, 0, 0);
  await openCustom(page);

  const dialog = page.getByRole('dialog', { name: 'Custom number format' });
  const code = dialog.getByLabel('Format code');
  await expect(code).toHaveValue('General');
  await expect(dialog.getByLabel('Preview')).toHaveText('-3');

  // A code that is not valid cannot be applied.
  await code.fill('0.0"');
  await expect(
    dialog.getByText('This format code is not valid.'),
  ).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled();
  // It looks disabled, even under the pointer.
  await dialog.getByRole('button', { name: 'OK' }).hover({ force: true });
  await expect(dialog.getByRole('button', { name: 'OK' })).toHaveCSS(
    'opacity',
    '0.5',
  );
  await expect(dialog.getByRole('button', { name: 'OK' })).toHaveCSS(
    'background-color',
    'rgba(0, 0, 0, 0)',
  );

  await code.fill('[Red]0.0;[Blue](0.0);"zero";"Text: "@');
  await expect(dialog.getByLabel('Preview')).toHaveText('(3.0)');
  await expect(dialog.getByLabel('Preview')).toHaveCSS(
    'color',
    'rgb(0, 0, 255)',
  );
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  await expectCell(page, 0, 0, 'A1 (3.0)');
  await expect(button(page, 'Number format')).toHaveText('Custom');

  // Cancel leaves the format as it was.
  await openCustom(page);
  await code.fill('0.000');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expectCell(page, 0, 0, 'A1 (3.0)');

  await page.keyboard.press('ControlOrMeta+z');
  await expectCell(page, 0, 0, 'A1 -3');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the number format controls are translated and dark', async ({
    page,
  }) => {
    await openApp(page);
    for (const name of [
      '数字格式',
      '增加小数位数',
      '减少小数位数',
      '千位分隔符',
    ]) {
      await expect(button(page, name)).toBeVisible();
    }
    await expect(button(page, '数字格式')).toHaveText('常规');
    await openCustom(page, '数字格式', '其他数字格式…');
    const dialog = page.getByRole('dialog', { name: '自定义数字格式' });
    await expect(dialog).toBeVisible();
    const background = await dialog.evaluate(
      (element) => getComputedStyle(element).backgroundColor,
    );
    expect(background).not.toBe('rgb(255, 255, 255)');
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });
});
