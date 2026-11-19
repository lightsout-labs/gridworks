import { expect, test, type Page } from '@playwright/test';
import { cellRect, UNIFORM } from '../../src/app/grid/geometry';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

const formulaBar = (page: Page) =>
  page.getByRole('textbox', { name: 'Formula bar' });
const suggestions = (page: Page) =>
  page.getByRole('listbox', { name: 'Functions' });
const hint = (page: Page) => page.getByRole('note', { name: 'Arguments' });
const description = (page: Page) =>
  page.getByRole('note', { name: 'Description' });

/** Where the middle of a cell is on the page, at the top of the sheet. */
async function cellPoint(page: Page, row: number, col: number) {
  const canvas = (await page.locator('canvas').boundingBox())!;
  const rect = cellRect(
    { row, col },
    { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM },
  );
  return {
    x: canvas.x + rect.x + rect.width / 2,
    y: canvas.y + rect.y + rect.height / 2,
  };
}

/** The color the canvas drew just inside the left edge of a cell. */
async function leftEdgeColor(page: Page, row: number, col: number) {
  const rect = cellRect(
    { row, col },
    { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM },
  );
  return page.locator('canvas').evaluate((canvas: HTMLCanvasElement, rect) => {
    const scale = canvas.width / canvas.clientWidth;
    const { data } = canvas
      .getContext('2d')!
      .getImageData(
        Math.floor((rect.x + 1) * scale),
        Math.floor((rect.y + rect.height / 2) * scale),
        1,
        1,
      );
    return `rgb(${data[0]}, ${data[1]}, ${data[2]})`;
  }, rect);
}

test('function names are completed and their arguments shown while typing', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '4');
  await enter(page, '6');
  await clickCell(page, 3, 0);
  await page.keyboard.type('=sumi');
  await expect(suggestions(page).getByRole('option')).toHaveText([
    'SUMIF',
    'SUMIFS',
  ]);
  await expect(
    suggestions(page).getByRole('option', { name: 'SUMIF', exact: true }),
  ).toHaveAttribute('aria-selected', 'true');
  // The arrows move through the list rather than leaving the cell.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Tab');
  await expect(formulaBar(page)).toHaveValue('=SUMIF(');
  await expect(suggestions(page)).toBeHidden();

  await expect(hint(page)).toHaveText('SUMIF(range, criteria, [sum_range])');
  await expect(hint(page).locator('strong')).toHaveText('range');
  await page.keyboard.type('A1:A2,">5",');
  await expect(hint(page).locator('strong')).toHaveText('[sum_range]');
  await page.keyboard.type('A1:A2)');
  await expect(hint(page)).toBeHidden();
  await page.keyboard.press('Enter');
  await expectCell(page, 3, 0, 'A4 6');

  // A click chooses too; a name no function starts with offers nothing,
  // and Escape closes the list but keeps the entry.
  await clickCell(page, 4, 0);
  await page.keyboard.type('=ro');
  await suggestions(page).getByRole('option', { name: 'ROUNDUP' }).click();
  await expect(formulaBar(page)).toHaveValue('=ROUNDUP(');
  await page.keyboard.type('2.1,0)+xyz');
  await expect(suggestions(page)).toBeHidden();
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.press('Backspace');
  await page.keyboard.type('ab');
  await expect(suggestions(page).getByRole('option')).toHaveText(['ABS']);
  await page.keyboard.press('Escape');
  await expect(suggestions(page)).toBeHidden();
  await expect(formulaBar(page)).toHaveValue('=ROUNDUP(2.1,0)+ab');
});

test('the list describes the highlighted function', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('=sumi');
  await expect(description(page)).toHaveText(
    'Adds the cells specified by a given condition or criteria',
  );
  await page.keyboard.press('ArrowDown');
  await expect(description(page)).toHaveText(
    'Adds the cells specified by a given set of conditions or criteria',
  );
  await page.keyboard.press('Escape');
  await expect(description(page)).toBeHidden();
});

test('ISFORMULA and FORMULATEXT are offered, and read the formula of a cell', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 2);
  await enter(page, '=1+1');
  await clickCell(page, 1, 0);
  await page.keyboard.type('=isfo');
  await expect(suggestions(page).getByRole('option')).toHaveText(['ISFORMULA']);
  await expect(description(page)).toHaveText(
    'Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE',
  );
  await page.keyboard.press('Tab');
  await expect(hint(page)).toHaveText('ISFORMULA(reference)');
  await page.keyboard.type('C1)');
  await page.keyboard.press('Enter');
  await expectCell(page, 1, 0, 'A2 TRUE');

  await clickCell(page, 2, 0);
  await page.keyboard.type('=formulat');
  await expect(description(page)).toHaveText('Returns a formula as a string');
  await page.keyboard.press('Tab');
  await expect(hint(page)).toHaveText('FORMULATEXT(reference)');
  await page.keyboard.type('C1)');
  await page.keyboard.press('Enter');
  await expectCell(page, 2, 0, 'A3 =1+1');

  await clickCell(page, 0, 2);
  await enter(page, '2');
  await expectCell(page, 1, 0, 'A2 FALSE');
  await expectCell(page, 2, 0, 'A3 #N/A');
});

/** Whether the highlighted entry lies wholly inside the list's visible area. */
function highlightedInView(page: Page) {
  return suggestions(page).evaluate((list) => {
    const option = list.querySelector('[aria-selected="true"]')!;
    const inner = list.getBoundingClientRect();
    const outer = option.getBoundingClientRect();
    return outer.top >= inner.top && outer.bottom <= inner.bottom;
  });
}

/** Presses an arrow `times` times, the highlighted entry staying in view. */
async function walk(page: Page, key: string, times: number) {
  for (let step = 0; step < times; step++) {
    await page.keyboard.press(key);
    await expect.poll(() => highlightedInView(page)).toBe(true);
  }
}

test('the list scrolls to keep the highlighted entry in view', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('=S');
  const options = suggestions(page).getByRole('option');
  await walk(page, 'ArrowDown', 12);
  await expect(options.nth(12)).toHaveAttribute('aria-selected', 'true');
  await expect(description(page)).toBeInViewport();
  await expect(description(page)).not.toBeEmpty();
  await walk(page, 'ArrowUp', 12);
  await expect(options.first()).toHaveAttribute('aria-selected', 'true');
  // Up from the first entry goes round to the last.
  await walk(page, 'ArrowUp', 1);
  await expect(options.last()).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');

  // A list of argument values and names, typed in the formula bar.
  await formulaBar(page).click();
  await page.keyboard.type('=VLOOKUP("Apple",A1:B10,2,');
  await walk(page, 'ArrowDown', 12);
  await expect(options.nth(12)).toHaveAttribute('aria-selected', 'true');
  await expect(description(page)).toBeInViewport();
  await walk(page, 'ArrowUp', 12);
  await expect(options.first()).toHaveAttribute('aria-selected', 'true');
});

test('arguments offer the values Excel offers, before the names', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 3, 3);
  await page.keyboard.type('=VLOOKUP("Apple",A1:B10,2,');
  const options = suggestions(page).getByRole('option');
  await expect(options.nth(0)).toHaveText('TRUE - Approximate match');
  await expect(options.nth(1)).toHaveText('FALSE - Exact match');
  await expect(options.nth(2)).toHaveText('ABS');
  await expect(hint(page)).toBeHidden();
  await page.keyboard.type('F');
  await expect(options.first()).toHaveText('FALSE - Exact match');
  await expect(options.nth(1)).toHaveText('FACT');
  await page.keyboard.press('Enter');
  await expect(formulaBar(page)).toHaveValue('=VLOOKUP("Apple",A1:B10,2,FALSE');
  // Typed in full, neither the value nor the function is offered.
  await expect(suggestions(page)).toBeHidden();
  await page.keyboard.press('Backspace');
  await page.keyboard.type('E');
  await expect(suggestions(page)).toBeHidden();
  // Enter commits the entry and moves down.
  await page.keyboard.press('Enter');
  await expect(page.getByRole('status')).toHaveText('D5');
  await expectCell(page, 3, 3, 'D4 =VLOOKUP("Apple",A1:B10,2,FALSE');

  // The function FALSE still takes its parenthesis.
  await clickCell(page, 4, 3);
  await page.keyboard.type('=FAL');
  await expect(options).toHaveText(['FALSE']);
  await page.keyboard.press('Tab');
  await expect(formulaBar(page)).toHaveValue('=FALSE(');
  await page.keyboard.press('Escape');

  // Typing narrows the values by their text, signs included.
  await page.keyboard.type('=MATCH(A1,B1:B9,');
  await expect(options.nth(0)).toHaveText('1 - Less than');
  await expect(options.nth(1)).toHaveText('0 - Exact match');
  await expect(options.nth(2)).toHaveText('-1 - Greater than');
  await page.keyboard.type('-');
  await expect(options).toHaveText(['-1 - Greater than']);
  await options.first().click();
  await expect(formulaBar(page)).toHaveValue('=MATCH(A1,B1:B9,-1');
  await page.keyboard.press('Escape');
});

test('clicking and dragging on the grid inserts references, on another sheet too', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await enter(page, '3');
  await clickCell(page, 0, 1);
  await enter(page, '10');

  await clickCell(page, 4, 0);
  await page.keyboard.type('=SUM(');
  const from = await cellPoint(page, 0, 0);
  const to = await cellPoint(page, 2, 1);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await expect(formulaBar(page)).toHaveValue('=SUM(A1');
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:B3');
  // The entry goes on, in the cell it started in.
  await expect(page.getByRole('textbox', { name: 'A5' })).toBeFocused();
  await page.keyboard.type(')');
  await page.keyboard.press('Enter');
  await expectCell(page, 4, 0, 'A5 16');

  // Where no reference can go, a click commits the entry and selects.
  await clickCell(page, 5, 0);
  await page.keyboard.type('=5');
  await clickCell(page, 0, 3);
  await expect(page.getByRole('status')).toHaveText('D1');
  await expectCell(page, 5, 0, 'A6 5');

  // A reference to another sheet.
  await page.getByRole('button', { name: 'New sheet' }).click();
  await clickCell(page, 2, 2);
  await enter(page, '7');
  await page.getByRole('tab', { name: 'Sheet1', exact: true }).click();
  await clickCell(page, 6, 0);
  await page.keyboard.type('=A1+');
  await page.getByRole('tab', { name: 'Sheet2', exact: true }).click();
  await clickCell(page, 2, 2);
  await expect(formulaBar(page)).toHaveValue('=A1+Sheet2!C3');
  await expect(formulaBar(page)).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(
    page.getByRole('tab', { name: 'Sheet1', exact: true }),
  ).toHaveAttribute('aria-selected', 'true');
  await expectCell(page, 6, 0, 'A7 8');
});

test('Shift and Ctrl with a click extend and add references', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 5, 5);
  await page.keyboard.type('=SUM(');
  await clickCell(page, 0, 0);
  await page.keyboard.down('Shift');
  await clickCell(page, 2, 2);
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:C3');
  await clickCell(page, 1, 1);
  await page.keyboard.up('Shift');
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:B2');

  // Ctrl or Cmd adds a comma and drags a new reference.
  const from = await cellPoint(page, 2, 2);
  const to = await cellPoint(page, 3, 3);
  await page.mouse.move(from.x, from.y);
  await page.keyboard.down('ControlOrMeta');
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
  await page.keyboard.up('ControlOrMeta');
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:B2,C3:D4');
  await expect(page.getByRole('textbox', { name: 'F6' })).toBeFocused();
  await page.keyboard.press('Escape');
});

test('the arrows point at references in an entry typed in a cell', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await enter(page, '3');

  await clickCell(page, 3, 1);
  await page.keyboard.type('=SUM(');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowLeft');
  await expect(formulaBar(page)).toHaveValue('=SUM(A3');
  await page.keyboard.press('Shift+ArrowUp');
  await page.keyboard.press('Shift+ArrowUp');
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:A3');
  const reference = page.locator('[data-place="cell"] [data-reference]');
  const color = await reference.evaluate(
    (span) => getComputedStyle(span).color,
  );
  expect(await leftEdgeColor(page, 0, 0)).toBe(color);

  // After an operator, the next arrow starts from the cell edited.
  await page.keyboard.type(')*');
  await page.keyboard.press('ArrowUp');
  await expect(formulaBar(page)).toHaveValue('=SUM(A1:A3)*B3');
  // After other text, an arrow commits and moves.
  await page.keyboard.type('+5');
  await page.keyboard.press('ArrowDown');
  await expect(page.getByRole('status')).toHaveText('B5');
  await expectCell(page, 3, 1, 'B4 5'); // B3 is empty.

  // In Edit mode, the arrows move the caret.
  await page.keyboard.press('F2');
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.type('1');
  await page.keyboard.press('Enter');
  await expectCell(page, 3, 1, 'B4 15');
});

test('each reference has its own color, and its range is outlined in it', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 3);
  await page.keyboard.type('=A1+B2:C3*a1');
  const editor = page.getByRole('textbox', { name: 'D1' });
  const references = page.locator('[data-place="cell"] [data-reference]');
  await expect(references).toHaveText(['A1', 'B2:C3', 'a1']);
  const colors = await references.evaluateAll((spans) =>
    spans.map((span) => getComputedStyle(span).color),
  );
  expect(colors[0]).toBe(colors[2]);
  expect(colors[0]).not.toBe(colors[1]);
  await expect(editor).toBeFocused();

  expect(await leftEdgeColor(page, 0, 0)).toBe(colors[0]);
  expect(await leftEdgeColor(page, 1, 1)).toBe(colors[1]);
  expect(await leftEdgeColor(page, 2, 1)).toBe(colors[1]);
  const plain = await leftEdgeColor(page, 5, 5);
  expect(await leftEdgeColor(page, 1, 2)).toBe(plain); // Inside the range.

  await page.keyboard.press('Escape');
  await expect(references).toHaveCount(0);
  expect(await leftEdgeColor(page, 0, 0)).toBe(plain);
});

test('F4 cycles the reference at the caret through absolute and relative', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('=1+B2');
  for (const text of ['=1+$B$2', '=1+B$2', '=1+$B2', '=1+B2']) {
    await page.keyboard.press('F4');
    await expect(formulaBar(page)).toHaveValue(text);
  }
  // In Edit mode the arrows move the caret onto another reference, and
  // the caret stays after the reference that F4 changes.
  await page.keyboard.type('+C3');
  await page.keyboard.press('F2');
  for (let left = 0; left < 3; left++) await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('F4');
  await page.keyboard.type('*2');
  await expect(formulaBar(page)).toHaveValue('=1+$B$2*2+C3');
});

test('Insert function on the Formulas tab lists every function for a formula', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '-3');
  await page.getByRole('tab', { name: 'Formulas' }).click();
  await clickCell(page, 1, 0);
  const insert = page.getByRole('button', { name: 'Insert function' });
  await insert.hover();
  await expect(
    page.getByRole('tooltip', { name: 'Choose a function to insert' }),
  ).toBeVisible();
  await insert.click();
  await expect(formulaBar(page)).toHaveValue('=');
  await expect(formulaBar(page)).toBeFocused();
  await expect(suggestions(page).getByRole('option').first()).toHaveText('ABS');
  await expect(suggestions(page).getByRole('option')).not.toHaveCount(1);
  await suggestions(page).getByRole('option', { name: 'ABS' }).click();
  await expect(formulaBar(page)).toHaveValue('=ABS(');
  await expect(hint(page)).toHaveText('ABS(number)');

  // In a formula, the function goes at the caret.
  await page.keyboard.type('A1)+');
  await insert.click();
  await expect(formulaBar(page)).toBeFocused();
  await page.keyboard.type('SQ');
  await expect(suggestions(page).getByRole('option')).toHaveText([
    'SQRT',
    'SQRTPI',
  ]);
  await page.keyboard.press('Enter');
  await expect(formulaBar(page)).toHaveValue('=ABS(A1)+SQRT(');
  await page.keyboard.type('16)');
  await page.keyboard.press('Enter');
  await expectCell(page, 1, 0, 'A2 7');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the list and the hint are translated and dark, with hover states', async ({
    page,
  }) => {
    await openApp(page);
    await clickCell(page, 0, 0);
    await page.keyboard.type('=ro');
    const list = page.getByRole('listbox', { name: '函数' });
    await expect(list).toBeVisible();
    const background = (selector: string) =>
      page
        .locator(selector)
        .first()
        .evaluate((node) => getComputedStyle(node).backgroundColor);
    const dark = (color: string) => {
      const [red, green, blue] = color.match(/\d+/g)!.map(Number);
      return (red! + green! + blue!) / 3 < 128;
    };
    expect(dark(await background('[role="listbox"]'))).toBe(true);
    await expect(page.getByRole('note', { name: '说明' })).toHaveText(
      '按指定的位数对数值进行四舍五入',
    );
    expect(dark(await background('[role="note"]'))).toBe(true);
    const option = list.getByRole('option').nth(1);
    const resting = await option.evaluate(
      (node) => getComputedStyle(node).backgroundColor,
    );
    await option.hover();
    await expect
      .poll(() =>
        option.evaluate((node) => getComputedStyle(node).backgroundColor),
      )
      .not.toBe(resting);
    await page.keyboard.press('Tab');
    await expect(page.getByRole('note', { name: '参数' })).toBeVisible();
    expect(dark(await background('[role="note"]'))).toBe(true);
    await page.keyboard.press('Escape');
    await page.keyboard.type('=VLOOKUP(A1,B1:C2,2,');
    await expect(list.getByRole('option').nth(1)).toHaveText(
      'FALSE - 精确匹配',
    );
    await page.getByRole('tab', { name: '公式' }).click();
    await expect(page.getByRole('button', { name: '插入函数' })).toBeVisible();
  });
});
