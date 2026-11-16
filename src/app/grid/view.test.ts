import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, SHEET_ROWS } from '../../engine/workbook';
import { pick, single } from './selection';
import { View } from './view';

function setup() {
  const workbook = createWorkbook();
  const view = new View(workbook);
  const sheet = workbook.sheets[0]!.id;
  return { workbook, view, sheet };
}

const down = { type: 'step', rows: 1, cols: 0, extend: false } as const;

describe('View', () => {
  it('starts on A1 of the first sheet at 100%, with no entry', () => {
    const { view, sheet } = setup();
    expect(view.state).toEqual({
      sheet,
      selection: single({ row: 0, col: 0 }),
      entry: null,
      zoom: 1,
      endMode: false,
      copied: null,
      filling: null,
      resizing: null,
      painter: null,
    });
  });

  it('commits the entry to the active cell', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 2, col: 1 }));
    view.type('4');
    view.type('42');
    expect(view.state.entry).toEqual({
      text: '42',
      mode: 'enter',
      place: 'cell',
      sheet,
      caret: 2,
      allFunctions: false,
      point: null,
    });
    expect(workbook.cell(sheet, 2, 1).display).toBe('');

    view.commit();
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 2, 1).display).toBe('42');
  });

  it('commits a waiting entry to its own cell when the selection changes', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.select(single({ row: 5, col: 5 }));
    expect(view.state).toMatchObject({
      selection: single({ row: 5, col: 5 }),
      entry: null,
    });
    expect(workbook.cell(sheet, 0, 0).display).toBe('7');

    view.type('8');
    view.navigate(down);
    expect(view.state.selection.active).toEqual({ row: 6, col: 5 });
    expect(workbook.cell(sheet, 5, 5).display).toBe('8');
  });

  it('navigates over the cells of the workbook', () => {
    const { workbook, view, sheet } = setup();
    for (const row of [0, 1, 2]) {
      workbook.execute({ type: 'setCell', sheet, row, col: 0, input: '1' });
    }
    view.navigate({ type: 'edge', rows: 1, cols: 0, extend: false });
    expect(view.state.selection.active).toEqual({ row: 2, col: 0 });
    view.navigate({ type: 'lastCell', extend: false });
    expect(view.state.selection.active).toEqual({ row: 2, col: 0 });
  });

  it('turns End mode on and off, and ends it with the next move', () => {
    const { view } = setup();
    view.navigate({ type: 'endMode' });
    expect(view.state.endMode).toBe(true);
    view.navigate({ type: 'endMode' });
    expect(view.state.endMode).toBe(false);
    view.navigate({ type: 'endMode' });
    view.navigate(down);
    expect(view.state.endMode).toBe(false);
    view.navigate({ type: 'endMode' });
    view.select(single({ row: 1, col: 1 }));
    expect(view.state.endMode).toBe(false);
  });

  it('edits the contents of the active cell, in the cell or the formula bar', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '12%' });
    view.edit('cell');
    expect(view.state.entry).toEqual({
      text: '12%',
      mode: 'edit',
      place: 'cell',
      sheet,
      caret: 3,
      allFunctions: false,
      point: null,
    });
    view.type('15%');
    view.commit();
    expect(workbook.cell(sheet, 0, 0).display).toBe('15%');

    view.edit('formulaBar');
    expect(view.state.entry).toEqual({
      text: '15%',
      mode: 'edit',
      place: 'formulaBar',
      sheet,
      caret: 3,
      allFunctions: false,
      point: null,
    });
  });

  it('keeps the text typed so far when an entry switches to Edit mode', () => {
    const { view, sheet } = setup();
    view.type('ab');
    view.edit('cell');
    expect(view.state.entry).toEqual({
      text: 'ab',
      mode: 'edit',
      place: 'cell',
      sheet,
      caret: 2,
      allFunctions: false,
      point: null,
    });
    view.edit('formulaBar');
    expect(view.state.entry?.place).toBe('formulaBar');
    view.type('abc');
    expect(view.state.entry).toEqual({
      text: 'abc',
      mode: 'edit',
      place: 'formulaBar',
      sheet,
      caret: 3,
      allFunctions: false,
      point: null,
    });
  });

  it('keeps the caret of the entry, at the end of what is typed unless moved', () => {
    const { view } = setup();
    view.type('=SUM()');
    expect(view.state.entry?.caret).toBe(6);
    view.moveCaret(5);
    expect(view.state.entry?.caret).toBe(5);
    view.type('=SUM(1)', 6);
    expect(view.state.entry?.caret).toBe(6);
  });

  it('inserts references picked on the grid into a formula', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 4, col: 0 }));
    expect(view.canPick).toBe(false);
    view.type('=SUM(');
    expect(view.canPick).toBe(true);
    view.pick({ row: 0, col: 0 });
    // Dragging replaces the reference picked.
    view.pick({ row: 2, col: 1 }, 'extend');
    expect(view.state.entry).toMatchObject({ text: '=SUM(A1:B3', caret: 10 });
    expect(view.state.selection).toEqual(single({ row: 4, col: 0 }));
    view.type('=SUM(A1:B3)');
    expect(view.canPick).toBe(false);
    view.commit();
    expect(workbook.cell(sheet, 4, 0).input).toBe('=SUM(A1:B3)');
  });

  it('extends the reference picked with Shift, from where it started', () => {
    const { view } = setup();
    view.select(single({ row: 5, col: 5 }));
    view.type('=');
    view.pick({ row: 0, col: 0 });
    view.pick({ row: 2, col: 2 }, 'extend');
    expect(view.state.entry?.text).toBe('=A1:C3');
    view.pick({ row: 1, col: 1 }, 'extend');
    expect(view.state.entry?.text).toBe('=A1:B2');

    // A drag from C3 to A1 started at C3.
    view.type('=');
    view.pick({ row: 2, col: 2 });
    view.pick({ row: 0, col: 0 }, 'extend');
    expect(view.state.entry?.text).toBe('=A1:C3');
    view.pick({ row: 3, col: 3 }, 'extend');
    expect(view.state.entry?.text).toBe('=C3:D4');

    // With no reference before the caret, just the cell.
    view.type('=1+');
    view.pick({ row: 1, col: 1 }, 'extend');
    expect(view.state.entry?.text).toBe('=1+B2');
  });

  it('adds a reference after a comma with Ctrl, or picks one without a reference before', () => {
    const { view } = setup();
    view.select(single({ row: 5, col: 5 }));
    view.type('=SUM(');
    view.pick({ row: 2, col: 2 }, 'add');
    expect(view.state.entry?.text).toBe('=SUM(C3');
    view.type('=SUM(A1');
    view.pick({ row: 2, col: 2 }, 'add');
    view.pick({ row: 3, col: 3 }, 'extend');
    expect(view.state.entry).toMatchObject({
      text: '=SUM(A1,C3:D4',
      caret: 13,
    });
  });

  it('points with the arrows from the cell edited, extends with Shift, and starts again after an operator', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 9, col: 1, input: '1' });
    view.select(single({ row: 1, col: 1 }));
    view.type('=');
    view.point({ type: 'step', rows: 0, cols: 1, extend: false });
    expect(view.state.entry?.text).toBe('=C2');
    view.point(down);
    expect(view.state.entry?.text).toBe('=C3');
    view.point({ type: 'step', rows: 0, cols: 1, extend: true });
    expect(view.state.entry).toMatchObject({ text: '=C3:D3', caret: 6 });
    expect(view.canPoint).toBe(true);
    // After a reference typed, or other text, the arrows commit.
    view.type('=C3:D3+A1');
    expect(view.canPoint).toBe(false);
    view.type('=C3:D3+1');
    expect(view.canPoint).toBe(false);
    view.type('=C3:D3+');
    expect(view.canPoint).toBe(true);
    view.point({ type: 'step', rows: -1, cols: 0, extend: false });
    expect(view.state.entry?.text).toBe('=C3:D3+B1');
    // Ctrl with an arrow goes to the edge of the data, as it selects.
    view.type('=C3:D3+B1*');
    view.point({ type: 'edge', rows: 1, cols: 0, extend: false });
    expect(view.state.entry?.text).toBe('=C3:D3+B1*B10');
    expect(view.state.selection).toEqual(single({ row: 1, col: 1 }));
  });

  it('goes on with a formula on another sheet, and commits it to its own cell', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    workbook.execute({ type: 'setSheet', sheet: second, name: 'My data' });
    view.select(single({ row: 1, col: 1 }));
    view.type('=1+');
    view.showSheet(second);
    expect(view.state.sheet).toBe(second);
    // The cell is not shown, so the entry goes on in the formula bar.
    expect(view.state.entry).toMatchObject({ place: 'formulaBar', sheet });
    expect(view.canPick).toBe(true);
    view.pick({ row: 2, col: 2 });
    expect(view.state.entry?.text).toBe("=1+'My data'!C3");
    view.navigate(down);
    expect(view.state.sheet).toBe(sheet);
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
    expect(workbook.cell(sheet, 1, 1).input).toBe("=1+'My data'!C3");

    // Back on its own sheet, a reference has no sheet name; cancelling
    // also goes back.
    view.type('=');
    view.showSheet(second);
    view.showSheet(sheet);
    view.pick({ row: 0, col: 0 });
    expect(view.state.entry?.text).toBe('=A1');
    view.showSheet(second);
    view.cancel();
    expect(view.state).toMatchObject({ sheet, entry: null });

    // Where no reference can go, a click on the other sheet commits the
    // formula to its own cell and selects there.
    view.select(single({ row: 1, col: 1 }));
    view.type('=SUM(');
    view.showSheet(second);
    view.pick({ row: 0, col: 0 });
    view.type("=SUM('My data'!A1)");
    expect(view.canPick).toBe(false);
    view.select(single({ row: 3, col: 3 }));
    expect(view.state).toMatchObject({
      sheet: second,
      entry: null,
      selection: single({ row: 3, col: 3 }),
    });
    expect(workbook.cell(sheet, 1, 1).input).toBe("=SUM('My data'!A1)");
  });

  it('cycles the reference at the caret between absolute and relative', () => {
    const { view } = setup();
    view.cycleReference(); // Nothing to cycle.
    expect(view.state.entry).toBeNull();
    view.type('=A1+2', 3);
    view.cycleReference();
    expect(view.state.entry).toMatchObject({ text: '=$A$1+2', caret: 5 });
  });

  it('completes the function name typed', () => {
    const { view } = setup();
    view.type('=1+ro');
    view.complete('ROUND');
    expect(view.state.entry).toMatchObject({ text: '=1+ROUND(', caret: 9 });
  });

  it('opens the list of every function, in a formula or a new one', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: 'x' });
    view.insertFunction();
    expect(view.state.entry).toMatchObject({
      text: '=',
      caret: 1,
      mode: 'edit',
      place: 'formulaBar',
      allFunctions: true,
    });
    view.complete('ABS');
    expect(view.state.entry).toMatchObject({
      text: '=ABS(',
      allFunctions: false,
    });
    view.type('=ABS(1)+', 8);
    view.insertFunction();
    expect(view.state.entry).toMatchObject({ text: '=ABS(1)+', caret: 8 });
    // Moving the caret or typing closes the list.
    view.moveCaret(2);
    expect(view.state.entry?.allFunctions).toBe(false);
    view.insertFunction();
    view.type('=ABS(1)+S');
    expect(view.state.entry?.allFunctions).toBe(false);
    view.cancel();
    view.type('abc');
    view.insertFunction();
    expect(view.state.entry).toMatchObject({ text: '=', mode: 'enter' });
  });

  it('clears the selected cells', () => {
    const { workbook, view, sheet } = setup();
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [3, 3],
    ] as const) {
      workbook.execute({ type: 'setCell', sheet, row, col, input: 'x' });
    }
    view.select({
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 3, left: 3, bottom: 3, right: 3 },
      ],
      active: { row: 0, col: 0 },
      anchor: { row: 3, col: 3 },
      focus: { row: 3, col: 3 },
    });
    view.clear();
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
  });

  it('abandons the entry on cancel', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.cancel();
    view.commit(); // Nothing is waiting any more.
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 0, 0).display).toBe('');
  });

  it('undoes and redoes a change, selecting the cells it changed', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 2, col: 1 }));
    view.type('4');
    view.commit();
    view.select(single({ row: 9, col: 9 }));
    view.undo();
    expect(workbook.cell(sheet, 2, 1).value).toBeNull();
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
    view.select(single({ row: 9, col: 9 }));
    view.redo();
    expect(workbook.cell(sheet, 2, 1).value).toBe(4);
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
  });

  it('selects every range an undone change cleared, the first cell of the last active', () => {
    const { view } = setup();
    const ranges = [
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 3, left: 3, bottom: 5, right: 4 },
    ];
    view.select({
      ranges,
      active: { row: 0, col: 0 },
      anchor: { row: 3, col: 3 },
      focus: { row: 5, col: 4 },
    });
    view.clear();
    view.select(single({ row: 9, col: 9 }));
    view.undo();
    expect(view.state.selection).toEqual({
      ranges,
      active: { row: 3, col: 3 },
      anchor: { row: 3, col: 3 },
      focus: { row: 5, col: 4 },
    });
  });

  it('abandons an entry to undo, and keeps the selection when there is nothing to undo or redo', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 1, col: 1 }));
    view.type('7');
    view.undo();
    view.redo();
    expect(view.state.entry).toBeNull();
    expect(view.state.selection).toEqual(single({ row: 1, col: 1 }));
    expect(workbook.cell(sheet, 1, 1).value).toBeNull();
  });

  it('shows another sheet at A1, committing the entry first', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    view.select(single({ row: 3, col: 3 }));
    view.type('5');
    view.showSheet(second);
    expect(workbook.cell(sheet, 3, 3).value).toBe(5);
    expect(view.state.sheet).toBe(second);
    expect(view.state.selection).toEqual(single({ row: 0, col: 0 }));
  });

  it('shows a sheet it adds, and the neighbor of the shown sheet it deletes', () => {
    const { workbook, view, sheet } = setup();
    view.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    expect(view.state.sheet).toBe(second);
    // A change to another sheet leaves the shown one.
    view.select(single({ row: 2, col: 2 }));
    view.execute({ type: 'setSheet', sheet, name: 'First' });
    view.execute({ type: 'setSheet', sheet: second, index: 0 });
    expect(view.state.sheet).toBe(second);
    expect(view.state.selection).toEqual(single({ row: 2, col: 2 }));
    view.execute({ type: 'deleteSheet', sheet: second });
    expect(view.state.sheet).toBe(sheet);
  });

  it('shows the sheet whose change it undoes or redoes', () => {
    const { workbook, view, sheet } = setup();
    // Whatever reads the shown sheet on a change, as the grid does, finds it.
    workbook.subscribe(() => workbook.cell(view.state.sheet, 0, 0));
    view.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    view.select(single({ row: 4, col: 4 }));
    view.undo();
    expect(view.state.sheet).toBe(sheet);
    expect(view.state.selection).toEqual(single({ row: 0, col: 0 }));
    view.redo();
    expect(view.state.sheet).toBe(second);
    // Undoing the addition of a sheet that is not shown stays where it is.
    view.showSheet(sheet);
    view.undo();
    expect(view.state.sheet).toBe(sheet);
    view.redo();
    view.showSheet(second);
    view.select(single({ row: 4, col: 4 }));
    view.execute({ type: 'setSheet', sheet: second, color: '#30a46c' });
    view.undo();
    // The same sheet: the selection stays.
    expect(view.state.selection).toEqual(single({ row: 4, col: 4 }));
  });

  it('zooms between 50% and 200%', () => {
    const { view } = setup();
    view.zoomTo(1.5);
    expect(view.state.zoom).toBe(1.5);
    view.zoomTo(3);
    expect(view.state.zoom).toBe(2);
    view.zoomTo(0.1);
    expect(view.state.zoom).toBe(0.5);
    // Steps of a tenth add up without rounding errors.
    view.zoomTo(0.5 + 0.1 + 0.1);
    expect(view.state.zoom).toBe(0.7);
  });

  it('tells its listeners of each change, with a new state', () => {
    const { view } = setup();
    const before = view.state;
    const listener = vi.fn();
    const unsubscribe = view.subscribe(listener);
    view.select(single({ row: 1, col: 0 }));
    expect(listener).toHaveBeenCalledTimes(1);
    expect(view.state).not.toBe(before);
    expect(view.state).toBe(view.state);

    unsubscribe();
    view.type('1');
    expect(listener).toHaveBeenCalledTimes(1);
  });

  describe('cut, copy and paste', () => {
    const range = (top: number, left: number, bottom = top, right = left) => ({
      top,
      left,
      bottom,
      right,
    });
    const selecting = (
      view: View,
      top: number,
      left: number,
      bottom = top,
      right = left,
    ) =>
      view.select({
        ...single({ row: top, col: left }),
        ranges: [range(top, left, bottom, right)],
        focus: { row: bottom, col: right },
      });

    it('copies the selection, and pastes it at the selection, which then holds the cells pasted', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '1' });
      workbook.execute({
        type: 'setCell',
        sheet,
        row: 1,
        col: 0,
        input: '=A1+1',
      });
      selecting(view, 0, 0, 1, 0);
      view.copy();
      expect(view.state.copied).toEqual({
        sheet,
        range: range(0, 0, 1, 0),
        cut: false,
      });
      view.select(single({ row: 4, col: 2 }));
      view.paste();
      expect(workbook.cell(sheet, 5, 2).formula).toBe('=C5+1');
      expect(view.state.selection.ranges).toEqual([range(4, 2, 5, 2)]);
      // A copy can be pasted again, here as values only.
      view.select(single({ row: 0, col: 4 }));
      view.paste('values');
      expect(workbook.cell(sheet, 1, 4)).toMatchObject({
        value: 2,
        formula: null,
      });
      expect(view.state.copied).not.toBeNull();
    });

    it('fills a larger selection with one cell copied', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: 'x' });
      view.copy();
      selecting(view, 2, 1, 4, 2);
      view.paste();
      expect(workbook.cell(sheet, 4, 2).value).toBe('x');
      expect(view.state.selection.ranges).toEqual([range(2, 1, 4, 2)]);
    });

    it('moves cut cells once, to another sheet too', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '7' });
      workbook.execute({ type: 'addSheet', index: 1 });
      const other = workbook.sheets[1]!.id;
      view.cut();
      expect(view.state.copied).toMatchObject({ cut: true });
      view.showSheet(other);
      view.select(single({ row: 1, col: 1 }));
      view.paste();
      expect(workbook.cell(sheet, 0, 0).value).toBeNull();
      expect(workbook.cell(other, 1, 1).value).toBe(7);
      expect(view.state.copied).toBeNull();
      view.paste();
      expect(workbook.canUndo).toBe(true);
      view.undo();
      expect(workbook.cell(sheet, 0, 0).value).toBe(7);
    });

    it('forgets what was copied on Escape, on another change, and on undo', () => {
      const { workbook, view, sheet } = setup();
      view.copy();
      view.cancel();
      expect(view.state.copied).toBeNull();

      view.copy();
      view.type('1');
      // An entry being typed keeps it, until it changes a cell.
      view.cancel();
      expect(view.state.copied).not.toBeNull();
      view.type('1');
      view.commit();
      expect(view.state.copied).toBeNull();

      view.copy();
      view.undo();
      expect(view.state.copied).toBeNull();
      expect(workbook.cell(sheet, 0, 0).value).toBeNull();
    });

    it('gives the cells copied or cut as they are shown, for the system clipboard', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({
        type: 'setCells',
        sheet,
        row: 0,
        col: 0,
        inputs: [
          ['12%', '=A1*2'],
          ['', 'x'],
        ],
      });
      selecting(view, 0, 0, 1, 1);
      expect(view.copy().text).toBe('12%\t0.24\n\tx\n');
      expect(view.cut().html).toContain('<td>0.24</td>');
    });

    it('pastes text from another application at the selection as typed, in one change', () => {
      const { workbook, view, sheet } = setup();
      selecting(view, 2, 1, 5, 5);
      view.pasteText('12%\t2024-01-02\r\n=B3*2\r\n');
      expect(workbook.cell(sheet, 2, 1).display).toBe('12%');
      expect(workbook.cell(sheet, 2, 2).input).toBe('1/2/2024');
      expect(workbook.cell(sheet, 3, 1).value).toBe(0.24);
      expect(view.state.selection.ranges).toEqual([range(2, 1, 3, 2)]);
      view.undo();
      expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
      expect(workbook.canUndo).toBe(false);
      // Nothing to paste changes nothing.
      view.pasteText('');
      expect(workbook.canUndo).toBe(false);
    });

    it('pastes the cells copied in Gridworks, formulas and all, when the text is theirs', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '1' });
      workbook.execute({
        type: 'setCell',
        sheet,
        row: 0,
        col: 1,
        input: '=A1+1',
      });
      selecting(view, 0, 0, 0, 1);
      const { text } = view.copy();
      view.select(single({ row: 2, col: 0 }));
      // As Windows puts it on the clipboard.
      view.pasteText(text.replaceAll('\n', '\r\n'));
      expect(workbook.cell(sheet, 2, 1).formula).toBe('=A3+1');
      // Other text is pasted as typed, even with cells copied.
      view.select(single({ row: 4, col: 0 }));
      view.pasteText('1\t=A1+1');
      expect(workbook.cell(sheet, 4, 1).formula).toBe('=A1+1');
    });

    it('commits the entry before copying or pasting', () => {
      const { workbook, view, sheet } = setup();
      view.type('5');
      view.copy();
      expect(workbook.cell(sheet, 0, 0).value).toBe(5);
      view.select(single({ row: 1, col: 0 }));
      view.type('6');
      view.paste();
      expect(workbook.cell(sheet, 1, 0).value).toBe(5);
    });
  });

  describe('fill', () => {
    const range = (top: number, left: number, bottom = top, right = left) => ({
      top,
      left,
      bottom,
      right,
    });
    const selecting = (
      view: View,
      top: number,
      left: number,
      bottom = top,
      right = left,
    ) =>
      view.select({
        ...single({ row: top, col: left }),
        ranges: [range(top, left, bottom, right)],
        focus: { row: bottom, col: right },
      });
    const set = (
      workbook: ReturnType<typeof setup>['workbook'],
      sheet: string,
      cells: Record<string, string>,
    ) => {
      for (const [place, input] of Object.entries(cells)) {
        const [row, col] = place.split(',').map(Number);
        workbook.execute({
          type: 'setCell',
          sheet,
          row: row!,
          col: col!,
          input,
        });
      }
    };

    it('drags the fill handle the way the pointer went furthest, and fills there', () => {
      const { workbook, view, sheet } = setup();
      set(workbook, sheet, { '2,2': '1', '3,2': '2' });
      selecting(view, 2, 2, 3, 2);
      view.previewFill({ row: 7, col: 4 });
      expect(view.state.filling).toEqual(range(2, 2, 7, 2));
      view.previewFill({ row: 4, col: 6 });
      expect(view.state.filling).toEqual(range(2, 2, 3, 6));
      view.previewFill({ row: 0, col: 1 });
      expect(view.state.filling).toEqual(range(0, 2, 3, 2));
      view.previewFill({ row: 3, col: 0 });
      expect(view.state.filling).toEqual(range(2, 0, 3, 2));
      // Inside the selection, there is nothing to fill.
      view.previewFill({ row: 3, col: 2 });
      expect(view.state.filling).toBeNull();
      view.fill();
      expect(workbook.canUndo).toBe(true);
      expect(workbook.cell(sheet, 4, 2).value).toBeNull();

      view.previewFill({ row: 5, col: 2 });
      view.fill();
      expect(view.state.filling).toBeNull();
      expect(workbook.cell(sheet, 5, 2).value).toBe(4);
      expect(view.state.selection.ranges).toEqual([range(2, 2, 5, 2)]);
      expect(view.state.selection.active).toEqual({ row: 2, col: 2 });
    });

    it('fills down as far as the column beside has data, on the left or else the right', () => {
      const { workbook, view, sheet } = setup();
      set(workbook, sheet, {
        '0,0': '1',
        '1,0': '2',
        '2,0': '3',
        '0,1': '=A1*2',
        '0,3': 'x',
        '1,3': 'y',
      });
      view.select(single({ row: 0, col: 1 }));
      view.autoFill();
      expect(workbook.cell(sheet, 2, 1).formula).toBe('=A3*2');
      expect(workbook.cell(sheet, 3, 1).value).toBeNull();
      expect(view.state.selection.ranges).toEqual([range(0, 1, 2, 1)]);

      set(workbook, sheet, { '0,2': 'z' });
      view.select(single({ row: 0, col: 2 }));
      view.autoFill();
      expect(workbook.cell(sheet, 2, 2).value).toBe('z');

      set(workbook, sheet, { '0,6': 'alone' });
      view.select(single({ row: 0, col: 6 }));
      view.autoFill();
      expect(workbook.cell(sheet, 1, 6).value).toBeNull();
    });

    it('fills down and right from the first row or column of the selection, or from the cells before it', () => {
      const { workbook, view, sheet } = setup();
      set(workbook, sheet, { '0,0': '1', '0,1': '=A1+1' });
      selecting(view, 0, 0, 2, 1);
      view.copy();
      view.fillDown();
      // Unlike a paste, a fill forgets what was copied, as in Excel.
      expect(view.state.copied).toBeNull();
      expect(workbook.cell(sheet, 2, 0).value).toBe(1);
      expect(workbook.cell(sheet, 2, 1).formula).toBe('=A3+1');
      expect(view.state.selection.ranges).toEqual([range(0, 0, 2, 1)]);

      selecting(view, 3, 0, 3, 1);
      view.fillDown();
      expect(workbook.cell(sheet, 3, 1).formula).toBe('=A4+1');

      // Numbers are copied, not counted on.
      selecting(view, 0, 0, 0, 3);
      view.fillRight();
      expect(workbook.cell(sheet, 0, 3).value).toBe(1);
      view.select(single({ row: 1, col: 2 }));
      view.fillRight();
      expect(workbook.cell(sheet, 1, 2).formula).toBe('=B2+1');

      // From the first row or column, there is nothing to fill from.
      const empty = setup();
      empty.view.fillDown();
      empty.view.fillRight();
      expect(empty.workbook.canUndo).toBe(false);
    });
  });

  describe('rows and columns', () => {
    const origin = { row: 0, col: 0 };
    /** Selects whole rows or columns, adding them to the selection when asked. */
    function selectLines(
      view: View,
      kind: 'row' | 'column',
      from: number,
      to = from,
      add = false,
    ) {
      const target = (index: number) =>
        kind === 'row'
          ? ({ kind, row: index } as const)
          : ({ kind, col: index } as const);
      const start = pick(
        view.state.selection,
        target(from),
        add ? 'add' : 'replace',
        origin,
      );
      view.select(pick(start, target(to), 'extend', origin));
    }

    it('inserts and deletes the selected rows or columns, each span at once', () => {
      const { workbook, view, sheet } = setup();
      for (let row = 0; row < 5; row++) {
        workbook.execute({
          type: 'setCell',
          sheet,
          row,
          col: 0,
          input: String(row + 1),
        });
      }
      selectLines(view, 'row', 1, 2);
      selectLines(view, 'row', 4, 4, true);
      view.insertLines('rows');
      const column = () =>
        Array.from(
          { length: 8 },
          (_, row) => workbook.cell(sheet, row, 0).display,
        );
      expect(column()).toEqual(['1', '', '', '2', '3', '4', '', '5']);
      // The rows inserted are selected, so deleting them takes them back.
      view.deleteLines('rows');
      expect(column()).toEqual(['1', '2', '3', '4', '5', '', '', '']);
      workbook.undo();
      workbook.undo();
      // With whole rows selected, cells selected beside them are left out.
      selectLines(view, 'row', 0);
      view.select(
        pick(
          view.state.selection,
          { kind: 'cell', row: 3, col: 1 },
          'add',
          origin,
        ),
      );
      view.deleteLines('rows');
      expect(column().slice(0, 4)).toEqual(['2', '3', '4', '5']);
      workbook.undo();
      selectLines(view, 'column', 0);
      view.insertLines('columns');
      expect(workbook.cell(sheet, 0, 1).display).toBe('1');
    });

    it('hides and shows again the selected rows or columns', () => {
      const { workbook, view, sheet } = setup();
      selectLines(view, 'column', 1, 2);
      expect(view.hasHidden('columns')).toBe(false);
      view.hideLines('columns', true);
      expect([...workbook.layout(sheet).columns.hidden]).toEqual([1, 2]);
      selectLines(view, 'column', 0, 3);
      expect(view.hasHidden('columns')).toBe(true);
      view.hideLines('columns', false);
      expect(workbook.layout(sheet).columns.hidden.size).toBe(0);
    });

    it('resizes a line dragged by its border, or every selected line with it', () => {
      const { workbook, view, sheet } = setup();
      view.previewResize('columns', 3, 150);
      expect(view.state.resizing).toEqual({
        direction: 'columns',
        index: 3,
        size: 150,
      });
      view.resize();
      expect(view.state.resizing).toBeNull();
      expect([...workbook.layout(sheet).columns.sizes]).toEqual([[3, 150]]);

      selectLines(view, 'row', 1, 2);
      selectLines(view, 'row', 5, 5, true);
      view.previewResize('rows', 2, 40);
      view.resize();
      expect([...workbook.layout(sheet).rows.sizes]).toEqual([
        [1, 40],
        [2, 40],
        [5, 40],
      ]);
      // A line outside the selection is resized alone, and dragged to
      // nothing it is hidden.
      view.previewResize('rows', 8, -3);
      view.resize();
      expect([...workbook.layout(sheet).rows.hidden]).toEqual([8]);
      expect(workbook.layout(sheet).rows.sizes.has(8)).toBe(false);
      // Without a drag, there is nothing to resize.
      const changes = workbook.canUndo;
      view.resize();
      expect(workbook.canUndo).toBe(changes);
    });

    it('fits columns to their widest text, and rows to the standard height', () => {
      const { workbook, view, sheet } = setup();
      const set = (row: number, col: number, input: string) =>
        workbook.execute({ type: 'setCell', sheet, row, col, input });
      set(0, 1, 'abc');
      set(3, 1, 'abcdefghij');
      set(0, 2, 'x');
      // Text in a hidden row does not count.
      set(6, 2, 'hidden text');
      workbook.execute({
        type: 'setLines',
        sheet,
        direction: 'rows',
        spans: [{ start: 6, end: 6, hidden: true }],
      });
      const measure = (text: string) => text.length * 10;
      selectLines(view, 'column', 1, 3);
      view.fitLines('columns', 2, measure);
      const { sizes } = workbook.layout(sheet).columns;
      expect(sizes.get(1)).toBe(108);
      expect(sizes.get(2)).toBe(18);
      // An empty column gets the standard width.
      expect(sizes.has(3)).toBe(false);

      // Wrapped text fits its column, so it does not widen it.
      view.select(single({ row: 0, col: 4 }));
      view.type('a long wrapped text');
      view.commit();
      view.select(single({ row: 0, col: 4 }));
      view.toggle('wrap');
      view.fitLines('columns', 4, measure);
      expect(workbook.layout(sheet).columns.sizes.has(4)).toBe(false);

      selectLines(view, 'row', 0);
      view.previewResize('rows', 0, 60);
      view.resize();
      view.fitLines('rows', 0, measure);
      expect(workbook.layout(sheet).rows.sizes.size).toBe(0);
    });

    it('selects the lines a change to rows or columns changed when it is undone', () => {
      const { workbook, view } = setup();
      selectLines(view, 'column', 2);
      view.deleteLines('columns');
      view.select(single({ row: 5, col: 5 }));
      view.undo();
      expect(view.state.selection.ranges).toEqual([
        { top: 0, left: 2, bottom: SHEET_ROWS - 1, right: 2 },
      ]);
      expect(workbook.canRedo).toBe(true);
    });
  });

  describe('formats', () => {
    it('formats the selected cells, after committing a waiting entry', () => {
      const { workbook, view, sheet } = setup();
      view.select(single({ row: 1, col: 1 }));
      view.type('x');
      view.select(
        pick(view.state.selection, { kind: 'cell', row: 3, col: 2 }, 'add', {
          row: 0,
          col: 0,
        }),
      );
      view.format({ fill: '#ffff00', horizontal: 'right' });
      expect(workbook.cell(sheet, 1, 1)).toMatchObject({
        display: 'x',
        style: { fill: '#ffff00', horizontal: 'right' },
      });
      expect(workbook.cell(sheet, 3, 2).style.fill).toBe('#ffff00');
      expect(workbook.cell(sheet, 2, 2).style).toEqual({});
    });

    it('toggles bold, italic, underline and wrap by the active cell, as Excel does', () => {
      const { workbook, view, sheet } = setup();
      view.select(single({ row: 1, col: 0 }));
      view.toggle('bold');
      // The active cell is bold, so the whole selection stops being bold.
      view.select(
        pick(view.state.selection, { kind: 'cell', row: 0, col: 0 }, 'extend', {
          row: 0,
          col: 0,
        }),
      );
      view.toggle('bold');
      expect(workbook.cell(sheet, 1, 0).style.bold).toBeUndefined();
      view.toggle('bold');
      expect(workbook.cell(sheet, 0, 0).style.bold).toBe(true);
      expect(workbook.cell(sheet, 1, 0).style.bold).toBe(true);
      view.toggle('wrap');
      expect(workbook.cell(sheet, 0, 0).style).toEqual({
        bold: true,
        wrap: true,
      });
    });
    it('adds and removes decimals and the thousands separator by the active cell', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({
        type: 'setCell',
        sheet,
        row: 0,
        col: 0,
        input: '1234.5',
      });
      workbook.execute({ type: 'setCell', sheet, row: 1, col: 0, input: '2' });
      view.select(
        pick(view.state.selection, { kind: 'cell', row: 1, col: 0 }, 'extend', {
          row: 0,
          col: 0,
        }),
      );
      view.adjustFormat('moreDecimals');
      expect(workbook.cell(sheet, 0, 0).display).toBe('1234.50');
      expect(workbook.cell(sheet, 1, 0).display).toBe('2.00');
      view.adjustFormat('thousands');
      view.adjustFormat('fewerDecimals');
      expect(workbook.cell(sheet, 0, 0).display).toBe('1,234.5');
      view.adjustFormat('thousands');
      expect(workbook.cell(sheet, 0, 0).style.numberFormat).toBe('0.0');
      workbook.undo();
      expect(workbook.cell(sheet, 0, 0).style.numberFormat).toBe('#,##0.0');

      // A date in the active cell changes nothing, in it or the others.
      workbook.execute({
        type: 'setCell',
        sheet,
        row: 0,
        col: 0,
        input: '1/2/2026',
      });
      const canRedo = workbook.canRedo;
      view.adjustFormat('moreDecimals');
      expect(workbook.cell(sheet, 0, 0).style.numberFormat).toBe('m/d/yyyy');
      expect(workbook.cell(sheet, 1, 0).style.numberFormat).toBe('#,##0.0');
      expect(workbook.canRedo).toBe(canRedo);
    });
  });

  describe('borders, merges and the format painter', () => {
    const at = (row: number, col: number) => ({ row, col });
    const range = (
      top: number,
      left: number,
      bottom: number,
      right: number,
    ) => ({
      top,
      left,
      bottom,
      right,
    });
    const step = (rows: number, cols: number, extend = false) =>
      ({ type: 'step', rows, cols, extend }) as const;

    function merged() {
      const { workbook, view, sheet } = setup();
      view.select(single(at(1, 1)));
      view.select(
        pick(
          view.state.selection,
          { kind: 'cell', ...at(2, 2) },
          'extend',
          at(0, 0),
        ),
      );
      view.merge();
      return { workbook, view, sheet };
    }

    it('sets borders on the selected cells', () => {
      const { workbook, view, sheet } = setup();
      view.select(
        pick(
          single(at(0, 0)),
          { kind: 'cell', ...at(1, 1) },
          'extend',
          at(0, 0),
        ),
      );
      view.setBorders('outside', { line: 'medium', color: '#0070c0' });
      expect(workbook.cell(sheet, 1, 1).style).toEqual({
        borderRight: { line: 'medium', color: '#0070c0' },
        borderBottom: { line: 'medium', color: '#0070c0' },
      });
    });

    it('merges the selected cells, and selects a merged cell as one', () => {
      const { workbook, view, sheet } = merged();
      expect(workbook.merges(sheet)).toEqual([range(1, 1, 2, 2)]);

      // A click inside it selects all of it, its top-left cell active.
      view.select(single(at(2, 2)));
      expect(view.state.selection.ranges).toEqual([range(1, 1, 2, 2)]);
      expect(view.state.selection.active).toEqual(at(1, 1));

      // A range that reaches into it grows to take it in.
      view.select(
        pick(
          single(at(0, 0)),
          { kind: 'cell', ...at(1, 1) },
          'extend',
          at(0, 0),
        ),
      );
      expect(view.state.selection.ranges).toEqual([range(0, 0, 2, 2)]);

      expect(view.canMerge).toBe(true);
      expect(view.canUnmerge).toBe(true);
      view.select(single(at(1, 1)));
      expect(view.canMerge).toBe(false);
      view.unmerge();
      expect(workbook.merges(sheet)).toEqual([]);
      expect(view.canUnmerge).toBe(false);
    });

    it('moves the keyboard over a merged cell as over one cell', () => {
      const { view } = merged();
      view.select(single(at(1, 0)));
      view.navigate(step(0, 1));
      expect(view.state.selection.ranges).toEqual([range(1, 1, 2, 2)]);
      view.navigate(step(0, 1));
      expect(view.state.selection.active).toEqual(at(1, 3));
      view.navigate(step(0, -1));
      view.navigate(step(1, 0));
      expect(view.state.selection.active).toEqual(at(3, 1));

      // Shift with an arrow grows the selection beyond the merged cell.
      view.select(single(at(1, 1)));
      view.navigate(step(1, 0, true));
      expect(view.state.selection.ranges).toEqual([range(1, 1, 3, 2)]);
    });

    it('paints the formats of a range onto the next range selected, once', () => {
      const { workbook, view, sheet } = setup();
      workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '5%' });
      workbook.execute({ type: 'setCell', sheet, row: 4, col: 4, input: 'x' });
      view.format({ bold: true, fill: '#ffff00' });
      view.setBorders('all', { line: 'thin' });
      const look = workbook.cell(sheet, 0, 0).style;

      view.startPainter();
      expect(view.state.painter).toEqual({ sheet, range: range(0, 0, 0, 0) });
      view.select(single(at(4, 4)));
      view.paint();
      expect(workbook.cell(sheet, 4, 4)).toMatchObject({
        display: 'x',
        style: look,
      });
      expect(view.state.painter).toBeNull();

      // Without the painter, selecting paints nothing; Escape puts it away.
      view.select(single(at(6, 6)));
      view.paint();
      expect(workbook.cell(sheet, 6, 6).style).toEqual({});
      view.startPainter();
      view.cancel();
      expect(view.state.painter).toBeNull();
      // Clicked again, the button puts it away too.
      view.startPainter();
      view.startPainter();
      expect(view.state.painter).toBeNull();

      workbook.undo();
      expect(workbook.cell(sheet, 4, 4).style).toEqual({});
    });
  });

  it('freezes the rows above and columns left of the active cell, the top row or the first column', () => {
    const { workbook, view, sheet } = setup();
    const frozen = () => workbook.layout(sheet).frozen;
    // At A1 there is nothing above or left to freeze, nor anything to unfreeze.
    expect(view.canFreeze).toBe(false);
    expect(view.canUnfreeze).toBe(false);
    view.type('1');
    view.select(single({ row: 3, col: 2 }));
    expect(view.canFreeze).toBe(true);
    view.freeze('panes');
    // A waiting entry is committed first.
    expect(workbook.cell(sheet, 0, 0).display).toBe('1');
    expect(frozen()).toEqual({ rows: 3, cols: 2 });
    expect(view.canUnfreeze).toBe(true);
    view.freeze('topRow');
    expect(frozen()).toEqual({ rows: 1, cols: 0 });
    view.freeze('firstColumn');
    expect(frozen()).toEqual({ rows: 0, cols: 1 });
    view.unfreeze();
    expect(frozen()).toBeUndefined();
    view.undo();
    expect(frozen()).toEqual({ rows: 0, cols: 1 });
    // Undo keeps the selection, as the change has no cells.
    expect(view.state.selection.active).toEqual({ row: 3, col: 2 });
  });
});
