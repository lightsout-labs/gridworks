import { describe, expect, it, vi } from 'vitest';
import {
  cellName,
  createWorkbook,
  SHEET_COLUMNS,
  SHEET_ROWS,
  Workbook,
  type Command,
  type Rule,
  type StyleChange,
} from './index';

function setup() {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  const set = (row: number, col: number, input: string) =>
    workbook.execute({ type: 'setCell', sheet, row, col, input });
  const show = (row: number, col: number) =>
    workbook.cell(sheet, row, col).display;
  return { workbook, sheet, set, show };
}

describe('workbook', () => {
  it('runs without a DOM', () => {
    expect(typeof document).toBe('undefined');
    expect(typeof window).toBe('undefined');
  });

  it('calculates a formula from another cell', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '2');
    set(1, 0, '=A1+1');
    expect(workbook.cell(sheet, 1, 0)).toEqual({
      value: 3,
      formula: '=A1+1',
      display: '3',
      input: '=A1+1',
      style: {},
    });
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 2,
      formula: null,
      display: '2',
      input: '2',
      style: {},
    });
  });

  it('recalculates a formula when a cell it reads changes', () => {
    const { set, show } = setup();
    set(1, 0, '=A1+1');
    set(2, 0, '=A2+1');
    expect(show(2, 0)).toBe('2');
    set(0, 0, '40');
    expect(show(2, 0)).toBe('42');
  });

  it('replaces a formula with a value, and a value with nothing', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '=B1+1');
    set(0, 0, 'text');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 'text',
      formula: null,
      display: 'text',
      input: 'text',
      style: {},
    });
    set(0, 0, '');
    expect(workbook.cell(sheet, 0, 0).value).toBeNull();
  });

  it('keeps an entry that is not a valid formula as text', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '=A1+');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: '=A1+',
      formula: null,
      display: '=A1+',
      input: '=A1+',
      style: {},
    });
  });

  it('reads a cell that was never set as empty', () => {
    const { workbook, sheet } = setup();
    expect(workbook.cell(sheet, 99_999, 99)).toEqual({
      value: null,
      formula: null,
      display: '',
      input: '',
      style: {},
    });
  });

  it('notifies subscribers of each change until they unsubscribe', () => {
    const { workbook, sheet, set } = setup();
    const listener = vi.fn();
    const unsubscribe = workbook.subscribe(listener);
    set(0, 0, '1');
    set(0, 0, '2');
    expect(listener).toHaveBeenCalledTimes(2);
    expect(listener).toHaveBeenLastCalledWith({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2',
    });
    unsubscribe();
    set(0, 0, '3');
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('restores values and formulas from a snapshot', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '2');
    set(1, 0, '=A1+1');
    set(4, 2, 'note');
    set(5, 0, '1/2/2026');
    set(6, 0, 'TRUE');
    const snapshot = JSON.parse(JSON.stringify(workbook.toSnapshot()));
    const restored = Workbook.fromSnapshot(snapshot);
    expect(restored.sheets).toEqual(workbook.sheets);
    expect(restored.cell(sheet, 0, 0).display).toBe('2');
    expect(restored.cell(sheet, 1, 0)).toEqual(workbook.cell(sheet, 1, 0));
    expect(restored.cell(sheet, 4, 2).display).toBe('note');
    expect(restored.cell(sheet, 5, 0)).toEqual(workbook.cell(sheet, 5, 0));
    expect(restored.cell(sheet, 6, 0).value).toBe(true);
    expect(restored.toSnapshot()).toEqual(snapshot);
  });

  it('refuses a snapshot of an unknown version', () => {
    const snapshot = { ...createWorkbook().toSnapshot(), version: 2 };
    expect(() => Workbook.fromSnapshot(snapshot as never)).toThrow(/version/);
  });

  it('refuses to restore a formula that does not parse', () => {
    const { workbook, sheet } = setup();
    const cell = { top: 0, left: 0, bottom: 0, right: 0 };
    expect(() =>
      workbook.execute({
        type: 'restore',
        sheet,
        ranges: [cell],
        cells: [{ row: 0, col: 0, value: null, formula: '=1+' }],
      }),
    ).toThrow(/=1\+/);
  });

  it('refuses a command for a sheet that does not exist', () => {
    const { workbook } = setup();
    expect(() =>
      workbook.execute({
        type: 'setCell',
        sheet: 'missing',
        row: 0,
        col: 0,
        input: '1',
      }),
    ).toThrow(/missing/);
  });

  it('names cells as Excel does', () => {
    expect(cellName(0, 0)).toBe('A1');
    expect(cellName(1, 25)).toBe('Z2');
    expect(cellName(9, 26)).toBe('AA10');
    expect(cellName(0, 701)).toBe('ZZ1');
    expect(cellName(0, 702)).toBe('AAA1');
  });

  it('totals the numbers of ranges and counts the cells that are not empty', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1');
    set(1, 0, '2');
    set(2, 0, 'text');
    set(0, 1, '=A1+A2');
    set(5, 5, '100'); // Outside both ranges.
    const ranges = [
      { top: 0, left: 0, bottom: 3, right: 0 },
      { top: 0, left: 1, bottom: 0, right: 1 },
    ];
    expect(workbook.aggregate(sheet, ranges)).toEqual({
      count: 4,
      numbers: 3,
      sum: 6,
    });
    expect(
      workbook.aggregate(sheet, [{ top: 10, left: 10, bottom: 20, right: 20 }]),
    ).toEqual({ count: 0, numbers: 0, sum: 0 });
  });

  it('totals a whole sheet of 100,000 by 100 cells quickly', () => {
    const columns = Array.from({ length: SHEET_COLUMNS }, () =>
      Array.from({ length: SHEET_ROWS }, () => 1),
    );
    const workbook = Workbook.fromSnapshot({
      version: 1,
      name: 'Big',
      sheets: [{ id: 's', name: 'S', columns, formulas: [] }],
    });
    const whole = {
      top: 0,
      left: 0,
      bottom: SHEET_ROWS - 1,
      right: SHEET_COLUMNS - 1,
    };
    const start = performance.now();
    expect(workbook.aggregate('s', [whole]).sum).toBe(
      SHEET_ROWS * SHEET_COLUMNS,
    );
    expect(performance.now() - start).toBeLessThan(500);
  });

  it('knows how far the cells in use extend', () => {
    const { workbook, sheet, set } = setup();
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
    set(4, 1, '1');
    set(2, 6, '=B5+1');
    expect(workbook.extent(sheet)).toEqual({ rows: 5, cols: 7 });
    set(4, 1, ''); // An emptied cell is no longer in use.
    expect(workbook.extent(sheet)).toEqual({ rows: 3, cols: 7 });
  });

  it('stores typed numbers, dates, times, percentages and booleans as values with their format', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(1, 0, '12%');
    set(2, 0, 'true');
    set(3, 0, '1,234.5');
    set(4, 0, '13:45');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 46024,
      formula: null,
      display: '1/2/2026',
      input: '1/2/2026',
      style: { numberFormat: 'm/d/yyyy' },
    });
    expect(workbook.cell(sheet, 1, 0)).toMatchObject({
      value: 0.12,
      display: '12%',
      input: '12%',
    });
    expect(workbook.cell(sheet, 2, 0)).toMatchObject({
      value: true,
      display: 'TRUE',
    });
    expect(workbook.cell(sheet, 3, 0)).toMatchObject({
      value: 1234.5,
      display: '1,234.50',
      input: '1234.5',
    });
    expect(workbook.cell(sheet, 4, 0).display).toBe('13:45');
  });

  it("keeps a cell's format when a plain number replaces its value, as Excel does", () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(0, 0, '46025');
    expect(workbook.cell(sheet, 0, 0).display).toBe('1/3/2026');
    set(0, 0, '50%');
    expect(workbook.cell(sheet, 0, 0).display).toBe('50%');
  });

  it('shows the result of a formula in the format of its cell', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(0, 0, '=1+46024');
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      display: '1/3/2026',
      input: '=1+46024',
    });
  });

  it('adds booleans as numbers in formulas', () => {
    const { set, show } = setup();
    set(0, 0, 'TRUE');
    set(0, 1, '=A1+1');
    expect(show(0, 1)).toBe('2');
  });

  it('clears the values and formulas of ranges, and keeps their formats', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '1');
    set(1, 1, '=A1+1');
    set(2, 0, '12%');
    set(5, 5, 'outside');
    const listener = vi.fn();
    workbook.subscribe(listener);
    const command = {
      type: 'clear',
      sheet,
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 2, left: 0, bottom: SHEET_ROWS - 1, right: 0 },
      ],
    } as const;
    workbook.execute(command);
    expect(listener).toHaveBeenCalledWith(command);
    expect(workbook.cell(sheet, 0, 0).value).toBeNull();
    expect(workbook.cell(sheet, 1, 1)).toMatchObject({
      value: null,
      formula: null,
    });
    expect(workbook.cell(sheet, 2, 0).value).toBeNull();
    expect(show(5, 5)).toBe('outside');
    expect(workbook.extent(sheet)).toEqual({ rows: 6, cols: 6 });
    set(2, 0, '0.5');
    expect(show(2, 0)).toBe('50%');
  });

  it('tells listeners the year it gave a date typed without one, which a replay keeps', () => {
    const { workbook, sheet } = setup();
    const listener = vi.fn();
    workbook.subscribe(listener);
    workbook.execute({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2-Jan',
    });
    const year = new Date().getFullYear();
    expect(listener).toHaveBeenCalledWith({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2-Jan',
      year,
    });

    const replay = createWorkbook();
    replay.execute({ ...listener.mock.calls[0]![0], year: 2030 });
    expect(replay.cell(sheet, 0, 0).input).toBe('1/2/2030');
  });
});

describe('setCells', () => {
  it('sets a block of typed inputs as one change, recognizing each as typed', () => {
    const { workbook, sheet, show } = setup();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    const change = workbook.execute({
      type: 'setCells',
      sheet,
      row: 1,
      col: 1,
      inputs: [
        ['12%', '1,234.5', '=B2*2'],
        ['2024-01-02', 'TRUE'],
        ['', 'text'],
      ],
    });
    expect(change).toEqual({
      sheet,
      ranges: [{ top: 1, left: 1, bottom: 3, right: 3 }],
    });
    expect(show(1, 1)).toBe('12%');
    expect(workbook.cell(sheet, 1, 2).value).toBe(1234.5);
    expect(workbook.cell(sheet, 1, 3).value).toBe(0.24);
    expect(workbook.cell(sheet, 2, 1).input).toBe('1/2/2024');
    expect(workbook.cell(sheet, 2, 2).value).toBe(true);
    expect(workbook.cell(sheet, 3, 2).value).toBe('text');
    expect(Workbook.fromSnapshot(replay.toSnapshot()).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );

    workbook.undo();
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
    expect(workbook.canUndo).toBe(false);
    workbook.redo();
    expect(show(1, 1)).toBe('12%');
  });

  it('refuses an empty block', () => {
    const { workbook, sheet } = setup();
    for (const inputs of [[], [[]]]) {
      expect(() =>
        workbook.execute({ type: 'setCells', sheet, row: 0, col: 0, inputs }),
      ).toThrow();
    }
    expect(workbook.canUndo).toBe(false);
  });

  it('keeps what fits at the edge of the sheet, and the year of a date typed without one', () => {
    const { workbook, sheet } = setup();
    const listener = vi.fn();
    workbook.subscribe(listener);
    const change = workbook.execute({
      type: 'setCells',
      sheet,
      row: SHEET_ROWS - 1,
      col: SHEET_COLUMNS - 1,
      inputs: [
        ['2-Jan', 'beyond'],
        ['beyond', 'beyond'],
      ],
    });
    const edge = { top: SHEET_ROWS - 1, left: SHEET_COLUMNS - 1 };
    expect(change.ranges).toEqual([
      { ...edge, bottom: edge.top, right: edge.left },
    ]);
    expect(workbook.extent(sheet)).toEqual({
      rows: SHEET_ROWS,
      cols: SHEET_COLUMNS,
    });
    expect(listener.mock.calls[0]![0]).toMatchObject({
      year: new Date().getFullYear(),
    });
  });
});

describe('undo and redo', () => {
  it('undoes a change to a value, a formula or a format, and redo it', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '12%');
    set(0, 0, '=1+1');
    set(0, 0, '1/2/2026');
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: 2,
      formula: '=1+1',
      display: '200%',
    });
    workbook.undo();
    expect(show(0, 0)).toBe('12%');
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: null,
      formula: null,
    });
    set(0, 0, '0.5');
    expect(show(0, 0)).toBe('0.5'); // The format is undone too.
    workbook.undo();
    workbook.redo();
    workbook.redo(); // Nothing more to redo after a new change.
    expect(show(0, 0)).toBe('0.5');
  });

  it('undoes clearing ranges, and redoes it', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '1');
    set(1, 1, '=A1+1');
    set(3, 0, '12%');
    set(5, 5, 'outside');
    workbook.execute({
      type: 'clear',
      sheet,
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 3, left: 0, bottom: SHEET_ROWS - 1, right: 0 },
      ],
    });
    workbook.undo();
    expect(show(0, 0)).toBe('1');
    expect(workbook.cell(sheet, 1, 1)).toMatchObject({
      formula: '=A1+1',
      display: '2',
    });
    expect(show(3, 0)).toBe('12%');
    workbook.redo();
    expect(workbook.extent(sheet)).toEqual({ rows: 6, cols: 6 });
    expect(workbook.cell(sheet, 1, 1).formula).toBeNull();
  });

  it('gives the cells that each undo and redo changed', () => {
    const { workbook, sheet, set } = setup();
    expect(workbook.undo()).toBeNull();
    set(2, 3, 'x');
    const ranges = [
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 4, left: 4, bottom: 4, right: 4 },
    ];
    workbook.execute({ type: 'clear', sheet, ranges });
    expect(workbook.undo()).toEqual({ sheet, ranges });
    const cell = { sheet, ranges: [{ top: 2, left: 3, bottom: 2, right: 3 }] };
    expect(workbook.undo()).toEqual(cell);
    expect(workbook.redo()).toEqual(cell);
    expect(workbook.redo()).toEqual({ sheet, ranges });
    expect(workbook.redo()).toBeNull();
  });

  it('undoes the last 100 changes in order, and no more', () => {
    const { workbook, sheet, set, show } = setup();
    for (let step = 1; step <= 101; step++) set(0, 0, String(step));
    for (let step = 100; step >= 1; step--) {
      expect(workbook.canUndo).toBe(true);
      workbook.undo();
      expect(show(0, 0)).toBe(String(step));
    }
    expect(workbook.canUndo).toBe(false);
    expect(workbook.undo()).toBeNull();
    expect(workbook.cell(sheet, 0, 0).value).toBe(1);
  });

  it('tells whether there is something to undo and to redo, already to its listeners', () => {
    const { workbook, set } = setup();
    const seen: boolean[][] = [];
    workbook.subscribe(() => seen.push([workbook.canUndo, workbook.canRedo]));
    expect([workbook.canUndo, workbook.canRedo]).toEqual([false, false]);
    set(0, 0, '1');
    workbook.undo();
    workbook.redo();
    expect(seen).toEqual([
      [true, false],
      [false, true],
      [true, false],
    ]);
  });

  it('discards what could be redone when a new change follows an undo', () => {
    const { workbook, set, show } = setup();
    set(0, 0, '1');
    set(0, 0, '2');
    workbook.undo();
    set(1, 0, '3');
    expect(workbook.canRedo).toBe(false);
    expect(workbook.redo()).toBeNull();
    expect(show(0, 0)).toBe('1');
  });

  it('tells listeners of each undo and redo as a command that replays it', () => {
    const { workbook, sheet, set } = setup();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1/2/2026');
    set(1, 0, '=A1+1');
    workbook.execute({
      type: 'clear',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
    });
    workbook.undo();
    workbook.undo();
    workbook.redo();
    workbook.undo();
    workbook.undo();
    workbook.redo();
    const stored = JSON.parse(JSON.stringify(replay.toSnapshot()));
    expect(Workbook.fromSnapshot(stored).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
    expect(replay.cell(sheet, 0, 0).display).toBe('1/2/2026');
  });

  it('opens a stored workbook with its recorded changes, which cannot be undone', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1');
    const opened = Workbook.fromSnapshot(workbook.toSnapshot(), [
      { type: 'setCell', sheet, row: 1, col: 0, input: '=A1+1' },
    ]);
    expect(opened.cell(sheet, 1, 0).display).toBe('2');
    expect(opened.canUndo).toBe(false);
  });
});

describe('sheets', () => {
  const names = (workbook: Workbook) =>
    workbook.sheets.map((sheet) => sheet.name);

  it('adds empty sheets with the next free names, where asked', () => {
    const { workbook, sheet } = setup();
    const added = workbook.execute({ type: 'addSheet', index: 1 });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet2']);
    expect(added).toEqual({ sheet: workbook.sheets[1]!.id, ranges: [] });
    workbook.execute({ type: 'addSheet', index: 0 });
    expect(names(workbook)).toEqual(['Sheet3', 'Sheet1', 'Sheet2']);
    expect(new Set(workbook.sheets.map((each) => each.id)).size).toBe(3);
    expect(workbook.cell(added.sheet, 0, 0).value).toBeNull();
    expect(workbook.sheets[1]!.id).toBe(sheet);
  });

  it('duplicates a sheet beside it, with its contents and formats', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '50%');
    set(1, 0, '=A1+1');
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 1,
      copyOf: sheet,
    });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet1 (2)']);
    expect(workbook.cell(copy, 0, 0).display).toBe('50%');
    expect(workbook.cell(copy, 1, 0).formula).toBe('=A1+1');
    // The copy is a sheet of its own.
    set(0, 0, '1');
    expect(workbook.cell(copy, 1, 0).value).toBe(1.5);
    workbook.execute({ type: 'addSheet', index: 2, copyOf: sheet });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet1 (2)', 'Sheet1 (3)']);
  });

  it('renames, recolors and moves a sheet', () => {
    const { workbook, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    workbook.execute({ type: 'setSheet', sheet, name: 'Sales' });
    workbook.execute({ type: 'setSheet', sheet, color: '#e5484d' });
    expect(workbook.sheets[0]).toEqual({
      id: sheet,
      name: 'Sales',
      color: '#e5484d',
    });
    workbook.execute({ type: 'setSheet', sheet, index: 1 });
    expect(names(workbook)).toEqual(['Sheet2', 'Sales']);
    workbook.execute({ type: 'setSheet', sheet, color: null });
    expect(workbook.sheets[1]).toEqual({ id: sheet, name: 'Sales' });
  });

  it('deletes a sheet, and shows the one that takes its place', () => {
    const { workbook, sheet } = setup();
    const { sheet: second } = workbook.execute({ type: 'addSheet', index: 1 });
    const { sheet: third } = workbook.execute({ type: 'addSheet', index: 2 });
    expect(workbook.execute({ type: 'deleteSheet', sheet: second })).toEqual({
      sheet: third,
      ranges: [],
    });
    expect(workbook.execute({ type: 'deleteSheet', sheet: third })).toEqual({
      sheet,
      ranges: [],
    });
    expect(names(workbook)).toEqual(['Sheet1']);
  });

  it('refuses to delete the last sheet', () => {
    const { workbook, sheet } = setup();
    expect(() => workbook.execute({ type: 'deleteSheet', sheet })).toThrow();
    expect(names(workbook)).toEqual(['Sheet1']);
    expect(workbook.canUndo).toBe(false);
  });

  it('tells what is wrong with a sheet name, as Excel does', () => {
    const { workbook, sheet } = setup();
    const { sheet: second } = workbook.execute({ type: 'addSheet', index: 1 });
    expect(workbook.sheetNameProblem('Budget 2026', sheet)).toBeNull();
    expect(workbook.sheetNameProblem('sheet1', sheet)).toBeNull();
    expect(workbook.sheetNameProblem('SHEET1', second)).toBe('taken');
    expect(workbook.sheetNameProblem('  ', sheet)).toBe('empty');
    expect(workbook.sheetNameProblem('x'.repeat(32), sheet)).toBe('tooLong');
    for (const name of ['a:b', 'a/b', 'a\\b', 'a?', 'a*', '[a]', "'a"]) {
      expect(workbook.sheetNameProblem(name, sheet)).toBe('character');
    }
    expect(() =>
      workbook.execute({ type: 'setSheet', sheet, name: 'Sheet2' }),
    ).toThrow();
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet2']);
  });

  it('undoes and redoes each change to the sheets, showing the sheet it changed', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '7');
    const before = workbook.toSnapshot();
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 1,
      copyOf: sheet,
    });
    workbook.execute({
      type: 'setSheet',
      sheet: copy,
      name: 'Copy',
      color: '#30a46c',
    });
    workbook.execute({ type: 'setSheet', sheet: copy, index: 0 });
    workbook.execute({ type: 'deleteSheet', sheet });
    const after = workbook.toSnapshot();

    expect(workbook.undo()).toEqual({ sheet, ranges: [] });
    expect(names(workbook)).toEqual(['Copy', 'Sheet1']);
    expect(workbook.cell(sheet, 0, 0).value).toBe(7);
    expect(workbook.undo()).toEqual({ sheet: copy, ranges: [] });
    expect(names(workbook)).toEqual(['Sheet1', 'Copy']);
    workbook.undo();
    expect(workbook.sheets[1]).toEqual({ id: copy, name: 'Sheet1 (2)' });
    expect(workbook.undo()).toEqual({ sheet, ranges: [] });
    expect(workbook.toSnapshot()).toEqual(before);

    for (let step = 0; step < 4; step++) workbook.redo();
    expect(workbook.toSnapshot()).toEqual(after);
  });

  it('tells listeners of sheet changes as commands that replay them', () => {
    const { workbook, sheet, set } = setup();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1');
    workbook.execute({ type: 'addSheet', index: 1 });
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 0,
      copyOf: sheet,
    });
    workbook.execute({ type: 'setSheet', sheet: copy, name: 'Copy' });
    workbook.execute({ type: 'deleteSheet', sheet });
    workbook.undo();
    workbook.undo();
    workbook.redo();
    const stored = JSON.parse(JSON.stringify(replay.toSnapshot()));
    expect(Workbook.fromSnapshot(stored).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
  });
});

/** A workbook whose column B holds a chain of `length` formulas from A1. */
function chain(length: number): Workbook {
  return Workbook.fromSnapshot({
    version: 1,
    name: 'Chain',
    sheets: [
      {
        id: 's',
        name: 'S',
        columns: [[1]],
        formulas: Array.from({ length }, (_, row): [number, number, string] => [
          row,
          1,
          row === 0 ? '=A1+1' : `=B${row}+1`,
        ]),
      },
    ],
  });
}

describe('recalculation', () => {
  it('recalculates the formulas that depend on a cell, directly or indirectly, and no others', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '1');
    set(0, 1, '=A1*2');
    set(0, 2, '=SUM(B1,A1:A3)');
    set(0, 3, '=D2+1');
    expect(show(0, 2)).toBe('3');
    expect(show(0, 3)).toBe('1');
    const before = workbook.calculations;
    set(2, 0, '10');
    expect(show(0, 2)).toBe('13');
    expect(show(0, 1)).toBe('2');
    expect(show(0, 3)).toBe('1');
    expect(workbook.calculations - before).toBe(1);
    set(0, 0, '5');
    expect(show(0, 2)).toBe('25');
    expect(workbook.calculations - before).toBe(3);
    expect(workbook.cell(sheet, 0, 1).value).toBe(10);
  });

  it('recalculates formulas over whole columns and rows', () => {
    const { set, show } = setup();
    set(0, 1, '=SUM(A:A)');
    set(1, 0, '=SUM(3:3)');
    set(2, 5, '4');
    expect(show(0, 1)).toBe('4');
    set(99_999, 0, '1');
    expect(show(0, 1)).toBe('5');
  });

  it('recalculates a chain of 100,000 formulas, calculating each once', () => {
    const workbook = chain(100_000);
    expect(workbook.cell('s', 99_999, 1).value).toBe(100_001);
    const before = workbook.calculations;
    workbook.execute({
      type: 'setCell',
      sheet: 's',
      row: 0,
      col: 0,
      input: '2',
    });
    expect(workbook.cell('s', 99_999, 1).value).toBe(100_002);
    expect(workbook.calculations - before).toBe(100_000);
  });

  it('calculates only the formulas read after a change, and the ones they need', () => {
    const workbook = chain(1000);
    expect(workbook.cell('s', 999, 1).value).toBe(1001);
    const before = workbook.calculations;
    workbook.execute({
      type: 'setCell',
      sheet: 's',
      row: 0,
      col: 0,
      input: '2',
    });
    expect(workbook.cell('s', 9, 1).value).toBe(12);
    expect(workbook.calculations - before).toBe(10);
  });

  it('reports a circular reference, gives its cells 0, and recovers when it is broken', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '=B1+1');
    expect(workbook.circularReference()).toBeNull();
    set(0, 1, '=A1+1');
    expect(workbook.circularReference()).toEqual({ sheet, row: 0, col: 1 });
    expect(show(0, 0)).toBe('0');
    expect(show(0, 1)).toBe('0');
    set(5, 5, '7');
    expect(workbook.circularReference()).toEqual({ sheet, row: 0, col: 1 });
    set(0, 1, '1');
    expect(workbook.circularReference()).toBeNull();
    expect(show(0, 0)).toBe('2');
    set(3, 3, '=SUM(D1:D10)');
    expect(workbook.circularReference()).toEqual({ sheet, row: 3, col: 3 });
    workbook.undo();
    expect(workbook.circularReference()).toBeNull();
  });

  it('reports a circular reference in a stored workbook when it opens', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '=A1');
    const opened = Workbook.fromSnapshot(workbook.toSnapshot());
    expect(opened.circularReference()).toEqual({ sheet, row: 0, col: 0 });
  });

  it('reads other sheets, and recalculates when they change', () => {
    const { workbook, set, show } = setup();
    const { sheet: other } = workbook.execute({ type: 'addSheet', index: 1 });
    workbook.execute({ type: 'setSheet', sheet: other, name: 'Q 1' });
    set(0, 0, "='Q 1'!A1*2+SUM('q 1'!B:B)");
    const setOther = (input: string) =>
      workbook.execute({
        type: 'setCell',
        sheet: other,
        row: 0,
        col: 0,
        input,
      });
    setOther('3');
    expect(show(0, 0)).toBe('6');
    setOther('4');
    expect(show(0, 0)).toBe('8');
    set(0, 1, '=Nowhere!A1');
    expect(show(0, 1)).toBe('#REF!');
  });

  it('rewrites references to a renamed sheet, and back on undo', () => {
    const { workbook, sheet, set, show } = setup();
    const { sheet: other } = workbook.execute({ type: 'addSheet', index: 1 });
    workbook.execute({
      type: 'setCell',
      sheet: other,
      row: 0,
      col: 0,
      input: '5',
    });
    set(0, 0, '=sheet2!A1+1');
    workbook.execute({ type: 'setSheet', sheet: other, name: 'Sales Data' });
    expect(workbook.cell(sheet, 0, 0).formula).toBe("='Sales Data'!A1+1");
    expect(show(0, 0)).toBe('6');
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0).formula).toBe('=Sheet2!A1+1');
    expect(show(0, 0)).toBe('6');
  });

  it('gives #REF! for references to a deleted sheet, until the deletion is undone', () => {
    const { workbook, set, show } = setup();
    const { sheet: other } = workbook.execute({ type: 'addSheet', index: 1 });
    workbook.execute({
      type: 'setCell',
      sheet: other,
      row: 0,
      col: 0,
      input: '5',
    });
    set(0, 0, '=Sheet2!A1+1');
    set(0, 1, '=A1*2');
    expect(show(0, 1)).toBe('12');
    workbook.execute({ type: 'deleteSheet', sheet: other });
    expect(show(0, 0)).toBe('#REF!');
    expect(show(0, 1)).toBe('#REF!');
    workbook.undo();
    expect(show(0, 1)).toBe('12');
  });

  it('stores formulas as Excel writes them', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '= sum( a1:b2 ) ');
    expect(workbook.cell(sheet, 0, 0).formula).toBe('=SUM(A1:B2)');
  });
});

describe('paste', () => {
  const range = (top: number, left: number, bottom = top, right = left) => ({
    top,
    left,
    bottom,
    right,
  });

  function pasting() {
    const { workbook, sheet, set, show } = setup();
    const formula = (row: number, col: number) =>
      workbook.cell(sheet, row, col).formula;
    const paste = (
      source: ReturnType<typeof range>,
      target: ReturnType<typeof range>,
      options: { cut?: boolean; only?: 'values' | 'formulas' | 'formats' } = {},
    ) =>
      workbook.execute({
        type: 'paste',
        sheet,
        target,
        source: { sheet, range: source },
        ...options,
      });
    return { workbook, sheet, set, show, formula, paste };
  }

  it('copies values and formulas, moving relative references and keeping absolute ones', () => {
    const { set, show, formula, paste } = pasting();
    set(0, 0, '1');
    set(1, 0, '2');
    set(0, 1, '=A1+$A$2+A$1+$A1');
    set(1, 1, 'text');
    const change = paste(range(0, 0, 1, 1), range(4, 2));
    expect(change.ranges).toEqual([range(4, 2, 5, 3)]);
    expect(show(4, 2)).toBe('1');
    expect(formula(4, 3)).toBe('=C5+$A$2+C$1+$A5');
    expect(show(5, 3)).toBe('text');
    // The source stays as it was.
    expect(formula(0, 1)).toBe('=A1+$A$2+A$1+$A1');
    // A reference moved off the sheet becomes #REF!.
    paste(range(0, 1), range(0, 4));
    paste(range(0, 4), range(0, 0));
    expect(formula(0, 0)).toBe('=#REF!+$A$2+#REF!+$A1');
  });

  it('carries formats with the values, and empties cells copied from empty ones', () => {
    const { workbook, sheet, set, show, paste } = pasting();
    set(0, 0, '12%');
    set(0, 2, 'old');
    set(1, 2, '1/2/2026');
    paste(range(0, 0, 1, 0), range(0, 2));
    expect(show(0, 2)).toBe('12%');
    expect(workbook.cell(sheet, 1, 2).value).toBeNull();
    set(1, 2, '3');
    expect(show(1, 2)).toBe('3');
  });

  it('fills a target of whole copies of the source, and puts one copy at any other', () => {
    const { set, show, formula, paste } = pasting();
    set(0, 0, '=B1*2');
    const change = paste(range(0, 0), range(2, 2, 4, 3));
    expect(change.ranges).toEqual([range(2, 2, 4, 3)]);
    expect(formula(2, 2)).toBe('=D3*2');
    expect(formula(4, 3)).toBe('=E5*2');
    set(10, 0, '1');
    set(11, 0, '2');
    paste(range(10, 0, 11, 0), range(10, 5, 12, 5));
    expect(show(10, 5)).toBe('1');
    expect(show(11, 5)).toBe('2');
    expect(show(12, 5)).toBe('');
    // A copy at the edge of the sheet keeps what fits.
    paste(range(10, 0, 11, 0), range(SHEET_ROWS - 1, 6));
    expect(show(SHEET_ROWS - 1, 6)).toBe('1');
  });

  it('pastes only the values, formula results among them, keeping the formats there', () => {
    const { workbook, sheet, set, show, formula, paste } = pasting();
    set(0, 0, '0.25');
    set(0, 1, '=A1*2');
    set(0, 2, '=1/0');
    set(1, 0, '10%');
    set(1, 1, '10%');
    set(1, 2, '10%');
    paste(range(0, 0, 0, 2), range(1, 0), { only: 'values' });
    expect(show(1, 0)).toBe('25%');
    expect(show(1, 1)).toBe('50%');
    expect(formula(1, 1)).toBeNull();
    // Gridworks has no error constants, so an error is kept as its text.
    expect(workbook.cell(sheet, 1, 2).value).toBe('#DIV/0!');
  });

  it('pastes only the formulas and values, keeping the formats there', () => {
    const { set, show, formula, paste } = pasting();
    set(0, 0, '300%');
    set(0, 1, '=A1*2');
    set(1, 0, '1/2/2026');
    set(1, 1, 'kept?');
    paste(range(0, 0, 0, 1), range(1, 0), { only: 'formulas' });
    expect(show(1, 0)).toBe('1/3/1900');
    expect(formula(1, 1)).toBe('=A2*2');
  });

  it('pastes only the formats, keeping the contents there', () => {
    const { workbook, set, show, paste } = pasting();
    set(0, 0, '12%');
    set(1, 0, '0.5');
    set(1, 1, '=A2');
    paste(range(0, 0), range(1, 0, 1, 1), { only: 'formats' });
    expect(show(1, 0)).toBe('50%');
    expect(show(1, 1)).toBe('50%');
    workbook.undo();
    expect(show(1, 0)).toBe('0.5');
    workbook.redo();
    expect(show(1, 1)).toBe('50%');
  });

  it('moves cut cells, and the references to them, leaving the source empty', () => {
    const { workbook, sheet, set, show, formula, paste } = pasting();
    set(0, 0, '1');
    set(1, 0, '12%');
    set(2, 0, '=A1+A2+B1');
    set(0, 3, '=SUM(A1:A2)+A1+$A$2+SUM(A1:A5)');
    const change = paste(range(0, 0, 2, 0), range(4, 1, 9, 9), {
      cut: true,
    });
    expect(change.ranges[0]).toEqual(range(4, 1, 6, 1));
    expect(workbook.cell(sheet, 0, 0).value).toBeNull();
    expect(workbook.cell(sheet, 2, 0).formula).toBeNull();
    expect(show(5, 1)).toBe('12%');
    // References inside the moved cells follow them; others stay.
    expect(formula(6, 1)).toBe('=B5+B6+B1');
    expect(formula(0, 3)).toBe('=SUM(B5:B6)+B5+$B$6+SUM(A1:A5)');
    // Setting the old place does not change the formulas.
    set(0, 0, '100');
    expect(show(6, 1)).toBe('1.12');
    // A cut pasted at the edge of the sheet moves back to keep every cell.
    paste(range(4, 1, 5, 1), range(SHEET_ROWS - 1, 0), { cut: true });
    expect(show(SHEET_ROWS - 2, 0)).toBe('1');
    expect(show(SHEET_ROWS - 1, 0)).toBe('12%');
  });

  it('moves cut cells to another sheet, with the references to and from them', () => {
    const { workbook, sheet, set, formula } = pasting();
    workbook.execute({ type: 'addSheet', index: 1, name: 'Data' });
    const data = workbook.sheets[1]!.id;
    set(0, 0, '1');
    set(1, 0, '=A1+C1');
    set(0, 1, '=A2*2');
    workbook.execute({
      type: 'setCell',
      sheet: data,
      row: 0,
      col: 0,
      input: '=Sheet1!A1',
    });
    workbook.execute({
      type: 'paste',
      sheet: data,
      target: range(5, 2),
      source: { sheet, range: range(0, 0, 1, 0) },
      cut: true,
    });
    expect(formula(0, 1)).toBe('=Data!C7*2');
    expect(workbook.cell(data, 6, 2).formula).toBe('=C6+Sheet1!C1');
    expect(workbook.cell(data, 0, 0).formula).toBe('=C6');
    expect(workbook.cell(data, 6, 2).value).toBe(1);
    workbook.undo();
    expect(formula(0, 1)).toBe('=A2*2');
    expect(formula(1, 0)).toBe('=A1+C1');
    expect(workbook.cell(data, 0, 0).formula).toBe('=Sheet1!A1');
    expect(workbook.cell(data, 6, 2).formula).toBeNull();
    workbook.redo();
    expect(formula(0, 1)).toBe('=Data!C7*2');
    expect(workbook.cell(data, 0, 0).formula).toBe('=C6');
  });

  it('undoes and redoes a paste, selecting the cells it changed, and tells listeners a command that replays it', () => {
    const { workbook, set, show, formula, paste } = pasting();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1');
    set(0, 1, '=A1+1');
    set(1, 0, '=TODAY()');
    set(3, 0, 'under');
    paste(range(1, 0), range(5, 5), { only: 'values' });
    paste(range(0, 0, 0, 1), range(0, 0, 3, 1));
    paste(range(0, 0, 1, 1), range(1, 0), { cut: true });
    expect(formula(2, 1)).toBe('=A3+1');
    const change = workbook.undo()!;
    expect(change.ranges).toEqual([range(1, 0, 2, 1), range(0, 0, 1, 1)]);
    expect(formula(0, 1)).toBe('=A1+1');
    expect(show(3, 0)).toBe('1');
    workbook.undo();
    workbook.undo();
    expect(show(3, 0)).toBe('under');
    workbook.redo();
    workbook.redo();
    workbook.redo();
    expect(formula(2, 1)).toBe('=A3+1');
    const stored = JSON.parse(JSON.stringify(replay.toSnapshot()));
    expect(Workbook.fromSnapshot(stored).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
  });
});

describe('fill', () => {
  const range = (top: number, left: number, bottom = top, right = left) => ({
    top,
    left,
    bottom,
    right,
  });

  function filling() {
    const { workbook, sheet, set, show } = setup();
    const formula = (row: number, col: number) =>
      workbook.cell(sheet, row, col).formula;
    const fill = (
      source: ReturnType<typeof range>,
      target: ReturnType<typeof range>,
    ) => workbook.execute({ type: 'fill', sheet, source, target });
    return { workbook, sheet, set, show, formula, fill };
  }

  it('repeats cells that make no series, with their formats, in each direction', () => {
    const { set, show, fill } = filling();
    set(2, 0, 'a');
    set(3, 0, '12%');
    const change = fill(range(2, 0, 3, 0), range(2, 0, 6, 0));
    expect(change.ranges).toEqual([range(4, 0, 6, 0)]);
    expect([4, 5, 6].map((row) => show(row, 0))).toEqual(['a', '12%', 'a']);
    // Upward, the pattern keeps its place: the cell above the first is the last.
    fill(range(2, 0, 3, 0), range(0, 0, 3, 0));
    expect([0, 1].map((row) => show(row, 0))).toEqual(['a', '12%']);
    set(0, 5, 'x');
    set(0, 6, '');
    fill(range(0, 5, 0, 6), range(0, 5, 0, 9));
    expect([7, 8, 9].map((col) => show(0, col))).toEqual(['x', '', 'x']);
    fill(range(0, 5, 0, 6), range(0, 3, 0, 6));
    expect([3, 4].map((col) => show(0, col))).toEqual(['x', '']);
    // One number is copied, not counted on.
    set(10, 0, '7');
    fill(range(10, 0), range(10, 0, 12, 0));
    expect(show(12, 0)).toBe('7');
  });

  it('moves the relative references of formulas by the distance filled', () => {
    const { set, formula, fill } = filling();
    set(5, 5, '=A1+$A$1+A$1+$A1');
    fill(range(5, 5), range(5, 5, 7, 5));
    expect(formula(7, 5)).toBe('=A3+$A$1+A$1+$A3');
    fill(range(5, 5), range(5, 3, 5, 5));
    expect(formula(5, 3)).toBe('=#REF!+$A$1+#REF!+$A1');
    fill(range(5, 5), range(4, 5, 5, 5));
    expect(formula(4, 5)).toBe('=#REF!+$A$1+A$1+#REF!');
  });

  it('extends numbers by their step, both ways', () => {
    const { show, set, fill } = filling();
    set(1, 0, '1');
    set(2, 0, '3');
    fill(range(1, 0, 2, 0), range(1, 0, 4, 0));
    fill(range(1, 0, 2, 0), range(0, 0, 2, 0));
    expect([0, 3, 4].map((row) => show(row, 0))).toEqual(['-1', '5', '7']);
    set(0, 2, '0.1');
    set(0, 3, '0.2');
    set(0, 4, '0.3');
    fill(range(0, 2, 0, 4), range(0, 2, 0, 6));
    expect(show(0, 6)).toBe('0.5');
  });

  it('extends dates by days, or by months when they fall on the same day', () => {
    const { show, set, fill } = filling();
    set(0, 0, '1/30/2024');
    fill(range(0, 0), range(0, 0, 2, 0));
    expect(show(2, 0)).toBe('2/1/2024');
    set(0, 1, '1/31/2024');
    set(1, 1, '3/31/2024');
    fill(range(0, 1, 1, 1), range(0, 1, 4, 1));
    // The day is kept where the month has it, and the month's last otherwise.
    expect([2, 3, 4].map((row) => show(row, 1))).toEqual([
      '5/31/2024',
      '7/31/2024',
      '9/30/2024',
    ]);
    set(0, 2, '1/1/2024');
    set(1, 2, '1/2/2024');
    fill(range(0, 2, 1, 2), range(0, 2, 2, 2));
    expect(show(2, 2)).toBe('1/3/2024');
    set(0, 3, '1/31/2024');
    set(1, 3, '2/29/2024');
    fill(range(0, 3, 1, 3), range(0, 3, 2, 3));
    expect(show(2, 3)).toBe('3/29/2024');
  });

  it('extends weekday and month names, keeping their case', () => {
    const { show, set, fill } = filling();
    set(0, 0, 'Sat');
    fill(range(0, 0), range(0, 0, 2, 0));
    expect([1, 2].map((row) => show(row, 0))).toEqual(['Sun', 'Mon']);
    set(0, 1, 'JANUARY');
    set(1, 1, 'MARCH');
    fill(range(0, 1, 1, 1), range(0, 1, 2, 1));
    expect(show(2, 1)).toBe('MAY');
    set(0, 2, 'may');
    fill(range(0, 2), range(0, 2, 0, 3));
    expect(show(0, 3)).toBe('jun');
    set(0, 4, 'Monday');
    fill(range(0, 4), range(0, 4, 0, 5));
    expect(show(0, 5)).toBe('Tuesday');
    // A name with something else is copied.
    set(5, 0, 'Mon');
    set(6, 0, '3');
    fill(range(5, 0, 6, 0), range(5, 0, 7, 0));
    expect(show(7, 0)).toBe('Mon');
  });

  it('counts on the number at the end of a text, keeping its digits, in each direction', () => {
    const { show, set, fill } = filling();
    set(0, 0, 'Item 1');
    fill(range(0, 0), range(0, 0, 2, 0));
    expect([1, 2].map((row) => show(row, 0))).toEqual(['Item 2', 'Item 3']);
    set(0, 1, 'Item 1');
    set(1, 1, 'Item 3');
    fill(range(0, 1, 1, 1), range(0, 1, 3, 1));
    expect([2, 3].map((row) => show(row, 1))).toEqual(['Item 5', 'Item 7']);
    set(5, 0, 'A001');
    fill(range(5, 0), range(5, 0, 5, 1));
    expect(show(5, 1)).toBe('A002');
    set(6, 0, 'Item 9');
    fill(range(6, 0), range(6, 0, 6, 1));
    expect(show(6, 1)).toBe('Item 10');
    set(10, 3, 'Item 3');
    fill(range(10, 3), range(8, 3, 10, 3));
    expect([8, 9].map((row) => show(row, 3))).toEqual(['Item 1', 'Item 2']);
    fill(range(10, 3), range(10, 1, 10, 3));
    expect(show(10, 1)).toBe('Item 1');
    // Different texts before the numbers, or uneven or no steps, are copied.
    set(16, 0, 'Item 1');
    set(17, 0, 'Item 1');
    fill(range(16, 0, 17, 0), range(16, 0, 18, 0));
    expect(show(18, 0)).toBe('Item 1');
    set(12, 0, 'Item 1');
    set(13, 0, 'Part 2');
    set(12, 1, 'Item 1');
    set(13, 1, 'Item 2');
    set(14, 1, 'Item 4');
    fill(range(12, 0, 13, 0), range(12, 0, 14, 0));
    fill(range(12, 1, 14, 1), range(12, 1, 15, 1));
    expect(show(14, 0)).toBe('Item 1');
    expect(show(15, 1)).toBe('Item 1');
  });

  it('copies, extends, or fills only the formats or only the contents, as asked', () => {
    const { workbook, sheet, show, set } = filling();
    const fillAs = (
      how: 'copy' | 'series' | 'formats' | 'values',
      source: ReturnType<typeof range>,
      target: ReturnType<typeof range>,
    ) => workbook.execute({ type: 'fill', sheet, source, target, how });
    set(0, 0, '1');
    set(1, 0, '2');
    fillAs('copy', range(0, 0, 1, 0), range(0, 0, 3, 0));
    expect([2, 3].map((row) => show(row, 0))).toEqual(['1', '2']);
    // One number counts on by one as a series.
    set(0, 1, '5');
    fillAs('series', range(0, 1), range(0, 1, 2, 1));
    expect(show(2, 1)).toBe('7');
    // A text that makes no series is still copied.
    set(0, 2, 'x');
    fillAs('series', range(0, 2), range(0, 2, 1, 2));
    expect(show(1, 2)).toBe('x');

    set(5, 0, '10%');
    set(6, 0, '0.5');
    set(7, 0, 'kept');
    fillAs('formats', range(5, 0), range(5, 0, 7, 0));
    expect([6, 7].map((row) => show(row, 0))).toEqual(['50%', 'kept']);
    set(9, 0, '1');
    set(9, 1, '=A10*2');
    set(10, 0, '25%');
    set(10, 1, '25%');
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [range(9, 0, 9, 1)],
      style: { bold: true },
    });
    fillAs('values', range(9, 0, 9, 1), range(9, 0, 10, 1));
    expect([show(10, 0), show(10, 1)]).toEqual(['100%', '200%']);
    expect(workbook.cell(sheet, 10, 1).formula).toBe('=A11*2');
    expect(workbook.cell(sheet, 10, 0).style.bold).toBeUndefined();
  });

  it('tells whether a fill extends a series', () => {
    const { workbook, sheet, set } = filling();
    const extends_ = (
      source: ReturnType<typeof range>,
      target: ReturnType<typeof range>,
      how?: 'series',
    ) =>
      workbook.fillsSeries({
        type: 'fill',
        sheet,
        source,
        target,
        ...(how && { how }),
      });
    set(0, 0, '1');
    set(0, 1, 'Item 1');
    set(0, 2, 'x');
    expect(extends_(range(0, 0), range(0, 0, 3, 0))).toBe(false);
    expect(extends_(range(0, 0), range(0, 0, 3, 0), 'series')).toBe(true);
    expect(extends_(range(0, 1), range(0, 1, 3, 1))).toBe(true);
    expect(extends_(range(0, 2), range(0, 2, 3, 2), 'series')).toBe(false);
    expect(extends_(range(0, 0, 0, 2), range(0, 0, 3, 2))).toBe(true);
  });

  it('is undone and redone as one change, and tells listeners a command that replays it', () => {
    const { workbook, set, show, fill } = filling();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1');
    set(1, 0, '2');
    set(3, 0, 'under');
    fill(range(0, 0, 1, 0), range(0, 0, 4, 0));
    expect(show(3, 0)).toBe('4');
    expect(workbook.undo()!.ranges).toEqual([range(2, 0, 4, 0)]);
    expect(show(3, 0)).toBe('under');
    expect(show(4, 0)).toBe('');
    workbook.redo();
    expect(show(4, 0)).toBe('5');
    expect(Workbook.fromSnapshot(replay.toSnapshot()).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
  });
});

describe('rows and columns', () => {
  function lines() {
    const { workbook, sheet, set, show } = setup();
    const formula = (row: number, col: number) =>
      workbook.cell(sheet, row, col).formula;
    const insert = (
      direction: 'rows' | 'columns',
      ...spans: [number, number][]
    ) =>
      workbook.execute({
        type: 'insertLines',
        sheet,
        direction,
        spans: spans.map(([start, end]) => ({ start, end })),
      });
    const remove = (
      direction: 'rows' | 'columns',
      ...spans: [number, number][]
    ) =>
      workbook.execute({
        type: 'deleteLines',
        sheet,
        direction,
        spans: spans.map(([start, end]) => ({ start, end })),
      });
    return { workbook, sheet, set, show, formula, insert, remove };
  }

  it('inserts rows, moving the cells below and every reference to them', () => {
    const { workbook, set, show, formula, insert } = lines();
    set(0, 0, '1');
    set(2, 0, '3');
    set(3, 0, '=A3*2');
    set(4, 0, '=SUM(A1:A4)+$A$3');
    set(0, 1, '=A3+A1');
    const other = workbook.execute({ type: 'addSheet', index: 1 }).sheet;
    workbook.execute({
      type: 'setCell',
      sheet: other,
      row: 0,
      col: 0,
      input: '=Sheet1!A3',
    });
    insert('rows', [1, 2]);
    expect(show(4, 0)).toBe('3');
    expect(show(1, 0)).toBe('');
    expect(formula(5, 0)).toBe('=A5*2');
    expect(formula(6, 0)).toBe('=SUM(A1:A6)+$A$5');
    expect(show(6, 0)).toBe('13');
    expect(formula(0, 1)).toBe('=A5+A1');
    expect(workbook.cell(other, 0, 0).formula).toBe('=Sheet1!A5');
    // The moved cells are read where they are now.
    set(4, 0, '10');
    expect(show(5, 0)).toBe('20');
    expect(show(0, 1)).toBe('11');
    expect(workbook.cell(other, 0, 0).display).toBe('10');
  });

  it('deletes rows, giving #REF! for references to deleted cells and shrinking ranges', () => {
    const { set, show, formula, remove } = lines();
    for (let row = 0; row < 5; row++) set(row, 0, String(row + 1));
    set(0, 1, '=A3');
    set(1, 1, '=SUM(A2:A5)');
    set(2, 1, 'gone');
    set(5, 1, '=A5+A6');
    set(6, 1, '=SUM(A2:A3)');
    remove('rows', [1, 2]);
    expect(show(1, 0)).toBe('4');
    expect(formula(0, 1)).toBe('=#REF!');
    expect(show(0, 1)).toBe('#REF!');
    // The formula in a deleted row is gone with it.
    expect(show(1, 1)).toBe('');
    expect(formula(3, 1)).toBe('=A3+A4');
    expect(formula(4, 1)).toBe('=SUM(#REF!)');
    expect(show(3, 1)).toBe('5');
  });

  it('inserts and deletes columns, with the references to them', () => {
    const { set, show, formula, insert, remove } = lines();
    set(0, 0, '1');
    set(0, 1, '2');
    set(0, 2, '=A1+B1');
    set(1, 0, '=SUM(B:C)');
    insert('columns', [1, 1]);
    expect(show(0, 2)).toBe('2');
    expect(formula(0, 3)).toBe('=A1+C1');
    expect(formula(1, 0)).toBe('=SUM(C:D)');
    remove('columns', [0, 0]);
    expect(formula(0, 2)).toBe('=#REF!+B1');
    expect(formula(1, 0)).toBeNull();
  });

  it('inserts and deletes at each of several spans at once', () => {
    const { set, show, insert, remove } = lines();
    for (let row = 0; row < 6; row++) set(row, 0, String(row + 1));
    remove('rows', [3, 3], [1, 1]);
    expect([0, 1, 2, 3].map((row) => show(row, 0))).toEqual([
      '1',
      '3',
      '5',
      '6',
    ]);
    insert('rows', [1, 1], [2, 3]);
    expect([0, 1, 2, 3, 4, 5, 6].map((row) => show(row, 0))).toEqual([
      '1',
      '',
      '3',
      '',
      '',
      '5',
      '6',
    ]);
  });

  it('sets the sizes of rows and columns and hides them, keeping them in the snapshot', () => {
    const { workbook, sheet } = lines();
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [
        { start: 1, end: 2, size: 150 },
        { start: 4, end: 4, size: 30 },
      ],
    });
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 3, end: 5, hidden: true }],
    });
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 2, end: 4, size: null }],
    });
    const { rows, columns } = workbook.layout(sheet);
    expect([...columns.sizes]).toEqual([[1, 150]]);
    expect([...rows.hidden].sort()).toEqual([3, 4, 5]);
    expect(rows.sizes.size).toBe(0);
    const reopened = Workbook.fromSnapshot(workbook.toSnapshot());
    expect(reopened.layout(sheet)).toEqual(workbook.layout(sheet));
    expect(workbook.layout(sheet)).toBe(workbook.layout(sheet));
  });

  it('freezes and unfreezes rows and columns, with undo, redo and the snapshot', () => {
    const { workbook, sheet, insert, remove } = lines();
    expect(workbook.layout(sheet).frozen).toBeUndefined();
    workbook.execute({ type: 'freeze', sheet, rows: 2, cols: 1 });
    expect(workbook.layout(sheet).frozen).toEqual({ rows: 2, cols: 1 });
    const reopened = Workbook.fromSnapshot(workbook.toSnapshot());
    expect(reopened.layout(sheet).frozen).toEqual({ rows: 2, cols: 1 });

    // Inserting and deleting lines keeps what is frozen.
    insert('rows', [0, 0]);
    remove('columns', [0, 0]);
    expect(workbook.layout(sheet).frozen).toEqual({ rows: 2, cols: 1 });
    workbook.undo();
    workbook.undo();

    workbook.execute({ type: 'freeze', sheet, rows: 0, cols: 0 });
    expect(workbook.layout(sheet).frozen).toBeUndefined();
    expect(workbook.toSnapshot().sheets[0]!.frozen).toBeUndefined();
    expect(workbook.undo()).toEqual({ sheet, ranges: [] });
    expect(workbook.layout(sheet).frozen).toEqual({ rows: 2, cols: 1 });
    workbook.undo();
    expect(workbook.layout(sheet).frozen).toBeUndefined();
    workbook.redo();
    expect(workbook.layout(sheet).frozen).toEqual({ rows: 2, cols: 1 });

    expect(() =>
      workbook.execute({ type: 'freeze', sheet, rows: SHEET_ROWS, cols: 0 }),
    ).toThrow();
    expect(() =>
      workbook.execute({ type: 'freeze', sheet, rows: 0, cols: -1 }),
    ).toThrow();
  });

  it('moves the sizes and hidden lines with the rows and columns', () => {
    const { workbook, sheet, insert, remove } = lines();
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'rows',
      spans: [
        { start: 2, end: 2, size: 40 },
        { start: 5, end: 5, hidden: true },
      ],
    });
    insert('rows', [0, 1]);
    expect([...workbook.layout(sheet).rows.sizes]).toEqual([[4, 40]]);
    expect([...workbook.layout(sheet).rows.hidden]).toEqual([7]);
    remove('rows', [4, 4]);
    expect(workbook.layout(sheet).rows.sizes.size).toBe(0);
    expect([...workbook.layout(sheet).rows.hidden]).toEqual([6]);
  });

  it('undoes and redoes each change to rows and columns, selecting the lines it changed', () => {
    const { workbook, sheet, set, show, formula, insert, remove } = lines();
    set(0, 0, '1');
    set(1, 0, '2');
    set(2, 0, '=A2*10');
    remove('rows', [1, 1]);
    expect(formula(1, 0)).toBe('=#REF!*10');
    expect(workbook.undo()!.ranges).toEqual([
      { top: 1, left: 0, bottom: 1, right: SHEET_COLUMNS - 1 },
    ]);
    expect(formula(2, 0)).toBe('=A2*10');
    expect(show(2, 0)).toBe('20');
    workbook.redo();
    expect(formula(1, 0)).toBe('=#REF!*10');
    workbook.undo();

    insert('columns', [0, 1]);
    expect(show(2, 2)).toBe('20');
    expect(workbook.undo()!.ranges).toEqual([
      { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: 1 },
    ]);
    expect(formula(2, 0)).toBe('=A2*10');
    set(1, 0, '3');
    expect(show(2, 0)).toBe('30');

    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 0, end: 1, size: 50, hidden: true }],
    });
    workbook.undo();
    expect(workbook.layout(sheet).rows.sizes.size).toBe(0);
    expect(workbook.layout(sheet).rows.hidden.size).toBe(0);
    workbook.redo();
    expect([...workbook.layout(sheet).rows.sizes]).toEqual([
      [0, 50],
      [1, 50],
    ]);
  });

  it('tells listeners of changes to rows and columns as commands that replay them', () => {
    const { workbook, sheet, set, insert, remove } = lines();
    const replay = Workbook.fromSnapshot(workbook.toSnapshot());
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1');
    set(3, 0, '=A1+A3');
    insert('rows', [0, 0]);
    remove('columns', [2, 3]);
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 0, end: 0, size: 120 }],
    });
    remove('rows', [2, 2]);
    workbook.undo();
    workbook.undo();
    workbook.redo();
    expect(replay.toSnapshot()).toEqual(workbook.toSnapshot());
  });

  it('refuses spans outside the sheet or that end before they start', () => {
    const { workbook, sheet } = lines();
    for (const spans of [
      [{ start: -1, end: 0 }],
      [{ start: 3, end: 2 }],
      [{ start: 0, end: SHEET_COLUMNS }],
    ]) {
      expect(() =>
        workbook.execute({
          type: 'deleteLines',
          sheet,
          direction: 'columns',
          spans,
        }),
      ).toThrow();
      expect(() =>
        workbook.execute({
          type: 'setLines',
          sheet,
          direction: 'columns',
          spans,
        }),
      ).toThrow();
    }
    expect(workbook.canUndo).toBe(false);
  });
});

describe('cell formats', () => {
  /** Ten pixels a character, whatever the font: wraps at 8 in a column. */
  const measure = (text: string) => text.length * 10;

  function formats() {
    const workbook = Workbook.fromSnapshot(
      createWorkbook().toSnapshot(),
      [],
      measure,
    );
    const sheet = workbook.sheets[0]!.id;
    const set = (row: number, col: number, input: string) =>
      workbook.execute({ type: 'setCell', sheet, row, col, input });
    const format = (
      range: { top: number; left: number; bottom: number; right: number },
      style: StyleChange,
    ) => workbook.execute({ type: 'setStyle', sheet, ranges: [range], style });
    const style = (row: number, col: number) =>
      workbook.cell(sheet, row, col).style;
    const fitted = (row: number) =>
      workbook.layout(sheet).rows.fitted?.get(row);
    return { workbook, sheet, set, format, style, fitted };
  }
  const A1B2 = { top: 0, left: 0, bottom: 1, right: 1 };

  it('sets the font, colors and alignment of ranges, keeping values and number formats', () => {
    const { workbook, sheet, set, format, style } = formats();
    set(0, 0, '50%');
    format(A1B2, { bold: true, font: 'Georgia', size: 14, color: '#c00000' });
    format(A1B2, {
      fill: '#ffff00',
      horizontal: 'center',
      vertical: 'top',
      wrap: true,
    });
    expect(style(0, 0)).toEqual({
      numberFormat: '0%',
      bold: true,
      font: 'Georgia',
      size: 14,
      color: '#c00000',
      fill: '#ffff00',
      horizontal: 'center',
      vertical: 'top',
      wrap: true,
    });
    expect(workbook.cell(sheet, 0, 0).display).toBe('50%');
    expect(style(1, 1)).toEqual({ ...style(0, 0), numberFormat: undefined });
    expect(style(2, 0)).toEqual({});

    // `null` takes a property back to the default.
    format(A1B2, { bold: null, fill: null, wrap: null });
    expect(style(0, 0).bold).toBeUndefined();
    expect(style(0, 0).fill).toBeUndefined();
    expect(style(0, 0).italic).toBeUndefined();
    expect(style(0, 0).size).toBe(14);
  });

  it('shows values in number formats, with the color a format gives them', () => {
    const { workbook, sheet, set, format } = formats();
    set(0, 0, '-1234.5');
    format(A1B2, { numberFormat: '#,##0.00;[Red](#,##0.00)' });
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      display: '(1,234.50)',
      color: '#ff0000',
      input: '-1234.5',
    });
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({ display: '-1234.5' });
    expect(workbook.cell(sheet, 0, 0).color).toBeUndefined();
  });

  it('keeps what is typed into a cell with the text format as text', () => {
    const { workbook, sheet, set, format, style } = formats();
    format(A1B2, { numberFormat: '@' });
    set(0, 0, '007');
    set(0, 1, '12%');
    set(1, 0, '=1+1');
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: '007',
      display: '007',
      input: '007',
    });
    expect(workbook.cell(sheet, 0, 1).value).toBe('12%');
    expect(style(0, 1).numberFormat).toBe('@');
    expect(workbook.cell(sheet, 1, 0)).toMatchObject({
      value: '=1+1',
      formula: null,
    });
    // A formula already there is replaced, and pasted text is kept too.
    set(1, 1, '=2+2');
    format(A1B2, { numberFormat: '@' });
    set(1, 1, '=A1');
    expect(workbook.cell(sheet, 1, 1)).toMatchObject({
      value: '=A1',
      formula: null,
    });
    workbook.execute({
      type: 'setCells',
      sheet,
      row: 0,
      col: 0,
      inputs: [['0012', '1/2/2026']],
    });
    expect(workbook.cell(sheet, 0, 0).value).toBe('0012');
    expect(workbook.cell(sheet, 0, 1).value).toBe('1/2/2026');
    // Elsewhere the same entry is a number.
    set(2, 0, '007');
    expect(workbook.cell(sheet, 2, 0).value).toBe(7);
  });

  it('keeps one style for the cells that look the same, however it was set', () => {
    const { workbook, format } = formats();
    format({ top: 0, left: 0, bottom: 0, right: 0 }, { bold: true });
    format({ top: 0, left: 0, bottom: 0, right: 0 }, { italic: true });
    format(
      { top: 1, left: 0, bottom: 1, right: 0 },
      { italic: true, bold: true },
    );
    format({ top: 2, left: 0, bottom: 2, right: 0 }, { underline: true });
    format({ top: 2, left: 0, bottom: 2, right: 0 }, { underline: null });
    const { styles, sheets } = workbook.toSnapshot();
    expect(sheets[0]!.styleIds![0]!.slice(0, 2)).toEqual([2, 2]);
    expect(styles![sheets[0]!.styleIds![0]![2]!]).toEqual({});
  });

  it('undoes and redoes a change of format, which listeners can replay', () => {
    const { workbook, sheet, set, format, style } = formats();
    const replay = Workbook.fromSnapshot(workbook.toSnapshot());
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '=1+1');
    format(A1B2, { strike: true, underline: true });
    expect(workbook.undo()).toEqual({ sheet, ranges: [A1B2] });
    expect(style(0, 0)).toEqual({});
    expect(workbook.cell(sheet, 0, 0).value).toBe(2);
    workbook.redo();
    expect(style(1, 0)).toEqual({ strike: true, underline: true });
    expect(replay.toSnapshot()).toEqual(workbook.toSnapshot());
  });

  it('formats the cells in use of whole columns and rows, at once', () => {
    const { workbook, sheet, set, format, style } = formats();
    set(2, 0, 'x');
    set(0, 3, 'y');
    const start = performance.now();
    format(
      { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: 0 },
      { bold: true },
    );
    format(
      { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: SHEET_COLUMNS - 1 },
      { italic: true },
    );
    // Not the 10 million cells of the sheet, which would take seconds.
    expect(performance.now() - start).toBeLessThan(1000);
    expect(style(2, 0)).toEqual({ bold: true, italic: true });
    expect(style(0, 3)).toEqual({ italic: true });
    expect(workbook.toSnapshot().sheets[0]!.styleIds![0]!.length).toBe(3);
    expect(workbook.cell(sheet, 3, 0).style).toEqual({});
  });

  it('formats nothing of whole lines where no cell is in use', () => {
    const { workbook, format, style } = formats();
    format(
      { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: SHEET_COLUMNS - 1 },
      { bold: true },
    );
    expect(style(0, 0)).toEqual({});
    expect(workbook.toSnapshot().sheets[0]!.styleIds).toEqual([]);
  });

  it('grows a row to fit wrapped text and large fonts, unless the row has a size of its own', () => {
    const { workbook, sheet, set, format, fitted } = formats();
    set(0, 0, 'aaaa bbbb cccc');
    expect(fitted(0)).toBeUndefined();
    // Three lines of 16 pixels and the padding.
    format({ top: 0, left: 0, bottom: 0, right: 0 }, { wrap: true });
    expect(fitted(0)).toBe(56);
    set(0, 0, 'aaaa bbbb');
    expect(fitted(0)).toBe(40);
    // A wider column takes the text on one line.
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 0, end: 0, size: 200 }],
    });
    expect(fitted(0)).toBeUndefined();

    // A line of 20 points is 32 pixels high.
    format({ top: 1, left: 1, bottom: 1, right: 1 }, { size: 20 });
    expect(fitted(1)).toBe(40);
    workbook.undo();
    expect(fitted(1)).toBeUndefined();
    workbook.redo();

    // A size set on the row is kept apart, and overrides the fit.
    workbook.execute({
      type: 'setLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 1, end: 1, size: 30 }],
    });
    expect(workbook.layout(sheet).rows.sizes.get(1)).toBe(30);
    expect(fitted(1)).toBe(40);

    // Opened again, or with rows inserted, the rows are fitted again.
    const reopened = Workbook.fromSnapshot(workbook.toSnapshot(), [], measure);
    expect(reopened.layout(sheet).rows.fitted?.get(1)).toBe(40);
    workbook.execute({
      type: 'insertLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 0, end: 0 }],
    });
    expect(fitted(2)).toBe(40);
    expect(fitted(1)).toBeUndefined();
  });
});

describe('borders', () => {
  const thin = { line: 'thin', color: '#ff0000' } as const;
  const A1B2 = { top: 0, left: 0, bottom: 1, right: 1 };

  function borders() {
    const { workbook, sheet, set } = setup();
    const style = (row: number, col: number) =>
      workbook.cell(sheet, row, col).style;
    return { workbook, sheet, set, style };
  }

  it('sets borders on every side of every cell, keeping the rest of their style', () => {
    const { workbook, sheet, style } = borders();
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [A1B2],
      style: { bold: true },
    });
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [A1B2],
      sides: 'all',
      border: thin,
    });
    for (const [row, col] of [
      [0, 0],
      [1, 1],
    ] as const) {
      expect(style(row, col)).toEqual({
        bold: true,
        borderTop: thin,
        borderRight: thin,
        borderBottom: thin,
        borderLeft: thin,
      });
    }
  });

  it('sets the outside borders and those of one side on the edges of each range', () => {
    const { workbook, sheet, style } = borders();
    const range = { top: 0, left: 0, bottom: 2, right: 2 };
    const thick = { line: 'thick' } as const;
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [range],
      sides: 'outside',
      border: thick,
    });
    expect(style(0, 0)).toEqual({ borderTop: thick, borderLeft: thick });
    expect(style(0, 1)).toEqual({ borderTop: thick });
    expect(style(2, 2)).toEqual({ borderRight: thick, borderBottom: thick });
    expect(style(1, 1)).toEqual({});

    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [range],
      sides: 'bottom',
      border: thin,
    });
    expect(style(2, 1)).toEqual({ borderBottom: thin });
    expect(style(2, 2)).toEqual({ borderRight: thick, borderBottom: thin });
    expect(style(1, 1)).toEqual({});
  });

  it('puts the top border of a whole column on its first cell', () => {
    const { workbook, sheet, set, style } = borders();
    set(4, 0, 'x');
    const column = { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: 0 };
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [column],
      sides: 'top',
      border: thin,
    });
    expect(style(0, 0)).toEqual({ borderTop: thin });
    expect(style(4, 0)).toEqual({});
    workbook.undo();
    expect(style(0, 0)).toEqual({});
  });

  it('removes every border with none, and undoes it', () => {
    const { workbook, sheet, style } = borders();
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [A1B2],
      sides: 'all',
      border: thin,
    });
    workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: [A1B2],
      sides: 'all',
      border: null,
    });
    expect(style(1, 0)).toEqual({});
    expect(workbook.undo()).toEqual({ sheet, ranges: [A1B2] });
    expect(style(1, 0).borderLeft).toEqual(thin);
    workbook.redo();
    expect(style(1, 0)).toEqual({});
  });
});

describe('merges', () => {
  function merges() {
    const { workbook, sheet, set, show } = setup();
    const merge = (top: number, left: number, bottom: number, right: number) =>
      workbook.execute({
        type: 'merge',
        sheet,
        ranges: [{ top, left, bottom, right }],
      });
    return { workbook, sheet, set, show, merge };
  }

  it('merges a range into its top-left cell, keeping only its value', () => {
    const { workbook, sheet, set, show, merge } = merges();
    set(0, 0, 'keep');
    set(1, 1, 'lose');
    set(0, 2, '=B2');
    set(5, 5, '1');
    merge(0, 0, 1, 1);
    expect(workbook.merges(sheet)).toEqual([
      { top: 0, left: 0, bottom: 1, right: 1 },
    ]);
    expect(show(0, 0)).toBe('keep');
    expect(show(1, 1)).toBe('');
    expect(show(0, 2)).toBe('0');
    expect(show(5, 5)).toBe('1');
  });

  it('keeps a formula in the top-left cell, and merges ranges that overlap as one', () => {
    const { workbook, sheet, set, show } = merges();
    set(5, 5, '7');
    set(0, 0, '=F6*2');
    workbook.execute({
      type: 'merge',
      sheet,
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 1, left: 1, bottom: 2, right: 2 },
      ],
    });
    expect(workbook.merges(sheet)).toEqual([
      { top: 0, left: 0, bottom: 2, right: 2 },
    ]);
    expect(show(0, 0)).toBe('14');
  });

  it('does not grow rows to fit the wrapped text of a merged cell', () => {
    const { workbook, sheet, set, merge } = merges();
    set(0, 0, 'a long text that wraps onto many lines in one column');
    merge(0, 0, 1, 3);
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
      style: { wrap: true },
    });
    expect(workbook.layout(sheet).rows.fitted?.get(0)).toBeUndefined();
  });

  it('pastes the merges of the cells copied with their formats', () => {
    const { workbook, sheet, merge } = merges();
    merge(0, 0, 0, 1);
    workbook.execute({
      type: 'paste',
      sheet,
      target: { top: 3, left: 0, bottom: 4, right: 1 },
      source: { sheet, range: { top: 0, left: 0, bottom: 0, right: 1 } },
      only: 'formats',
    });
    expect(workbook.merges(sheet)).toEqual([
      { top: 0, left: 0, bottom: 0, right: 1 },
      { top: 3, left: 0, bottom: 3, right: 1 },
      { top: 4, left: 0, bottom: 4, right: 1 },
    ]);
    workbook.undo();
    expect(workbook.merges(sheet)).toHaveLength(1);
  });

  it('merges over the merges a range overlaps, and unmerges those it touches', () => {
    const { workbook, sheet, merge } = merges();
    merge(0, 0, 1, 1);
    merge(1, 1, 3, 1);
    expect(workbook.merges(sheet)).toEqual([
      { top: 0, left: 0, bottom: 3, right: 1 },
    ]);
    merge(5, 0, 5, 3);
    workbook.execute({
      type: 'unmerge',
      sheet,
      ranges: [{ top: 2, left: 0, bottom: 5, right: 0 }],
    });
    expect(workbook.merges(sheet)).toEqual([]);
  });

  it('undoes and redoes a merge and an unmerge, which listeners can replay', () => {
    const { workbook, sheet, set, show, merge } = merges();
    const replay = Workbook.fromSnapshot(workbook.toSnapshot());
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, 'a');
    set(0, 1, 'b');
    merge(0, 0, 0, 1);
    const range = { top: 0, left: 0, bottom: 0, right: 1 };
    expect(workbook.undo()).toEqual({ sheet, ranges: [range] });
    expect(workbook.merges(sheet)).toEqual([]);
    expect(show(0, 1)).toBe('b');
    workbook.redo();
    expect(workbook.merges(sheet)).toEqual([range]);
    expect(show(0, 1)).toBe('');

    workbook.execute({ type: 'unmerge', sheet, ranges: [range] });
    expect(workbook.merges(sheet)).toEqual([]);
    workbook.undo();
    expect(workbook.merges(sheet)).toEqual([range]);
    workbook.redo();
    expect(replay.toSnapshot()).toEqual(workbook.toSnapshot());
    expect(Workbook.fromSnapshot(workbook.toSnapshot()).merges(sheet)).toEqual(
      [],
    );
  });

  it('keeps merges in snapshots and copies of sheets', () => {
    const { workbook, sheet, merge } = merges();
    merge(2, 2, 4, 3);
    const range = { top: 2, left: 2, bottom: 4, right: 3 };
    expect(Workbook.fromSnapshot(workbook.toSnapshot()).merges(sheet)).toEqual([
      range,
    ]);
    const copy = workbook.execute({
      type: 'addSheet',
      index: 1,
      copyOf: sheet,
    });
    expect(workbook.merges(copy.sheet)).toEqual([range]);
  });

  it('moves, grows, shrinks and drops merges with rows and columns inserted and deleted', () => {
    const { workbook, sheet, merge } = merges();
    merge(2, 0, 4, 0);
    merge(0, 2, 0, 3);
    const lines = (
      type: 'insertLines' | 'deleteLines',
      direction: 'rows' | 'columns',
      start: number,
      end: number,
    ) =>
      workbook.execute({
        type,
        sheet,
        direction,
        spans: [{ start, end }],
      });
    lines('insertLines', 'rows', 0, 0);
    lines('insertLines', 'rows', 4, 5);
    expect(workbook.merges(sheet)).toContainEqual({
      top: 3,
      left: 0,
      bottom: 7,
      right: 0,
    });
    lines('deleteLines', 'rows', 2, 6);
    expect(workbook.merges(sheet)).toEqual([
      { top: 1, left: 2, bottom: 1, right: 3 },
    ]);
    lines('deleteLines', 'columns', 3, 3);
    expect(workbook.merges(sheet)).toEqual([]);
    workbook.undo();
    expect(workbook.merges(sheet)).toEqual([
      { top: 1, left: 2, bottom: 1, right: 3 },
    ]);
  });
});

describe('conditional formats', () => {
  const RED = { color: '#9c0006', fill: '#ffc7ce' } as const;
  const range = (top: number, left: number, bottom = top, right = left) => ({
    top,
    left,
    bottom,
    right,
  });
  /** A column A of values and a rule over A1:A6. */
  function rules(condition: Rule['condition'], values: string[]) {
    const { workbook, sheet, set } = setup();
    values.forEach((input, row) => set(row, 0, input));
    workbook.execute({
      type: 'setRules',
      sheet,
      rules: [{ ranges: [range(0, 0, 5, 0)], condition, format: RED }],
    });
    /** The rows of column A that the rule formats. */
    const formatted = () =>
      values.flatMap((_, row) =>
        workbook.cell(sheet, row, 0).conditional ? [row] : [],
      );
    return { workbook, sheet, set, formatted };
  }

  it('formats the cells greater than, less than, between or equal to a value', () => {
    const values = ['5', '10', '15', 'ten', '', '=A2*2'];
    expect(rules({ type: 'greater', value: '10' }, values).formatted()).toEqual(
      // Text sorts after numbers, as in Excel; empty cells are left alone.
      [2, 3, 5],
    );
    expect(rules({ type: 'less', value: '10' }, values).formatted()).toEqual([
      0,
    ]);
    expect(
      rules({ type: 'between', value: '20', other: '10' }, values).formatted(),
    ).toEqual([1, 2, 5]);
    expect(rules({ type: 'equal', value: 'TEN' }, values).formatted()).toEqual([
      3,
    ]);
    expect(rules({ type: 'equal', value: '15' }, values).formatted()).toEqual([
      2,
    ]);
  });

  it('formats the cells whose text contains a string, in any case', () => {
    const values = ['North', 'south', 'Northeast', '1200', 'west', ''];
    expect(
      rules({ type: 'contains', value: 'NORTH' }, values).formatted(),
    ).toEqual([0, 2]);
    expect(
      rules({ type: 'contains', value: '20' }, values).formatted(),
    ).toEqual([3]);
  });

  it('formats duplicate values and the top or bottom values, ties included', () => {
    const values = ['3', 'a', '7', 'A', '3', '9'];
    expect(rules({ type: 'duplicate' }, values).formatted()).toEqual([
      0, 1, 3, 4,
    ]);
    expect(rules({ type: 'top', count: 2 }, values).formatted()).toEqual([
      2, 5,
    ]);
    expect(rules({ type: 'bottom', count: 2 }, values).formatted()).toEqual([
      0, 4,
    ]);
  });

  it('gives the format of the first rule that matches, over the cell style, and leaves the style alone', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '50');
    workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: [range(0, 0)],
      style: { bold: true, fill: '#00b050' },
    });
    workbook.execute({
      type: 'setRules',
      sheet,
      rules: [
        {
          ranges: [range(0, 0)],
          condition: { type: 'greater', value: '10' },
          format: { color: '#c00000', italic: true },
        },
        {
          ranges: [range(0, 0, 9, 0)],
          condition: { type: 'greater', value: '0' },
          format: { color: '#0070c0', fill: '#ffff00' },
        },
      ],
    });
    const cell = workbook.cell(sheet, 0, 0);
    expect(cell.conditional).toEqual({
      color: '#c00000',
      italic: true,
      fill: '#ffff00',
    });
    expect(cell.style).toEqual({ bold: true, fill: '#00b050' });
  });

  it('follows changed values and formula results', () => {
    const { set, formatted } = rules({ type: 'top', count: 1 }, ['4', '=A1*2']);
    expect(formatted()).toEqual([1]);
    set(0, 0, '-4');
    expect(formatted()).toEqual([0]);
    set(0, 0, '1');
    expect(formatted()).toEqual([1]);
  });

  it('undoes and redoes a change of rules, keeps them in snapshots, and tells listeners', () => {
    const { workbook, sheet } = setup();
    const heard: Command[] = [];
    workbook.subscribe((command) => heard.push(command));
    const rule: Rule = {
      ranges: [range(0, 0, 2, 1)],
      condition: { type: 'duplicate' },
      format: { bold: true },
    };
    const change = workbook.execute({ type: 'setRules', sheet, rules: [rule] });
    expect(change).toEqual({ sheet, ranges: [range(0, 0, 2, 1)] });
    expect(workbook.rules(sheet)).toEqual([rule]);
    expect(heard).toEqual([{ type: 'setRules', sheet, rules: [rule] }]);
    expect(Workbook.fromSnapshot(workbook.toSnapshot()).rules(sheet)).toEqual([
      rule,
    ]);
    workbook.undo();
    expect(workbook.rules(sheet)).toEqual([]);
    workbook.redo();
    expect(workbook.rules(sheet)).toEqual([rule]);
  });

  it('works out duplicates on a copy of a sheet from its own values', () => {
    const { workbook, sheet, set, formatted } = rules({ type: 'duplicate' }, [
      '1',
      '1',
    ]);
    expect(formatted()).toEqual([0, 1]);
    workbook.execute({ type: 'addSheet', index: 1, copyOf: sheet });
    const copy = workbook.sheets[1]!.id;
    workbook.execute({
      type: 'setCell',
      sheet: copy,
      row: 1,
      col: 0,
      input: '2',
    });
    set(0, 0, '1');
    expect(formatted()).toEqual([0, 1]);
    expect(workbook.cell(copy, 0, 0).conditional).toBeUndefined();
    expect(workbook.cell(sheet, 0, 0).conditional).toBeDefined();
    // Read the copy first, then the original, with nothing run in between.
    workbook.execute({
      type: 'setCell',
      sheet: copy,
      row: 1,
      col: 0,
      input: '2',
    });
    expect(workbook.cell(copy, 0, 0).conditional).toBeUndefined();
    expect(workbook.cell(sheet, 0, 0).conditional).toBeDefined();
  });

  it('moves, grows and shrinks the ranges of rules with rows and columns, and drops those deleted', () => {
    const { workbook, sheet } = setup();
    const format = { bold: true } as const;
    const condition = { type: 'duplicate' } as const;
    workbook.execute({
      type: 'setRules',
      sheet,
      rules: [
        { ranges: [range(2, 0, 5, 0), range(9, 0)], condition, format },
        { ranges: [range(1, 1)], condition, format },
      ],
    });
    workbook.execute({
      type: 'insertLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 3, end: 4 }],
    });
    expect(workbook.rules(sheet).map((rule) => rule.ranges)).toEqual([
      [range(2, 0, 7, 0), range(11, 0)],
      [range(1, 1)],
    ]);
    workbook.execute({
      type: 'deleteLines',
      sheet,
      direction: 'rows',
      spans: [{ start: 0, end: 2 }],
    });
    expect(workbook.rules(sheet).map((rule) => rule.ranges)).toEqual([
      [range(0, 0, 4, 0), range(8, 0)],
    ]);
    workbook.execute({
      type: 'insertLines',
      sheet,
      direction: 'columns',
      spans: [{ start: 0, end: 0 }],
    });
    expect(workbook.rules(sheet).map((rule) => rule.ranges)).toEqual([
      [range(0, 1, 4, 1), range(8, 1)],
    ]);
  });
});
