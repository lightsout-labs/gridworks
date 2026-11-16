import {
  editText,
  formatColor,
  formatValue,
  GENERAL,
  isTextFormat,
  parseInput,
} from '../format';
import {
  columnName,
  parse,
  print,
  shift,
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Ast,
  type Move,
} from '../formula';
import { isError, type Value } from '../value';
import {
  Calculation,
  emptySheet,
  key,
  NO_LINES,
  place,
  type Constant,
  type Frozen,
  type Lines,
  type Sheet,
  type SheetLayout,
} from './calc';
import { Rules, type Rule, type RuleFormat } from './rules';
import { series } from './series';
import {
  cellHeight,
  changedStyle,
  estimatedWidth,
  mayGrowRow,
  STANDARD_COLUMN_WIDTH,
  STANDARD_ROW_HEIGHT,
  type Border,
  type BorderSides,
  type Style,
  type StyleChange,
  type TextWidth,
} from './style';

export type { Value } from '../value';
export type { Frozen, Lines, SheetLayout } from './calc';
export type { Condition, Rule, RuleFormat } from './rules';
export {
  BORDER_LINES,
  DEFAULT_FONT_SIZE,
  FONT_SIZES,
  FONTS,
  fontPixels,
  lineHeight,
  STANDARD_COLUMN_WIDTH,
  STANDARD_ROW_HEIGHT,
  TEXT_PADDING,
  wrapText,
  type Border,
  type BorderSides,
  type Style,
  type StyleChange,
  type TextWidth,
} from './style';
export {
  columnName,
  quoteSheetName,
  SHEET_COLUMNS,
  SHEET_ROWS,
} from '../formula';

/** An edit to the workbook. Rows and columns count from zero. */
export type Command =
  | {
      readonly type: 'setCell';
      readonly sheet: string;
      readonly row: number;
      readonly col: number;
      /**
       * What the user typed: a number, a date, a time, a percentage, a
       * boolean, text, or a formula starting with `=`.
       */
      readonly input: string;
      /**
       * The year of a date typed without one. When absent, the current
       * year, which the workbook adds to the command its listeners get.
       */
      readonly year?: number;
    }
  | {
      /**
       * Sets a block of cells from its top left as if each input were
       * typed, as text pasted from another application is. Rows may differ
       * in length; what falls outside the sheet is left out.
       */
      readonly type: 'setCells';
      readonly sheet: string;
      readonly row: number;
      readonly col: number;
      readonly inputs: readonly (readonly string[])[];
      /** As in `setCell`, for every date typed without a year. */
      readonly year?: number;
    }
  | {
      /**
       * Changes the style of the cells of ranges. Ranges over whole rows or
       * columns change only the cells in use there.
       */
      readonly type: 'setStyle';
      readonly sheet: string;
      readonly ranges: readonly Range[];
      readonly style: StyleChange;
    }
  | {
      /**
       * Puts a border on sides of the cells of ranges, or takes it off
       * when `null`, keeping the rest of their style. Ranges over whole
       * rows or columns change only the cells in use there.
       */
      readonly type: 'setBorders';
      readonly sheet: string;
      readonly ranges: readonly Range[];
      readonly sides: BorderSides;
      readonly border: Border | null;
    }
  | {
      /**
       * Merges each range into one cell, with the merges it overlaps: the
       * top-left cell keeps its value, and the others are emptied.
       */
      readonly type: 'merge';
      readonly sheet: string;
      readonly ranges: readonly Range[];
    }
  | {
      /** Unmerges the merged cells that ranges overlap. */
      readonly type: 'unmerge';
      readonly sheet: string;
      readonly ranges: readonly Range[];
    }
  | {
      /** Empties the cells of ranges, keeping their formats. */
      readonly type: 'clear';
      readonly sheet: string;
      readonly ranges: readonly Range[];
    }
  | Restore
  | {
      /**
       * Copies the cells of a range, or moves them when `cut`, to a sheet.
       * A target made of whole copies of the source is filled with them;
       * any other gets one copy at its top left, and a cut always one. A
       * copy moves the relative references of its formulas; a cut moves
       * every reference to its cells. `only` pastes the values (formula
       * results among them), the formulas and values, or the formats,
       * keeping the rest of the target. It gives the cells pasted first,
       * and listeners get it as a `group`.
       */
      readonly type: 'paste';
      readonly sheet: string;
      readonly target: Range;
      readonly source: { readonly sheet: string; readonly range: Range };
      readonly cut?: boolean;
      readonly only?: 'values' | 'formulas' | 'formats';
    }
  | {
      /**
       * Fills `target` from `source`, a range at one of its edges, as the
       * fill handle does: each row or column of the source that starts a
       * series is extended as one, and any other is repeated, its formulas'
       * relative references moved, with the formats of its cells. The
       * target spans the source and extends it on one side. `how` repeats
       * every one (`copy`); extends one number too, counting by one
       * (`series`); or fills only the formats (`formats`) or only the
       * contents (`values`), keeping the rest of the target. It gives the
       * cells filled, and listeners get it as a `group`.
       */
      readonly type: 'fill';
      readonly sheet: string;
      readonly source: Range;
      readonly target: Range;
      readonly how?: 'copy' | 'series' | 'formats' | 'values';
    }
  | MoveLines
  | ({
      /**
       * Sets the size of the rows or columns of each span, or the default
       * size when `null`, and hides or shows them; what a span leaves out
       * stays as it is.
       */
      readonly type: 'setLines';
    } & LinesOf<LineSpan>)
  | {
      /**
       * Freezes the rows above `rows` and the columns left of `cols`, so
       * that they stay in place while the others scroll; 0 and 0 unfreeze.
       */
      readonly type: 'freeze';
      readonly sheet: string;
      readonly rows: number;
      readonly cols: number;
    }
  | {
      /**
       * Gives a sheet these conditional formats in place of its own, the
       * first taking priority where they differ (R20).
       */
      readonly type: 'setRules';
      readonly sheet: string;
      readonly rules: readonly Rule[];
    }
  | {
      /**
       * Puts sheets back as they were, giving the cells of `ranges` of
       * `sheet` as changed: how the workbook undoes a change to rows or
       * columns.
       */
      readonly type: 'restoreSheets';
      readonly sheet: string;
      readonly ranges: readonly Range[];
      readonly sheets: readonly SheetSnapshot[];
    }
  | {
      /**
       * Restores of different cells made as one change: how a paste is
       * recorded, undone and redone. It gives the cells of the first.
       */
      readonly type: 'group';
      readonly restores: readonly Restore[];
    }
  | {
      /**
       * Adds an empty sheet at a place in the tab order, or a copy of the
       * sheet `copyOf`. When absent, the workbook gives the identifier and
       * the next free name, and adds them to the command its listeners get.
       */
      readonly type: 'addSheet';
      readonly index: number;
      readonly copyOf?: string;
      readonly sheet?: string;
      readonly name?: string;
    }
  | {
      /** Deletes a sheet; the last one cannot be deleted (R10). */
      readonly type: 'deleteSheet';
      readonly sheet: string;
    }
  | {
      /** Puts a deleted sheet back: how the workbook undoes a deletion. */
      readonly type: 'restoreSheet';
      readonly index: number;
      readonly contents: SheetSnapshot;
    }
  | {
      /**
       * Renames, recolors or moves a sheet to a place in the tab order. A
       * `null` color removes the color.
       */
      readonly type: 'setSheet';
      readonly sheet: string;
      readonly name?: string;
      readonly color?: string | null;
      readonly index?: number;
    };

/** Rows or columns. */
export type Direction = 'rows' | 'columns';

/** A run of rows or columns, its ends included. */
export interface Span {
  readonly start: number;
  readonly end: number;
}

/**
 * Inserts before each span as many empty rows or columns as it has, as
 * Excel inserts before the lines selected, or deletes the lines of each
 * span. Cells and the sizes of lines move with them, and so does every
 * reference to them; a reference to deleted cells becomes `#REF!`. What an
 * insertion pushes off the sheet is lost.
 */
export type MoveLines =
  | (LinesOf<Span> & { readonly type: 'insertLines' })
  | (LinesOf<Span> & { readonly type: 'deleteLines' });

/** Spans of the rows or columns of a sheet. */
interface LinesOf<T extends Span> {
  readonly sheet: string;
  readonly direction: Direction;
  readonly spans: readonly T[];
}

/** A run of rows or columns, with the size and visibility to give them. */
export interface LineSpan extends Span {
  /** In CSS pixels at 100%; `null` for the default size. */
  readonly size?: number | null;
  readonly hidden?: boolean;
}

/** Why a sheet name cannot be used, as Excel refuses it. */
export type SheetNameProblem = 'empty' | 'tooLong' | 'character' | 'taken';

/** Excel's limit on the length of a sheet name. */
const SHEET_NAME_LIMIT = 31;

/**
 * Makes ranges hold exactly the cells listed, with their contents and
 * styles, and nothing else: how the workbook undoes and redoes.
 */
export interface Restore {
  readonly type: 'restore';
  readonly sheet: string;
  readonly ranges: readonly Range[];
  readonly cells: readonly CellContents[];
  /**
   * When present, the merges that overlap the ranges are replaced by
   * these; otherwise they stay as they are.
   */
  readonly merges?: readonly Range[];
}

/** A cell's contents and style, as a `restore` command puts them back. */
export interface CellContents {
  readonly row: number;
  readonly col: number;
  readonly value: Constant;
  /** The formula with its leading `=`, in place of the value. */
  readonly formula?: string;
  /** Absent for the default style. */
  readonly style?: Style;
}

/**
 * The cells that a change changed, or, for a change to the sheets, no
 * ranges and the sheet to show: the one added, changed or put back, or the
 * one that took the place of a deleted one.
 */
export interface Change {
  readonly sheet: string;
  readonly ranges: readonly Range[];
}

/** How many changes can be undone (R5). */
const UNDO_LIMIT = 100;

/** A rectangle of cells, its bounds included. */
export interface Range {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** What a set of cells adds up to, as the status bar shows it. */
export interface Aggregate {
  /** The cells that are not empty. */
  readonly count: number;
  /** The cells that hold a number, and their total. */
  readonly numbers: number;
  readonly sum: number;
}

/** What a cell holds and shows. */
export interface Cell {
  readonly value: Value;
  /** The formula with its leading `=`, or `null` for a plain value. */
  readonly formula: string | null;
  /** The value shown in the cell's format. */
  readonly display: string;
  /**
   * The color, as `#rrggbb`, that the cell's format gives its value, as
   * `[Red]` does, over the style's; absent when it gives none.
   */
  readonly color?: string;
  /** The text that edits the cell, as the formula bar shows it. */
  readonly input: string;
  readonly style: Style;
  /**
   * What the conditional formats change of the style, drawn over it and
   * over the format's color; absent when none applies (R20).
   */
  readonly conditional?: RuleFormat;
}

export interface SheetInfo {
  readonly id: string;
  readonly name: string;
  /** The color of the tab, as `#rrggbb`; absent for none. */
  readonly color?: string;
}

export interface SheetSnapshot extends SheetInfo {
  /** Values by column, then by row; `null` where a cell is empty. */
  readonly columns: Constant[][];
  /** Formulas as `[row, column, text]`, the text with its leading `=`. */
  readonly formulas: [number, number, string][];
  /**
   * Indexes in the workbook's styles by column, then by row; 0 where a cell
   * has the default style. Absent in snapshots saved before cells had
   * styles.
   */
  readonly styleIds?: number[][];
  /** Sizes as `[index, size]`, in CSS pixels at 100%; absent for none. */
  readonly rowHeights?: [number, number][];
  readonly columnWidths?: [number, number][];
  /** Absent for none. */
  readonly hiddenRows?: number[];
  readonly hiddenColumns?: number[];
  /** The ranges merged into one cell; absent for none. */
  readonly merges?: Range[];
  /** The rows and columns frozen; absent for none. */
  readonly frozen?: Frozen;
  /** The conditional formats, the first first; absent for none. */
  readonly conditionalFormats?: Rule[];
}

/**
 * The versioned, JSON-compatible form in which a workbook leaves the
 * engine, for storage and files.
 */
export interface Snapshot {
  readonly version: 1;
  readonly name: string;
  readonly sheets: SheetSnapshot[];
  /**
   * The distinct styles of the cells, the default one first. Absent in
   * snapshots saved before cells had styles.
   */
  readonly styles?: Style[];
}

type Paste = Extract<Command, { type: 'paste' }>;
type SheetChange = Extract<
  Command,
  { type: 'addSheet' | 'deleteSheet' | 'restoreSheet' | 'setSheet' }
>;
type SetLines = Extract<Command, { type: 'setLines' }>;
type RestoreSheets = Extract<Command, { type: 'restoreSheets' }>;
type Fill = Extract<Command, { type: 'fill' }>;
type CellCommand = Extract<
  Command,
  {
    type:
      | 'setCell'
      | 'setCells'
      | 'clear'
      | 'restore'
      | 'setStyle'
      | 'setBorders'
      | 'merge'
      | 'unmerge';
  }
>;

/** Whether two ranges share a cell. */
function overlap(a: Range, b: Range): boolean {
  return (
    a.top <= b.bottom &&
    b.top <= a.bottom &&
    a.left <= b.right &&
    b.left <= a.right
  );
}

/**
 * Ranges grown over the merges they overlap, and those that then overlap
 * each other made one, as each becomes one merged cell.
 */
function combined(ranges: readonly Range[], merges: readonly Range[]): Range[] {
  const result: Range[] = [];
  for (const range of ranges) {
    let grown = withMerges(range, merges);
    for (let index = 0; index < result.length;) {
      if (overlap(result[index]!, grown)) {
        grown = withMerges(withMerges(grown, [result[index]!]), merges);
        result.splice(index, 1);
        index = 0;
      } else {
        index++;
      }
    }
    result.push(grown);
  }
  return result;
}

/** The range of one cell. */
function cellArea(row: number, col: number): Range {
  return { top: row, left: col, bottom: row, right: col };
}

/** A range grown until it takes in every merge it overlaps. */
export function withMerges(range: Range, merges: readonly Range[]): Range {
  let result = range;
  for (let grown = true; grown;) {
    grown = false;
    for (const merge of merges) {
      if (!overlap(result, merge)) continue;
      const top = Math.min(result.top, merge.top);
      const left = Math.min(result.left, merge.left);
      const bottom = Math.max(result.bottom, merge.bottom);
      const right = Math.max(result.right, merge.right);
      if (
        top !== result.top ||
        left !== result.left ||
        bottom !== result.bottom ||
        right !== result.right
      ) {
        result = { top, left, bottom, right };
        grown = true;
      }
    }
  }
  return result;
}

/** The cells a `setCells` command sets, inside the sheet. */
function blockOf(command: Extract<Command, { type: 'setCells' }>): Range {
  const { row, col, inputs } = command;
  const width = Math.max(0, ...inputs.map((each) => each.length));
  if (width === 0) throw new Error('An empty block');
  return {
    top: row,
    left: col,
    bottom: Math.min(row + inputs.length, SHEET_ROWS) - 1,
    right: Math.min(col + width, SHEET_COLUMNS) - 1,
  };
}

/** The name of a cell, as in `A1`. */
export function cellName(row: number, col: number): string {
  return columnName(col) + (row + 1);
}

/** A new workbook with one empty sheet. */
export function createWorkbook(): Workbook {
  return Workbook.fromSnapshot({
    version: 1,
    name: 'Book1',
    sheets: [{ id: 'sheet1', name: 'Sheet1', columns: [], formulas: [] }],
  });
}

export class Workbook {
  private readonly listeners = new Set<(command: Command) => void>();
  /** The formulas, what they read, and their results. */
  private readonly calc = new Calculation(() => this.sheetList);
  /** The index of each style in `styles`, by its JSON text. */
  private readonly styleIndex: Map<string, number>;
  /** The commands that undo the last changes, the latest last. */
  private readonly undoList: Command[] = [];
  /** The commands that redo the changes undone, the latest undone last. */
  private readonly redoList: Command[] = [];
  /** The formats that conditional formats give cells. */
  private readonly conditions = new Rules();

  private constructor(
    private readonly name: string,
    private readonly sheetList: Sheet[],
    private readonly styles: Style[],
    private readonly measure: TextWidth,
  ) {
    this.styleIndex = new Map(
      styles.map((style, index) => [JSON.stringify(style), index]),
    );
  }

  /**
   * The workbook of a snapshot, with `changes` made after it replayed.
   * There is nothing to undo in it. `measure` gives the width of text, to
   * fit rows to wrapped text; without fonts, a guess.
   */
  static fromSnapshot(
    snapshot: Snapshot,
    changes: readonly Command[] = [],
    measure: TextWidth = estimatedWidth,
  ): Workbook {
    if (snapshot.version !== 1) {
      throw new Error(`Unknown snapshot version: ${String(snapshot.version)}`);
    }
    const sheets = snapshot.sheets.map(sheetOf);
    const workbook = new Workbook(
      snapshot.name,
      sheets,
      snapshot.styles ?? [{}],
      measure,
    );
    workbook.calc.rebuild();
    workbook.fitNewSheets();
    for (const command of changes) workbook.apply(command);
    return workbook;
  }

  toSnapshot(): Snapshot {
    return {
      version: 1,
      name: this.name,
      sheets: this.sheetList.map(snapshotOf),
      styles: [...this.styles],
    };
  }

  /** The sheets in tab order. */
  get sheets(): readonly SheetInfo[] {
    return this.sheetList.map(({ id, name, color }) => ({
      id,
      name,
      ...(color !== undefined && { color }),
    }));
  }

  /**
   * Why `name` cannot be the name of the sheet `sheetId`, or `null` when it
   * can. Names differ when they differ in more than case, as in Excel.
   */
  sheetNameProblem(name: string, sheetId: string): SheetNameProblem | null {
    if (name.trim() === '') return 'empty';
    if (name.length > SHEET_NAME_LIMIT) return 'tooLong';
    if (/[:\\/?*[\]]/.test(name) || /^'|'$/.test(name)) return 'character';
    const folded = name.toLowerCase();
    const taken = this.sheetList.some(
      (sheet) => sheet.id !== sheetId && sheet.name.toLowerCase() === folded,
    );
    return taken ? 'taken' : null;
  }

  /**
   * Makes a change, which can then be undone, and gives what it changed;
   * what could be redone is forgotten.
   */
  execute(command: Command): Change {
    this.redoList.length = 0;
    return this.apply(command, this.undoList);
  }

  /** Undoes the last change, and gives the cells it changed. */
  undo(): Change | null {
    return this.reverse(this.undoList, this.redoList);
  }

  /** Redoes the last change undone, and gives the cells it changed. */
  redo(): Change | null {
    return this.reverse(this.redoList, this.undoList);
  }

  get canUndo(): boolean {
    return this.undoList.length > 0;
  }

  get canRedo(): boolean {
    return this.redoList.length > 0;
  }

  /** How many formula results have been calculated, to measure recalculation. */
  get calculations(): number {
    return this.calc.calculations;
  }

  /** A cell of a circular reference, or `null` when there is none (R15). */
  circularReference(): { sheet: string; row: number; col: number } | null {
    const formula = this.calc.circularReference();
    return formula
      ? { sheet: formula.sheet.id, row: formula.row, col: formula.col }
      : null;
  }

  cell(sheetId: string, row: number, col: number): Cell {
    const sheet = this.sheet(sheetId);
    const value = this.calc.cell(sheet, row, col);
    const formula =
      sheet.formulas.size === 0
        ? null
        : (sheet.formulas.get(key(row, col))?.text ?? null);
    const style = this.styles[sheet.styleIds[col]?.[row] ?? 0]!;
    const code = style.numberFormat ?? GENERAL;
    const color = code === GENERAL ? undefined : formatColor(value, code);
    const conditional =
      sheet.rules.length === 0
        ? undefined
        : this.conditions.format(sheet.rules, row, col, value, (area, visit) =>
            this.calc.cells(sheet, area, visit),
          );
    return {
      value,
      formula,
      style,
      display: formatValue(value, code),
      ...(color !== undefined && { color }),
      ...(conditional !== undefined && { conditional }),
      // Only the formula bar and the editor need it, not every drawn cell.
      get input() {
        return formula ?? editText(value, code);
      },
    };
  }

  /**
   * The count, and the total of the numbers, of the cells of each range. A
   * cell in two ranges counts twice, as in `SUM`.
   */
  aggregate(sheetId: string, ranges: readonly Range[]): Aggregate {
    const sheet = this.sheet(sheetId);
    let count = 0;
    let numbers = 0;
    let sum = 0;
    const add = (value: Value) => {
      if (value === null) return;
      count++;
      if (typeof value === 'number') {
        numbers++;
        sum += value;
      }
    };
    for (const { top, left, bottom, right } of ranges) {
      // Formula cells hold null among the values, so they are added apart.
      const last = Math.min(right, sheet.columns.length - 1);
      for (let col = left; col <= last; col++) {
        const column = sheet.columns[col];
        if (!column) continue;
        const end = Math.min(bottom, column.length - 1);
        // The hot loop of a whole sheet, so `add` is written out here.
        for (let row = top; row <= end; row++) {
          const value = column[row];
          if (typeof value === 'number') {
            count++;
            numbers++;
            sum += value;
          } else if (value != null) {
            count++;
          }
        }
      }
      for (const { row, col } of sheet.formulas.values()) {
        if (row >= top && row <= bottom && col >= left && col <= right) {
          add(this.calc.cell(sheet, row, col));
        }
      }
    }
    return { count, numbers, sum };
  }

  /**
   * The sizes of the rows and columns of a sheet and the ones hidden: the
   * same object until they change.
   */
  layout(sheetId: string): SheetLayout {
    return this.sheet(sheetId).layout;
  }

  /** The conditional formats of a sheet, the first first. */
  rules(sheetId: string): readonly Rule[] {
    return this.sheet(sheetId).rules;
  }

  /** The ranges merged into one cell: the same array until they change. */
  merges(sheetId: string): readonly Range[] {
    return this.sheet(sheetId).merges;
  }

  /** How many rows and columns, from A1, the cells in use extend over. */
  extent(sheetId: string): { rows: number; cols: number } {
    const sheet = this.sheet(sheetId);
    let rows = 0;
    let cols = 0;
    sheet.columns.forEach((column, col) => {
      let row = column.length - 1;
      while (row >= 0 && column[row] == null) row--;
      if (row < 0) return;
      rows = Math.max(rows, row + 1);
      cols = col + 1;
    });
    for (const { row, col } of sheet.formulas.values()) {
      rows = Math.max(rows, row + 1);
      cols = Math.max(cols, col + 1);
    }
    return { rows, cols };
  }

  /**
   * Calls `listener` with each command after it is executed, until the
   * returned function is called.
   */
  subscribe(listener: (command: Command) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** Runs the last command of `from`, and keeps its reverse in `to`. */
  private reverse(from: Command[], to: Command[]): Change | null {
    const command = from.pop();
    return command ? this.apply(command, to) : null;
  }

  /**
   * Runs a command, keeps the command that reverses it in `history`, tells
   * the listeners, and gives what it changed.
   */
  private apply(command: Command, history?: Command[]): Change {
    let applied: Applied;
    if (
      command.type === 'setCell' ||
      command.type === 'setCells' ||
      command.type === 'clear' ||
      command.type === 'restore' ||
      command.type === 'setStyle' ||
      command.type === 'setBorders' ||
      command.type === 'merge' ||
      command.type === 'unmerge'
    ) {
      applied = this.changeCells(command);
      const sheet = this.sheet(command.sheet);
      this.calc.changed(sheet, applied.change.ranges);
      this.calc.settle();
      this.fitRows(sheet, applied.change.ranges);
    } else if (
      command.type === 'paste' ||
      command.type === 'fill' ||
      command.type === 'group'
    ) {
      const restores =
        command.type === 'paste'
          ? this.pasted(command)
          : command.type === 'fill'
            ? [this.filled(command)]
            : command.restores;
      const each = restores.map((restore) => {
        const done = this.changeCells(restore);
        this.calc.changed(this.sheet(restore.sheet), done.change.ranges);
        return done;
      });
      this.calc.settle();
      for (const { change } of each) {
        this.fitRows(this.sheet(change.sheet), change.ranges);
      }
      if (each.length === 0) throw new Error('An empty group');
      // The restores change different cells, so they undo in any order.
      applied = {
        executed: { type: 'group', restores },
        reverse: {
          type: 'group',
          restores: each.map((done) => done.reverse as Restore),
        },
        change: each[0]!.change,
      };
    } else if (
      command.type === 'insertLines' ||
      command.type === 'deleteLines' ||
      command.type === 'restoreSheets'
    ) {
      applied =
        command.type === 'restoreSheets'
          ? this.restoreSheets(command)
          : this.moveLines(command);
      // Formulas have moved, so what each reads is found again.
      this.calc.rebuild();
      this.fitNewSheets();
    } else if (command.type === 'freeze') {
      applied = this.freeze(command);
    } else if (command.type === 'setRules') {
      applied = this.setRules(command);
    } else if (command.type === 'setLines') {
      applied = this.setLines(command);
      // Wrapped text takes more or fewer lines in another width.
      if (command.direction === 'columns') {
        this.fitRows(this.sheet(command.sheet));
      }
    } else {
      applied = this.changeSheets(command);
      this.fitNewSheets();
      // References to sheets by name may now find another sheet, or none.
      if (command.type !== 'setSheet' || command.name !== undefined) {
        this.calc.rebuild();
      }
    }
    // Values, formulas or rules may have changed what the rules find.
    this.conditions.forget();
    const { executed, reverse, change } = applied;
    if (history) {
      history.push(reverse);
      if (history.length > UNDO_LIMIT) history.shift();
    }
    for (const listener of this.listeners) listener(executed);
    return change;
  }

  private changeCells(command: CellCommand): Applied {
    const sheet = this.sheet(command.sheet);
    const ranges =
      command.type === 'setStyle' || command.type === 'setBorders'
        ? command.ranges.flatMap((range) => this.inUse(sheet, range) ?? [])
        : command.type === 'merge'
          ? combined(command.ranges, sheet.merges)
          : command.type === 'unmerge'
            ? this.mergesIn(sheet, command.ranges)
            : command.type === 'setCell'
              ? [
                  {
                    top: command.row,
                    left: command.col,
                    bottom: command.row,
                    right: command.col,
                  },
                ]
              : command.type === 'setCells'
                ? [blockOf(command)]
                : command.ranges;
    const merging =
      command.type === 'merge' ||
      command.type === 'unmerge' ||
      (command.type === 'restore' && command.merges !== undefined);
    const reverse: Restore = {
      type: 'restore',
      sheet: sheet.id,
      ranges,
      cells: this.contents(sheet, ranges),
      ...(merging && {
        merges: sheet.merges.filter((merge) =>
          ranges.some((range) => overlap(range, merge)),
        ),
      }),
    };
    let executed: Command = command;
    if (command.type === 'setCell' || command.type === 'setCells') {
      const year =
        command.type === 'setCell'
          ? this.setCell(sheet, command)
          : this.setCells(sheet, command, ranges[0]!);
      // Listeners that record the command replay it with the same date.
      if (year !== undefined) executed = { ...command, year };
    } else if (command.type === 'clear') {
      for (const range of command.ranges) this.clear(sheet, range);
    } else if (command.type === 'setStyle') {
      for (const range of ranges) this.setStyle(sheet, range, command.style);
    } else if (command.type === 'setBorders') {
      for (const range of ranges) this.setBorders(sheet, range, command);
    } else if (command.type === 'merge') {
      this.merge(sheet, ranges);
    } else if (command.type === 'unmerge') {
      this.unmerge(sheet, ranges);
    } else {
      this.restore(sheet, command);
    }
    return { executed, reverse, change: { sheet: sheet.id, ranges } };
  }

  private changeSheets(command: SheetChange): Applied {
    const list = this.sheetList;
    const shown = (sheet: string): Change => ({ sheet, ranges: [] });
    switch (command.type) {
      case 'addSheet': {
        const source = command.copyOf ? this.sheet(command.copyOf) : null;
        const sheet = command.sheet ?? this.freeId();
        const name =
          command.name ??
          (source ? this.copyName(source.name) : this.freeName());
        const contents = source
          ? { ...snapshotOf(source), id: sheet, name }
          : { id: sheet, name, columns: [], formulas: [] };
        this.insert(command.index, sheetOf(contents));
        return {
          executed: { ...command, sheet, name },
          reverse: { type: 'deleteSheet', sheet },
          change: shown(sheet),
        };
      }
      case 'restoreSheet':
        this.insert(command.index, sheetOf(command.contents));
        return {
          executed: command,
          reverse: { type: 'deleteSheet', sheet: command.contents.id },
          change: shown(command.contents.id),
        };
      case 'deleteSheet': {
        const sheet = this.sheet(command.sheet);
        if (list.length === 1) throw new Error('The last sheet stays');
        const index = list.indexOf(sheet);
        list.splice(index, 1);
        return {
          executed: command,
          reverse: { type: 'restoreSheet', index, contents: snapshotOf(sheet) },
          change: shown(list[Math.min(index, list.length - 1)]!.id),
        };
      }
      case 'setSheet': {
        const sheet = this.sheet(command.sheet);
        const { name, color, index } = command;
        if (name !== undefined) {
          const problem = this.sheetNameProblem(name, sheet.id);
          if (problem) throw new Error(`Sheet name ${problem}: ${name}`);
        }
        const reverse: Command = {
          type: 'setSheet',
          sheet: sheet.id,
          ...(name !== undefined && { name: sheet.name }),
          ...(color !== undefined && { color: sheet.color ?? null }),
          ...(index !== undefined && { index: list.indexOf(sheet) }),
        };
        if (name !== undefined) this.rename(sheet, name);
        if (color === null) delete sheet.color;
        else if (color !== undefined) sheet.color = color;
        if (index !== undefined) {
          list.splice(list.indexOf(sheet), 1);
          this.insert(index, sheet);
        }
        return { executed: command, reverse, change: shown(sheet.id) };
      }
    }
  }

  /**
   * Inserts or deletes rows or columns: replaces the sheet with one made
   * from its snapshot with the lines moved, and the other sheets whose
   * formulas refer to the lines moved with ones whose formulas follow.
   */
  private moveLines(command: MoveLines): Applied {
    const sheet = this.sheet(command.sheet);
    const { direction } = command;
    const rows = direction === 'rows';
    const limit = rows ? SHEET_ROWS : SHEET_COLUMNS;
    const spans = merged(checked(command));
    const insert = command.type === 'insertLines';
    // From the last span to the first, so that each keeps its place.
    const shifts = spans
      .map(({ start, end }) => ({
        at: start,
        count: insert ? end - start + 1 : start - end - 1,
      }))
      .reverse();
    /** Where a line goes, or `null` when it is deleted or pushed off. */
    const moved = (index: number): number | null => {
      let place = index;
      for (const { at, count } of shifts) {
        if (place < at) continue;
        if (count < 0 && place < at - count) return null;
        place += count;
      }
      return place < limit ? place : null;
    };
    const follow = (own: Sheet, ast: Ast): Ast =>
      shifts.reduce(
        (each, { at, count }) =>
          shift(each, {
            type: 'lines',
            sheet: sheet.name,
            own: own.name,
            direction,
            at,
            count,
          }),
        ast,
      );

    const before: SheetSnapshot[] = [];
    this.sheetList.forEach((each, index) => {
      const formulas: [number, number, string][] = [];
      let changed = each === sheet;
      for (const { row, col, ast, text } of each.formulas.values()) {
        const ast2 = follow(each, ast);
        if (ast2 !== ast) changed = true;
        const formula = ast2 === ast ? text : `=${print(ast2)}`;
        if (each !== sheet) {
          formulas.push([row, col, formula]);
          continue;
        }
        const place = moved(rows ? row : col);
        if (place === null) continue;
        formulas.push(rows ? [place, col, formula] : [row, place, formula]);
      }
      if (!changed) return;
      const snapshot = snapshotOf(each);
      before.push(snapshot);
      const after: SheetSnapshot =
        each === sheet
          ? { ...movedLines(snapshot, rows, shifts, limit, moved), formulas }
          : { ...snapshot, formulas };
      this.sheetList[index] = sheetOf(after);
    });
    // The lines inserted, where they end up, or the lines deleted.
    let added = 0;
    const lines = spans.map(({ start, end }) => {
      const span = { start: start + added, end: end + added };
      if (insert) added += end - start + 1;
      return span;
    });
    const ranges = lines.map((span) => lineRange(direction, span));
    return {
      executed: command,
      reverse: {
        type: 'restoreSheets',
        sheet: sheet.id,
        ranges,
        sheets: before,
      },
      change: { sheet: sheet.id, ranges },
    };
  }

  /** Puts sheets back from their snapshots, in their places. */
  private restoreSheets(command: RestoreSheets): Applied {
    const before = command.sheets.map((contents) => {
      const index = this.sheetList.indexOf(this.sheet(contents.id));
      const snapshot = snapshotOf(this.sheetList[index]!);
      this.sheetList[index] = sheetOf(contents);
      return snapshot;
    });
    const { sheet, ranges } = command;
    return {
      executed: command,
      reverse: { ...command, sheets: before },
      change: { sheet, ranges },
    };
  }

  /** Freezes rows and columns, or unfreezes them. */
  private freeze(command: Extract<Command, { type: 'freeze' }>): Applied {
    const sheet = this.sheet(command.sheet);
    const { rows, cols } = command;
    const valid = (count: number, limit: number) =>
      Number.isInteger(count) && count >= 0 && count < limit;
    if (!valid(rows, SHEET_ROWS) || !valid(cols, SHEET_COLUMNS)) {
      throw new Error(`Invalid frozen panes: ${rows} rows, ${cols} columns`);
    }
    const old = sheet.layout.frozen ?? { rows: 0, cols: 0 };
    const { rows: rowLines, columns } = sheet.layout;
    sheet.layout = {
      rows: rowLines,
      columns,
      ...((rows > 0 || cols > 0) && { frozen: { rows, cols } }),
    };
    return {
      executed: command,
      reverse: { ...command, ...old },
      change: { sheet: sheet.id, ranges: [] },
    };
  }

  /** Replaces the conditional formats of a sheet. */
  private setRules(command: Extract<Command, { type: 'setRules' }>): Applied {
    const sheet = this.sheet(command.sheet);
    const old = sheet.rules;
    sheet.rules = command.rules;
    return {
      executed: command,
      reverse: { ...command, rules: old },
      change: {
        sheet: sheet.id,
        ranges: [...old, ...command.rules].flatMap((rule) => rule.ranges),
      },
    };
  }

  /** Sets the sizes and visibility of rows or columns. */
  private setLines(command: SetLines): Applied {
    const sheet = this.sheet(command.sheet);
    const { direction } = command;
    const old = sheet.layout[direction];
    const sizes = new Map(old.sizes);
    const hidden = new Set(old.hidden);
    // The lines as they were, in runs that look the same.
    const reverse: LineSpan[] = [];
    for (const span of checked(command)) {
      for (let index = span.start; index <= span.end; index++) {
        const size = old.sizes.get(index) ?? null;
        const hides = old.hidden.has(index);
        const last = reverse[reverse.length - 1];
        if (
          last?.end === index - 1 &&
          last.size === size &&
          last.hidden === hides
        ) {
          reverse[reverse.length - 1] = { ...last, end: index };
        } else {
          reverse.push({ start: index, end: index, size, hidden: hides });
        }
        if (span.size === null) sizes.delete(index);
        else if (span.size !== undefined) sizes.set(index, span.size);
        if (span.hidden === true) hidden.add(index);
        else if (span.hidden === false) hidden.delete(index);
      }
    }
    sheet.layout = { ...sheet.layout, [direction]: { ...old, sizes, hidden } };
    const ranges = command.spans.map((span) => lineRange(direction, span));
    return {
      executed: command,
      reverse: { ...command, spans: reverse },
      change: { sheet: sheet.id, ranges },
    };
  }

  /** The restores that make a paste, read before any of them runs. */
  private pasted(command: Paste): Restore[] {
    const { target, source, cut, only } = command;
    const from = this.sheet(source.sheet);
    const to = this.sheet(command.sheet);
    const { range } = source;
    const height = range.bottom - range.top + 1;
    const width = range.right - range.left + 1;
    const whole =
      !cut &&
      (target.bottom - target.top + 1) % height === 0 &&
      (target.right - target.left + 1) % width === 0;
    // A cut keeps every cell, moved back from the edges of the sheet; a
    // copy keeps what fits.
    const top = cut ? Math.min(target.top, SHEET_ROWS - height) : target.top;
    const left = cut
      ? Math.min(target.left, SHEET_COLUMNS - width)
      : target.left;
    const area: Range = {
      top,
      left,
      bottom: Math.min(
        SHEET_ROWS - 1,
        whole ? target.bottom : top + height - 1,
      ),
      right: Math.min(
        SHEET_COLUMNS - 1,
        whole ? target.right : left + width - 1,
      ),
    };
    const cells = this.contents(from, [range]);
    if (cut) return this.moved(from, range, to, area, cells);

    // What the target keeps: its formats when contents are pasted, its
    // contents when formats are, and nothing when both are.
    const pasted = new Map<string, CellContents>();
    if (only) {
      for (const cell of this.contents(to, [area])) {
        const { row, col, value, formula, style } = cell;
        pasted.set(
          key(row, col),
          only === 'formats'
            ? { row, col, value, ...(formula !== undefined && { formula }) }
            : { row, col, value: null, ...(style && { style }) },
        );
      }
    }
    for (let top = area.top; top <= area.bottom; top += height) {
      for (let left = area.left; left <= area.right; left += width) {
        const rows = top - range.top;
        const cols = left - range.left;
        for (const cell of cells) {
          const row = cell.row + rows;
          const col = cell.col + cols;
          if (row > area.bottom || col > area.right) continue;
          const kept = pasted.get(key(row, col));
          let contents: Pick<CellContents, 'value' | 'formula'>;
          if (only === 'formats') {
            contents = kept ?? { value: null };
          } else if (cell.formula === undefined) {
            contents = { value: cell.value };
          } else if (only === 'values') {
            const value = this.calc.cell(from, cell.row, cell.col);
            // Gridworks has no error constants, so an error is kept as text.
            contents = { value: isError(value) ? value.error : value };
          } else {
            const { ast } = from.formulas.get(key(cell.row, cell.col))!;
            const moved = shift(ast, { type: 'offset', rows, cols });
            contents = { value: null, formula: `=${print(moved)}` };
          }
          const style =
            only === 'values' || only === 'formulas' ? kept?.style : cell.style;
          pasted.set(key(row, col), {
            row,
            col,
            ...contents,
            ...(style && { style }),
          });
        }
      }
    }
    const filled = [...pasted.values()].filter(
      ({ value, formula, style }) =>
        value !== null || formula !== undefined || style !== undefined,
    );
    // Formats bring the merges of the cells copied, as in Excel.
    const merges: Range[] = [];
    if (only === 'formats') {
      const inside = from.merges.filter(
        (merge) =>
          merge.top >= range.top &&
          merge.bottom <= range.bottom &&
          merge.left >= range.left &&
          merge.right <= range.right,
      );
      for (let top = area.top; top <= area.bottom; top += height) {
        for (let left = area.left; left <= area.right; left += width) {
          for (const merge of inside) {
            const moved = {
              top: merge.top + top - range.top,
              left: merge.left + left - range.left,
              bottom: merge.bottom + top - range.top,
              right: merge.right + left - range.left,
            };
            if (moved.bottom <= area.bottom && moved.right <= area.right) {
              merges.push(moved);
            }
          }
        }
      }
    }
    return [
      {
        type: 'restore',
        sheet: to.id,
        ranges: [area],
        cells: filled,
        ...(only === 'formats' && { merges }),
      },
    ];
  }

  /** Whether a fill extends a series in any row or column of its source. */
  fillsSeries(fill: Fill): boolean {
    return this.fillLines(fill).lines.some((line) => line.extended);
  }

  /** The restore that makes a fill, read before it runs. */
  private filled(fill: Fill): Restore {
    const { sheet, area, places, lines } = this.fillLines(fill);
    const { how } = fill;
    // What the target keeps: its contents when formats are filled, and its
    // formats when contents are.
    const kept = new Map<string, CellContents>();
    if (how === 'formats' || how === 'values') {
      for (const cell of this.contents(sheet, [area])) {
        kept.set(key(cell.row, cell.col), cell);
      }
    }
    const cells: CellContents[] = [];
    for (const { seeds, extended, at } of lines) {
      const length = seeds.length;
      for (const place of places) {
        const { row, col } = at(place);
        const seed = seeds[((place % length) + length) % length];
        const keep = kept.get(key(row, col));
        let contents: Pick<CellContents, 'value' | 'formula'>;
        if (how === 'formats') {
          contents = {
            value: keep?.value ?? null,
            ...(keep?.formula !== undefined && { formula: keep.formula }),
          };
        } else if (extended) {
          contents = { value: extended(place) };
        } else if (seed?.formula !== undefined) {
          const { ast } = sheet.formulas.get(key(seed.row, seed.col))!;
          const move: Move = {
            type: 'offset',
            rows: row - seed.row,
            cols: col - seed.col,
          };
          contents = { value: null, formula: `=${print(shift(ast, move))}` };
        } else {
          contents = { value: seed?.value ?? null };
        }
        const style = how === 'values' ? keep?.style : seed?.style;
        if (contents.value === null && !contents.formula && !style) continue;
        cells.push({ row, col, ...contents, ...(style && { style }) });
      }
    }
    return { type: 'restore', sheet: sheet.id, ranges: [area], cells };
  }

  /**
   * The area a fill changes, the places it fills along the way it goes,
   * the source's cells being at 0, 1 and on, and each row or column of the
   * source: its cells, the series they start, if any, and where each place
   * of it is.
   */
  private fillLines({ sheet: id, source, target, how }: Fill) {
    const sheet = this.sheet(id);
    const contains =
      target.top <= source.top &&
      target.left <= source.left &&
      target.bottom >= source.bottom &&
      target.right >= source.right;
    const sides = [
      source.top > target.top,
      target.bottom > source.bottom,
      source.left > target.left,
      target.right > source.right,
    ];
    if (!contains || sides.filter(Boolean).length !== 1) {
      throw new Error('A fill extends its source on one side');
    }
    const [up, down, left] = sides;
    const vertical = up || down;
    const area: Range = up
      ? { ...target, bottom: source.top - 1 }
      : down
        ? { ...target, top: source.bottom + 1 }
        : left
          ? { ...target, right: source.left - 1 }
          : { ...target, left: source.right + 1 };
    const seeds = new Map(
      this.contents(sheet, [source]).map((cell) => [
        key(cell.row, cell.col),
        cell,
      ]),
    );
    const length = vertical
      ? source.bottom - source.top + 1
      : source.right - source.left + 1;
    const [first, last] = vertical
      ? [source.left, source.right]
      : [source.top, source.bottom];
    const [from, to] = vertical
      ? [area.top - source.top, area.bottom - source.top]
      : [area.left - source.left, area.right - source.left];
    const places = Array.from({ length: to - from + 1 }, (_, i) => from + i);
    const lines = [];
    for (let line = first; line <= last; line++) {
      const at = (place: number) =>
        vertical
          ? { row: source.top + place, col: line }
          : { row: line, col: source.left + place };
      const lineSeeds = Array.from({ length }, (_, place) => {
        const { row, col } = at(place);
        return seeds.get(key(row, col));
      });
      const extended =
        how === 'copy' || how === 'formats'
          ? null
          : series(
              lineSeeds.map((cell) => ({
                value: cell?.value ?? null,
                formula: cell?.formula !== undefined,
                format: cell?.style?.numberFormat ?? GENERAL,
              })),
              how === 'series',
            );
      lines.push({ seeds: lineSeeds, extended, at });
    }
    return { sheet, area, places, lines };
  }

  /**
   * The restores that move cut cells to an area, and rewrite the formulas
   * elsewhere that refer to them.
   */
  private moved(
    from: Sheet,
    range: Range,
    to: Sheet,
    area: Range,
    cells: readonly CellContents[],
  ): Restore[] {
    const rows = area.top - range.top;
    const cols = area.left - range.left;
    const move = (own: Sheet, newOwn: Sheet): Move => ({
      type: 'moveCells',
      sheet: from.name,
      range,
      to: to.name,
      rows,
      cols,
      own: own.name,
      newOwn: newOwn.name,
    });
    const inside = (
      cell: { row: number; col: number },
      { top, left, bottom, right }: Range,
    ) =>
      cell.row >= top &&
      cell.row <= bottom &&
      cell.col >= left &&
      cell.col <= right;
    const placed = cells.map((cell) => {
      const row = cell.row + rows;
      const col = cell.col + cols;
      if (cell.formula === undefined) return { ...cell, row, col };
      const { ast } = from.formulas.get(key(cell.row, cell.col))!;
      const formula = `=${print(shift(ast, move(from, to)))}`;
      return { ...cell, row, col, formula };
    });
    const restores: Restore[] =
      from === to
        ? [
            {
              type: 'restore',
              sheet: to.id,
              ranges: [area, range],
              cells: placed,
            },
          ]
        : [
            { type: 'restore', sheet: to.id, ranges: [area], cells: placed },
            { type: 'restore', sheet: from.id, ranges: [range], cells: [] },
          ];
    for (const sheet of this.sheetList) {
      const changed: CellContents[] = [];
      for (const { row, col, ast } of sheet.formulas.values()) {
        if (sheet === from && inside({ row, col }, range)) continue;
        if (sheet === to && inside({ row, col }, area)) continue;
        const moved = shift(ast, move(sheet, sheet));
        if (moved === ast) continue;
        const id = sheet.styleIds[col]?.[row] ?? 0;
        changed.push({
          row,
          col,
          value: null,
          formula: `=${print(moved)}`,
          ...(id !== 0 && { style: this.styles[id] }),
        });
      }
      if (changed.length === 0) continue;
      restores.push({
        type: 'restore',
        sheet: sheet.id,
        ranges: changed.map(({ row, col }) => ({
          top: row,
          left: col,
          bottom: row,
          right: col,
        })),
        cells: changed,
      });
    }
    return restores;
  }

  /** Renames a sheet, and the references to it in every formula. */
  private rename(sheet: Sheet, name: string): void {
    const move = { type: 'renameSheet', from: sheet.name, to: name } as const;
    for (const each of this.sheetList) {
      const moved = [...each.formulas.values()].flatMap((formula) => {
        const ast = shift(formula.ast, move);
        return ast === formula.ast ? [] : [{ ...formula, ast }];
      });
      for (const { row, col, ast } of moved) {
        this.calc.put(each, row, col, `=${print(ast)}`, ast);
      }
    }
    sheet.name = name;
  }

  /** Puts a sheet at a place in the tab order, kept within the tabs. */
  private insert(index: number, sheet: Sheet): void {
    const list = this.sheetList;
    if (list.some((each) => each.id === sheet.id)) {
      throw new Error(`Sheet exists: ${sheet.id}`);
    }
    list.splice(Math.max(0, Math.min(index, list.length)), 0, sheet);
  }

  /** The first identifier `sheetN` that no sheet has. */
  private freeId(): string {
    const ids = new Set(this.sheetList.map((sheet) => sheet.id));
    let number = 1;
    while (ids.has(`sheet${number}`)) number++;
    return `sheet${number}`;
  }

  /** The first name `SheetN` that is free, from the number of sheets on. */
  private freeName(): string {
    let number = this.sheetList.length + 1;
    while (this.sheetNameProblem(`Sheet${number}`, '')) number++;
    return `Sheet${number}`;
  }

  /** The name of a copy, as Excel gives it: `Sheet1 (2)`. */
  private copyName(name: string): string {
    const base = name.replace(/ \(\d+\)$/, '');
    for (let number = 2; ; number++) {
      const suffix = ` (${number})`;
      const copy = base.slice(0, SHEET_NAME_LIMIT - suffix.length) + suffix;
      if (!this.sheetNameProblem(copy, '')) return copy;
    }
  }

  /** The cells of ranges that have contents or a style other than the default. */
  private contents(sheet: Sheet, ranges: readonly Range[]): CellContents[] {
    const cells = new Map<string, CellContents>();
    for (const { top, left, bottom, right } of ranges) {
      const last = Math.min(
        right,
        Math.max(sheet.columns.length, sheet.styleIds.length) - 1,
      );
      for (let col = left; col <= last; col++) {
        const values = sheet.columns[col] ?? [];
        const ids = sheet.styleIds[col] ?? [];
        const end = Math.min(bottom, Math.max(values.length, ids.length) - 1);
        for (let row = top; row <= end; row++) {
          const value = values[row] ?? null;
          const id = ids[row] ?? 0;
          if (value === null && id === 0) continue;
          cells.set(key(row, col), {
            row,
            col,
            value,
            ...(id !== 0 && { style: this.styles[id] }),
          });
        }
      }
      for (const [cellKey, { row, col, text }] of sheet.formulas) {
        if (row < top || row > bottom || col < left || col > right) continue;
        const id = sheet.styleIds[col]?.[row] ?? 0;
        cells.set(cellKey, {
          row,
          col,
          value: null,
          formula: text,
          ...(id !== 0 && { style: this.styles[id] }),
        });
      }
    }
    return [...cells.values()];
  }

  /** Empties the ranges of a `restore`, styles too, and fills in its cells. */
  private restore(sheet: Sheet, { ranges, cells, merges }: Restore): void {
    for (const range of ranges) {
      this.clear(sheet, range);
      const { top, left, bottom, right } = range;
      const last = Math.min(right, sheet.styleIds.length - 1);
      for (let col = left; col <= last; col++) {
        const ids = sheet.styleIds[col];
        if (!ids) continue;
        const end = Math.min(bottom, ids.length - 1);
        for (let row = top; row <= end; row++) ids[row] = 0;
      }
    }
    for (const { row, col, value, formula, style } of cells) {
      if (formula !== undefined) {
        const ast = parse(formula.slice(1));
        // A command read back from storage may be damaged.
        if (!ast) throw new Error(`Invalid formula in restore: ${formula}`);
        this.calc.put(sheet, row, col, formula, ast);
      }
      (sheet.columns[col] ??= [])[row] = value;
      if (style) (sheet.styleIds[col] ??= [])[row] = this.styleId(style);
    }
    if (merges) {
      this.unmerge(sheet, ranges);
      sheet.merges = [...sheet.merges, ...merges];
    }
  }

  /** Sets a cell, and gives the year it gave a date typed without one. */
  private setCell(
    sheet: Sheet,
    command: Extract<Command, { type: 'setCell' }>,
  ): number | undefined {
    const { row, col, input } = command;
    const column = (sheet.columns[col] ??= []);
    const code = this.styles[sheet.styleIds[col]?.[row] ?? 0]!.numberFormat;
    if (code !== undefined && isTextFormat(code)) {
      // A cell with the text format keeps what is typed, as in Excel.
      this.calc.remove(sheet, row, col);
      column[row] = input === '' ? null : input;
      return undefined;
    }
    const parsed = input.startsWith('=') ? parse(input.slice(1)) : null;
    if (parsed) {
      // Sheet names are written as the sheets have them, as in Excel.
      const ast = this.sheetList.reduce(
        (ast: Ast, { name }) =>
          shift(ast, { type: 'renameSheet', from: name, to: name }),
        parsed,
      );
      this.calc.put(sheet, row, col, `=${print(ast)}`, ast);
      column[row] = null;
      return undefined;
    }
    this.calc.remove(sheet, row, col);
    const { value, format, year } = parseInput(input, command.year);
    column[row] = value;
    if (format === undefined) return year;
    // The way the value was typed sets the cell's format, as in Excel.
    const ids = (sheet.styleIds[col] ??= []);
    const style = { ...this.styles[ids[row] ?? 0], numberFormat: format };
    ids[row] = this.styleId(style);
    return year;
  }

  /**
   * Sets the cells of a block that fall inside `range`, and gives the year
   * it gave a date typed without one.
   */
  private setCells(
    sheet: Sheet,
    command: Extract<Command, { type: 'setCells' }>,
    range: Range,
  ): number | undefined {
    let year = command.year;
    command.inputs.forEach((inputs, index) => {
      const row = command.row + index;
      if (row > range.bottom) return;
      inputs.forEach((input, offset) => {
        const col = command.col + offset;
        if (col > range.right) return;
        year =
          this.setCell(sheet, {
            ...command,
            type: 'setCell',
            row,
            col,
            input,
            year,
          }) ?? year;
      });
    });
    return year;
  }

  /** Changes the style of the cells of a range. */
  private setStyle(sheet: Sheet, range: Range, change: StyleChange): void {
    for (let col = range.left; col <= range.right; col++) {
      const ids = (sheet.styleIds[col] ??= []);
      for (let row = range.top; row <= range.bottom; row++) {
        ids[row] = this.styleId(
          changedStyle(this.styles[ids[row] ?? 0]!, change),
        );
      }
    }
  }

  /**
   * Puts a border on sides of the cells of a range: those at its edges
   * for the sides of the range, every side for `all`.
   */
  private setBorders(
    sheet: Sheet,
    range: Range,
    { sides, border }: Extract<Command, { type: 'setBorders' }>,
  ): void {
    // Kept in one order, so that equal borders make equal styles.
    const line = border && {
      line: border.line,
      ...(border.color !== undefined && { color: border.color }),
    };
    const { top, left, bottom, right } = range;
    const edges: [BorderSides, keyof Style, Range][] = [
      ['top', 'borderTop', { ...range, bottom: top }],
      ['right', 'borderRight', { ...range, left: right }],
      ['bottom', 'borderBottom', { ...range, top: bottom }],
      ['left', 'borderLeft', { ...range, right: left }],
    ];
    for (const [side, property, edge] of edges) {
      if (sides === 'all') {
        this.setStyle(sheet, range, { [property]: line });
      } else if (sides === 'outside' || sides === side) {
        this.setStyle(sheet, edge, { [property]: line });
      }
    }
  }

  /** Merges ranges, emptying all but their top-left cells. */
  private merge(sheet: Sheet, ranges: readonly Range[]): void {
    this.unmerge(sheet, ranges);
    for (const range of ranges) {
      const { top, left } = range;
      const kept = sheet.formulas.get(key(top, left))?.text;
      const value = sheet.columns[left]?.[top] ?? null;
      this.clear(sheet, range);
      if (kept !== undefined) {
        this.calc.put(sheet, top, left, kept, parse(kept.slice(1))!);
      } else if (value !== null) {
        sheet.columns[left]![top] = value;
      }
    }
    const merged = ranges.filter(
      ({ top, left, bottom, right }) => top !== bottom || left !== right,
    );
    sheet.merges = [...sheet.merges, ...merged];
  }

  /** Unmerges the merges that ranges overlap. */
  private unmerge(sheet: Sheet, ranges: readonly Range[]): void {
    const kept = sheet.merges.filter(
      (merge) => !ranges.some((range) => overlap(range, merge)),
    );
    if (kept.length !== sheet.merges.length) sheet.merges = kept;
  }

  /** The merges that ranges overlap, or the ranges when there are none. */
  private mergesIn(sheet: Sheet, ranges: readonly Range[]): Range[] {
    const merges = sheet.merges.filter((merge) =>
      ranges.some((range) => overlap(range, merge)),
    );
    return merges.length > 0 ? merges : [...ranges];
  }

  /**
   * A range with its whole rows or columns cut down to the cells in use:
   * those with a value, a formula or a style; `null` when there are none.
   */
  private inUse(sheet: Sheet, range: Range): Range | null {
    let { bottom, right } = range;
    if (range.top === 0 && bottom === SHEET_ROWS - 1) {
      let rows = 0;
      for (const column of [...sheet.columns, ...sheet.styleIds]) {
        rows = Math.max(rows, column?.length ?? 0);
      }
      for (const { row } of sheet.formulas.values()) {
        rows = Math.max(rows, row + 1);
      }
      if (rows === 0) return null;
      bottom = rows - 1;
    }
    if (range.left === 0 && right === SHEET_COLUMNS - 1) {
      let cols = Math.max(sheet.columns.length, sheet.styleIds.length);
      for (const { col } of sheet.formulas.values()) {
        cols = Math.max(cols, col + 1);
      }
      if (cols === 0) return null;
      right = cols - 1;
    }
    return { ...range, bottom, right };
  }

  /** Fits the rows of every sheet that has not been fitted yet. */
  private fitNewSheets(): void {
    for (const sheet of this.sheetList) {
      if (!sheet.layout.rows.fitted) this.fitRows(sheet);
    }
  }

  /**
   * Gives the rows of ranges, or every row of a sheet, the height their
   * text needs, kept apart from the sizes set on rows.
   */
  private fitRows(sheet: Sheet, ranges?: readonly Range[]): void {
    const { rows, columns } = sheet.layout;
    const growing = this.styles.map(mayGrowRow);
    const old = rows.fitted ?? new Map<number, number>();
    if (!growing.includes(true) && old.size === 0) {
      if (!rows.fitted) {
        sheet.layout = { ...sheet.layout, rows: { ...rows, fitted: old } };
      }
      return;
    }
    const fitted = new Map(ranges ? old : []);
    const heights = new Map<number, number>();
    const height = (row: number, col: number, id: number) => {
      // Merged cells are not fitted to their text, as in Excel.
      if (sheet.merges.some((merge) => overlap(merge, cellArea(row, col)))) {
        return;
      }
      const width = columns.hidden.has(col)
        ? 0
        : (columns.sizes.get(col) ?? STANDARD_COLUMN_WIDTH);
      if (width === 0) return;
      const { display } = this.cell(sheet.id, row, col);
      const needed = cellHeight(display, this.styles[id]!, width, this.measure);
      heights.set(row, Math.max(heights.get(row) ?? 0, needed));
    };
    sheet.styleIds.forEach((ids, col) => {
      if (!ids) return;
      if (!ranges) {
        ids.forEach((id, row) => {
          if (growing[id]) height(row, col, id);
        });
        return;
      }
      for (const { top, bottom } of ranges) {
        const end = Math.min(bottom, ids.length - 1);
        for (let row = top; row <= end; row++) {
          const id = ids[row] ?? 0;
          if (growing[id]) height(row, col, id);
        }
      }
    });
    for (const { top, bottom } of ranges ?? []) {
      for (const row of old.keys()) {
        if (row >= top && row <= bottom) fitted.delete(row);
      }
    }
    for (const [row, needed] of heights) {
      if (needed > STANDARD_ROW_HEIGHT) fitted.set(row, Math.ceil(needed));
    }
    const same =
      rows.fitted !== undefined &&
      fitted.size === old.size &&
      [...fitted].every(([row, size]) => old.get(row) === size);
    if (!same) {
      sheet.layout = { ...sheet.layout, rows: { ...rows, fitted } };
    }
  }

  /** Empties the cells of a range of a sheet, leaving their styles. */
  private clear(sheet: Sheet, range: Range): void {
    const { top, left, bottom, right } = range;
    const last = Math.min(right, sheet.columns.length - 1);
    for (let col = left; col <= last; col++) {
      const column = sheet.columns[col];
      if (!column) continue;
      const end = Math.min(bottom, column.length - 1);
      for (let row = top; row <= end; row++) column[row] = null;
    }
    const inside = [...sheet.formulas.values()].filter(
      ({ row, col }) =>
        row >= top && row <= bottom && col >= left && col <= right,
    );
    for (const { row, col } of inside) this.calc.remove(sheet, row, col);
  }

  /** The index of a style in the table, added if it is new. */
  private styleId(style: Style): number {
    const text = JSON.stringify(style);
    let id = this.styleIndex.get(text);
    if (id === undefined) {
      id = this.styles.push(style) - 1;
      this.styleIndex.set(text, id);
    }
    return id;
  }

  private sheet(id: string): Sheet {
    const sheet = this.sheetList.find((each) => each.id === id);
    if (!sheet) throw new Error(`No such sheet: ${id}`);
    return sheet;
  }
}

/** A command, as its listeners get it, and the command that reverses it. */
interface Applied {
  readonly executed: Command;
  readonly reverse: Command;
  readonly change: Change;
}

/**
 * Spans sorted, with those that overlap made one. Those that only touch
 * stay apart: an insertion goes before each.
 */
function merged(spans: readonly Span[]): Span[] {
  const sorted = [...spans].sort((a, b) => a.start - b.start);
  const result: Span[] = [];
  for (const span of sorted) {
    const last = result[result.length - 1];
    if (last && span.start <= last.end) {
      result[result.length - 1] = {
        start: last.start,
        end: Math.max(last.end, span.end),
      };
    } else {
      result.push(span);
    }
  }
  return result;
}

/**
 * The spans of a command, which must lie within the sheet, each from its
 * start to its end: a damaged command read back from storage fails.
 */
function checked<T extends Span>(command: {
  readonly direction: Direction;
  readonly spans: readonly T[];
}): readonly T[] {
  const limit = command.direction === 'rows' ? SHEET_ROWS : SHEET_COLUMNS;
  for (const { start, end } of command.spans) {
    if (!(
      Number.isInteger(start) &&
      start >= 0 &&
      start <= end &&
      end < limit
    )) {
      throw new Error(`Invalid span: ${start} to ${end}`);
    }
  }
  return command.spans;
}

/** The cells of whole rows or columns. */
function lineRange(direction: Direction, { start, end }: Span): Range {
  return direction === 'rows'
    ? { top: start, left: 0, bottom: end, right: SHEET_COLUMNS - 1 }
    : { top: 0, left: start, bottom: SHEET_ROWS - 1, right: end };
}

/**
 * A snapshot's cells and line sizes with rows or columns inserted and
 * deleted: `count` lines before `at` for each shift, or as many deleted
 * when negative, the last first. `moved` gives where a line goes.
 */
function movedLines(
  snapshot: SheetSnapshot,
  rows: boolean,
  shifts: readonly { at: number; count: number }[],
  limit: number,
  moved: (index: number) => number | null,
): SheetSnapshot {
  const splice = <T>(list: T[], empty: T): T[] => {
    let result = list;
    for (const { at, count } of shifts) {
      if (at >= result.length) continue;
      result =
        count > 0
          ? result
              .slice(0, at)
              .concat(Array<T>(count).fill(empty), result.slice(at))
          : result.slice(0, at).concat(result.slice(at - count));
    }
    return result.slice(0, limit);
  };
  const cells = <T>(columns: T[][], empty: T): T[][] =>
    rows ? columns.map((column) => splice(column, empty)) : splice(columns, []);
  const sizes = (pairs: [number, number][] | undefined) =>
    pairs?.flatMap(([index, size]): [number, number][] => {
      const place = moved(index);
      return place === null ? [] : [[place, size]];
    });
  const hidden = (indexes: number[] | undefined) =>
    indexes?.flatMap((index) => {
      const place = moved(index);
      return place === null ? [] : [place];
    });
  // A merge keeps the lines of it that stay, and grows with those inserted
  // inside it; one left with a single cell is no longer a merge.
  const span = (start: number, end: number): [number, number] | null => {
    let first = start;
    while (first <= end && moved(first) === null) first++;
    let last = end;
    while (last >= first && moved(last) === null) last--;
    if (first > last) return null;
    return [moved(first)!, moved(last)!];
  };
  const merges = snapshot.merges?.flatMap((merge) => {
    const lines = rows
      ? span(merge.top, merge.bottom)
      : span(merge.left, merge.right);
    if (!lines) return [];
    const [start, end] = lines;
    const result = rows
      ? { ...merge, top: start, bottom: end }
      : { ...merge, left: start, right: end };
    return result.top === result.bottom && result.left === result.right
      ? []
      : [result];
  });
  // A rule keeps the lines of its ranges that stay, and is dropped with them.
  const conditionalFormats = snapshot.conditionalFormats?.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((range) => {
      const lines = rows
        ? span(range.top, range.bottom)
        : span(range.left, range.right);
      if (!lines) return [];
      const [start, end] = lines;
      return [
        rows
          ? { ...range, top: start, bottom: end }
          : { ...range, left: start, right: end },
      ];
    });
    return ranges.length === 0 ? [] : [{ ...rule, ranges }];
  });
  return {
    ...snapshot,
    ...(merges && { merges }),
    ...(conditionalFormats && { conditionalFormats }),
    columns: cells(snapshot.columns, null),
    ...(snapshot.styleIds && { styleIds: cells(snapshot.styleIds, 0) }),
    ...(rows
      ? {
          rowHeights: sizes(snapshot.rowHeights),
          hiddenRows: hidden(snapshot.hiddenRows),
        }
      : {
          columnWidths: sizes(snapshot.columnWidths),
          hiddenColumns: hidden(snapshot.hiddenColumns),
        }),
  };
}

/** A sheet of the workbook, made from a snapshot of it. */
function sheetOf(snapshot: SheetSnapshot): Sheet {
  const sheet = emptySheet(
    snapshot.id,
    snapshot.name,
    snapshot.columns.map((column) => [...column]),
    (snapshot.styleIds ?? []).map((column) => [...column]),
  );
  if (snapshot.color !== undefined) sheet.color = snapshot.color;
  const lines = (
    sizes: [number, number][] | undefined,
    hidden: number[] | undefined,
  ): Lines =>
    sizes?.length || hidden?.length
      ? { sizes: new Map(sizes), hidden: new Set(hidden) }
      : NO_LINES;
  sheet.layout = {
    rows: lines(snapshot.rowHeights, snapshot.hiddenRows),
    columns: lines(snapshot.columnWidths, snapshot.hiddenColumns),
    ...(snapshot.frozen && { frozen: { ...snapshot.frozen } }),
  };
  if (snapshot.merges)
    sheet.merges = snapshot.merges.map((each) => ({ ...each }));
  if (snapshot.conditionalFormats)
    sheet.rules = [...snapshot.conditionalFormats];
  for (const [row, col, text] of snapshot.formulas) {
    const ast = parse(text.slice(1));
    if (!ast) throw new Error(`Invalid formula in snapshot: ${text}`);
    place(sheet, row, col, text, ast);
  }
  return sheet;
}

function snapshotOf(sheet: Sheet): SheetSnapshot {
  const { rows, columns, frozen } = sheet.layout;
  return {
    id: sheet.id,
    name: sheet.name,
    ...(sheet.color !== undefined && { color: sheet.color }),
    columns: Array.from(sheet.columns, (column) =>
      Array.from(column ?? [], (value) => value ?? null),
    ),
    formulas: Array.from(
      sheet.formulas.values(),
      ({ row, col, text }): [number, number, string] => [row, col, text],
    ),
    styleIds: Array.from(sheet.styleIds, (column) =>
      Array.from(column ?? [], (id) => id ?? 0),
    ),
    ...(rows.sizes.size > 0 && { rowHeights: [...rows.sizes] }),
    ...(columns.sizes.size > 0 && { columnWidths: [...columns.sizes] }),
    ...(rows.hidden.size > 0 && { hiddenRows: [...rows.hidden] }),
    ...(columns.hidden.size > 0 && { hiddenColumns: [...columns.hidden] }),
    ...(sheet.merges.length > 0 && { merges: [...sheet.merges] }),
    ...(frozen && { frozen: { ...frozen } }),
    ...(sheet.rules.length > 0 && { conditionalFormats: [...sheet.rules] }),
  };
}
