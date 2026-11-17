import {
  cellName,
  columnName,
  quoteSheetName,
  type Change,
  type Command,
  type Direction,
  type LineSpan,
  type NameProblem,
  type Border,
  type BorderSides,
  type ColumnFilter,
  type Criterion,
  type Range,
  type Search,
  type SortKey,
  type Span,
  type StyleChange,
  type TextWidth,
  type Workbook,
  SHEET_COLUMNS,
  SHEET_ROWS,
} from '../../engine/workbook';
import {
  complete,
  cycleReference,
  insertReference,
  pickable,
  startsOperand,
  type Edit,
} from './formula';
import { changedFormat, GENERAL, type FormatChange } from '../../engine/format';
import type { ClipboardContents } from '../platform';
import { fromClipboard, toClipboard } from './clipboard';
import type { EntryMode } from './keys';
import { navigate, type Navigation, type SheetContents } from './navigation';
import type { CellPosition } from './geometry';
import {
  covering,
  extendTo,
  overMerges,
  single,
  type Selection,
} from './selection';
import { TEXT_PADDING } from './text';

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2;

const START = { row: 0, col: 0 };

/** A change to the sheets of a workbook, or to a sheet's rules. */
export type SheetCommand = Extract<
  Command,
  { type: 'addSheet' | 'deleteSheet' | 'setSheet' | 'setRules' }
>;

/** An entry being typed into the active cell. */
export interface Entry {
  readonly text: string;
  readonly mode: EntryMode;
  /** Where it is typed: in the cell or in the formula bar. */
  readonly place: 'cell' | 'formulaBar';
  /**
   * The sheet of the cell it goes to, the active one there. Another sheet
   * can be shown while a formula is typed, to pick references on it.
   */
  readonly sheet: string;
  /** Where the caret is in the text. */
  readonly caret: number;
  /**
   * Whether the list of every function is open, as Insert function opens
   * it, until the text or the caret changes.
   */
  readonly allFunctions: boolean;
  /**
   * The cells of the reference just picked or pointed at, on the shown
   * sheet, which Shift and the arrows move from, or `null`.
   */
  readonly point: Selection | null;
}

export interface ViewState {
  /** The identifier of the sheet that is shown. */
  readonly sheet: string;
  readonly selection: Selection;
  /** The entry being typed, or `null` when there is none. */
  readonly entry: Entry | null;
  /** The size of the grid, where 1 is 100%. */
  readonly zoom: number;
  /** Whether End was pressed, so that the next arrow jumps as Ctrl+arrow. */
  readonly endMode: boolean;
  /** The cells copied or cut, to paste, or `null` when there are none. */
  readonly copied: Copied | null;
  /**
   * The range that the fill handle is dragged over, the selection
   * included, or `null` when it is not dragged beyond the selection.
   */
  readonly filling: Range | null;
  /**
   * The last fill of the fill handle, which can be redone another way,
   * until the workbook changes otherwise, or `null`.
   */
  readonly filled: Filled | null;
  /** The size a row or column is dragged to by its border, or `null`. */
  readonly resizing: Resizing | null;
  /**
   * The range whose formats the format painter gives the next range
   * selected, or `null` when it is put away.
   */
  readonly painter: { readonly sheet: string; readonly range: Range } | null;
  /**
   * The cell that find went to last, which the grid scrolls into view, or
   * `null`; a new object each time.
   */
  readonly reveal: CellPosition | null;
  /**
   * Whether the entry was rejected for breaking its cell's rule of data
   * validation, until it is tried again or abandoned.
   */
  readonly rejected: boolean;
}

/** What find looks for, on the shown sheet or in the whole workbook. */
export interface Finding extends Omit<Search, 'sheets' | 'only'> {
  readonly workbook?: boolean;
}

/** Which of the matches find went to, from 1, and how many there are. */
export interface Found {
  readonly index: number;
  readonly count: number;
}

/** A row or column dragged by its border to a size, in CSS pixels at 100%. */
export interface Resizing {
  readonly direction: Direction;
  readonly index: number;
  readonly size: number;
}

/** A fill of the fill handle, and whether it can extend a series. */
export interface Filled {
  readonly sheet: string;
  readonly source: Range;
  readonly target: Range;
  readonly series: boolean;
}

/** How a fill is redone. */
export type FillHow = 'copy' | 'series' | 'formats' | 'values';

/** A range copied or cut, which a moving border shows until it is forgotten. */
export interface Copied {
  readonly sheet: string;
  readonly range: Range;
  readonly cut: boolean;
}

/**
 * What the user is looking at and doing in a workbook: the sheet, the
 * selection, the entry being typed and the zoom. The grid edits it, and
 * the parts of the window around the grid read it.
 */
export class View {
  private current: ViewState;
  private readonly listeners = new Set<() => void>();
  /** The sheets in tab order as of the last change to the workbook. */
  private sheetIds: string[];
  /** Whether a paste is running, which keeps what was copied. */
  private pasting = false;
  /** Whether a fill is running, which keeps the last fill. */
  private filling = false;

  /**
   * Make it right after the workbook, so that it hears of each change
   * first: it leaves a deleted sheet before anything reads it.
   */
  constructor(private readonly workbook: Workbook) {
    this.sheetIds = workbook.sheets.map((sheet) => sheet.id);
    workbook.subscribe(() => {
      this.keepSheet();
      // Any other change forgets what was copied, as in Excel.
      if (!this.pasting && this.current.copied) this.change({ copied: null });
      if (!this.filling && this.current.filled) this.change({ filled: null });
    });
    this.current = {
      sheet: workbook.sheets[0]!.id,
      selection: single(START),
      entry: null,
      zoom: 1,
      endMode: false,
      copied: null,
      filling: null,
      filled: null,
      resizing: null,
      painter: null,
      reveal: null,
      rejected: false,
    };
  }

  /** The same object until the next change. */
  get state(): ViewState {
    return this.current;
  }

  /** Calls `listener` after every change, until the returned function is called. */
  readonly subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  /** Changes the selection, after committing an entry that is waiting. */
  select(selection: Selection): void {
    // A formula typed on another sheet goes to its own cell, and the
    // selection is that of the sheet shown.
    const { sheet } = this.current;
    if (!this.commit()) return;
    this.change({
      sheet,
      selection: overMerges(selection, this.workbook.merges(sheet)),
      endMode: false,
    });
  }

  /** Moves the selection by the keyboard, after committing a waiting entry. */
  navigate(navigation: Navigation): void {
    if (navigation.type === 'endMode') {
      this.change({ endMode: !this.current.endMode });
      return;
    }
    if (!this.commit()) return;
    const { selection } = this.current;
    this.select(navigate(selection, navigation, this.sheetContents()));
  }

  /**
   * Starts an entry in Enter mode that replaces the contents of the active
   * cell, or replaces the text of the entry there; the caret is at the end
   * unless given.
   */
  type(text: string, caret = text.length): void {
    const { entry, sheet } = this.current;
    this.change({
      entry: entry
        ? { ...entry, text, caret, allFunctions: false, point: null }
        : {
            text,
            mode: 'enter',
            place: 'cell',
            sheet,
            caret,
            allFunctions: false,
            point: null,
          },
    });
  }

  /** Moves the caret of the entry. */
  moveCaret(caret: number): void {
    const { entry } = this.current;
    if (entry && entry.caret !== caret) {
      this.change({
        entry: { ...entry, caret, allFunctions: false, point: null },
      });
    }
  }

  /**
   * Whether a click on the grid picks a reference for the formula being
   * typed rather than selecting: where a reference can go at the caret.
   */
  get canPick(): boolean {
    const { entry } = this.current;
    return entry !== null && pickable(entry.text, entry.caret);
  }

  /**
   * Whether the arrows point at a reference for the formula being typed
   * rather than commit: where an operand starts, or after the reference
   * they or a click put there, not after one typed.
   */
  get canPoint(): boolean {
    const { entry } = this.current;
    if (entry === null) return false;
    return entry.point !== null || startsOperand(entry.text, entry.caret);
  }

  /**
   * Puts a reference to a cell of the shown sheet at the caret of the
   * entry, in place of the one picked just before (`replace`, a click);
   * stretched from where that one started (`extend`, Shift or a drag);
   * or after it and a comma (`add`, Ctrl or Cmd). Without a reference
   * before the caret, each picks the cell alone.
   */
  pick(
    cell: CellPosition,
    mode: 'replace' | 'extend' | 'add' = 'replace',
  ): void {
    const point = this.current.entry?.point;
    this.pointAt(
      mode === 'extend' && point ? extendTo(point, cell) : single(cell),
      mode === 'add',
    );
  }

  /**
   * Moves the reference pointed at for the formula by the keyboard, or
   * extends it, as the arrows move the selection; the first starts from
   * the cell edited.
   */
  point(navigation: Navigation): void {
    const { entry, selection } = this.current;
    if (!entry) return;
    const from = entry.point ?? single(selection.active);
    this.pointAt(navigate(from, navigation, this.sheetContents()));
  }

  /** Cycles the reference at the caret between absolute and relative, as F4. */
  cycleReference(): void {
    const { entry } = this.current;
    const edit = entry && cycleReference(entry.text, entry.caret);
    if (edit) this.edited(edit);
  }

  /**
   * Puts a function in place of the start of its name typed at the caret,
   * or a defined name when `defined`.
   */
  complete(name: string, defined = false): void {
    const { entry } = this.current;
    if (entry) this.edited(complete(entry.text, entry.caret, name, defined));
  }

  /** The workbook's defined names, in alphabetical order. */
  get definedNames(): string[] {
    return this.workbook.names.map((each) => each.name);
  }

  /**
   * What the name box shows: the defined name of the range selected, when
   * one range is selected and a name refers to exactly it, or else the
   * active cell.
   */
  get selectionName(): string {
    const { sheet, selection } = this.current;
    const [only] = selection.ranges;
    const found =
      selection.ranges.length === 1 &&
      this.workbook.names.find(({ refersTo }) => {
        const place = this.workbook.locate(refersTo, sheet);
        return place?.sheet === sheet && sameRange(place.range, only!);
      });
    const { row, col } = selection.active;
    return found ? found.name : cellName(row, col);
  }

  /**
   * Goes to a cell address, a range or a defined name typed in the name
   * box, on its sheet; or names the last range selected with a new name,
   * as Excel does (R27). `invalid` when the text is none of these.
   */
  goTo(text: string): 'invalid' | null {
    if (!this.commit()) return null;
    const { sheet, selection } = this.current;
    const place = this.workbook.locate(text, sheet);
    if (place) {
      const { top, left } = place.range;
      this.change({
        sheet: place.sheet,
        selection: overMerges(
          covering([place.range]),
          this.workbook.merges(place.sheet),
        ),
        endMode: false,
        reveal: { row: top, col: left },
      });
      return null;
    }
    const name = text.trim();
    if (this.workbook.nameProblem(name) !== null) return 'invalid';
    const range = selection.ranges[selection.ranges.length - 1]!;
    this.workbook.execute({
      type: 'setNames',
      names: [
        ...this.workbook.names,
        { name, refersTo: this.absoluteReference(sheet, range) },
      ],
    });
    return null;
  }

  /**
   * Gives the defined name `old` a new name and the range of a reference
   * typed with or without its `=`, on the shown sheet unless it names one;
   * formulas that use it follow a new name. Gives why it cannot, the
   * reference being `reference`.
   */
  setName(
    old: string,
    name: string,
    reference: string,
  ): NameProblem | 'reference' | null {
    const problem = this.workbook.nameProblem(name, old);
    if (problem) return problem;
    const place = this.workbook.locate(
      reference.replace(/^\s*=/, ''),
      this.current.sheet,
    );
    if (!place) return 'reference';
    if (!this.commit()) return null;
    const folded = old.toLowerCase();
    this.workbook.execute({
      type: 'setNames',
      names: [
        ...this.workbook.names.filter(
          (each) => each.name.toLowerCase() !== folded,
        ),
        { name, refersTo: this.absoluteReference(place.sheet, place.range) },
      ],
      ...(name !== old && { renamed: { from: old, to: name } }),
    });
    return null;
  }

  /** Deletes a defined name; formulas that use it give `#NAME?`. */
  deleteName(name: string): void {
    if (!this.commit()) return;
    this.workbook.execute({
      type: 'setNames',
      names: this.workbook.names.filter((each) => each.name !== name),
    });
  }

  /** A range as a defined name refers to it, as in `Sheet1!$A$1:$B$2`. */
  private absoluteReference(sheetId: string, range: Range): string {
    const { top, left, bottom, right } = range;
    const col = (index: number) => `$${columnName(index)}`;
    const row = (index: number) => `$${index + 1}`;
    const text =
      top === 0 && bottom === SHEET_ROWS - 1
        ? `${col(left)}:${col(right)}`
        : left === 0 && right === SHEET_COLUMNS - 1
          ? `${row(top)}:${row(bottom)}`
          : top === bottom && left === right
            ? col(left) + row(top)
            : `${col(left)}${row(top)}:${col(right)}${row(bottom)}`;
    const { name } = this.workbook.sheets.find((each) => each.id === sheetId)!;
    return `${quoteSheetName(name)}!${text}`;
  }

  /**
   * Opens the list of every function at the caret of a formula, starting
   * one in the formula bar in place of anything else.
   */
  insertFunction(): void {
    if (!this.current.entry) this.edit('formulaBar');
    const entry = this.current.entry!;
    const formula = entry.text.startsWith('=');
    this.change({
      entry: {
        ...entry,
        text: formula ? entry.text : '=',
        caret: formula ? entry.caret : 1,
        allFunctions: true,
      },
    });
  }

  /**
   * Starts an entry in Edit mode that holds the contents of the active cell,
   * or moves the entry there to Edit mode, in the cell or the formula bar.
   */
  edit(place: Entry['place']): void {
    const { sheet, selection, entry } = this.current;
    const { row, col } = selection.active;
    const text = entry?.text ?? this.workbook.cell(sheet, row, col).input;
    this.change({
      entry: {
        text,
        mode: 'edit',
        place,
        sheet: entry?.sheet ?? sheet,
        caret: entry?.caret ?? text.length,
        allFunctions: entry?.allFunctions ?? false,
        point: entry?.point ?? null,
      },
    });
  }

  /**
   * Writes the entry to its cell, and shows that cell's sheet again, and
   * gives whether there is no entry left. An entry that breaks its cell's
   * rule of data validation stays, and is `rejected` until it is tried
   * again or abandoned (R26).
   */
  commit(): boolean {
    const { selection, entry } = this.current;
    if (entry === null) return true;
    const { sheet, text: input } = entry;
    const { row, col } = selection.active;
    if (!this.workbook.allows(sheet, row, col, input)) {
      this.change({ rejected: true });
      return false;
    }
    this.change({ entry: null, sheet });
    this.workbook.execute({ type: 'setCell', sheet, row, col, input });
    return true;
  }

  /** Goes back to an entry that was rejected, in Edit mode, as Excel does. */
  retry(): void {
    const { entry } = this.current;
    this.change({
      rejected: false,
      entry: entry && { ...entry, mode: 'edit' },
    });
  }

  /**
   * Gives the selected cells a rule of data validation, or none, after
   * committing an entry (R26).
   */
  validate(criterion: Criterion | null): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    this.workbook.execute({
      type: 'validate',
      sheet,
      ranges: selection.ranges,
      criterion,
    });
  }

  /** Sets the active cell to an item of its list, abandoning an entry. */
  choose(item: string): void {
    if (this.current.entry) this.cancel();
    const { sheet, selection } = this.current;
    const { row, col } = selection.active;
    this.workbook.execute({ type: 'setCell', sheet, row, col, input: item });
  }

  /**
   * Shows another sheet, at A1, after committing a waiting entry; or
   * keeps a formula that can take a reference going in the formula bar,
   * to pick one on that sheet.
   */
  showSheet(sheet: string): void {
    const { entry } = this.current;
    if (entry && this.canPick) {
      this.change({
        sheet,
        entry: { ...entry, place: 'formulaBar', point: null },
      });
      return;
    }
    if (!this.commit()) return;
    this.change({ sheet, selection: single(START), endMode: false });
  }

  /**
   * Makes a change to the sheets or their conditional formats, after
   * committing a waiting entry, and shows a sheet it adds.
   */
  execute(command: SheetCommand): void {
    if (!this.commit()) return;
    const change = this.workbook.execute(command);
    if (command.type === 'addSheet') this.showSheet(change.sheet);
  }

  /** Changes the style of the selected cells, after committing an entry. */
  format(style: StyleChange): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    this.workbook.execute({
      type: 'setStyle',
      sheet,
      ranges: selection.ranges,
      style,
    });
  }

  /**
   * Gives the selected cells the format of the active cell with one more or
   * one fewer decimal place, or the thousands separator turned on or off,
   * as Excel does; nothing changes when that of the active cell would not.
   */
  adjustFormat(change: FormatChange): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    const { row, col } = selection.active;
    const { value, style } = this.workbook.cell(sheet, row, col);
    const code = style.numberFormat ?? GENERAL;
    const changed = changedFormat(code, value, change);
    // A date, a time or text keeps its format, and so do the other cells.
    if (changed !== code) this.format({ numberFormat: changed });
  }

  /**
   * Turns a property on for the selected cells, or off when the active
   * cell has it, as Excel does.
   */
  toggle(property: 'bold' | 'italic' | 'underline' | 'strike' | 'wrap'): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    const { row, col } = selection.active;
    const on = this.workbook.cell(sheet, row, col).style[property] === true;
    this.format({ [property]: on ? null : true });
  }

  /** Puts a border on sides of the selected ranges, or takes it off. */
  setBorders(sides: BorderSides, border: Border | null): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    this.workbook.execute({
      type: 'setBorders',
      sheet,
      ranges: selection.ranges,
      sides,
      border,
    });
  }

  /**
   * Whether a selected range spans more than one cell and is not already
   * merged into one, to merge.
   */
  get canMerge(): boolean {
    const { sheet, selection } = this.current;
    const merges = this.workbook.merges(sheet);
    return selection.ranges.some(
      (range) =>
        (range.top !== range.bottom || range.left !== range.right) &&
        !merges.some(
          (merge) =>
            merge.top === range.top &&
            merge.left === range.left &&
            merge.bottom === range.bottom &&
            merge.right === range.right,
        ),
    );
  }

  /** Whether the selection holds a merged cell, to unmerge. */
  get canUnmerge(): boolean {
    const { sheet, selection } = this.current;
    return this.workbook
      .merges(sheet)
      .some((merge) =>
        selection.ranges.some(
          (range) =>
            range.top <= merge.bottom &&
            merge.top <= range.bottom &&
            range.left <= merge.right &&
            merge.left <= range.right,
        ),
      );
  }

  /** Whether there are rows above or columns left of the active cell. */
  get canFreeze(): boolean {
    const { active } = this.current.selection;
    return active.row > 0 || active.col > 0;
  }

  /** Whether the shown sheet has frozen rows or columns. */
  get canUnfreeze(): boolean {
    return this.workbook.layout(this.current.sheet).frozen !== undefined;
  }

  /**
   * Freezes the rows above and the columns left of the active cell, the
   * top row, or the first column, in place of what was frozen, after
   * committing.
   */
  freeze(what: 'panes' | 'topRow' | 'firstColumn'): void {
    const { active } = this.current.selection;
    this.setFrozen(
      what === 'panes'
        ? { rows: active.row, cols: active.col }
        : what === 'topRow'
          ? { rows: 1, cols: 0 }
          : { rows: 0, cols: 1 },
    );
  }

  /** Unfreezes the rows and columns of the shown sheet, after committing. */
  unfreeze(): void {
    this.setFrozen({ rows: 0, cols: 0 });
  }

  /**
   * The range a sort sorts: the one range selected, its whole rows or
   * columns cut down to the cells in use, or `null` when that is a single
   * cell or none, there are several ranges, or it holds merged cells,
   * which Excel refuses to sort when they differ in size.
   */
  get sortRange(): Range | null {
    const { sheet, selection } = this.current;
    if (selection.ranges.length !== 1) return null;
    const { rows, cols } = this.workbook.extent(sheet);
    const selected = selection.ranges[0]!;
    const range = {
      ...selected,
      bottom: Math.min(selected.bottom, rows - 1),
      right: Math.min(selected.right, cols - 1),
    };
    const merged = this.workbook
      .merges(sheet)
      .some(
        (merge) =>
          range.top <= merge.bottom &&
          merge.top <= range.bottom &&
          range.left <= merge.right &&
          merge.left <= range.right,
      );
    const empty = range.bottom < range.top || range.right < range.left;
    const one = range.top === range.bottom && range.left === range.right;
    return empty || one || merged ? null : range;
  }

  /**
   * Sorts the rows of `sortRange` by columns, leaving the first row in
   * place when it holds `headers`, after committing an entry.
   */
  sort(keys: readonly SortKey[], headers = false): void {
    if (!this.commit()) return;
    const range = this.sortRange;
    if (!range) return;
    const { sheet } = this.current;
    this.workbook.execute({
      type: 'sort',
      sheet,
      range,
      keys,
      ...(headers && { headers }),
    });
  }

  /**
   * Sorts `sortRange` by the column of the active cell. One right of the
   * cells in use is empty, and sorting by it changes nothing, as in Excel.
   */
  sortByActive(descending: boolean): void {
    const { col } = this.current.selection.active;
    if (col > (this.sortRange?.right ?? -1)) return;
    this.sort([{ col, ...(descending && { descending }) }]);
  }

  /** Whether the Filter button can turn the sheet's filter on or off. */
  get canFilter(): boolean {
    const { sheet } = this.current;
    return this.workbook.filter(sheet) !== null || this.filterRange() !== null;
  }

  /**
   * Turns the sheet's filter off, or on over `filterRange` with nothing
   * filtered yet, after committing an entry.
   */
  toggleFilter(): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const range = this.filterRange();
    if (this.workbook.filter(sheet)) {
      this.workbook.execute({ type: 'filter', sheet, filter: null });
    } else if (range) {
      this.workbook.execute({
        type: 'filter',
        sheet,
        filter: { range, columns: [] },
      });
    }
  }

  /**
   * Filters a column of the sheet's filter by `column`, in place of what it
   * filtered by.
   */
  filterBy(column: ColumnFilter): void {
    this.refilter(column.col, column);
  }

  /** Stops filtering the column `col` of the sheet's filter. */
  unfilter(col: number): void {
    this.refilter(col, null);
  }

  private refilter(col: number, column: ColumnFilter | null): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const filter = this.workbook.filter(sheet);
    if (!filter) return;
    const columns = filter.columns.filter((each) => each.col !== col);
    this.workbook.execute({
      type: 'filter',
      sheet,
      filter: { ...filter, columns: column ? [...columns, column] : columns },
    });
  }

  /** Whether the sheet's filter filters any column. */
  get canClearFilter(): boolean {
    const { sheet } = this.current;
    return (this.workbook.filter(sheet)?.columns.length ?? 0) > 0;
  }

  /** Shows every row of the sheet's filter again, keeping the filter. */
  clearFilter(): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const filter = this.workbook.filter(sheet);
    if (!filter || filter.columns.length === 0) return;
    this.workbook.execute({
      type: 'filter',
      sheet,
      filter: { ...filter, columns: [] },
    });
  }

  /** Merges each selected range into one cell, after committing. */
  merge(): void {
    this.mergeCommand('merge');
  }

  /** Unmerges the merged cells of the selection, after committing. */
  unmerge(): void {
    this.mergeCommand('unmerge');
  }

  /**
   * Takes up the format painter for the selected range, the last one, or
   * puts it away when it is taken up.
   */
  startPainter(): void {
    if (!this.commit()) return;
    const { sheet, painter } = this.current;
    this.change({
      painter: painter ? null : { sheet, range: this.lastRange() },
    });
  }

  /**
   * Gives the selected range, the last one, the formats of the painter's
   * range, as a paste of formats does, and puts the painter away. Does
   * nothing without the painter.
   */
  paint(): void {
    const { painter, sheet } = this.current;
    if (!painter) return;
    this.change({ painter: null });
    const change = this.workbook.execute({
      type: 'paste',
      sheet,
      target: this.lastRange(),
      source: painter,
      only: 'formats',
    });
    this.change({ selection: covering([change.ranges[0]!]), endMode: false });
  }

  /** Empties the selected cells. */
  clear(): void {
    const { sheet, selection } = this.current;
    this.workbook.execute({ type: 'clear', sheet, ranges: selection.ranges });
  }

  /**
   * Copies the selected range, the last one, to paste, after committing a
   * waiting entry, and gives its cells as they are shown, for the system
   * clipboard; `null` when the entry is rejected.
   */
  copy(): ClipboardContents | null {
    return this.copyRange(false);
  }

  /** Cuts the selected range, to move it with the next paste, as `copy`. */
  cut(): ClipboardContents | null {
    return this.copyRange(true);
  }

  /**
   * Pastes the cells copied or cut at the selected range, after committing
   * a waiting entry, and selects them. `only` pastes the values, the
   * formulas or the formats of a copy. Cut cells are pasted once.
   */
  paste(only?: 'values' | 'formulas' | 'formats'): void {
    const { copied } = this.current;
    if (!copied) return;
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    const target = selection.ranges[selection.ranges.length - 1]!;
    this.pasting = true;
    try {
      const change = this.workbook.execute({
        type: 'paste',
        sheet,
        target,
        source: { sheet: copied.sheet, range: copied.range },
        ...(copied.cut ? { cut: true } : only && { only }),
      });
      this.change({
        selection: covering([change.ranges[0]!]),
        endMode: false,
        ...(copied.cut && { copied: null }),
      });
    } finally {
      this.pasting = false;
    }
  }

  /**
   * Pastes text from the system clipboard at the selected range. Text that
   * `copy` or `cut` gave for the cells copied pastes those cells; other
   * text, from another application, sets the cells from the top left of
   * the range as if typed, a row for each line and a cell for each tab.
   */
  pasteText(text: string): void {
    const { copied } = this.current;
    // Windows ends lines with CR LF on the clipboard.
    const lines = (text: string) => text.replace(/\r\n?/g, '\n');
    if (copied && lines(text) === lines(this.contents(copied).text)) {
      this.paste();
      return;
    }
    if (text === '') return;
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    const { top: row, left: col } =
      selection.ranges[selection.ranges.length - 1]!;
    const change = this.workbook.execute({
      type: 'setCells',
      sheet,
      row,
      col,
      inputs: fromClipboard(text),
    });
    this.change({ selection: covering(change.ranges), endMode: false });
  }

  /**
   * Drags the fill handle of the selected range, the last one, to a cell:
   * the range to fill extends the selection to it, down, up, right or left,
   * whichever way the cell is furthest beyond it.
   */
  previewFill({ row, col }: CellPosition): void {
    const range = this.lastRange();
    const rows = Math.max(row - range.bottom, range.top - row, 0);
    const cols = Math.max(col - range.right, range.left - col, 0);
    let filling: Range | null = null;
    if (rows > 0 && rows >= cols) {
      filling =
        row > range.bottom ? { ...range, bottom: row } : { ...range, top: row };
    } else if (cols > 0) {
      filling =
        col > range.right ? { ...range, right: col } : { ...range, left: col };
    }
    this.change({ filling });
  }

  /**
   * Fills the range the fill handle was dragged over from the selection,
   * extending its series, and selects it. With `swap`, as a drag with Ctrl
   * does, it copies cells that would extend a series, and otherwise counts
   * one number on by one.
   */
  fill(swap = false): void {
    const { filling, sheet } = this.current;
    if (!filling) return;
    this.change({ filling: null });
    const source = this.lastRange();
    const command = { type: 'fill', sheet, source, target: filling } as const;
    const how = swap
      ? this.workbook.fillsSeries(command)
        ? 'copy'
        : 'series'
      : undefined;
    this.fillFrom(source, filling, how);
  }

  /**
   * Undoes the last fill and does it again another way, as one change to
   * undo.
   */
  refill(how: FillHow): void {
    const { filled } = this.current;
    if (!filled) return;
    this.filling = true;
    try {
      this.workbook.undo();
      const { sheet, source, target } = filled;
      this.workbook.execute({ type: 'fill', sheet, source, target, how });
    } finally {
      this.filling = false;
    }
  }

  /**
   * Fills down from the selected range, as a double-click on the fill
   * handle does, as far as the column on its left has data, or else the
   * column on its right.
   */
  autoFill(): void {
    const range = this.lastRange();
    const { sheet } = this.current;
    const filled = (row: number, col: number) =>
      col >= 0 && this.workbook.cell(sheet, row, col).value !== null;
    const below = range.bottom + 1;
    const col = filled(below, range.left - 1)
      ? range.left - 1
      : range.right + 1;
    let bottom = range.bottom;
    while (bottom < SHEET_ROWS - 1 && filled(bottom + 1, col)) bottom++;
    if (bottom > range.bottom) this.fillFrom(range, { ...range, bottom });
  }

  /**
   * Copies the first row of the selected range into the rest of it, or the
   * row above a range of one row into it, as Ctrl+D does.
   */
  fillDown(): void {
    const range = this.lastRange();
    const top = range.top === range.bottom ? range.top - 1 : range.top;
    if (top < 0) return;
    this.copyInto({ ...range, top, bottom: top }, { ...range, top: top + 1 });
  }

  /** Copies the first column into the rest, as `fillDown` does rows. */
  fillRight(): void {
    const range = this.lastRange();
    const left = range.left === range.right ? range.left - 1 : range.left;
    if (left < 0) return;
    this.copyInto(
      { ...range, left, right: left },
      { ...range, left: left + 1 },
    );
  }

  /**
   * Inserts before each selected range as many rows or columns as it
   * spans, after committing a waiting entry, and selects them.
   */
  insertLines(direction: Direction): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const spans = this.selectedSpans(direction);
    const change = this.workbook.execute({
      type: 'insertLines',
      sheet,
      direction,
      spans,
    });
    this.change({ selection: covering(change.ranges), endMode: false });
  }

  /** Deletes the rows or columns of the selection, after committing. */
  deleteLines(direction: Direction): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const spans = this.selectedSpans(direction);
    this.workbook.execute({ type: 'deleteLines', sheet, direction, spans });
  }

  /** Hides, or shows again, the rows or columns of the selection. */
  hideLines(direction: Direction, hidden: boolean): void {
    const spans = this.selectedSpans(direction).map((span) => ({
      ...span,
      hidden,
    }));
    this.setLines(direction, spans);
  }

  /** Drags the border of a row or column to give it a size. */
  previewResize(direction: Direction, index: number, size: number): void {
    this.change({ resizing: { direction, index, size } });
  }

  /**
   * Gives the size a border was dragged to to its row or column, or to
   * every whole line selected with it; dragged to nothing, they are
   * hidden, as in Excel.
   */
  resize(): void {
    const { resizing } = this.current;
    if (!resizing) return;
    this.change({ resizing: null });
    const { direction, index } = resizing;
    const size = Math.round(resizing.size);
    this.setLines(
      direction,
      this.linesWith(direction, index).map((span) =>
        size > 0 ? { ...span, size, hidden: false } : { ...span, hidden: true },
      ),
    );
  }

  /**
   * Fits a row or column, or every whole line selected with it, to its
   * contents, as a double-click on its border does: a column to its
   * widest text shown, as `measure` gives it at 100%, and a row to the
   * height its text needs.
   */
  fitLines(direction: Direction, index: number, measure: TextWidth): void {
    const { sheet } = this.current;
    const { rows } = this.workbook.extent(sheet);
    // The text of hidden rows is not shown, so it does not count.
    const { hidden } = this.workbook.layout(sheet).rows;
    const spans: LineSpan[] = [];
    for (const { start, end } of this.linesWith(direction, index)) {
      for (let line = start; line <= end; line++) {
        let widest = 0;
        for (let row = 0; direction === 'columns' && row < rows; row++) {
          if (hidden.has(row)) continue;
          const { display, style } = this.workbook.cell(sheet, row, line);
          if (display !== '' && !style.wrap) {
            widest = Math.max(widest, measure(display, style));
          }
        }
        const size = widest > 0 ? Math.ceil(widest + 2 * TEXT_PADDING) : null;
        spans.push({ start: line, end: line, size });
      }
    }
    this.setLines(direction, spans);
  }

  /**
   * Goes to the first match after the active cell, row by row and then
   * sheet by sheet, or back to the first; `null` when there is none (R25).
   */
  findNext(finding: Finding): Found | null {
    if (!this.commit()) return null;
    const found = this.workbook.find(this.search(finding));
    if (found.length === 0) return null;
    const order = this.workbook.sheets.map((sheet) => sheet.id);
    const { sheet, selection } = this.current;
    const from = [
      order.indexOf(sheet),
      selection.active.row,
      selection.active.col,
    ];
    const after = found.findIndex((place) => {
      const at = [order.indexOf(place.sheet), place.row, place.col];
      const index = at.findIndex((value, part) => value !== from[part]);
      return index >= 0 && at[index]! > from[index]!;
    });
    const index = Math.max(after, 0);
    const { row, col } = found[index]!;
    this.change({
      sheet: found[index]!.sheet,
      selection: overMerges(
        single({ row, col }),
        this.workbook.merges(found[index]!.sheet),
      ),
      endMode: false,
      reveal: { row, col },
    });
    return { index: index + 1, count: found.length };
  }

  /**
   * Replaces the search in the active cell when it holds it, looking in
   * formulas, and goes to the next match.
   */
  replace(finding: Finding, replacement: string): Found | null {
    if (!this.commit()) return null;
    const inFormulas = { ...finding, lookIn: 'formulas' } as const;
    const { sheet, selection } = this.current;
    const search = {
      ...this.search(inFormulas),
      only: { sheet, ...selection.active },
    };
    if (this.workbook.find(search).length > 0) {
      this.workbook.execute({ type: 'replace', search, replacement });
    }
    return this.findNext(inFormulas);
  }

  /** Replaces every match, looking in formulas, and gives how many cells. */
  replaceAll(finding: Finding, replacement: string): number {
    if (!this.commit()) return 0;
    const search = this.search({ ...finding, lookIn: 'formulas' });
    const count = this.workbook.find(search).length;
    if (count > 0)
      this.workbook.execute({ type: 'replace', search, replacement });
    return count;
  }

  /** The search of a finding, over the sheets it covers in tab order. */
  private search({ workbook, ...search }: Finding): Search {
    const sheets = workbook
      ? this.workbook.sheets.map((sheet) => sheet.id)
      : [this.current.sheet];
    return { ...search, sheets };
  }

  /**
   * Undoes the last change, abandoning the entry, and selects the cells it
   * changed.
   */
  undo(): void {
    this.reverse(() => this.workbook.undo());
  }

  /** Redoes the last change undone, as `undo` undoes. */
  redo(): void {
    this.reverse(() => this.workbook.redo());
  }

  /**
   * Abandons the entry, and shows its cell's sheet again; without an entry,
   * forgets what was copied and puts the format painter away.
   */
  cancel(): void {
    const { entry, sheet } = this.current;
    this.change(
      entry
        ? { entry: null, sheet: entry.sheet, rejected: false }
        : { entry: null, sheet, copied: null, painter: null },
    );
  }

  /** Sets the zoom, kept between 50% and 200% and in whole percents. */
  zoomTo(zoom: number): void {
    const percent = Math.round(
      Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom)) * 100,
    );
    this.change({ zoom: percent / 100 });
  }

  private setLines(direction: Direction, spans: LineSpan[]): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    this.workbook.execute({ type: 'setLines', sheet, direction, spans });
  }

  /** The rows or columns that the ranges of the selection span. */
  /**
   * The rows or columns that the whole lines selected span, as a header's
   * menu selects them, or else those of every range selected.
   */
  private selectedSpans(direction: Direction): Span[] {
    const whole = this.wholeSpans(direction);
    return whole.length > 0
      ? whole
      : this.current.selection.ranges.map((range) =>
          direction === 'rows'
            ? { start: range.top, end: range.bottom }
            : { start: range.left, end: range.right },
        );
  }

  /** The spans of the selected ranges that are whole rows or columns. */
  private wholeSpans(direction: Direction): Span[] {
    const rows = direction === 'rows';
    return this.current.selection.ranges
      .filter((range) =>
        rows
          ? range.left === 0 && range.right === SHEET_COLUMNS - 1
          : range.top === 0 && range.bottom === SHEET_ROWS - 1,
      )
      .map((range) =>
        rows
          ? { start: range.top, end: range.bottom }
          : { start: range.left, end: range.right },
      );
  }

  /** Whether any of the rows or columns selected is hidden, to show again. */
  hasHidden(direction: Direction): boolean {
    const { hidden } = this.workbook.layout(this.current.sheet)[direction];
    return this.selectedSpans(direction).some(({ start, end }) =>
      [...hidden].some((index) => index >= start && index <= end),
    );
  }

  /**
   * A row or column, or, when it is among whole lines selected, all of
   * those: the lines a drag or double-click on its border changes.
   */
  private linesWith(direction: Direction, index: number): Span[] {
    const spans = this.wholeSpans(direction);
    return spans.some((span) => index >= span.start && index <= span.end)
      ? spans
      : [{ start: index, end: index }];
  }

  private setFrozen({ rows, cols }: { rows: number; cols: number }): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    this.workbook.execute({ type: 'freeze', sheet, rows, cols });
  }

  private mergeCommand(type: 'merge' | 'unmerge'): void {
    if (!this.commit()) return;
    const { sheet, selection } = this.current;
    this.workbook.execute({ type, sheet, ranges: selection.ranges });
  }

  private lastRange(): Range {
    const { ranges } = this.current.selection;
    return ranges[ranges.length - 1]!;
  }

  private fillFrom(source: Range, target: Range, how?: FillHow): void {
    if (!this.commit()) return;
    const { sheet } = this.current;
    const command = { type: 'fill', sheet, source, target } as const;
    this.filling = true;
    try {
      this.workbook.execute({ ...command, ...(how && { how }) });
    } finally {
      this.filling = false;
    }
    const series = this.workbook.fillsSeries({ ...command, how: 'series' });
    this.change({
      selection: covering([target]),
      endMode: false,
      filled: { sheet, source, target, series },
    });
  }

  /** Pastes a range copied into another of whole copies of it. */
  private copyInto(source: Range, target: Range): void {
    const { sheet } = this.current;
    this.workbook.execute({
      type: 'paste',
      sheet,
      target,
      source: { sheet, range: source },
    });
  }

  private copyRange(cut: boolean): ClipboardContents | null {
    if (!this.commit()) return null;
    const { sheet, selection } = this.current;
    const range = selection.ranges[selection.ranges.length - 1]!;
    const copied = { sheet, range, cut };
    this.change({ copied });
    return this.contents(copied);
  }

  /** The cells of a range as they are shown, for the system clipboard. */
  private contents({ sheet, range }: Copied): ClipboardContents {
    const { top, left, bottom, right } = range;
    const rows: string[][] = [];
    for (let row = top; row <= bottom; row++) {
      const cells: string[] = [];
      for (let col = left; col <= right; col++) {
        cells.push(this.workbook.cell(sheet, row, col).display);
      }
      rows.push(cells);
    }
    return toClipboard(rows);
  }

  private edited({ text, caret }: Edit, point: Selection | null = null): void {
    const entry = {
      ...this.current.entry!,
      text,
      caret,
      allFunctions: false,
      point,
    };
    this.change({ entry });
  }

  /** Puts a reference to the last range of `point` at the caret. */
  private pointAt(point: Selection, add = false): void {
    const { entry, sheet } = this.current;
    if (!entry) return;
    const { top, left, bottom, right } = point.ranges[point.ranges.length - 1]!;
    const name = this.workbook.sheets.find((each) => each.id === sheet)!.name;
    const prefix = sheet === entry.sheet ? '' : `${quoteSheetName(name)}!`;
    const end =
      top === bottom && left === right ? '' : `:${cellName(bottom, right)}`;
    const reference = prefix + cellName(top, left) + end;
    this.edited(
      insertReference(entry.text, entry.caret, reference, add),
      point,
    );
  }

  /**
   * The range a filter is turned on over: the one range selected, or the
   * data around the active cell when one cell is selected, cut down to the
   * cells in use; `null` when there is none.
   */
  private filterRange(): Range | null {
    const { sheet, selection } = this.current;
    if (selection.ranges.length !== 1) return null;
    let selected = selection.ranges[0]!;
    if (selected.top === selected.bottom && selected.left === selected.right) {
      const { row, col } = selection.active;
      if (this.workbook.cell(sheet, row, col).value === null) return null;
      selected = navigate(selection, { type: 'all' }, this.sheetContents())
        .ranges[0]!;
    }
    const { rows, cols } = this.workbook.extent(sheet);
    const range = {
      ...selected,
      bottom: Math.min(selected.bottom, rows - 1),
      right: Math.min(selected.right, cols - 1),
    };
    return range.bottom < range.top || range.right < range.left ? null : range;
  }

  /** What navigation needs to know of the shown sheet. */
  private sheetContents(): SheetContents {
    const { workbook } = this;
    const { sheet } = this.current;
    return {
      isEmpty: (row, col) => workbook.cell(sheet, row, col).value === null,
      extent: () => workbook.extent(sheet),
      merges: workbook.merges(sheet),
      hidden: {
        rows: workbook.layout(sheet).rows.hidden,
        cols: workbook.layout(sheet).columns.hidden,
      },
    };
  }

  private reverse(run: () => Change | null): void {
    this.cancel();
    const count = this.workbook.sheets.length;
    const change = run();
    if (!change) return;
    // A deleted sheet leaves the shown one, unless it was the shown one,
    // which `keepSheet` has already left.
    if (this.workbook.sheets.length < count) return;
    if (change.ranges.length > 0) {
      this.change({
        sheet: change.sheet,
        selection: covering(change.ranges),
        endMode: false,
      });
    } else if (change.sheet !== this.current.sheet) {
      this.showSheet(change.sheet);
    }
  }

  /** When the shown sheet is deleted, shows the one that takes its place. */
  private keepSheet(): void {
    const ids = this.workbook.sheets.map((sheet) => sheet.id);
    const { sheet } = this.current;
    if (!ids.includes(sheet)) {
      const index = Math.min(this.sheetIds.indexOf(sheet), ids.length - 1);
      this.change({
        sheet: ids[index]!,
        selection: single(START),
        entry: null,
        endMode: false,
      });
    }
    this.sheetIds = ids;
  }

  private change(change: Partial<ViewState>): void {
    this.current = { ...this.current, ...change };
    for (const listener of this.listeners) listener();
  }
}

function sameRange(a: Range, b: Range): boolean {
  return (
    a.top === b.top &&
    a.left === b.left &&
    a.bottom === b.bottom &&
    a.right === b.right
  );
}
