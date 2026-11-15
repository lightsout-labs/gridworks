import {
  cellName,
  quoteSheetName,
  type Change,
  type Command,
  type Range,
  type Workbook,
  SHEET_ROWS,
} from '../../engine/workbook';
import {
  complete,
  cycleReference,
  insertReference,
  pickable,
  type Edit,
} from './formula';
import type { ClipboardContents } from '../platform';
import { fromClipboard, toClipboard } from './clipboard';
import type { EntryMode } from './keys';
import { navigate, type Navigation, type SheetContents } from './navigation';
import type { CellPosition } from './geometry';
import { covering, single, type Selection } from './selection';

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2;

const START = { row: 0, col: 0 };

/** A change to the sheets of a workbook. */
export type SheetCommand = Extract<
  Command,
  { type: 'addSheet' | 'deleteSheet' | 'setSheet' }
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
}

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
    });
    this.current = {
      sheet: workbook.sheets[0]!.id,
      selection: single(START),
      entry: null,
      zoom: 1,
      endMode: false,
      copied: null,
      filling: null,
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
    this.commit();
    this.change({ sheet, selection, endMode: false });
  }

  /** Moves the selection by the keyboard, after committing a waiting entry. */
  navigate(navigation: Navigation): void {
    if (navigation.type === 'endMode') {
      this.change({ endMode: !this.current.endMode });
      return;
    }
    this.commit();
    const { workbook } = this;
    const { sheet, selection } = this.current;
    const contents: SheetContents = {
      isEmpty: (row, col) => workbook.cell(sheet, row, col).value === null,
      extent: () => workbook.extent(sheet),
    };
    this.select(navigate(selection, navigation, contents));
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
        ? { ...entry, text, caret, allFunctions: false }
        : {
            text,
            mode: 'enter',
            place: 'cell',
            sheet,
            caret,
            allFunctions: false,
          },
    });
  }

  /** Moves the caret of the entry. */
  moveCaret(caret: number): void {
    const { entry } = this.current;
    if (entry && entry.caret !== caret) {
      this.change({ entry: { ...entry, caret, allFunctions: false } });
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
   * Puts a reference to a range of the shown sheet at the caret of the
   * entry, in place of the one picked just before.
   */
  pick({ top, left, bottom, right }: Range): void {
    const { entry, sheet } = this.current;
    if (!entry) return;
    const name = this.workbook.sheets.find((each) => each.id === sheet)!.name;
    const prefix = sheet === entry.sheet ? '' : `${quoteSheetName(name)}!`;
    const end =
      top === bottom && left === right ? '' : `:${cellName(bottom, right)}`;
    const reference = prefix + cellName(top, left) + end;
    this.edited(insertReference(entry.text, entry.caret, reference));
  }

  /** Cycles the reference at the caret between absolute and relative, as F4. */
  cycleReference(): void {
    const { entry } = this.current;
    const edit = entry && cycleReference(entry.text, entry.caret);
    if (edit) this.edited(edit);
  }

  /** Puts a function in place of the start of its name typed at the caret. */
  complete(name: string): void {
    const { entry } = this.current;
    if (entry) this.edited(complete(entry.text, entry.caret, name));
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
      },
    });
  }

  /**
   * Writes the entry to its cell, and shows that cell's sheet again. Does
   * nothing without an entry.
   */
  commit(): void {
    const { selection, entry } = this.current;
    if (entry === null) return;
    const { sheet, text: input } = entry;
    this.change({ entry: null, sheet });
    const { row, col } = selection.active;
    this.workbook.execute({ type: 'setCell', sheet, row, col, input });
  }

  /**
   * Shows another sheet, at A1, after committing a waiting entry; or
   * keeps a formula that can take a reference going in the formula bar,
   * to pick one on that sheet.
   */
  showSheet(sheet: string): void {
    const { entry } = this.current;
    if (entry && this.canPick) {
      this.change({ sheet, entry: { ...entry, place: 'formulaBar' } });
      return;
    }
    this.commit();
    this.change({ sheet, selection: single(START), endMode: false });
  }

  /**
   * Makes a change to the sheets, after committing a waiting entry, and
   * shows a sheet it adds.
   */
  execute(command: SheetCommand): void {
    this.commit();
    const change = this.workbook.execute(command);
    if (command.type === 'addSheet') this.showSheet(change.sheet);
  }

  /** Empties the selected cells. */
  clear(): void {
    const { sheet, selection } = this.current;
    this.workbook.execute({ type: 'clear', sheet, ranges: selection.ranges });
  }

  /**
   * Copies the selected range, the last one, to paste, after committing a
   * waiting entry, and gives its cells as they are shown, for the system
   * clipboard.
   */
  copy(): ClipboardContents {
    return this.copyRange(false);
  }

  /** Cuts the selected range, to move it with the next paste, as `copy`. */
  cut(): ClipboardContents {
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
    this.commit();
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
    this.commit();
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
   * extending its series, and selects it.
   */
  fill(): void {
    const { filling } = this.current;
    if (!filling) return;
    this.change({ filling: null });
    this.fillFrom(this.lastRange(), filling);
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
   * forgets what was copied.
   */
  cancel(): void {
    const { entry, sheet } = this.current;
    this.change(
      entry
        ? { entry: null, sheet: entry.sheet }
        : { entry: null, sheet, copied: null },
    );
  }

  /** Sets the zoom, kept between 50% and 200% and in whole percents. */
  zoomTo(zoom: number): void {
    const percent = Math.round(
      Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom)) * 100,
    );
    this.change({ zoom: percent / 100 });
  }

  private lastRange(): Range {
    const { ranges } = this.current.selection;
    return ranges[ranges.length - 1]!;
  }

  private fillFrom(source: Range, target: Range): void {
    this.commit();
    const { sheet } = this.current;
    this.workbook.execute({ type: 'fill', sheet, source, target });
    this.change({ selection: covering([target]), endMode: false });
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

  private copyRange(cut: boolean): ClipboardContents {
    this.commit();
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

  private edited({ text, caret }: Edit): void {
    const entry = { ...this.current.entry!, text, caret, allFunctions: false };
    this.change({ entry });
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
