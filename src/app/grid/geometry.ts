import {
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Direction,
  type Frozen,
  type Lines,
  type Range,
  type SheetLayout,
} from '../../engine/workbook';

/** Sizes of the grid at 100%, in CSS pixels. */
export const ROW_HEADER_WIDTH = 46;
export const COLUMN_HEADER_HEIGHT = 24;
export const COLUMN_WIDTH = 96;
export const ROW_HEIGHT = 24;

/** A cell by its row and column, counted from zero. */
export interface CellPosition {
  readonly row: number;
  readonly col: number;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * The sizes of the rows or the columns of a sheet at 100%, a hidden one
 * of size 0. Finding a line's place or the line at a place takes a time
 * that grows with the logarithm of the lines of other sizes.
 */
export class Axis {
  /** The lines of other sizes than the standard one, in order. */
  private readonly lines: number[];
  /** For each of them, how much the lines before it add to the standard. */
  private readonly extra: number[];

  constructor(
    private readonly count: number,
    private readonly standard: number,
    private readonly layout: Lines,
  ) {
    this.lines = [
      ...new Set([
        ...layout.sizes.keys(),
        ...layout.hidden,
        ...(layout.fitted?.keys() ?? []),
      ]),
    ]
      .filter((index) => index < count)
      .sort((a, b) => a - b);
    this.extra = [0];
    for (const index of this.lines) {
      this.extra.push(
        this.extra[this.extra.length - 1]! + this.size(index) - standard,
      );
    }
  }

  size(index: number): number {
    if (this.layout.hidden.has(index)) return 0;
    return (
      this.layout.sizes.get(index) ??
      this.layout.fitted?.get(index) ??
      this.standard
    );
  }

  /** Where a line starts. */
  start(index: number): number {
    let low = 0;
    let high = this.lines.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (this.lines[middle]! < index) low = middle + 1;
      else high = middle;
    }
    return index * this.standard + this.extra[low]!;
  }

  /** The size of all the lines. */
  get total(): number {
    return this.start(this.count);
  }

  /**
   * The line at a place: the last that starts there or before, which is
   * not a hidden one but at the end. Places outside give the nearest line.
   */
  at(place: number): number {
    let low = 0;
    let high = this.count - 1;
    while (low < high) {
      const middle = (low + high + 1) >> 1;
      if (this.start(middle) <= place) low = middle;
      else high = middle - 1;
    }
    return low;
  }
}

/**
 * The rows and columns of a sheet, as the grid places them, and how many
 * of them are frozen.
 */
export interface Layout {
  readonly rows: Axis;
  readonly cols: Axis;
  readonly frozen: Frozen;
}

const UNFROZEN: Frozen = { rows: 0, cols: 0 };

const layouts = new WeakMap<SheetLayout, Layout>();

/** The layout of a sheet's lines: the same object for the same lines. */
export function layoutOf(sheet: SheetLayout): Layout {
  let layout = layouts.get(sheet);
  if (!layout) {
    layout = {
      rows: new Axis(SHEET_ROWS, ROW_HEIGHT, sheet.rows),
      cols: new Axis(SHEET_COLUMNS, COLUMN_WIDTH, sheet.columns),
      frozen: sheet.frozen ?? UNFROZEN,
    };
    layouts.set(sheet, layout);
  }
  return layout;
}

const NONE: Lines = { sizes: new Map(), hidden: new Set() };

/** Every row and column of the standard size. */
export const UNIFORM = layoutOf({ rows: NONE, columns: NONE });

/**
 * Which part of the sheet the canvas shows: how far it is scrolled, in
 * CSS pixels at the zoom, the zoom, where 1 is 100%, and the sizes of the
 * rows and columns.
 */
export interface Viewport {
  readonly scrollX: number;
  readonly scrollY: number;
  readonly zoom: number;
  readonly layout: Layout;
}

/** What is under a point of the canvas. */
export type Target =
  | { readonly kind: 'cell'; readonly row: number; readonly col: number }
  | { readonly kind: 'row'; readonly row: number }
  | { readonly kind: 'column'; readonly col: number }
  | { readonly kind: 'all' };

/** The size of the whole sheet, headers included, at a zoom. */
export function sheetSize({
  zoom,
  layout,
}: Pick<Viewport, 'zoom' | 'layout'>): { width: number; height: number } {
  return {
    width: (ROW_HEADER_WIDTH + layout.cols.total) * zoom,
    height: (COLUMN_HEADER_HEIGHT + layout.rows.total) * zoom,
  };
}

/**
 * Where a cell is drawn on the canvas. The headers, and the frozen rows
 * and columns, stay in place.
 */
export function cellRect(
  { row, col }: CellPosition,
  { scrollX, scrollY, zoom, layout }: Viewport,
): Rect {
  const { rows, cols, frozen } = layout;
  return {
    x:
      (ROW_HEADER_WIDTH + cols.start(col)) * zoom -
      (col < frozen.cols ? 0 : scrollX),
    y:
      (COLUMN_HEADER_HEIGHT + rows.start(row)) * zoom -
      (row < frozen.rows ? 0 : scrollY),
    width: cols.size(col) * zoom,
    height: rows.size(row) * zoom,
  };
}

/**
 * The cell or header under a point of the canvas. A point beyond the last
 * row or column is taken to be on it.
 */
export function hitTest(x: number, y: number, viewport: Viewport): Target {
  const { zoom } = viewport;
  const inRowHeader = x < ROW_HEADER_WIDTH * zoom;
  const inColumnHeader = y < COLUMN_HEADER_HEIGHT * zoom;
  const { row, col } = cellAt(x, y, viewport);
  if (inRowHeader && inColumnHeader) return { kind: 'all' };
  if (inRowHeader) return { kind: 'row', row };
  if (inColumnHeader) return { kind: 'column', col };
  return { kind: 'cell', row, col };
}

/**
 * The cell at a point as if the headers were not drawn over the cells: over
 * a header, the cell scrolled behind it, or the first one when lines are
 * frozen. A drag past the headers or the edges of the canvas reaches the
 * cells beyond them this way.
 */
export function cellAt(x: number, y: number, viewport: Viewport): CellPosition {
  return { row: rowAt(y, viewport), col: colAt(x, viewport) };
}

function rowAt(y: number, { scrollY, zoom, layout }: Viewport): number {
  const frozen = layout.frozen.rows > 0 && y < frozenEnd(layout, 'rows', zoom);
  const scroll = frozen ? 0 : scrollY;
  return layout.rows.at((y + scroll) / zoom - COLUMN_HEADER_HEIGHT);
}

function colAt(x: number, { scrollX, zoom, layout }: Viewport): number {
  const frozen =
    layout.frozen.cols > 0 && x < frozenEnd(layout, 'columns', zoom);
  const scroll = frozen ? 0 : scrollX;
  return layout.cols.at((x + scroll) / zoom - ROW_HEADER_WIDTH);
}

/**
 * Where the frozen rows or columns end on the canvas, after the headers;
 * where the headers end when none are.
 */
function frozenEnd(layout: Layout, direction: Direction, zoom: number) {
  return direction === 'rows'
    ? (COLUMN_HEADER_HEIGHT + layout.rows.start(layout.frozen.rows)) * zoom
    : (ROW_HEADER_WIDTH + layout.cols.start(layout.frozen.cols)) * zoom;
}

/**
 * A part of the canvas drawn with its own scroll position: with frozen
 * rows or columns, up to four of them, and only one without. Within its
 * viewport, nothing is frozen.
 */
export interface Pane {
  readonly viewport: Viewport;
  /** Where it is on the canvas, the headers left out. */
  readonly clip: Rect;
  /** The cells it shows, partial ones included. */
  readonly range: Range;
}

/**
 * The panes of a canvas of this size: the frozen rows and columns, which
 * do not scroll, the frozen rows, which scroll sideways, the frozen
 * columns, which scroll up and down, and the rest, in that order.
 */
export function panes(
  viewport: Viewport,
  width: number,
  height: number,
): Pane[] {
  const { zoom, layout } = viewport;
  const parts = (
    direction: Direction,
    header: number,
    size: number,
    scroll: number,
    frozen: number,
  ) => {
    // Frozen lines past the canvas are not drawn.
    const end = Math.min(frozenEnd(layout, direction, zoom), size);
    const rest = { start: end, size: size - end, scroll };
    return frozen > 0
      ? [{ start: header * zoom, size: end - header * zoom, scroll: 0 }, rest]
      : [rest];
  };
  const unfrozen = { ...viewport, layout: { ...layout, frozen: UNFROZEN } };
  const result: Pane[] = [];
  for (const y of parts(
    'rows',
    COLUMN_HEADER_HEIGHT,
    height,
    viewport.scrollY,
    layout.frozen.rows,
  )) {
    for (const x of parts(
      'columns',
      ROW_HEADER_WIDTH,
      width,
      viewport.scrollX,
      layout.frozen.cols,
    )) {
      const pane = { ...unfrozen, scrollX: x.scroll, scrollY: y.scroll };
      const clip = { x: x.start, y: y.start, width: x.size, height: y.size };
      result.push({
        viewport: pane,
        clip,
        range: {
          top: rowAt(clip.y, pane),
          left: colAt(clip.x, pane),
          bottom: rowAt(clip.y + clip.height - 1, pane),
          right: colAt(clip.x + clip.width - 1, pane),
        },
      });
    }
  }
  return result;
}

/** How near a header border a point must be to drag it, in CSS pixels. */
const BORDER_REACH = 4;

/**
 * The row or column whose border, at its end in the headers, is near a
 * point of the canvas, to resize it; a hidden line's border is that of
 * the shown line before it.
 */
export function borderAt(
  x: number,
  y: number,
  viewport: Viewport,
): { direction: Direction; index: number } | null {
  const { zoom, layout } = viewport;
  const inRowHeader = x < ROW_HEADER_WIDTH * zoom;
  const inColumnHeader = y < COLUMN_HEADER_HEIGHT * zoom;
  if (inRowHeader === inColumnHeader) return null;
  const direction = inColumnHeader ? 'columns' : 'rows';
  const axis = inColumnHeader ? layout.cols : layout.rows;
  const point = inColumnHeader ? x : y;
  const index = inColumnHeader ? colAt(x, viewport) : rowAt(y, viewport);
  const rect = cellRect(
    inColumnHeader ? { row: 0, col: index } : { row: index, col: 0 },
    viewport,
  );
  const [start, size] = inColumnHeader
    ? [rect.x, rect.width]
    : [rect.y, rect.height];
  if (Math.abs(point - start - size) <= BORDER_REACH) {
    return { direction, index };
  }
  if (Math.abs(point - start) > BORDER_REACH) return null;
  let before = index - 1;
  while (before >= 0 && axis.size(before) === 0) before--;
  return before >= 0 ? { direction, index: before } : null;
}

/** The cells that a canvas of this size shows, partial ones included. */
export function visibleRange(
  width: number,
  height: number,
  viewport: Viewport,
): Range {
  const { zoom } = viewport;
  return {
    top: rowAt(COLUMN_HEADER_HEIGHT * zoom, viewport),
    left: colAt(ROW_HEADER_WIDTH * zoom, viewport),
    bottom: rowAt(height - 1, viewport),
    right: colAt(width - 1, viewport),
  };
}

/**
 * The scroll position closest to the current one at which a canvas of
 * this size shows the whole of a cell, or as much of it as fits.
 */
export function scrollToShow(
  cell: CellPosition,
  viewport: Viewport,
  width: number,
  height: number,
): { scrollX: number; scrollY: number } {
  const { scrollX, scrollY, zoom, layout } = viewport;
  const rect = cellRect(cell, viewport);
  // A frozen cell is always shown; the others, past the frozen ones.
  const show = (start: number, size: number, from: number, to: number) => {
    if (start < from) return start - from;
    if (start + size > to) return Math.min(start + size - to, start - from);
    return 0;
  };
  return {
    scrollX:
      cell.col < layout.frozen.cols
        ? scrollX
        : scrollX +
          show(rect.x, rect.width, frozenEnd(layout, 'columns', zoom), width),
    scrollY:
      cell.row < layout.frozen.rows
        ? scrollY
        : scrollY +
          show(rect.y, rect.height, frozenEnd(layout, 'rows', zoom), height),
  };
}

/** The size of the square drawn at the bottom right of the selection. */
export const FILL_HANDLE = 6;

/**
 * Whether a point of the canvas is on the fill handle of a range: within
 * a few pixels of its bottom right corner.
 */
export function isOnFillHandle(
  x: number,
  y: number,
  range: Range,
  viewport: Viewport,
): boolean {
  // Its own cell's corner, which stays with it when lines are frozen.
  const cell = cellRect({ row: range.bottom, col: range.right }, viewport);
  const corner = { x: cell.x + cell.width, y: cell.y + cell.height };
  const reach = FILL_HANDLE / 2 + 2;
  return Math.abs(x - corner.x) <= reach && Math.abs(y - corner.y) <= reach;
}
