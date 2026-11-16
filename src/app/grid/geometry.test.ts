import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../../engine/workbook';
import {
  barRect,
  borderAt,
  cellAt,
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  hitTest,
  isOnFillHandle,
  layoutOf,
  panes,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  scrollToShow,
  sheetSize,
  UNIFORM,
  visibleRange,
  type Viewport,
} from './geometry';

const top: Viewport = { scrollX: 0, scrollY: 0, zoom: 1, layout: UNIFORM };

describe('grid geometry', () => {
  it('places A1 after the headers and other cells by their size', () => {
    expect(cellRect({ row: 0, col: 0 }, top)).toEqual({
      x: ROW_HEADER_WIDTH,
      y: COLUMN_HEADER_HEIGHT,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
    });
    expect(cellRect({ row: 3, col: 2 }, top)).toMatchObject({
      x: ROW_HEADER_WIDTH + 2 * COLUMN_WIDTH,
      y: COLUMN_HEADER_HEIGHT + 3 * ROW_HEIGHT,
    });
  });

  it('moves the cells, not the headers, by the scroll position', () => {
    const scrolled = {
      scrollX: 10,
      scrollY: 2 * ROW_HEIGHT,
      zoom: 1,
      layout: UNIFORM,
    };
    expect(cellRect({ row: 2, col: 0 }, scrolled)).toMatchObject({
      x: ROW_HEADER_WIDTH - 10,
      y: COLUMN_HEADER_HEIGHT,
    });
    expect(
      hitTest(ROW_HEADER_WIDTH + 1, COLUMN_HEADER_HEIGHT + 1, scrolled),
    ).toEqual({ kind: 'cell', row: 2, col: 0 });
  });

  it('scales everything, headers included, by the zoom', () => {
    const zoomed = { scrollX: 0, scrollY: 0, zoom: 2, layout: UNIFORM };
    expect(cellRect({ row: 1, col: 1 }, zoomed)).toEqual({
      x: 2 * (ROW_HEADER_WIDTH + COLUMN_WIDTH),
      y: 2 * (COLUMN_HEADER_HEIGHT + ROW_HEIGHT),
      width: 2 * COLUMN_WIDTH,
      height: 2 * ROW_HEIGHT,
    });
    expect(sheetSize({ zoom: 0.5, layout: UNIFORM })).toEqual({
      width: (ROW_HEADER_WIDTH + SHEET_COLUMNS * COLUMN_WIDTH) / 2,
      height: (COLUMN_HEADER_HEIGHT + SHEET_ROWS * ROW_HEIGHT) / 2,
    });
  });

  it('finds the cell or header under a point', () => {
    const rect = cellRect({ row: 3, col: 2 }, top);
    expect(hitTest(rect.x, rect.y, top)).toEqual({
      kind: 'cell',
      row: 3,
      col: 2,
    });
    expect(
      hitTest(rect.x + rect.width - 1, rect.y + rect.height - 1, top),
    ).toEqual({ kind: 'cell', row: 3, col: 2 });
    expect(hitTest(ROW_HEADER_WIDTH - 1, rect.y, top)).toEqual({
      kind: 'row',
      row: 3,
    });
    expect(hitTest(rect.x, COLUMN_HEADER_HEIGHT - 1, top)).toEqual({
      kind: 'column',
      col: 2,
    });
    expect(hitTest(1, 1, top)).toEqual({ kind: 'all' });
  });

  it('finds the cell scrolled behind a header, for a drag past it', () => {
    const scrolled = {
      scrollX: 3 * COLUMN_WIDTH,
      scrollY: 3 * ROW_HEIGHT,
      zoom: 1,
      layout: UNIFORM,
    };
    expect(
      cellAt(ROW_HEADER_WIDTH - 1, COLUMN_HEADER_HEIGHT - 1, scrolled),
    ).toEqual({
      row: 2,
      col: 2,
    });
    expect(cellAt(-1000, -1000, scrolled)).toEqual({ row: 0, col: 0 });
  });

  it('finds the nearest cell for a point beyond the sheet', () => {
    const end = { scrollX: 1e9, scrollY: 1e9, zoom: 1, layout: UNIFORM };
    expect(hitTest(5000, 5000, end)).toEqual({
      kind: 'cell',
      row: SHEET_ROWS - 1,
      col: SHEET_COLUMNS - 1,
    });
    expect(hitTest(-5, -5, top)).toEqual({ kind: 'all' });
  });

  it('gives the cells a canvas shows, including partial ones', () => {
    const width = ROW_HEADER_WIDTH + 2.5 * COLUMN_WIDTH;
    const height = COLUMN_HEADER_HEIGHT + 4 * ROW_HEIGHT;
    expect(visibleRange(width, height, top)).toEqual({
      top: 0,
      left: 0,
      bottom: 3,
      right: 2,
    });
    const scrolled = {
      scrollX: COLUMN_WIDTH * 0.75,
      scrollY: 0,
      zoom: 1,
      layout: UNIFORM,
    };
    expect(visibleRange(width, height, scrolled)).toMatchObject({
      left: 0,
      right: 3,
    });
    const end = {
      scrollX: 0,
      scrollY: SHEET_ROWS * ROW_HEIGHT,
      zoom: 1,
      layout: UNIFORM,
    };
    expect(visibleRange(width, height, end)).toMatchObject({
      top: SHEET_ROWS - 1,
      bottom: SHEET_ROWS - 1,
    });
  });

  it('scrolls as little as it can to show a cell', () => {
    const width = ROW_HEADER_WIDTH + 4 * COLUMN_WIDTH;
    const height = COLUMN_HEADER_HEIGHT + 10 * ROW_HEIGHT;
    // Already shown: no change.
    expect(scrollToShow({ row: 5, col: 2 }, top, width, height)).toEqual({
      scrollX: 0,
      scrollY: 0,
    });
    // Below and to the right: the cell ends at the far edges.
    expect(scrollToShow({ row: 20, col: 6 }, top, width, height)).toEqual({
      scrollX: 3 * COLUMN_WIDTH,
      scrollY: 11 * ROW_HEIGHT,
    });
    // Above and to the left: the cell starts just after the headers.
    const scrolled = { scrollX: 1000, scrollY: 1000, zoom: 1, layout: UNIFORM };
    expect(scrollToShow({ row: 2, col: 1 }, scrolled, width, height)).toEqual({
      scrollX: COLUMN_WIDTH,
      scrollY: 2 * ROW_HEIGHT,
    });
    // At twice the size, the distances double.
    const zoomed = { scrollX: 0, scrollY: 0, zoom: 2, layout: UNIFORM };
    expect(
      scrollToShow({ row: 20, col: 0 }, zoomed, width, height).scrollY,
    ).toBe(2 * (COLUMN_HEADER_HEIGHT + 21 * ROW_HEIGHT) - height);
  });

  it('finds the fill handle at the bottom right corner of a range', () => {
    const range = { top: 1, left: 1, bottom: 2, right: 3 };
    const corner = cellRect({ row: 3, col: 4 }, top);
    expect(isOnFillHandle(corner.x, corner.y, range, top)).toBe(true);
    expect(isOnFillHandle(corner.x - 3, corner.y + 3, range, top)).toBe(true);
    expect(isOnFillHandle(corner.x - 8, corner.y, range, top)).toBe(false);
    const scrolled = { scrollX: 10, scrollY: 20, zoom: 2, layout: UNIFORM };
    const zoomed = cellRect({ row: 3, col: 4 }, scrolled);
    expect(isOnFillHandle(zoomed.x, zoomed.y, range, scrolled)).toBe(true);
  });

  it('places rows and columns of other sizes, and hidden ones, by their sizes', () => {
    const lines = {
      rows: { sizes: new Map([[1, 50]]), hidden: new Set([3, 4]) },
      columns: { sizes: new Map([[0, 30]]), hidden: new Set<number>() },
    };
    const layout = layoutOf(lines);
    // The same lines give the same layout, which the grid draws at once.
    expect(layoutOf(lines)).toBe(layout);
    const at = { scrollX: 0, scrollY: 0, zoom: 1, layout };
    expect(cellRect({ row: 2, col: 1 }, at)).toEqual({
      x: ROW_HEADER_WIDTH + 30,
      y: COLUMN_HEADER_HEIGHT + ROW_HEIGHT + 50,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
    });
    expect(cellRect({ row: 3, col: 0 }, at).height).toBe(0);
    // The cells shown, and the scroll that shows a cell, follow the sizes.
    expect(
      visibleRange(
        ROW_HEADER_WIDTH + 30 + COLUMN_WIDTH,
        COLUMN_HEADER_HEIGHT + 75,
        at,
      ),
    ).toEqual({ top: 0, left: 0, bottom: 2, right: 1 });
    expect(
      scrollToShow({ row: 5, col: 0 }, at, 500, COLUMN_HEADER_HEIGHT + 75)
        .scrollY,
    ).toBe(47);
    expect(cellRect({ row: 5, col: 0 }, at).y).toBe(
      COLUMN_HEADER_HEIGHT + 2 * ROW_HEIGHT + 50,
    );
    // A point after row 2 is on row 5, past the hidden ones.
    const below = COLUMN_HEADER_HEIGHT + 2 * ROW_HEIGHT + 51;
    expect(hitTest(ROW_HEADER_WIDTH + 1, below, at)).toEqual({
      kind: 'cell',
      row: 5,
      col: 0,
    });
    expect(hitTest(ROW_HEADER_WIDTH + 1, below - 2, at)).toMatchObject({
      row: 2,
    });
    expect(sheetSize({ zoom: 1, layout })).toEqual({
      width: ROW_HEADER_WIDTH + SHEET_COLUMNS * COLUMN_WIDTH - 66,
      height: COLUMN_HEADER_HEIGHT + (SHEET_ROWS - 1) * ROW_HEIGHT + 2,
    });
  });

  it('gives rows the height their text needs, unless they have a size of their own', () => {
    const layout = layoutOf({
      rows: {
        sizes: new Map([[2, 30]]),
        hidden: new Set<number>(),
        fitted: new Map([
          [1, 56],
          [2, 40],
        ]),
      },
      columns: { sizes: new Map(), hidden: new Set<number>() },
    });
    expect(layout.rows.size(1)).toBe(56);
    expect(layout.rows.size(2)).toBe(30);
    expect(layout.rows.start(3)).toBe(ROW_HEIGHT + 56 + 30);
  });

  it('finds the border of a row or column header near a point, to resize it', () => {
    const layout = layoutOf({
      rows: { sizes: new Map(), hidden: new Set([2]) },
      columns: { sizes: new Map(), hidden: new Set() },
    });
    const at = { scrollX: 0, scrollY: 0, zoom: 1, layout };
    const b1 = cellRect({ row: 0, col: 1 }, at);
    const end = b1.x + b1.width;
    expect(borderAt(end - 2, 5, at)).toEqual({
      direction: 'columns',
      index: 1,
    });
    expect(borderAt(end + 3, 5, at)).toEqual({
      direction: 'columns',
      index: 1,
    });
    expect(borderAt(end - 20, 5, at)).toBeNull();
    // The left border of the first column is the header's, not a column's.
    expect(borderAt(ROW_HEADER_WIDTH + 1, 5, at)).toBeNull();
    // Over the cells, there is no border to drag.
    expect(borderAt(end, COLUMN_HEADER_HEIGHT + 5, at)).toBeNull();
    // Row 2 is hidden, so the border after row 1 is row 1's.
    const row3 = cellRect({ row: 3, col: 0 }, at);
    expect(borderAt(5, row3.y + 1, at)).toEqual({
      direction: 'rows',
      index: 1,
    });
  });

  it('keeps frozen rows and columns in place while the others scroll under them', () => {
    const NONE = { sizes: new Map(), hidden: new Set<number>() };
    const layout = layoutOf({
      rows: NONE,
      columns: NONE,
      frozen: { rows: 2, cols: 1 },
    });
    const scrolled = { scrollX: 300, scrollY: 240, zoom: 1, layout };
    const right = ROW_HEADER_WIDTH + COLUMN_WIDTH;
    const below = COLUMN_HEADER_HEIGHT + 2 * ROW_HEIGHT;
    // A frozen cell stays where it is; the others move by the scroll.
    expect(cellRect({ row: 1, col: 0 }, scrolled)).toMatchObject({
      x: ROW_HEADER_WIDTH,
      y: COLUMN_HEADER_HEIGHT + ROW_HEIGHT,
    });
    expect(cellRect({ row: 12, col: 4 }, scrolled)).toMatchObject({
      x: ROW_HEADER_WIDTH + 4 * COLUMN_WIDTH - 300,
      y: COLUMN_HEADER_HEIGHT + 12 * ROW_HEIGHT - 240,
    });
    // Hit testing finds the frozen cells, and those scrolled past them.
    expect(hitTest(ROW_HEADER_WIDTH + 1, below - 1, scrolled)).toEqual({
      kind: 'cell',
      row: 1,
      col: 0,
    });
    expect(hitTest(right + 1, below + 1, scrolled)).toEqual({
      kind: 'cell',
      row: 12,
      col: 4,
    });
    expect(hitTest(5, below - 1, scrolled)).toEqual({ kind: 'row', row: 1 });
    expect(hitTest(right + 1, 5, scrolled)).toEqual({
      kind: 'column',
      col: 4,
    });
    // The fill handle of a range ending on the last frozen row is on it.
    const handle = cellRect({ row: 1, col: 0 }, scrolled);
    expect(
      isOnFillHandle(
        handle.x + handle.width,
        handle.y + handle.height,
        { top: 0, left: 0, bottom: 1, right: 0 },
        scrolled,
      ),
    ).toBe(true);
    // A cell is shown past the frozen ones, and a frozen one needs no scroll.
    const width = right + 4 * COLUMN_WIDTH;
    const height = below + 5 * ROW_HEIGHT;
    expect(scrollToShow({ row: 3, col: 2 }, scrolled, width, height)).toEqual({
      scrollX: COLUMN_WIDTH,
      scrollY: ROW_HEIGHT,
    });
    expect(scrollToShow({ row: 0, col: 0 }, scrolled, width, height)).toEqual({
      scrollX: 300,
      scrollY: 240,
    });
    // The canvas is drawn in four panes, each scrolled its own way.
    expect(panes(scrolled, width, height)).toMatchObject([
      {
        viewport: { ...scrolled, scrollX: 0, scrollY: 0, layout: UNIFORM },
        clip: {
          x: ROW_HEADER_WIDTH,
          y: COLUMN_HEADER_HEIGHT,
          width: COLUMN_WIDTH,
          height: 2 * ROW_HEIGHT,
        },
      },
      {
        viewport: { ...scrolled, scrollY: 0, layout: UNIFORM },
        clip: {
          x: right,
          y: COLUMN_HEADER_HEIGHT,
          width: width - right,
          height: 2 * ROW_HEIGHT,
        },
      },
      {
        viewport: { ...scrolled, scrollX: 0, layout: UNIFORM },
        clip: {
          x: ROW_HEADER_WIDTH,
          y: below,
          width: COLUMN_WIDTH,
          height: height - below,
        },
      },
      {
        viewport: { ...scrolled, layout: UNIFORM },
        clip: {
          x: right,
          y: below,
          width: width - right,
          height: height - below,
        },
      },
    ]);
    // Frozen lines past the canvas are left out, and nothing scrolls.
    const tall = {
      ...top,
      layout: layoutOf({
        rows: NONE,
        columns: NONE,
        frozen: { rows: SHEET_ROWS - 1, cols: 0 },
      }),
    };
    expect(panes(tall, width, height)).toMatchObject([
      {
        clip: {
          y: COLUMN_HEADER_HEIGHT,
          height: height - COLUMN_HEADER_HEIGHT,
        },
        range: { top: 0, bottom: 6 },
      },
      { clip: { y: height, height: 0 } },
    ]);
    // Without frozen lines, there is one pane.
    expect(panes(top, width, height)).toMatchObject([
      {
        viewport: top,
        clip: {
          x: ROW_HEADER_WIDTH,
          y: COLUMN_HEADER_HEIGHT,
          width: width - ROW_HEADER_WIDTH,
          height: height - COLUMN_HEADER_HEIGHT,
        },
      },
    ]);
  });
});

describe('barRect', () => {
  const box = { x: 10, y: 20, width: 100, height: 24 };

  it('puts a bar inside the lines of its cell, as long as its share of it', () => {
    expect(barRect(box, 0.5, 1)).toEqual({
      x: 12,
      y: 22,
      width: 48,
      height: 20,
    });
    expect(barRect(box, 1, 2)).toEqual({ x: 14, y: 24, width: 92, height: 16 });
  });

  it('draws nothing in a hidden row or column, or for a bar of no length', () => {
    expect(barRect({ ...box, height: 0 }, 1, 1)).toBeNull();
    expect(barRect({ ...box, width: 0 }, 1, 1)).toBeNull();
    expect(barRect(box, 0, 1)).toBeNull();
  });
});
