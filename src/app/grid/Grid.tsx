import * as ContextMenu from '@radix-ui/react-context-menu';
import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from 'react';
import {
  cellName,
  columnName,
  fontPixels,
  lineHeight,
  type Border,
  type Direction,
  type Range,
  type Style,
  type TextWidth,
  type Workbook,
  SHEET_COLUMNS,
  SHEET_ROWS,
} from '../../engine/workbook';
import {
  borderAt,
  cellAt,
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  FILL_HANDLE,
  hitTest,
  isOnFillHandle,
  layoutOf,
  panes,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  scrollToShow,
  sheetSize,
  visibleRange,
  type CellPosition,
  type Target,
  type Viewport,
} from './geometry';
import { formulaHelp, type Highlight } from './formula';
import { FormulaEditor } from './FormulaEditor';
import styles from './Grid.module.css';
import { keyAction } from './keys';
import { mergeAt, pick } from './selection';
import {
  cellText,
  cssFont,
  rowTexts,
  TEXT_PADDING,
  textColorOn,
  type RowText,
} from './text';
import type { Entry, View, ViewState } from './view';

const FONT_SIZE = 13;
export const MAC = /Mac|iPhone|iPad/.test(navigator.platform);

let measuring: CanvasRenderingContext2D | null = null;

/** The width of a text in a cell's font at 100%, as the grid draws it. */
export const measureText: TextWidth = (text: string, style: Style) => {
  measuring ??= document.createElement('canvas').getContext('2d')!;
  measuring.font = cssFont(style, 1);
  return measuring.measureText(text).width;
};

/** How often the border of the cells copied moves, in milliseconds. */
const MARCH = 100;

/**
 * The cells of the sheet that the view shows, drawn on a canvas that stays
 * in place while the element around it scrolls over the whole sheet. The
 * mouse and the keyboard select and edit cells as in Excel, and dragging
 * or double-clicking a header border resizes its row or column.
 * Right-clicking a cell, or a row or column header, opens `menu` for it,
 * the content of a context menu.
 */
export function Grid({
  workbook,
  view,
  menu,
}: {
  workbook: Workbook;
  view: View;
  menu: (lines?: Direction) => ReactNode;
}) {
  const scroller = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  /** The row or column whose border is dragged, and where it starts. */
  const resizeDrag = useRef<{
    direction: Direction;
    index: number;
    start: number;
  } | null>(null);
  /** Whose menu right-clicking opened: the cells', or the headers'. */
  const [menuLines, setMenuLines] = useState<Direction | undefined>();
  /**
   * What a drag selects, the fill handle it drags, or the corner of the
   * range it picks for a formula.
   */
  const drag = useRef<Target['kind'] | 'fill' | 'point' | null>(null);
  const state = useSyncExternalStore(view.subscribe, () => view.state);
  const { sheet, selection, entry, zoom, copied } = state;
  const { active } = selection;
  const layout = layoutOf(
    useSyncExternalStore(workbook.subscribe.bind(workbook), () =>
      workbook.layout(sheet),
    ),
  );
  const activeDisplay = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => workbook.cell(sheet, active.row, active.col).display,
  );

  const viewport = (): Viewport => ({
    scrollX: scroller.current!.scrollLeft,
    scrollY: scroller.current!.scrollTop,
    zoom: view.state.zoom,
    layout: layoutOf(workbook.layout(view.state.sheet)),
  });
  /** Where a mouse event is on the canvas. */
  const point = (event: { clientX: number; clientY: number }) => {
    const bounds = canvas.current!.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  };
  const size = () => ({
    width: scroller.current!.clientWidth,
    height: scroller.current!.clientHeight,
  });
  const topLeft = (): CellPosition => {
    const { width, height } = size();
    const { top, left } = visibleRange(width, height, viewport());
    return { row: top, col: left };
  };
  const show = (cell: CellPosition) => {
    const { width, height } = size();
    const { scrollX, scrollY } = scrollToShow(cell, viewport(), width, height);
    scroller.current!.scrollTo(scrollX, scrollY);
  };

  useEffect(() => {
    const element = scroller.current!;
    const paint = () => {
      const { width, height } = size();
      draw(canvas.current!, workbook, view.state, viewport(), width, height);
    };
    paint();
    const unsubscribeWorkbook = workbook.subscribe(paint);
    const unsubscribeView = view.subscribe(paint);
    // Drawn as soon as the position changes, which is once a frame at most.
    element.addEventListener('scroll', paint);
    const resize = new ResizeObserver(paint);
    resize.observe(element);
    // The colors are those of the theme, which is set on the root element.
    const theme = new MutationObserver(paint);
    theme.observe(document.documentElement, {
      attributeFilter: ['data-theme'],
    });
    // The border of the cells copied moves while they are shown.
    const march = window.setInterval(() => {
      if (view.state.copied?.sheet === view.state.sheet) paint();
    }, MARCH);
    return () => {
      window.clearInterval(march);
      unsubscribeWorkbook();
      unsubscribeView();
      element.removeEventListener('scroll', paint);
      resize.disconnect();
      theme.disconnect();
    };
    // `size` and `viewport` read the elements and the view when called.
  }, [workbook, view]);

  // The editor of a frozen cell moves with the scroll, to stay on it.
  const [, setScroll] = useState(0);
  const editsFrozen =
    entry !== null &&
    (active.row < layout.frozen.rows || active.col < layout.frozen.cols);
  useEffect(() => {
    if (!editsFrozen) return;
    const element = scroller.current!;
    const onScroll = () => setScroll((count) => count + 1);
    element.addEventListener('scroll', onScroll);
    return () => element.removeEventListener('scroll', onScroll);
  }, [editsFrozen]);

  // A new zoom keeps the same cell at the top left.
  const shownZoom = useRef(zoom);
  useLayoutEffect(() => {
    const element = scroller.current!;
    const ratio = zoom / shownZoom.current;
    shownZoom.current = zoom;
    element.scrollTo(element.scrollLeft * ratio, element.scrollTop * ratio);
  }, [zoom]);

  // Another sheet opens at its top left, ready for the keyboard unless a
  // formula is being typed in the formula bar.
  const shownSheet = useRef(sheet);
  useLayoutEffect(() => {
    if (shownSheet.current === sheet) return;
    shownSheet.current = sheet;
    scroller.current!.scrollTo(0, 0);
    if (!view.state.entry) scroller.current!.focus();
  }, [sheet, view]);

  // When an entry in the formula bar ends, the keyboard comes back to the
  // grid, which shows the cell it moved to.
  useEffect(() => {
    let previous: Entry | null = view.state.entry;
    return view.subscribe(() => {
      const { entry, selection } = view.state;
      if (previous?.place === 'formulaBar' && entry === null) {
        scroller.current!.focus();
        show(selection.active);
      }
      previous = entry;
    });
    // `show` reads the elements and the view when called.
  }, [view]);

  // Dragging goes on outside the canvas, and ends wherever the button is
  // released.
  useEffect(() => {
    const onMove = (event: globalThis.MouseEvent) => {
      const resizing = resizeDrag.current;
      if (resizing) {
        const { x, y } = point(event);
        const at = resizing.direction === 'rows' ? y : x;
        view.previewResize(
          resizing.direction,
          resizing.index,
          (at - resizing.start) / view.state.zoom,
        );
        return;
      }
      if (!drag.current || drag.current === 'all') return;
      const bounds = canvas.current!.getBoundingClientRect();
      // Past the headers or the edges, the drag reaches the cells beyond
      // them, and the grid scrolls to show them.
      const { row, col } = cellAt(
        event.clientX - bounds.left,
        event.clientY - bounds.top,
        viewport(),
      );
      if (drag.current === 'fill') {
        view.previewFill({ row, col });
        show({ row, col });
        return;
      }
      if (drag.current === 'point') {
        view.pick({ row, col }, 'extend');
        return;
      }
      const target: Target =
        drag.current === 'row'
          ? { kind: 'row', row }
          : drag.current === 'column'
            ? { kind: 'column', col }
            : { kind: 'cell', row, col };
      const next = pick(view.state.selection, target, 'extend', topLeft());
      if (
        next.focus.row === view.state.selection.focus.row &&
        next.focus.col === view.state.selection.focus.col
      ) {
        return;
      }
      view.select(next);
      // Over the headers, the grid scrolls along them only: the focus of
      // a whole row is in its last column.
      const shown = topLeft();
      show(
        drag.current === 'row'
          ? { row: next.focus.row, col: shown.col }
          : drag.current === 'column'
            ? { row: shown.row, col: next.focus.col }
            : next.focus,
      );
    };
    const onUp = () => {
      if (drag.current === 'fill') view.fill();
      // The format painter paints the range a click or drag selects.
      else if (drag.current && drag.current !== 'point') view.paint();
      if (resizeDrag.current) view.resize();
      drag.current = null;
      resizeDrag.current = null;
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [view]);

  /** Whether a mouse event is on the fill handle of the selection. */
  const onFillHandle = (event: MouseEvent) => {
    const { ranges } = view.state.selection;
    if (view.state.entry || ranges.length !== 1) return false;
    const bounds = canvas.current!.getBoundingClientRect();
    return isOnFillHandle(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
      ranges[0]!,
      viewport(),
    );
  };

  /** The header border under a mouse event, to resize its line. */
  const onBorder = (event: MouseEvent) => {
    if (view.state.entry) return null;
    const { x, y } = point(event);
    return borderAt(x, y, viewport());
  };

  const onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return;
    const border = onBorder(event);
    if (border) {
      const { direction, index } = border;
      const rect = cellRect(
        direction === 'rows' ? { row: index, col: 0 } : { row: 0, col: index },
        viewport(),
      );
      resizeDrag.current = {
        direction,
        index,
        start: direction === 'rows' ? rect.y : rect.x,
      };
      event.preventDefault();
      scroller.current!.focus();
      return;
    }
    if (onFillHandle(event)) {
      drag.current = 'fill';
      event.preventDefault();
      scroller.current!.focus();
      return;
    }
    const bounds = canvas.current!.getBoundingClientRect();
    const target = hitTest(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
      viewport(),
    );
    const mode = event.shiftKey
      ? 'extend'
      : (MAC ? event.metaKey : event.ctrlKey)
        ? 'add'
        : 'replace';
    if (target.kind === 'cell' && view.canPick) {
      view.pick({ row: target.row, col: target.col }, mode);
      drag.current = 'point';
      // The editor keeps the keyboard.
      event.preventDefault();
      return;
    }
    view.select(pick(view.state.selection, target, mode, topLeft()));
    drag.current = target.kind;
    // Keep the keyboard on the grid, which the browser would otherwise
    // take away from the editor that is closing.
    event.preventDefault();
    scroller.current!.focus();
  };

  // Right-clicking a cell or header outside the selection selects it
  // first, as in Excel; the corner has no menu.
  const onContextMenu = (event: MouseEvent) => {
    const { x, y } = point(event);
    const target = hitTest(x, y, viewport());
    if (target.kind === 'all') {
      event.preventDefault();
      return;
    }
    const last = { row: SHEET_ROWS - 1, col: SHEET_COLUMNS - 1 };
    // A header's menu is for its whole row or column.
    const [top, left, bottom, right] =
      target.kind === 'cell'
        ? [target.row, target.col, target.row, target.col]
        : target.kind === 'row'
          ? [target.row, 0, target.row, last.col]
          : [0, target.col, last.row, target.col];
    const inside = view.state.selection.ranges.some(
      (each) =>
        each.top <= top &&
        each.bottom >= bottom &&
        each.left <= left &&
        each.right >= right,
    );
    if (!inside) {
      view.select(pick(view.state.selection, target, 'replace', topLeft()));
    }
    setMenuLines(
      target.kind === 'row'
        ? 'rows'
        : target.kind === 'column'
          ? 'columns'
          : undefined,
    );
  };

  const onDoubleClick = (event: MouseEvent) => {
    const border = onBorder(event);
    if (border) {
      view.fitLines(border.direction, border.index, measureText);
      return;
    }
    if (onFillHandle(event)) {
      view.autoFill();
      return;
    }
    const bounds = canvas.current!.getBoundingClientRect();
    const target = hitTest(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
      viewport(),
    );
    if (target.kind === 'cell') view.edit('cell');
  };

  const onKeyDown = (event: KeyboardEvent) => {
    const { width, height } = size();
    const page = {
      rows: Math.max(
        1,
        Math.floor((height / zoom - COLUMN_HEADER_HEIGHT) / ROW_HEIGHT),
      ),
      cols: Math.max(
        1,
        Math.floor((width / zoom - ROW_HEADER_WIDTH) / COLUMN_WIDTH),
      ),
    };
    const action = keyAction(
      {
        key: event.key,
        shift: event.shiftKey,
        ctrl: event.ctrlKey,
        meta: event.metaKey,
        alt: event.altKey,
      },
      {
        mac: MAC,
        mode: entry?.mode ?? null,
        endMode: state.endMode,
        page,
        canPoint: view.canPick,
      },
    );
    if (!action) return;
    event.preventDefault();
    if (action.type === 'type') {
      // The editor opens holding this first character.
      view.type(action.text);
    } else if (action.type === 'edit') {
      view.edit('cell');
    } else if (action.type === 'clear') {
      view.clear();
    } else if (action.type === 'cycleReference') {
      view.cycleReference();
    } else if (action.type === 'cancel') {
      view.cancel();
      scroller.current!.focus();
    } else if (action.type === 'point') {
      view.point(action.navigation);
      show(view.state.entry!.point!.focus);
    } else {
      const { navigation } = action;
      // The editor closes, and the keyboard stays on the grid.
      scroller.current!.focus();
      view.navigate(navigation);
      if (navigation.type === 'endMode' || navigation.type === 'all') return;
      // A page key turns the page, so the cell keeps its place on screen.
      if (event.key === 'PageDown' || event.key === 'PageUp') {
        scroller.current!.scrollBy(
          navigation.type === 'step'
            ? navigation.cols * COLUMN_WIDTH * zoom
            : 0,
          navigation.type === 'step' ? navigation.rows * ROW_HEIGHT * zoom : 0,
        );
      }
      const { selection } = view.state;
      show(navigation.extend ? selection.focus : selection.active);
    }
  };

  const sheetArea = sheetSize({ zoom, layout });
  // A merged cell is edited in the whole of its box, placed on the sheet
  // that scrolls, so a frozen one is placed where the scroll has got to.
  const scrolled = scroller.current;
  const editor = areaRect(
    mergeAt(workbook.merges(sheet), active) ?? cellRange(active),
    {
      scrollX:
        active.col < layout.frozen.cols ? -(scrolled?.scrollLeft ?? 0) : 0,
      scrollY:
        active.row < layout.frozen.rows ? -(scrolled?.scrollTop ?? 0) : 0,
      zoom,
      layout: { ...layout, frozen: { rows: 0, cols: 0 } },
    },
  );
  const shownCopy =
    copied && copied.sheet === sheet ? rangeName(copied.range) : undefined;
  return (
    <ContextMenu.Root>
      <ContextMenu.Trigger asChild>
        <div
          ref={scroller}
          className={styles.grid}
          tabIndex={0}
          onKeyDown={onKeyDown}
          data-copied={shownCopy}
        >
          <div style={{ width: sheetArea.width, height: sheetArea.height }}>
            <canvas
              ref={canvas}
              className={styles.canvas}
              onMouseDown={onMouseDown}
              onMouseMove={(event) => {
                const border = onBorder(event);
                canvas.current!.style.cursor = border
                  ? border.direction === 'rows'
                    ? 'row-resize'
                    : 'col-resize'
                  : onFillHandle(event)
                    ? 'crosshair'
                    : view.state.painter
                      ? 'copy'
                      : '';
              }}
              onDoubleClick={onDoubleClick}
              onContextMenu={onContextMenu}
            />
          </div>
          {entry !== null && entry.sheet === sheet && (
            // The editor keeps the browser's menu for its text.
            <div
              className={styles.editorMenu}
              onContextMenu={(event) => event.stopPropagation()}
            >
              <FormulaEditor
                view={view}
                place="cell"
                label={cellName(active.row, active.col)}
                className={styles.editorBox}
                inputClassName={styles.editor}
                mirrorClassName={styles.editorMirror}
                style={{
                  left: editor.x,
                  top: editor.y,
                  width: editor.width + 1,
                  height: editor.height + 1,
                  fontSize: FONT_SIZE * zoom,
                }}
              />
            </div>
          )}
          <div role="status" className={styles.announcement}>
            {`${cellName(active.row, active.col)} ${activeDisplay}`.trim()}
          </div>
        </div>
      </ContextMenu.Trigger>
      {menu(menuLines)}
    </ContextMenu.Root>
  );
}

/** The range of one cell. */
function cellRange({ row, col }: CellPosition): Range {
  return { top: row, left: col, bottom: row, right: col };
}

/** Where a range of cells is on the canvas. */
function areaRect({ top, left, bottom, right }: Range, viewport: Viewport) {
  const start = cellRect({ row: top, col: left }, viewport);
  const end = cellRect({ row: bottom, col: right }, viewport);
  return {
    x: start.x,
    y: start.y,
    width: end.x + end.width - start.x,
    height: end.y + end.height - start.y,
  };
}

function sameRange(a: Range, b: Range): boolean {
  return (
    a.top === b.top &&
    a.left === b.left &&
    a.bottom === b.bottom &&
    a.right === b.right
  );
}

/** The widths of the lines of borders, in CSS pixels whatever the zoom. */
const BORDER_WIDTHS: Record<Border['line'], number> = {
  thin: 1,
  medium: 2,
  thick: 3,
  dashed: 1,
  dotted: 1,
  double: 1,
};

/**
 * Draws a border on a side of a box, over its grid line: solid, in dashes
 * or dots, or as two lines.
 */
function drawBorder(
  context: CanvasRenderingContext2D,
  { line, color }: Border,
  side: 'top' | 'right' | 'bottom' | 'left',
  box: { x: number; y: number; width: number; height: number },
  automatic: string,
): void {
  context.fillStyle = color ?? automatic;
  const across = side === 'top' || side === 'bottom';
  const at = Math.round(
    side === 'top'
      ? box.y
      : side === 'bottom'
        ? box.y + box.height
        : side === 'left'
          ? box.x
          : box.x + box.width,
  );
  const from = Math.round(across ? box.x : box.y);
  const to = Math.round(across ? box.x + box.width : box.y + box.height);
  const width = BORDER_WIDTHS[line];
  const stroke = (start: number, length: number, offset: number) => {
    if (across) context.fillRect(start, at + offset, length, width);
    else context.fillRect(at + offset, start, width, length);
  };
  if (line === 'double') {
    stroke(from, to - from + 1, -1);
    stroke(from, to - from + 1, 1);
  } else if (line === 'dashed' || line === 'dotted') {
    const [dash, gap] = line === 'dashed' ? [4, 2] : [1, 1];
    for (let start = from; start <= to; start += dash + gap) {
      stroke(start, Math.min(dash, to - start + 1), 0);
    }
  } else {
    stroke(from, to - from + 1, -Math.floor((width - 1) / 2));
  }
}

/** The name of a range, as in `A1:B2`, or of its one cell. */
function rangeName({ top, left, bottom, right }: Range): string {
  const start = cellName(top, left);
  return isSingleCell({ top, left, bottom, right })
    ? start
    : `${start}:${cellName(bottom, right)}`;
}

/**
 * Draws the lines of a cell's text in its box, aligned as `align` says
 * and as its style says vertically, underlined or struck through.
 */
function drawText(
  context: CanvasRenderingContext2D,
  lines: readonly string[],
  style: Style,
  align: RowText['align'],
  box: { x: number; y: number; width: number; height: number },
  {
    zoom,
    padding,
    color,
    setFont,
  }: {
    zoom: number;
    padding: number;
    color: string;
    setFont: (style: Style) => void;
  },
): void {
  setFont(style);
  context.fillStyle = color;
  context.textAlign = align;
  const line = lineHeight(style) * zoom;
  const size = fontPixels(style) * zoom;
  const block = line * lines.length;
  const top =
    style.vertical === 'top'
      ? box.y + padding
      : style.vertical === 'middle'
        ? box.y + (box.height - block) / 2
        : box.y + box.height - padding - block;
  const x =
    align === 'left'
      ? box.x + padding
      : align === 'right'
        ? box.x + box.width - padding
        : box.x + box.width / 2;
  const thickness = Math.max(1, Math.round(size / 14));
  lines.forEach((text, index) => {
    const y = top + line * (index + 0.5) + zoom;
    context.fillText(text, x, y);
    if (!style.underline && !style.strike) return;
    const width = context.measureText(text).width;
    const left =
      align === 'left' ? x : align === 'right' ? x - width : x - width / 2;
    if (style.underline) {
      context.fillRect(left, Math.round(y + size * 0.4), width, thickness);
    }
    if (style.strike) {
      context.fillRect(left, Math.round(y - size * 0.05), width, thickness);
    }
  });
}

function draw(
  canvas: HTMLCanvasElement,
  workbook: Workbook,
  { sheet, selection, entry, copied, filling, resizing }: ViewState,
  viewport: Viewport,
  width: number,
  height: number,
): void {
  const scale = window.devicePixelRatio;
  if (canvas.width !== Math.round(width * scale)) {
    canvas.width = Math.round(width * scale);
    canvas.style.width = `${width}px`;
  }
  if (canvas.height !== Math.round(height * scale)) {
    canvas.height = Math.round(height * scale);
    canvas.style.height = `${height}px`;
  }
  const context = canvas.getContext('2d')!;
  context.setTransform(scale, 0, 0, scale, 0, 0);
  const style = getComputedStyle(canvas);
  const color = (name: string) => style.getPropertyValue(name);
  const { zoom, layout } = viewport;
  const headerWidth = ROW_HEADER_WIDTH * zoom;
  const headerHeight = COLUMN_HEADER_HEIGHT * zoom;
  const { ranges, active } = selection;
  // The active cell is drawn as its merged cell, wherever that is.
  const activeArea =
    mergeAt(workbook.merges(sheet), active) ?? cellRange(active);
  const font = `${FONT_SIZE * zoom}px ${style.fontFamily}`;

  context.fillStyle = color('--background');
  context.fillRect(0, 0, width, height);

  let cellFont = '';
  const setFont = (each: Style) => {
    const wanted = cssFont(each, zoom);
    if (wanted !== cellFont) context.font = cellFont = wanted;
  };
  const measure = (text: string, each: Style) => {
    setFont(each);
    return context.measureText(text).width / zoom;
  };
  const background = color('--background');
  // The styles of the cells read, shared by the panes.
  const looks = new Map<string, Style>();
  const lookOf = (row: number, col: number): Style => {
    const key = `${row},${col}`;
    let look = looks.get(key);
    if (!look) {
      look = workbook.cell(sheet, row, col).style;
      looks.set(key, look);
    }
    return look;
  };
  const fillOf = (row: number, col: number) =>
    lookOf(row, col).fill ?? background;
  // While a formula picks references on another sheet, the selection is
  // that of the formula's sheet.
  const selected = entry === null || entry.sheet === sheet;
  const references = shownReferences(workbook, sheet, entry);
  const shown = panes(viewport, width, height);

  // The cells of each pane, kept inside it, each pane scrolled its own way.
  for (const { viewport: at, clip, range } of shown) {
    if (clip.width <= 0 || clip.height <= 0) continue;
    const rect = (cell: CellPosition) => cellRect(cell, at);
    const rangeRect = (each: Range) => areaRect(each, at);
    // The merged cells in sight, each drawn as one cell.
    const merges = workbook
      .merges(sheet)
      .filter(
        (merge) =>
          merge.top <= range.bottom &&
          merge.bottom >= range.top &&
          merge.left <= range.right &&
          merge.right >= range.left,
      );

    // The texts of each row, and the styles of the cells read for them.
    const rows: { row: number; texts: RowText[] }[] = [];
    for (let row = range.top; row <= range.bottom; row++) {
      const texts = rowTexts(
        (col) => {
          const cell = workbook.cell(sheet, row, col);
          looks.set(`${row},${col}`, cell.style);
          return cell;
        },
        range.left,
        range.right,
        measure,
        (col) => layout.cols.size(col),
        (col) => mergeAt(merges, { row, col }) !== undefined,
      );
      rows.push({ row, texts });
    }

    context.save();
    context.beginPath();
    context.rect(clip.x, clip.y, clip.width, clip.height);
    context.clip();

    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++) {
        if (mergeAt(merges, { row, col })) continue;
        const fill = fillOf(row, col);
        if (fill === background) continue;
        const box = rect({ row, col });
        context.fillStyle = fill;
        context.fillRect(box.x, box.y, box.width, box.height);
      }
    }

    context.fillStyle = color('--selection');
    for (const each of selected ? ranges : []) {
      const box = rangeRect(each);
      context.fillRect(box.x, box.y, box.width, box.height);
    }
    const shadesActive =
      selected && (ranges.length > 1 || !sameRange(ranges[0]!, activeArea));
    if (shadesActive) {
      // The active cell is not shaded, as in Excel.
      const box = rangeRect(activeArea);
      context.fillStyle = fillOf(active.row, active.col);
      context.fillRect(box.x, box.y, box.width, box.height);
    }

    context.strokeStyle = color('--border');
    context.lineWidth = 1;
    context.beginPath();
    for (let row = range.top; row <= range.bottom + 1; row++) {
      const y = Math.round(rect({ row, col: 0 }).y) + 0.5;
      context.moveTo(0, y);
      context.lineTo(width, y);
    }
    for (let col = range.left; col <= range.right + 1; col++) {
      const x = Math.round(rect({ row: 0, col }).x) + 0.5;
      context.moveTo(x, 0);
      context.lineTo(x, height);
    }
    context.stroke();

    context.textBaseline = 'middle';
    const padding = TEXT_PADDING * zoom;
    const textColor = (look: Style, format?: string) =>
      format ??
      look.color ??
      (look.fill !== undefined ? textColorOn(look.fill) : color('--text'));

    // Each merged cell, over the grid lines inside it, with the fill and
    // text of its top-left cell.
    for (const merge of merges) {
      const box = rangeRect(merge);
      if (box.width <= 0 || box.height <= 0) continue;
      const { top: row, left: col } = merge;
      context.save();
      context.beginPath();
      context.rect(box.x + 1, box.y + 1, box.width - 1, box.height - 1);
      context.clip();
      context.fillStyle = fillOf(row, col);
      context.fillRect(box.x, box.y, box.width, box.height);
      const shaded =
        selected &&
        ranges.some((each) => overlaps(each, merge)) &&
        !(shadesActive && sameRange(merge, activeArea)) &&
        !(ranges.length === 1 && sameRange(ranges[0]!, merge));
      if (shaded) {
        context.fillStyle = color('--selection');
        context.fillRect(box.x, box.y, box.width, box.height);
      }
      const cell = workbook.cell(sheet, row, col);
      if (cell.display !== '') {
        const { lines, align } = cellText(cell, box.width / zoom, measure);
        drawText(context, lines, cell.style, align, box, {
          zoom,
          padding,
          color: textColor(cell.style, cell.color),
          setFont,
        });
      }
      context.restore();
    }
    for (const { row, texts } of rows) {
      for (const {
        col,
        lines,
        style: look,
        color: format,
        align,
        from,
        to,
      } of texts) {
        const box = rect({ row, col });
        const start = rect({ row, col: from }).x;
        const end = rect({ row, col: to + 1 }).x;
        context.save();
        context.beginPath();
        context.rect(start + 1, box.y + 1, end - start - 1, box.height - 1);
        context.clip();
        // The grid lines inside text that runs over are hidden, as in Excel.
        for (let each = from + 1; each <= to; each++) {
          const x = Math.round(rect({ row, col: each }).x);
          const shaded = (c: number) =>
            selected &&
            ranges.some(
              (r) =>
                row >= r.top && row <= r.bottom && c >= r.left && c <= r.right,
            );
          context.fillStyle = fillOf(row, each - 1);
          context.fillRect(x, box.y + 1, 1, box.height - 1);
          if (shaded(each - 1) && shaded(each)) {
            context.fillStyle = color('--selection');
            context.fillRect(x, box.y + 1, 1, box.height - 1);
          }
        }
        drawText(context, lines, look, align, box, {
          zoom,
          padding,
          color: textColor(look, format),
          setFont,
        });
        context.restore();
      }
    }

    // The borders of the cells, but not on the sides inside a merged cell.
    const automatic = color('--text');
    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++) {
        const look = lookOf(row, col);
        const sides = [
          ['top', look.borderTop],
          ['right', look.borderRight],
          ['bottom', look.borderBottom],
          ['left', look.borderLeft],
        ] as const;
        if (sides.every(([, border]) => border === undefined)) continue;
        const box = rect({ row, col });
        if (box.width === 0 || box.height === 0) continue;
        const merge = mergeAt(merges, { row, col });
        for (const [side, border] of sides) {
          if (!border) continue;
          const inside =
            merge &&
            (side === 'top'
              ? row > merge.top
              : side === 'bottom'
                ? row < merge.bottom
                : side === 'left'
                  ? col > merge.left
                  : col < merge.right);
          if (!inside) drawBorder(context, border, side, box, automatic);
        }
      }
    }

    // The ranges of the formula being typed, each in its reference's color.
    context.lineWidth = 2;
    for (const { area, color: index } of references) {
      const box = rangeRect(area);
      context.strokeStyle = color(`--reference-${index}`);
      context.strokeRect(box.x + 1, box.y + 1, box.width - 1, box.height - 1);
    }

    context.strokeStyle = color('--accent');
    context.lineWidth = 2;
    if (selected && ranges.length === 1) {
      const box = rangeRect(ranges[0]!);
      context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
      if (entry === null) {
        // The fill handle, on a gap in the border.
        const x = box.x + box.width + 0.5;
        const y = box.y + box.height + 0.5;
        const size = FILL_HANDLE;
        context.fillStyle = color('--background');
        context.fillRect(
          x - size / 2 - 1,
          y - size / 2 - 1,
          size + 2,
          size + 2,
        );
        context.fillStyle = color('--accent');
        context.fillRect(x - size / 2, y - size / 2, size, size);
      }
    } else if (selected) {
      const box = rangeRect(activeArea);
      context.lineWidth = 1;
      context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
    }

    // The range the fill handle is dragged over.
    if (filling) {
      const box = rangeRect(filling);
      context.strokeStyle = color('--muted');
      context.lineWidth = 1;
      context.setLineDash([3, 3]);
      context.strokeRect(
        box.x + 1.5,
        box.y + 1.5,
        box.width - 2,
        box.height - 2,
      );
      context.setLineDash([]);
    }

    // The cells copied, with a border that moves.
    if (copied && copied.sheet === sheet) {
      const box = rangeRect(copied.range);
      context.strokeStyle = color('--accent');
      context.lineWidth = 2;
      context.setLineDash([4, 4]);
      context.lineDashOffset = -Math.floor(performance.now() / MARCH) % 8;
      context.strokeRect(box.x + 1, box.y + 1, box.width - 1, box.height - 1);
      context.setLineDash([]);
    }
    context.restore();
  }

  // The headers, those of selected rows and columns marked, each part
  // scrolled as the pane beside it.
  const covers = (each: Range, row: number | null, col: number | null) =>
    selected &&
    (row === null || (row >= each.top && row <= each.bottom)) &&
    (col === null || (col >= each.left && col <= each.right));
  context.fillStyle = color('--surface');
  context.fillRect(0, 0, width, headerHeight);
  context.fillRect(0, 0, headerWidth, height);
  const first = shown[0]!;
  context.font = font;
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  for (const { viewport: at, clip, range } of shown) {
    const rect = (cell: CellPosition) => cellRect(cell, at);
    if (clip.x === first.clip.x && clip.height > 0) {
      context.save();
      context.beginPath();
      context.rect(0, clip.y, headerWidth, clip.height);
      context.clip();
      context.fillStyle = color('--accent-soft');
      for (let row = range.top; row <= range.bottom; row++) {
        if (!ranges.some((each) => covers(each, row, null))) continue;
        const { y, height: h } = rect({ row, col: 0 });
        context.fillRect(0, y, headerWidth, h);
      }
      context.strokeStyle = color('--border');
      context.lineWidth = 1;
      context.beginPath();
      for (let row = range.top; row <= range.bottom + 1; row++) {
        const y = Math.round(rect({ row, col: 0 }).y) + 0.5;
        context.moveTo(0, y);
        context.lineTo(headerWidth, y);
      }
      context.stroke();
      context.fillStyle = color('--muted');
      for (let row = range.top; row <= range.bottom; row++) {
        const { y, height: h } = rect({ row, col: 0 });
        if (h > 0) {
          context.fillText(String(row + 1), headerWidth / 2, y + h / 2);
        }
      }
      context.restore();
    }
    if (clip.y === first.clip.y && clip.width > 0) {
      context.save();
      context.beginPath();
      context.rect(clip.x, 0, clip.width, headerHeight);
      context.clip();
      context.fillStyle = color('--accent-soft');
      for (let col = range.left; col <= range.right; col++) {
        if (!ranges.some((each) => covers(each, null, col))) continue;
        const { x, width: w } = rect({ row: 0, col });
        context.fillRect(x, 0, w, headerHeight);
      }
      context.strokeStyle = color('--border');
      context.lineWidth = 1;
      context.beginPath();
      for (let col = range.left; col <= range.right + 1; col++) {
        const x = Math.round(rect({ row: 0, col }).x) + 0.5;
        context.moveTo(x, 0);
        context.lineTo(x, headerHeight);
      }
      context.stroke();
      context.fillStyle = color('--muted');
      for (let col = range.left; col <= range.right; col++) {
        const { x, width: w } = rect({ row: 0, col });
        if (w > 0) {
          context.fillText(columnName(col), x + w / 2, headerHeight / 2);
        }
      }
      context.restore();
    }
  }
  context.strokeStyle = color('--border');
  context.lineWidth = 1;
  context.beginPath();
  context.moveTo(0, Math.round(headerHeight) + 0.5);
  context.lineTo(width, Math.round(headerHeight) + 0.5);
  context.moveTo(Math.round(headerWidth) + 0.5, 0);
  context.lineTo(Math.round(headerWidth) + 0.5, height);
  context.stroke();

  // A line where the frozen rows and columns end.
  const last = shown[shown.length - 1]!.clip;
  context.strokeStyle = color('--muted');
  context.beginPath();
  if (layout.frozen.rows > 0) {
    context.moveTo(0, Math.round(last.y) + 0.5);
    context.lineTo(width, Math.round(last.y) + 0.5);
  }
  if (layout.frozen.cols > 0) {
    context.moveTo(Math.round(last.x) + 0.5, 0);
    context.lineTo(Math.round(last.x) + 0.5, height);
  }
  context.stroke();

  // Where the border of a row or column being resized is dragged to.
  if (resizing) {
    const { direction, index, size } = resizing;
    const start = cellRect({ row: index, col: index }, viewport);
    const at = Math.round(
      (direction === 'rows' ? start.y : start.x) + Math.max(0, size) * zoom,
    );
    context.strokeStyle = color('--accent');
    context.lineWidth = 1;
    context.setLineDash([3, 3]);
    context.beginPath();
    if (direction === 'rows') {
      context.moveTo(0, at + 0.5);
      context.lineTo(width, at + 0.5);
    } else {
      context.moveTo(at + 0.5, 0);
      context.lineTo(at + 0.5, height);
    }
    context.stroke();
    context.setLineDash([]);
  }
}

/** The references of the entry to cells of the shown sheet. */
function shownReferences(
  workbook: Workbook,
  sheet: string,
  entry: Entry | null,
): Highlight[] {
  if (!entry) return [];
  const name = workbook.sheets
    .find((each) => each.id === sheet)!
    .name.toLowerCase();
  return formulaHelp(entry.text, entry.caret).references.filter((reference) =>
    reference.sheet === undefined
      ? entry.sheet === sheet
      : reference.sheet.toLowerCase() === name,
  );
}

function overlaps(a: Range, b: Range): boolean {
  return (
    a.top <= b.bottom &&
    b.top <= a.bottom &&
    a.left <= b.right &&
    b.left <= a.right
  );
}

function isSingleCell({ top, left, bottom, right }: Range): boolean {
  return top === bottom && left === right;
}
