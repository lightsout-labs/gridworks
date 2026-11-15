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
import { pick } from './selection';
import {
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
  const drag = useRef<Target['kind'] | 'fill' | CellPosition | null>(null);
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
      if (typeof drag.current === 'object') {
        const corner = drag.current;
        view.pick({
          top: Math.min(corner.row, row),
          left: Math.min(corner.col, col),
          bottom: Math.max(corner.row, row),
          right: Math.max(corner.col, col),
        });
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
    if (target.kind === 'cell' && view.canPick) {
      const { row, col } = target;
      view.pick({ top: row, left: col, bottom: row, right: col });
      drag.current = { row, col };
      // The editor keeps the keyboard.
      event.preventDefault();
      return;
    }
    const mode = event.shiftKey
      ? 'extend'
      : (MAC ? event.metaKey : event.ctrlKey)
        ? 'add'
        : 'replace';
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
      { mac: MAC, mode: entry?.mode ?? null, endMode: state.endMode, page },
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
  const editor = cellRect(active, { scrollX: 0, scrollY: 0, zoom, layout });
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
  const range = visibleRange(width, height, viewport);
  const { ranges, active } = selection;
  const rect = (cell: CellPosition) => cellRect(cell, viewport);
  const rangeRect = ({ top, left, bottom, right }: Range) => {
    const start = rect({ row: top, col: left });
    const end = rect({ row: bottom, col: right });
    return {
      x: start.x,
      y: start.y,
      width: end.x + end.width - start.x,
      height: end.y + end.height - start.y,
    };
  };
  const font = `${FONT_SIZE * zoom}px ${style.fontFamily}`;

  context.fillStyle = color('--background');
  context.fillRect(0, 0, width, height);

  // The texts of each row, and the fills of the cells read for them.
  let cellFont = '';
  const setFont = (each: Style) => {
    const wanted = cssFont(each, zoom);
    if (wanted !== cellFont) context.font = cellFont = wanted;
  };
  const measure = (text: string, each: Style) => {
    setFont(each);
    return context.measureText(text).width / zoom;
  };
  const fills = new Map<string, string>();
  const rows: { row: number; texts: RowText[] }[] = [];
  for (let row = range.top; row <= range.bottom; row++) {
    const texts = rowTexts(
      (col) => {
        const cell = workbook.cell(sheet, row, col);
        if (cell.style.fill !== undefined && col >= range.left) {
          fills.set(`${row},${col}`, cell.style.fill);
        }
        return cell;
      },
      range.left,
      range.right,
      measure,
      (col) => layout.cols.size(col),
    );
    rows.push({ row, texts });
  }
  // Text that runs over reads the cells it covers only to the first that
  // is not empty, so the fills of the empty ones are read here.
  const fillOf = (row: number, col: number) => {
    const key = `${row},${col}`;
    if (!fills.has(key)) {
      fills.set(key, workbook.cell(sheet, row, col).style.fill ?? '');
    }
    return fills.get(key) || color('--background');
  };

  // The cells, kept out of the headers.
  context.save();
  context.beginPath();
  context.rect(headerWidth, headerHeight, width, height);
  context.clip();

  for (let row = range.top; row <= range.bottom; row++) {
    for (let col = range.left; col <= range.right; col++) {
      const fill = fillOf(row, col);
      if (fill === color('--background')) continue;
      const box = rect({ row, col });
      context.fillStyle = fill;
      context.fillRect(box.x, box.y, box.width, box.height);
    }
  }

  // While a formula picks references on another sheet, the selection is
  // that of the formula's sheet.
  const selected = entry === null || entry.sheet === sheet;
  context.fillStyle = color('--selection');
  for (const each of selected ? ranges : []) {
    const box = rangeRect(each);
    context.fillRect(box.x, box.y, box.width, box.height);
  }
  if (selected && (ranges.length > 1 || !isSingleCell(ranges[0]!))) {
    // The active cell is not shaded, as in Excel.
    const box = rect(active);
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
  for (const { row, texts } of rows) {
    for (const { col, lines, style: look, align, from, to } of texts) {
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
        color:
          look.color ??
          (look.fill !== undefined ? textColorOn(look.fill) : color('--text')),
        setFont,
      });
      context.restore();
    }
  }

  // The ranges of the formula being typed, each in its reference's color.
  context.lineWidth = 2;
  for (const { area, color: index } of shownReferences(
    workbook,
    sheet,
    entry,
  )) {
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
      context.fillRect(x - size / 2 - 1, y - size / 2 - 1, size + 2, size + 2);
      context.fillStyle = color('--accent');
      context.fillRect(x - size / 2, y - size / 2, size, size);
    }
  } else if (selected) {
    const box = rect(active);
    context.lineWidth = 1;
    context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
  }

  // The range the fill handle is dragged over.
  if (filling) {
    const box = rangeRect(filling);
    context.strokeStyle = color('--muted');
    context.lineWidth = 1;
    context.setLineDash([3, 3]);
    context.strokeRect(box.x + 1.5, box.y + 1.5, box.width - 2, box.height - 2);
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

  // The headers, those of selected rows and columns marked.
  const covers = (each: Range, row: number | null, col: number | null) =>
    selected &&
    (row === null || (row >= each.top && row <= each.bottom)) &&
    (col === null || (col >= each.left && col <= each.right));
  context.fillStyle = color('--surface');
  context.fillRect(0, 0, width, headerHeight);
  context.fillRect(0, 0, headerWidth, height);
  context.fillStyle = color('--accent-soft');
  for (let row = range.top; row <= range.bottom; row++) {
    if (!ranges.some((each) => covers(each, row, null))) continue;
    const { y, height: h } = rect({ row, col: 0 });
    context.fillRect(0, Math.max(y, headerHeight), headerWidth, h);
  }
  for (let col = range.left; col <= range.right; col++) {
    if (!ranges.some((each) => covers(each, null, col))) continue;
    const { x, width: w } = rect({ row: 0, col });
    context.fillRect(Math.max(x, headerWidth), 0, w, headerHeight);
  }

  context.strokeStyle = color('--border');
  context.lineWidth = 1;
  context.beginPath();
  for (let row = range.top; row <= range.bottom + 1; row++) {
    const y = Math.round(rect({ row, col: 0 }).y) + 0.5;
    if (y < headerHeight) continue;
    context.moveTo(0, y);
    context.lineTo(headerWidth, y);
  }
  for (let col = range.left; col <= range.right + 1; col++) {
    const x = Math.round(rect({ row: 0, col }).x) + 0.5;
    if (x < headerWidth) continue;
    context.moveTo(x, 0);
    context.lineTo(x, headerHeight);
  }
  context.moveTo(0, Math.round(headerHeight) + 0.5);
  context.lineTo(width, Math.round(headerHeight) + 0.5);
  context.moveTo(Math.round(headerWidth) + 0.5, 0);
  context.lineTo(Math.round(headerWidth) + 0.5, height);
  context.stroke();

  context.font = font;
  context.textAlign = 'center';
  context.fillStyle = color('--muted');
  context.save();
  context.beginPath();
  context.rect(0, headerHeight, headerWidth, height);
  context.clip();
  for (let row = range.top; row <= range.bottom; row++) {
    const { y, height: h } = rect({ row, col: 0 });
    if (h > 0) context.fillText(String(row + 1), headerWidth / 2, y + h / 2);
  }
  context.restore();
  context.save();
  context.beginPath();
  context.rect(headerWidth, 0, width, headerHeight);
  context.clip();
  for (let col = range.left; col <= range.right; col++) {
    const { x, width: w } = rect({ row: 0, col });
    if (w > 0) context.fillText(columnName(col), x + w / 2, headerHeight / 2);
  }
  context.restore();

  // Where the border of a row or column being resized is dragged to.
  if (resizing) {
    const { direction, index, size } = resizing;
    const start = rect({ row: index, col: index });
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

function isSingleCell({ top, left, bottom, right }: Range): boolean {
  return top === bottom && left === right;
}
