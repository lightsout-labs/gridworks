import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  Baseline,
  Bold,
  Check,
  ChevronDown,
  Grid2x2,
  Italic,
  PaintBucket,
  PanelBottom,
  PanelLeft,
  PanelRight,
  PanelTop,
  Square,
  SquareDashed,
  Strikethrough,
  TableCellsMerge,
  TableCellsSplit,
  TextAlignCenter,
  TextAlignEnd,
  TextAlignStart,
  TextWrap,
  Underline,
  type LucideIcon,
} from 'lucide-react';
import {
  useCallback,
  useEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from 'react';
import {
  BORDER_LINES,
  DEFAULT_FONT_SIZE,
  type Border,
  type BorderSides,
  FONT_SIZES,
  FONTS,
  type Style,
  type StyleChange,
  type Workbook,
} from '../../engine/workbook';
import { MAC } from '../grid/Grid';
import { cssFont } from '../grid/text';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

/** The colors offered for text and fills, by the key of their name. */
const PALETTE: readonly [string, Key][] = [
  ['#000000', 'color.black'],
  ['#7f7f7f', 'color.gray'],
  ['#ffffff', 'color.white'],
  ['#c00000', 'color.darkRed'],
  ['#ff0000', 'color.red'],
  ['#ffc000', 'color.orange'],
  ['#ffff00', 'color.yellow'],
  ['#92d050', 'color.lightGreen'],
  ['#00b050', 'color.green'],
  ['#00b0f0', 'color.lightBlue'],
  ['#0070c0', 'color.blue'],
  ['#002060', 'color.darkBlue'],
  ['#7030a0', 'color.purple'],
];

/**
 * A value read from the workbook and the view, which renders again when
 * either changes.
 */
function useWorkbookView<T>(workbook: Workbook, view: View, read: () => T): T {
  const subscribe = useCallback(
    (listener: () => void) => {
      const unsubscribeWorkbook = workbook.subscribe(listener);
      const unsubscribeView = view.subscribe(listener);
      return () => {
        unsubscribeWorkbook();
        unsubscribeView();
      };
    },
    [workbook, view],
  );
  return useSyncExternalStore(subscribe, read);
}

/** The style of the active cell, which the controls show. */
function useActiveStyle(workbook: Workbook, view: View): Style {
  return useWorkbookView(workbook, view, () => {
    const { sheet, selection } = view.state;
    const { row, col } = selection.active;
    return workbook.cell(sheet, row, col).style;
  });
}

/** The sides of the Borders menu, as Excel orders them, with their icons. */
const SIDES: readonly [BorderSides | 'none', Key, LucideIcon][] = [
  ['bottom', 'borders.bottom', PanelBottom],
  ['top', 'borders.top', PanelTop],
  ['left', 'borders.left', PanelLeft],
  ['right', 'borders.right', PanelRight],
  ['none', 'borders.none', SquareDashed],
  ['all', 'borders.all', Grid2x2],
  ['outside', 'borders.outside', Square],
];

/** How each line looks in the menu, as a CSS border. */
const LINE_PREVIEWS: Record<Border['line'], string> = {
  thin: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  dashed: '1px dashed',
  dotted: '1px dotted',
  double: '3px double',
};

/**
 * The Borders menu: the sides to draw on the selected cells, or none, and
 * the line and color to draw them with, which it keeps for the next time.
 */
export function BorderControls({ view }: { view: View }) {
  const t = useTranslation();
  const [line, setLine] = useState<Border['line']>('thin');
  const [color, setColor] = useState<string | undefined>();
  const keepOpen = (event: Event) => event.preventDefault();
  return (
    <DropdownMenu.Root>
      <Tooltip text={t('borders.tooltip')}>
        <DropdownMenu.Trigger asChild>
          <button
            type="button"
            className={styles.button}
            aria-label={t('borders')}
          >
            <span className={styles.colorIcon}>
              <Grid2x2 size={ICON_SIZE} aria-hidden />
              <span
                className={styles.colorBar}
                style={{ background: color ?? 'var(--text)' }}
              />
            </span>
            <ChevronDown size={12} aria-hidden />
          </button>
        </DropdownMenu.Trigger>
      </Tooltip>
      <DropdownMenu.Portal>
        <DropdownMenu.Content
          className={`${styles.menu} ${styles.scrollingMenu}`}
          align="start"
          sideOffset={4}
        >
          {SIDES.map(([sides, name, Icon]) => (
            <DropdownMenu.Item
              key={sides}
              className={styles.menuItem}
              onSelect={() =>
                sides === 'none'
                  ? view.setBorders('all', null)
                  : view.setBorders(sides, {
                      line,
                      ...(color !== undefined && { color }),
                    })
              }
            >
              <Icon size={ICON_SIZE} aria-hidden />
              {t(name)}
            </DropdownMenu.Item>
          ))}
          <DropdownMenu.Separator className={styles.menuSeparator} />
          <DropdownMenu.Label className={styles.menuLabel}>
            {t('borders.lineStyle')}
          </DropdownMenu.Label>
          <DropdownMenu.RadioGroup
            value={line}
            onValueChange={(value) => setLine(value as Border['line'])}
          >
            {BORDER_LINES.map((each) => (
              <DropdownMenu.RadioItem
                key={each}
                value={each}
                className={styles.menuItem}
                aria-label={t(`line.${each}`)}
                onSelect={keepOpen}
              >
                <span
                  className={styles.linePreview}
                  style={{ borderTop: LINE_PREVIEWS[each] }}
                  aria-hidden
                />
                <DropdownMenu.ItemIndicator className={styles.menuCheck}>
                  <Check size={ICON_SIZE} aria-hidden />
                </DropdownMenu.ItemIndicator>
              </DropdownMenu.RadioItem>
            ))}
          </DropdownMenu.RadioGroup>
          <DropdownMenu.Separator className={styles.menuSeparator} />
          <DropdownMenu.Label className={styles.menuLabel}>
            {t('borders.lineColor')}
          </DropdownMenu.Label>
          <DropdownMenu.Item
            className={styles.menuItem}
            onSelect={(event) => {
              keepOpen(event);
              setColor(undefined);
            }}
          >
            {t('color.automatic')}
            {color === undefined && (
              <Check
                size={ICON_SIZE}
                className={styles.menuCheck}
                aria-hidden
              />
            )}
          </DropdownMenu.Item>
          <div className={styles.palette}>
            {PALETTE.map(([value, name]) => (
              <Tooltip key={value} text={t(name)}>
                <DropdownMenu.Item
                  className={styles.swatch}
                  style={{ background: value }}
                  aria-label={t(name)}
                  data-chosen={color === value || undefined}
                  onSelect={(event) => {
                    keepOpen(event);
                    setColor(value);
                  }}
                />
              </Tooltip>
            ))}
          </div>
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}

/** Merge cells and Unmerge cells, each enabled when it can do something. */
export function MergeControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const canMerge = useWorkbookView(workbook, view, () => view.canMerge);
  const canUnmerge = useWorkbookView(workbook, view, () => view.canUnmerge);
  return (
    <>
      <Tooltip text={t('merge.tooltip')}>
        <button
          type="button"
          className={styles.button}
          aria-label={t('merge')}
          disabled={!canMerge}
          onClick={() => view.merge()}
        >
          <TableCellsMerge size={ICON_SIZE} aria-hidden />
        </button>
      </Tooltip>
      <Tooltip text={t('unmerge.tooltip')}>
        <button
          type="button"
          className={styles.button}
          aria-label={t('unmerge')}
          disabled={!canUnmerge}
          onClick={() => view.unmerge()}
        >
          <TableCellsSplit size={ICON_SIZE} aria-hidden />
        </button>
      </Tooltip>
    </>
  );
}

/** The font, its size and look, and the colors of the selected cells. */
export function FontControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const style = useActiveStyle(workbook, view);
  const key = MAC ? '⌘' : 'Ctrl+';
  const format = (change: StyleChange) => view.format(change);
  return (
    <>
      <Menu
        label={t('font.family')}
        tooltip={t('font.family.tooltip')}
        shown={style.font ?? t('font.default')}
        wide
      >
        <DropdownMenu.RadioGroup
          value={style.font ?? ''}
          onValueChange={(font) => format({ font: font || null })}
        >
          {[['', t('font.default')] as const, ...FONTS.map((f) => [f, f])].map(
            ([font, name]) => (
              <DropdownMenu.RadioItem
                key={font}
                value={font}
                className={styles.menuItem}
                style={{ font: cssFont(font ? { font } : {}, 1) }}
              >
                {name}
              </DropdownMenu.RadioItem>
            ),
          )}
        </DropdownMenu.RadioGroup>
      </Menu>
      <Menu
        label={t('font.size')}
        tooltip={t('font.size.tooltip')}
        shown={String(style.size ?? DEFAULT_FONT_SIZE)}
      >
        <DropdownMenu.RadioGroup
          value={String(style.size ?? DEFAULT_FONT_SIZE)}
          onValueChange={(size) =>
            format({
              size: Number(size) === DEFAULT_FONT_SIZE ? null : Number(size),
            })
          }
        >
          {FONT_SIZES.map((size) => (
            <DropdownMenu.RadioItem
              key={size}
              value={String(size)}
              className={styles.menuItem}
            >
              {size}
            </DropdownMenu.RadioItem>
          ))}
        </DropdownMenu.RadioGroup>
      </Menu>
      <Toggle
        icon={Bold}
        label={t('bold')}
        tooltip={`${t('bold.tooltip')} (${key}B)`}
        on={style.bold === true}
        onToggle={() => view.toggle('bold')}
      />
      <Toggle
        icon={Italic}
        label={t('italic')}
        tooltip={`${t('italic.tooltip')} (${key}I)`}
        on={style.italic === true}
        onToggle={() => view.toggle('italic')}
      />
      <Toggle
        icon={Underline}
        label={t('underline')}
        tooltip={`${t('underline.tooltip')} (${key}U)`}
        on={style.underline === true}
        onToggle={() => view.toggle('underline')}
      />
      <Toggle
        icon={Strikethrough}
        label={t('strike')}
        tooltip={t('strike.tooltip')}
        on={style.strike === true}
        onToggle={() => view.toggle('strike')}
      />
      <ColorMenu
        icon={Baseline}
        label={t('color.text')}
        tooltip={t('color.text.tooltip')}
        color={style.color}
        none={t('color.automatic')}
        onColor={(color) => format({ color })}
      />
      <ColorMenu
        icon={PaintBucket}
        label={t('color.fill')}
        tooltip={t('color.fill.tooltip')}
        color={style.fill}
        none={t('color.noFill')}
        onColor={(fill) => format({ fill })}
      />
    </>
  );
}

/** The horizontal and vertical alignment of the selected cells, and wrap. */
export function AlignmentControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const style = useActiveStyle(workbook, view);
  const horizontal = (
    icon: LucideIcon,
    align: NonNullable<Style['horizontal']>,
  ) => (
    <Toggle
      icon={icon}
      label={t(`align.${align}`)}
      tooltip={t(`align.${align}.tooltip`)}
      on={style.horizontal === align}
      // Turned off, the cells align by the type of their values again.
      onToggle={() =>
        view.format({ horizontal: style.horizontal === align ? null : align })
      }
    />
  );
  const vertical = style.vertical ?? 'bottom';
  const verticalToggle = (
    icon: LucideIcon,
    align: 'top' | 'middle' | 'bottom',
  ) => (
    <Toggle
      icon={icon}
      label={t(`align.${align}`)}
      tooltip={t(`align.${align}.tooltip`)}
      on={vertical === align}
      onToggle={() =>
        view.format({ vertical: align === 'bottom' ? null : align })
      }
    />
  );
  return (
    <>
      {horizontal(TextAlignStart, 'left')}
      {horizontal(TextAlignCenter, 'center')}
      {horizontal(TextAlignEnd, 'right')}
      {verticalToggle(AlignVerticalJustifyStart, 'top')}
      {verticalToggle(AlignVerticalJustifyCenter, 'middle')}
      {verticalToggle(AlignVerticalJustifyEnd, 'bottom')}
      <Toggle
        icon={TextWrap}
        label={t('wrap')}
        tooltip={t('wrap.tooltip')}
        on={style.wrap === true}
        onToggle={() => view.toggle('wrap')}
      />
    </>
  );
}

/** A button that turns a property of the selected cells on or off. */
function Toggle({
  icon: Icon,
  label,
  tooltip,
  on,
  onToggle,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  on: boolean;
  onToggle: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-label={label}
        aria-pressed={on}
        onClick={onToggle}
      >
        <Icon size={ICON_SIZE} aria-hidden />
      </button>
    </Tooltip>
  );
}

/** A button that shows a value and opens a menu of the others. */
function Menu({
  label,
  tooltip,
  shown,
  wide = false,
  children,
}: {
  label: string;
  tooltip: string;
  shown: string;
  wide?: boolean;
  children: ReactNode;
}) {
  return (
    <DropdownMenu.Root>
      <Tooltip text={tooltip}>
        <DropdownMenu.Trigger asChild>
          <button
            type="button"
            className={`${styles.button} ${styles.picker} ${wide ? styles.wide : ''}`}
            aria-label={label}
          >
            <span className={styles.pickerValue}>{shown}</span>
            <ChevronDown size={ICON_SIZE} aria-hidden />
          </button>
        </DropdownMenu.Trigger>
      </Tooltip>
      <DropdownMenu.Portal>
        <DropdownMenu.Content
          className={`${styles.menu} ${styles.scrollingMenu}`}
          align="start"
          sideOffset={4}
        >
          {children}
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}

/**
 * A button that opens a palette of colors, the default one (`none`), and
 * a custom color, and shows the color chosen under its icon.
 */
function ColorMenu({
  icon: Icon,
  label,
  tooltip,
  color,
  none,
  onColor,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  color: string | undefined;
  none: string;
  onColor: (color: string | null) => void;
}) {
  const t = useTranslation();
  const latest = useRef(onColor);
  useEffect(() => {
    latest.current = onColor;
  });
  // The color picker's input stays outside the menu, which closes when the
  // picker opens. The color is taken once it is chosen, not at each step
  // of the picker, so it listens to `change` rather than React's `onChange`.
  const custom = useRef<HTMLInputElement>(null);
  useEffect(() => {
    const input = custom.current!;
    const choose = () => latest.current(input.value);
    input.addEventListener('change', choose);
    return () => input.removeEventListener('change', choose);
  }, []);
  return (
    <>
      <input
        ref={custom}
        type="color"
        className={styles.colorInput}
        aria-label={`${label}: ${t('color.custom')}`}
        tabIndex={-1}
        defaultValue={color ?? '#000000'}
      />
      <DropdownMenu.Root>
        <Tooltip text={tooltip}>
          <DropdownMenu.Trigger asChild>
            <button type="button" className={styles.button} aria-label={label}>
              <span className={styles.colorIcon}>
                <Icon size={ICON_SIZE} aria-hidden />
                <span
                  className={styles.colorBar}
                  style={{ background: color ?? 'transparent' }}
                />
              </span>
              <ChevronDown size={12} aria-hidden />
            </button>
          </DropdownMenu.Trigger>
        </Tooltip>
        <DropdownMenu.Portal>
          <DropdownMenu.Content
            className={styles.menu}
            align="start"
            sideOffset={4}
          >
            <DropdownMenu.Item
              className={styles.menuItem}
              onSelect={() => onColor(null)}
            >
              {none}
            </DropdownMenu.Item>
            <div className={styles.palette}>
              {PALETTE.map(([value, name]) => (
                <Tooltip key={value} text={t(name)}>
                  <DropdownMenu.Item
                    className={styles.swatch}
                    style={{ background: value }}
                    aria-label={t(name)}
                    data-chosen={color === value || undefined}
                    onSelect={() => onColor(value)}
                  />
                </Tooltip>
              ))}
            </div>
            <DropdownMenu.Item
              className={styles.menuItem}
              onSelect={() => custom.current!.click()}
            >
              {t('color.custom')}
            </DropdownMenu.Item>
          </DropdownMenu.Content>
        </DropdownMenu.Portal>
      </DropdownMenu.Root>
    </>
  );
}
