import * as Dialog from '@radix-ui/react-dialog';
import {
  ArrowDownAZ,
  ArrowDownZA,
  ArrowUpDown,
  Trash2,
  type LucideIcon,
} from 'lucide-react';
import { useCallback, useId, useState, useSyncExternalStore } from 'react';
import {
  columnName,
  type Range,
  type SortKey,
  type Workbook,
} from '../../engine/workbook';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import { rangesText } from './ranges';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

/**
 * Sort A to Z and Z to A by the active cell's column, and Custom sort for a
 * dialog that sorts by several columns, with a header row or none (R23).
 * Each is enabled when the selection can be sorted.
 */
export function SortControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
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
  // A string, so that the same range gives the same snapshot.
  const rangeText = useSyncExternalStore(subscribe, () =>
    JSON.stringify(view.sortRange),
  );
  const range = JSON.parse(rangeText) as Range | null;
  const [open, setOpen] = useState(false);
  return (
    <>
      <SortButton
        icon={ArrowDownAZ}
        label={t('sort.ascending')}
        tooltip={t('sort.ascending.tooltip')}
        disabled={!range}
        onClick={() => view.sortByActive(false)}
      />
      <SortButton
        icon={ArrowDownZA}
        label={t('sort.descending')}
        tooltip={t('sort.descending.tooltip')}
        disabled={!range}
        onClick={() => view.sortByActive(true)}
      />
      <SortButton
        icon={ArrowUpDown}
        label={t('sort.custom')}
        tooltip={t('sort.custom.tooltip')}
        disabled={!range}
        onClick={() => setOpen(true)}
      />
      <Dialog.Root open={open && range !== null} onOpenChange={setOpen}>
        <Dialog.Portal>
          <Dialog.Overlay className={styles.overlay} />
          <Dialog.Content className={`${styles.dialog} ${styles.wideDialog}`}>
            <Dialog.Title className={styles.dialogTitle}>
              {t('sort')}
            </Dialog.Title>
            {range && (
              <SortForm
                workbook={workbook}
                sheet={view.state.sheet}
                range={range}
                active={view.state.selection.active.col}
                onSort={(keys, headers) => {
                  view.sort(keys, headers);
                  setOpen(false);
                }}
                onCancel={() => setOpen(false)}
              />
            )}
          </Dialog.Content>
        </Dialog.Portal>
      </Dialog.Root>
    </>
  );
}

function SortButton({
  icon: Icon,
  label,
  tooltip,
  disabled,
  onClick,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  disabled: boolean;
  onClick: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon size={ICON_SIZE} aria-hidden />
        {label}
      </button>
    </Tooltip>
  );
}

/**
 * The levels of a sort, each a column and an order, and whether the first
 * row holds headers, which then name the columns.
 */
function SortForm({
  workbook,
  sheet,
  range,
  active,
  onSort,
  onCancel,
}: {
  workbook: Workbook;
  sheet: string;
  range: Range;
  /** The active cell's column, sorted by first. */
  active: number;
  onSort: (keys: SortKey[], headers: boolean) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [headers, setHeaders] = useState(false);
  const [keys, setKeys] = useState<SortKey[]>([{ col: active }]);
  const columns = Array.from(
    { length: range.right - range.left + 1 },
    (_, index) => range.left + index,
  );
  const columnLabel = (col: number) => {
    const header = headers && workbook.cell(sheet, range.top, col).display;
    return header || t('sort.column').replace('{name}', columnName(col));
  };
  const change = (index: number, key: SortKey) =>
    setKeys(keys.map((each, at) => (at === index ? key : each)));
  // The next column not sorted by yet, as Excel adds a level.
  const next = columns.find((col) => !keys.some((key) => key.col === col));
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        onSort(keys, headers);
      }}
    >
      <Dialog.Description className={styles.dialogText}>
        {t('sort.description').replace('{range}', rangesText([range]))}
      </Dialog.Description>
      <label className={styles.checkField}>
        <input
          type="checkbox"
          className={styles.checkbox}
          checked={headers}
          onChange={(event) => setHeaders(event.target.checked)}
        />
        {t('sort.headers')}
      </label>
      <ul className={styles.ruleList}>
        {keys.map((key, index) => (
          <li key={index} className={styles.sortLevel}>
            <label className={styles.sortLabel} htmlFor={`${id}-${index}`}>
              {t(index === 0 ? 'sort.by' : 'sort.then')}
            </label>
            <select
              id={`${id}-${index}`}
              className={`${styles.textInput} ${styles.ruleSelect}`}
              value={key.col}
              onChange={(event) =>
                change(index, { ...key, col: Number(event.target.value) })
              }
            >
              {columns.map((col) => (
                <option key={col} value={col}>
                  {columnLabel(col)}
                </option>
              ))}
            </select>
            <select
              aria-label={t('sort.order')}
              className={`${styles.textInput} ${styles.ruleSelect}`}
              value={key.descending ? 'descending' : 'ascending'}
              onChange={(event) =>
                change(index, {
                  col: key.col,
                  ...(event.target.value === 'descending' && {
                    descending: true,
                  }),
                })
              }
            >
              <option value="ascending">{t('sort.az')}</option>
              <option value="descending">{t('sort.za')}</option>
            </select>
            <Tooltip text={t('sort.delete.tooltip')}>
              <button
                type="button"
                className={styles.iconButton}
                aria-label={t('sort.delete')}
                disabled={keys.length === 1}
                onClick={() => setKeys(keys.filter((_, at) => at !== index))}
              >
                <Trash2 size={ICON_SIZE} aria-hidden />
              </button>
            </Tooltip>
          </li>
        ))}
      </ul>
      <div className={styles.dialogButtons}>
        <Tooltip text={t('sort.add.tooltip')}>
          <button
            type="button"
            className={styles.button}
            disabled={next === undefined}
            onClick={() =>
              next !== undefined && setKeys([...keys, { col: next }])
            }
          >
            {t('sort.add')}
          </button>
        </Tooltip>
        <button type="button" className={styles.button} onClick={onCancel}>
          {t('cancel')}
        </button>
        <button type="submit" className={`${styles.button} ${styles.primary}`}>
          {t('ok')}
        </button>
      </div>
    </form>
  );
}
