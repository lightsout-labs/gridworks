import * as Dialog from '@radix-ui/react-dialog';
import { ChevronDown, Funnel, FunnelX } from 'lucide-react';
import {
  useCallback,
  useId,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import {
  columnName,
  type ColumnFilter,
  type NumberTest,
  type TextTest,
  type Workbook,
} from '../../engine/workbook';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;
/** The width of the filter dropdown, which stays inside the window. */
const PANEL_WIDTH = 260;

/** Calls a listener on every change to the workbook or the view. */
function useChanges(workbook: Workbook, view: View) {
  return useCallback(
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
}

/**
 * The Data tab's Filter button, which turns the sheet's filter on over the
 * selection or off, and Clear, which shows every row it hides (R24).
 */
export function FilterControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const subscribe = useChanges(workbook, view);
  const on = useSyncExternalStore(
    subscribe,
    () => workbook.filter(view.state.sheet) !== null,
  );
  const canFilter = useSyncExternalStore(subscribe, () => view.canFilter);
  const canClear = useSyncExternalStore(subscribe, () => view.canClearFilter);
  return (
    <>
      <Tooltip text={t('filter.tooltip')}>
        <button
          type="button"
          className={styles.button}
          aria-pressed={on}
          disabled={!canFilter}
          onClick={() => view.toggleFilter()}
        >
          <Funnel size={ICON_SIZE} aria-hidden />
          {t('filter')}
        </button>
      </Tooltip>
      <Tooltip text={t('filter.clear.tooltip')}>
        <button
          type="button"
          className={styles.button}
          disabled={!canClear}
          onClick={() => view.clearFilter()}
        >
          <FunnelX size={ICON_SIZE} aria-hidden />
          {t('filter.clear')}
        </button>
      </Tooltip>
    </>
  );
}

/**
 * The dropdown button in the header of a column of the sheet's filter,
 * showing whether the column is filtered, and the panel it opens.
 */
export function FilterButton({
  workbook,
  view,
  col,
}: {
  workbook: Workbook;
  view: View;
  col: number;
}) {
  const t = useTranslation();
  const subscribe = useChanges(workbook, view);
  const column = useSyncExternalStore(
    subscribe,
    () =>
      workbook
        .filter(view.state.sheet)
        ?.columns.find((each) => each.col === col) ?? null,
  );
  const button = useRef<HTMLButtonElement>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(
    null,
  );
  const title = t('filter.column').replace('{name}', columnName(col));
  const label = column ? `${title} ${t('filter.filtered')}` : title;
  const Icon = column ? Funnel : ChevronDown;
  return (
    <Dialog.Root
      open={place !== null}
      onOpenChange={(open) => {
        if (!open) return setPlace(null);
        const bounds = button.current!.getBoundingClientRect();
        setPlace({
          left: Math.max(
            8,
            Math.min(bounds.right - PANEL_WIDTH, innerWidth - PANEL_WIDTH - 8),
          ),
          top: bounds.bottom + 4,
        });
      }}
    >
      <Tooltip text={label}>
        <Dialog.Trigger asChild>
          <button
            ref={button}
            type="button"
            className={styles.filterButton}
            aria-label={label}
            data-filtered={column !== null}
          >
            <Icon size={12} aria-hidden />
          </button>
        </Dialog.Trigger>
      </Tooltip>
      <Dialog.Portal>
        <Dialog.Content
          className={styles.filterPanel}
          style={{ ...place, width: PANEL_WIDTH }}
          aria-describedby={undefined}
        >
          <Dialog.Title className={styles.dialogTitle}>{title}</Dialog.Title>
          <FilterForm
            values={() => workbook.filterValues(view.state.sheet, col)}
            column={column}
            onFilter={(filter) => {
              if (filter) view.filterBy({ ...filter, col });
              else view.unfilter(col);
              setPlace(null);
            }}
            onCancel={() => setPlace(null)}
          />
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

type Kind = 'values' | `text.${TextTest}` | `number.${NumberTest}`;

const TEXT_TESTS: readonly TextTest[] = [
  'equals',
  'contains',
  'begins',
  'ends',
];
const NUMBER_TESTS: readonly NumberTest[] = [
  'equals',
  'greater',
  'less',
  'between',
];

/** A column filter without its column, which the button knows. */
type Criterion = DistributiveOmit<ColumnFilter, 'col'>;
type DistributiveOmit<T, K extends keyof T> = T extends unknown
  ? Omit<T, K>
  : never;

/**
 * What a column is filtered by: a checklist of its values, with a search
 * box, or a text or number condition.
 */
function FilterForm({
  values: readValues,
  column,
  onFilter,
  onCancel,
}: {
  /**
   * The column's values that the other columns let through, read once when
   * the panel opens rather than on every render of the button.
   */
  values: () => readonly string[];
  column: ColumnFilter | null;
  /** Filters by a criterion, or by nothing. */
  onFilter: (criterion: Criterion | null) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [values] = useState(readValues);
  const [kind, setKind] = useState<Kind>(
    column && column.type !== 'values'
      ? (`${column.type}.${column.test}` as Kind)
      : 'values',
  );
  const [text, setText] = useState(
    column?.type === 'text'
      ? column.text
      : column?.type === 'number'
        ? String(column.value)
        : '',
  );
  const [to, setTo] = useState(
    column?.type === 'number' && column.to !== undefined
      ? String(column.to)
      : '',
  );
  const [search, setSearch] = useState('');
  const [checked, setChecked] = useState<ReadonlySet<string>>(
    () => new Set(column?.type === 'values' ? column.values : values),
  );
  const shown = values.filter((value) =>
    value.toLowerCase().includes(search.toLowerCase()),
  );
  const allShown = shown.every((value) => checked.has(value));
  const number = (input: string) => (input.trim() === '' ? NaN : Number(input));
  /** Whether a number typed is not one; nothing typed is not yet wrong. */
  const wrong = (input: string) =>
    input.trim() !== '' && Number.isNaN(number(input));
  const [type, test] = kind.split('.') as ['text' | 'number', string];
  const invalid =
    type === 'number' &&
    (Number.isNaN(number(text)) ||
      (test === 'between' && Number.isNaN(number(to))));

  const criterion = (): Criterion | null => {
    if (kind === 'values') {
      // A search keeps only the values it finds, as in Excel.
      const chosen = shown.filter((value) => checked.has(value));
      return chosen.length === values.length
        ? null
        : { type: 'values', values: chosen };
    }
    if (type === 'text') {
      return { type, test: test as TextTest, text };
    }
    return {
      type,
      test: test as NumberTest,
      value: number(text),
      ...(test === 'between' && { to: number(to) }),
    };
  };
  const toggle = (value: string, on: boolean) => {
    const next = new Set(checked);
    if (on) next.add(value);
    else next.delete(value);
    setChecked(next);
  };
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        if (!invalid) onFilter(criterion());
      }}
    >
      <Tooltip text={t('filter.clearColumn.tooltip')}>
        <button
          type="button"
          className={styles.button}
          disabled={!column}
          onClick={() => onFilter(null)}
        >
          <FunnelX size={ICON_SIZE} aria-hidden />
          {t('filter.clearColumn')}
        </button>
      </Tooltip>
      <label className={styles.field} htmlFor={`${id}-kind`}>
        {t('filter.where')}
      </label>
      <select
        id={`${id}-kind`}
        className={styles.textInput}
        value={kind}
        onChange={(event) => setKind(event.target.value as Kind)}
      >
        <option value="values">{t('filter.values')}</option>
        <optgroup label={t('filter.text')}>
          {TEXT_TESTS.map((each) => (
            <option key={each} value={`text.${each}`}>
              {t(`filter.${each}` as Key)}
            </option>
          ))}
        </optgroup>
        <optgroup label={t('filter.number')}>
          {NUMBER_TESTS.map((each) => (
            <option key={each} value={`number.${each}`}>
              {t(`filter.${each}` as Key)}
            </option>
          ))}
        </optgroup>
      </select>
      {kind === 'values' ? (
        <>
          <input
            type="search"
            aria-label={t('filter.search')}
            placeholder={t('filter.search')}
            className={`${styles.textInput} ${styles.filterSearch}`}
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
          <ul className={styles.filterList}>
            <li>
              <label className={styles.checkField}>
                <input
                  type="checkbox"
                  className={styles.checkbox}
                  checked={allShown}
                  onChange={(event) => {
                    const next = new Set(checked);
                    for (const value of shown) {
                      if (event.target.checked) next.add(value);
                      else next.delete(value);
                    }
                    setChecked(next);
                  }}
                />
                {t('filter.all')}
              </label>
            </li>
            {shown.map((value) => (
              <li key={value}>
                <label className={styles.checkField}>
                  <input
                    type="checkbox"
                    className={styles.checkbox}
                    checked={checked.has(value)}
                    onChange={(event) => toggle(value, event.target.checked)}
                  />
                  {value === '' ? t('filter.blanks') : value}
                </label>
              </li>
            ))}
          </ul>
        </>
      ) : (
        <>
          <label className={styles.field} htmlFor={`${id}-value`}>
            {t('filter.value')}
          </label>
          <input
            id={`${id}-value`}
            className={styles.textInput}
            inputMode={type === 'number' ? 'decimal' : 'text'}
            aria-invalid={type === 'number' && wrong(text)}
            value={text}
            onChange={(event) => setText(event.target.value)}
          />
          {kind === 'number.between' && (
            <>
              <label className={styles.field} htmlFor={`${id}-to`}>
                {t('filter.and')}
              </label>
              <input
                id={`${id}-to`}
                className={styles.textInput}
                inputMode="decimal"
                aria-invalid={wrong(to)}
                value={to}
                onChange={(event) => setTo(event.target.value)}
              />
            </>
          )}
          {type === 'number' && (wrong(text) || wrong(to)) && (
            <p className={styles.fieldError}>{t('filter.invalid')}</p>
          )}
        </>
      )}
      <div className={styles.dialogButtons}>
        <button type="button" className={styles.button} onClick={onCancel}>
          {t('cancel')}
        </button>
        <button
          type="submit"
          className={`${styles.button} ${styles.primary}`}
          disabled={invalid}
        >
          {t('ok')}
        </button>
      </div>
    </form>
  );
}
