import * as Dialog from '@radix-ui/react-dialog';
import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import { ChevronDown, ListChecks } from 'lucide-react';
import { useCallback, useId, useState, useSyncExternalStore } from 'react';
import { formatValue, parseDate, parseInput } from '../../engine/format';
import { parse } from '../../engine/formula';
import type { Criterion, Workbook } from '../../engine/workbook';
import { EDITOR } from '../grid/FormulaEditor';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;
/** How dates are shown and read in the dialog, as typed in a cell. */
const DATE_FORMAT = 'm/d/yyyy';

/** What a rule allows, as the dialog offers it; `any` for no rule. */
type Allow = 'any' | Criterion['type'];

const ALLOWS: readonly Allow[] = ['any', 'list', 'number', 'date', 'length'];

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
 * The Data tab's Data validation button, and the dialog it opens: the rule
 * of the active cell, to give the selected cells, change or remove (R26).
 */
export function ValidationControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const [open, setOpen] = useState(false);
  const { sheet, selection } = view.state;
  return (
    <>
      <Tooltip text={t('validation.tooltip')}>
        <button
          type="button"
          className={styles.button}
          // A rejected entry keeps the dialog closed, so that the rule it
          // would set is not lost.
          onClick={() => {
            if (view.commit()) setOpen(true);
          }}
        >
          <ListChecks size={ICON_SIZE} aria-hidden />
          {t('validation')}
        </button>
      </Tooltip>
      <Dialog.Root open={open} onOpenChange={setOpen}>
        <Dialog.Portal>
          <Dialog.Overlay className={styles.overlay} />
          <Dialog.Content className={styles.dialog}>
            <Dialog.Title className={styles.dialogTitle}>
              {t('validation')}
            </Dialog.Title>
            <ValidationForm
              criterion={workbook.validation(
                sheet,
                selection.active.row,
                selection.active.col,
              )}
              onSave={(criterion) => {
                view.validate(criterion);
                setOpen(false);
              }}
              onCancel={() => setOpen(false)}
            />
          </Dialog.Content>
        </Dialog.Portal>
      </Dialog.Root>
    </>
  );
}

/** A bound as the dialog shows it. */
function boundText(type: Allow, bound: number | undefined): string {
  if (bound === undefined) return '';
  return type === 'date' ? formatValue(bound, DATE_FORMAT) : String(bound);
}

/** A bound read from the dialog: `null` for none, `undefined` if unreadable. */
function readBound(type: Allow, text: string): number | null | undefined {
  if (text.trim() === '') return null;
  if (type === 'date') return parseDate(text);
  const { value } = parseInput(text.trim());
  if (typeof value !== 'number') return undefined;
  return type === 'length' && (!Number.isInteger(value) || value < 0)
    ? undefined
    : value;
}

/**
 * What the rule allows, with its list's source or its bounds; the problem
 * with what is typed is shown once OK is pressed.
 */
function ValidationForm({
  criterion,
  onSave,
  onCancel,
}: {
  criterion: Criterion | null;
  onSave: (criterion: Criterion | null) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [type, setType] = useState<Allow>(criterion?.type ?? 'any');
  const [source, setSource] = useState(
    criterion?.type === 'list' ? criterion.source : '',
  );
  const bounded = criterion?.type === 'list' ? undefined : criterion;
  const [min, setMin] = useState(boundText(type, bounded?.min));
  const [max, setMax] = useState(boundText(type, bounded?.max));
  const [tried, setTried] = useState(false);

  /** The rule typed, or the key of what is wrong with it. */
  const read = (): { criterion: Criterion | null } | { problem: Key } => {
    if (type === 'any') return { criterion: null };
    if (type === 'list') {
      const text = source.trim();
      if (text === '' || text === '=')
        return { problem: 'validation.noSource' };
      if (text.startsWith('=') && parse(text.slice(1))?.type !== 'range') {
        return { problem: 'validation.badRange' };
      }
      return { criterion: { type, source: text } };
    }
    const low = readBound(type, min);
    const high = readBound(type, max);
    if (low === undefined || high === undefined) {
      return { problem: `validation.bad.${type}` };
    }
    if (low === null && high === null) return { problem: 'validation.noBound' };
    if (low !== null && high !== null && high < low) {
      return { problem: 'validation.order' };
    }
    return {
      criterion: {
        type,
        ...(low !== null && { min: low }),
        ...(high !== null && { max: high }),
      },
    };
  };
  const result = read();
  const problem = tried && 'problem' in result ? result.problem : null;
  const field = (
    name: string,
    label: string,
    text: string,
    onText: (text: string) => void,
  ) => (
    <>
      <label className={styles.field} htmlFor={`${id}-${name}`}>
        {label}
      </label>
      <input
        id={`${id}-${name}`}
        className={styles.textInput}
        value={text}
        spellCheck={false}
        autoComplete="off"
        aria-invalid={problem !== null}
        aria-describedby={problem ? `${id}-problem` : undefined}
        onChange={(event) => onText(event.target.value)}
      />
    </>
  );
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        setTried(true);
        if ('criterion' in result) onSave(result.criterion);
      }}
    >
      <Dialog.Description className={styles.dialogText}>
        {t('validation.description')}
      </Dialog.Description>
      <label className={styles.field} htmlFor={`${id}-type`}>
        {t('validation.allow')}
      </label>
      <select
        id={`${id}-type`}
        className={`${styles.textInput} ${styles.ruleSelect}`}
        value={type}
        onChange={(event) => {
          setType(event.target.value as Allow);
          setTried(false);
        }}
      >
        {ALLOWS.map((each) => (
          <option key={each} value={each}>
            {t(`validation.allow.${each}`)}
          </option>
        ))}
      </select>
      {type === 'list' &&
        field('source', t('validation.source'), source, setSource)}
      {type !== 'any' && type !== 'list' && (
        <>
          {field('min', t('validation.min'), min, setMin)}
          {field('max', t('validation.max'), max, setMax)}
        </>
      )}
      {problem && (
        <p id={`${id}-problem`} className={styles.problem}>
          {t(problem)}
        </p>
      )}
      <div className={styles.dialogButtons}>
        <Tooltip text={t('validation.clear.tooltip')}>
          <button
            type="button"
            className={styles.button}
            onClick={() => onSave(null)}
          >
            {t('validation.clear')}
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

/** The dropdown of the active cell's list, which sets the cell to an item. */
export function ListButton({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const subscribe = useChanges(workbook, view);
  // A string, so that the same items give the same snapshot.
  const itemsText = useSyncExternalStore(subscribe, () => {
    const { sheet, selection } = view.state;
    const { row, col } = selection.active;
    return JSON.stringify(workbook.listItems(sheet, row, col));
  });
  const items = JSON.parse(itemsText) as string[];
  return (
    <DropdownMenu.Root>
      <Tooltip text={t('validation.choose')}>
        <DropdownMenu.Trigger asChild>
          <button
            type="button"
            className={styles.filterButton}
            aria-label={t('validation.choose')}
          >
            <ChevronDown size={14} aria-hidden />
          </button>
        </DropdownMenu.Trigger>
      </Tooltip>
      <DropdownMenu.Portal>
        <DropdownMenu.Content
          className={`${styles.menu} ${styles.listMenu}`}
          align="end"
          sideOffset={4}
        >
          {items.length === 0 ? (
            <DropdownMenu.Item className={styles.menuItem} disabled>
              {t('validation.empty')}
            </DropdownMenu.Item>
          ) : (
            items.map((item, index) => (
              <DropdownMenu.Item
                key={index}
                className={styles.menuItem}
                onSelect={() => view.choose(item)}
              >
                {item}
              </DropdownMenu.Item>
            ))
          )}
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}

/**
 * The message that an entry breaks its cell's rule: Retry goes back to the
 * entry, and Cancel, or Escape, abandons it (R26).
 */
export function RejectedEntry({ view }: { view: View }) {
  const t = useTranslation();
  const rejected = useSyncExternalStore(
    view.subscribe,
    () => view.state.rejected,
  );
  return (
    <Dialog.Root
      open={rejected}
      onOpenChange={(open) => {
        if (!open) view.cancel();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content
          role="alertdialog"
          className={styles.dialog}
          onCloseAutoFocus={(event) => {
            // The editor of an entry tried again takes the keyboard; the
            // grid takes it back from one abandoned.
            event.preventDefault();
            const { entry } = view.state;
            if (!entry) return;
            document
              .querySelector<HTMLElement>(`[${EDITOR}="${entry.place}"]`)
              ?.focus();
          }}
        >
          <Dialog.Title className={styles.dialogTitle}>
            {t('validation.rejected')}
          </Dialog.Title>
          <Dialog.Description className={styles.dialogText}>
            {t('validation.rejected.text')}
          </Dialog.Description>
          <div className={styles.dialogButtons}>
            <button
              type="button"
              className={styles.button}
              onClick={() => view.cancel()}
            >
              {t('cancel')}
            </button>
            <button
              type="button"
              className={`${styles.button} ${styles.primary}`}
              onClick={() => view.retry()}
            >
              {t('validation.retry')}
            </button>
          </div>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}
