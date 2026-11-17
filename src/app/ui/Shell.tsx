import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import { ChevronDown, Minus, Plus, X } from 'lucide-react';
import {
  useCallback,
  useEffect,
  useMemo,
  useReducer,
  useState,
  useSyncExternalStore,
} from 'react';
import type { Workbook } from '../../engine/workbook';
import { FormulaEditor } from '../grid/FormulaEditor';
import { focusGrid, MAC } from '../grid/Grid';
import { keyAction } from '../grid/keys';
import { ZOOM_MAX, ZOOM_MIN, type View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import styles from './Shell.module.css';
import { circularCell, summarize } from './summary';

const ICON_SIZE = 14;
const ZOOM_STEP = 0.1;
const ONE_CELL = { rows: 1, cols: 1 };

/**
 * The name box, which names the active cell or the range selected and
 * goes to a range or a name typed or chosen in it, and, beside it, what
 * the active cell holds, where it can be edited. `onInvalid` tells that a
 * text typed in the name box is neither a reference nor a new name, or a
 * name whose cells were deleted.
 */
export function FormulaBar({
  workbook,
  view,
  onInvalid,
}: {
  workbook: Workbook;
  view: View;
  onInvalid: (text: string, problem: 'invalid' | 'deleted') => void;
}) {
  const t = useTranslation();
  const { sheet, selection } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  const { active } = selection;
  const contents = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => workbook.cell(sheet, active.row, active.col).input,
  );
  return (
    <div className={styles.formulaLine}>
      <NameBox workbook={workbook} view={view} onInvalid={onInvalid} />
      <Tooltip text={t('formulaBar.tooltip')}>
        <FormulaEditor
          view={view}
          place="formulaBar"
          contents={contents}
          label={t('formulaBar')}
          className={styles.formulaBox}
          inputClassName={styles.formulaBar}
          mirrorClassName={styles.formulaMirror}
          onKeyDown={(event) => {
            // Only Enter, Tab, Escape and F4 act here: it is always Edit mode.
            const action = keyAction(
              {
                key: event.key,
                shift: event.shiftKey,
                ctrl: event.ctrlKey,
                meta: event.metaKey,
                alt: event.altKey,
              },
              { mac: MAC, mode: 'edit', endMode: false, page: ONE_CELL },
            );
            if (action?.type === 'navigate') view.navigate(action.navigation);
            else if (action?.type === 'cancel') view.cancel();
            else if (action?.type === 'cycleReference') view.cycleReference();
            else return;
            event.preventDefault();
          }}
        />
      </Tooltip>
    </div>
  );
}

/**
 * A box that shows the name of the selection and goes where a text typed
 * in it says, giving the keyboard back to the grid, and a list of the
 * defined names to go to.
 */
function NameBox({
  workbook,
  view,
  onInvalid,
}: {
  workbook: Workbook;
  view: View;
  onInvalid: (text: string, problem: 'invalid' | 'deleted') => void;
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
  const shown = useSyncExternalStore(subscribe, () => view.selectionName);
  const names = useSyncExternalStore(subscribe, () => workbook.names);
  // What is typed, until it is gone to or abandoned.
  const [typed, setTyped] = useState<string | null>(null);
  const go = (text: string) => {
    const problem = view.goTo(text);
    if (problem) {
      onInvalid(text, problem);
      return;
    }
    setTyped(null);
    focusGrid();
  };
  return (
    <div className={styles.nameBox}>
      <Tooltip text={t('names.box.tooltip')}>
        <input
          className={styles.nameBoxInput}
          role="combobox"
          aria-label={t('nameBox')}
          aria-expanded={false}
          aria-autocomplete="none"
          spellCheck={false}
          value={typed ?? shown}
          onFocus={(event) => event.target.select()}
          onChange={(event) => setTyped(event.target.value)}
          onBlur={() => setTyped(null)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') go(typed ?? shown);
            else if (event.key === 'Escape') setTyped(null);
            else return;
            event.preventDefault();
          }}
        />
      </Tooltip>
      <DropdownMenu.Root>
        <Tooltip text={t('names.list.tooltip')}>
          <DropdownMenu.Trigger
            className={styles.iconButton}
            aria-label={t('names.list')}
            disabled={names.length === 0}
          >
            <ChevronDown size={ICON_SIZE} aria-hidden />
          </DropdownMenu.Trigger>
        </Tooltip>
        <DropdownMenu.Portal>
          <DropdownMenu.Content
            className={styles.menu}
            align="start"
            sideOffset={4}
            onCloseAutoFocus={(event) => {
              event.preventDefault();
              focusGrid();
            }}
          >
            {names.map(({ name }) => (
              <DropdownMenu.Item
                key={name}
                className={styles.menuItem}
                onSelect={() => go(name)}
              >
                {name}
              </DropdownMenu.Item>
            ))}
          </DropdownMenu.Content>
        </DropdownMenu.Portal>
      </DropdownMenu.Root>
    </div>
  );
}

/**
 * Whether the grid is waiting or an entry is being typed, a circular
 * reference, the average, count and sum of the selected cells, and the zoom.
 */
export function StatusBar({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet, selection, entry, endMode, zoom } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  // The figures change with the cells as well as with the selection.
  const [revision, changed] = useReducer((count: number) => count + 1, 0);
  useEffect(() => workbook.subscribe(changed), [workbook]);
  const summary = useMemo(
    () => summarize(workbook, sheet, selection),
    // `revision` stands for the contents of the workbook.
    [workbook, sheet, selection, revision],
  );
  const circular = useMemo(
    () => circularCell(workbook, sheet),
    // `revision` stands for the contents of the workbook.
    [workbook, sheet, revision],
  );
  const percent = Math.round(zoom * 100);
  return (
    <footer className={styles.statusBar}>
      <span>
        {t(
          entry === null
            ? 'status.ready'
            : entry.mode === 'enter'
              ? 'status.enter'
              : 'status.edit',
        )}
      </span>
      {endMode && <span>{t('status.endMode')}</span>}
      {circular && <span>{`${t('status.circular')}: ${circular}`}</span>}
      <div className={styles.summary} role="group" aria-label={t('summary')}>
        {summary?.average != null && (
          <span>{`${t('summary.average')}: ${summary.average}`}</span>
        )}
        {summary && <span>{`${t('summary.count')}: ${summary.count}`}</span>}
        {summary?.sum != null && (
          <span>{`${t('summary.sum')}: ${summary.sum}`}</span>
        )}
      </div>
      <div className={styles.zoom}>
        <Tooltip text={t('zoom.out.tooltip')}>
          <button
            type="button"
            className={styles.iconButton}
            aria-label={t('zoom.out')}
            disabled={zoom <= ZOOM_MIN}
            onClick={() => view.zoomTo(zoom - ZOOM_STEP)}
          >
            <Minus size={ICON_SIZE} aria-hidden />
          </button>
        </Tooltip>
        <Tooltip text={t('zoom.tooltip')}>
          <input
            type="range"
            className={styles.slider}
            aria-label={t('zoom')}
            min={ZOOM_MIN * 100}
            max={ZOOM_MAX * 100}
            step={10}
            value={percent}
            onChange={(event) => view.zoomTo(Number(event.target.value) / 100)}
          />
        </Tooltip>
        <Tooltip text={t('zoom.in.tooltip')}>
          <button
            type="button"
            className={styles.iconButton}
            aria-label={t('zoom.in')}
            disabled={zoom >= ZOOM_MAX}
            onClick={() => view.zoomTo(zoom + ZOOM_STEP)}
          >
            <Plus size={ICON_SIZE} aria-hidden />
          </button>
        </Tooltip>
        <span className={styles.percent}>{percent}%</span>
      </div>
    </footer>
  );
}

/**
 * A message about the last action, such as a file that could not be
 * opened, until the user closes it.
 */
export function MessageBar({
  text,
  error,
  onClose,
}: {
  text: string;
  error: boolean;
  onClose: () => void;
}) {
  const t = useTranslation();
  return (
    <div
      className={styles.message}
      data-kind={error ? 'error' : 'info'}
      role="alert"
    >
      <span className={styles.messageText}>{text}</span>
      <Tooltip text={t('message.close.tooltip')}>
        <button
          type="button"
          className={styles.button}
          aria-label={t('message.close')}
          onClick={onClose}
        >
          <X size={ICON_SIZE} aria-hidden />
        </button>
      </Tooltip>
    </div>
  );
}
