import * as Dialog from '@radix-ui/react-dialog';
import { Link2, MessageSquareX, StickyNote } from 'lucide-react';
import { useCallback, useId, useState, useSyncExternalStore } from 'react';
import type { Workbook } from '../../engine/workbook';
import { focusGrid } from '../grid/Grid';
import type { LinkTarget, View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

/** What is being edited in a dialog: the active cell's note or link. */
export type Editing = 'note' | 'link';

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
 * The Review tab's New note, or Edit note when the active cell has one,
 * and Delete note (R28).
 */
export function NoteControls({
  workbook,
  view,
  onEdit,
}: {
  workbook: Workbook;
  view: View;
  onEdit: (editing: Editing) => void;
}) {
  const t = useTranslation();
  const has = useSyncExternalStore(
    useChanges(workbook, view),
    () => view.note !== '',
  );
  return (
    <>
      <Tooltip text={t(has ? 'note.edit.tooltip' : 'note.new.tooltip')}>
        <button
          type="button"
          className={styles.button}
          onClick={() => onEdit('note')}
        >
          <StickyNote size={ICON_SIZE} aria-hidden />
          {t(has ? 'note.edit' : 'note.new')}
        </button>
      </Tooltip>
      <Tooltip text={t('note.delete.tooltip')}>
        <button
          type="button"
          className={styles.button}
          disabled={!has}
          onClick={() => view.setNote('')}
        >
          <MessageSquareX size={ICON_SIZE} aria-hidden />
          {t('note.delete')}
        </button>
      </Tooltip>
    </>
  );
}

/** The Insert tab's Link button, which inserts or edits a link (R29). */
export function LinkControls({
  onEdit,
}: {
  onEdit: (editing: Editing) => void;
}) {
  const t = useTranslation();
  return (
    <Tooltip text={t('link.tooltip')}>
      <button
        type="button"
        className={styles.button}
        onClick={() => onEdit('link')}
      >
        <Link2 size={ICON_SIZE} aria-hidden />
        {t('link')}
      </button>
    </Tooltip>
  );
}

/**
 * The dialog of the note or the link of the active cell, opened from the
 * ribbon or the cell's menu, which gives the keyboard back to the grid.
 */
export function CellDialog({
  workbook,
  view,
  editing,
  onClose,
}: {
  workbook: Workbook;
  view: View;
  editing: Editing | null;
  onClose: () => void;
}) {
  const t = useTranslation();
  const note = editing === 'note' ? view.note : '';
  const link = editing === 'link' ? view.link : null;
  const title: Key =
    editing === 'note'
      ? note
        ? 'note.edit'
        : 'note.new'
      : link
        ? 'link.edit'
        : 'link.insert';
  return (
    <Dialog.Root
      open={editing !== null}
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content
          className={styles.dialog}
          aria-describedby={undefined}
          onCloseAutoFocus={(event) => {
            event.preventDefault();
            focusGrid();
          }}
        >
          <Dialog.Title className={styles.dialogTitle}>{t(title)}</Dialog.Title>
          {editing === 'note' && (
            <NoteForm
              note={note}
              onSave={(text) => {
                view.setNote(text);
                onClose();
              }}
              onCancel={onClose}
            />
          )}
          {editing === 'link' && (
            <LinkForm
              workbook={workbook}
              view={view}
              link={link}
              onDone={onClose}
            />
          )}
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function NoteForm({
  note,
  onSave,
  onCancel,
}: {
  note: string;
  onSave: (text: string) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [text, setText] = useState(note);
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        onSave(text);
      }}
    >
      <label className={styles.field} htmlFor={id}>
        {t('note')}
      </label>
      <textarea
        id={id}
        className={`${styles.textInput} ${styles.noteInput}`}
        value={text}
        onChange={(event) => setText(event.target.value)}
        autoFocus
      />
      <Buttons onCancel={onCancel} />
    </form>
  );
}

/**
 * The text a link shows, and where it goes: a web address, or a cell
 * reference on a sheet; the problem with what is typed is shown once OK
 * is pressed.
 */
function LinkForm({
  workbook,
  view,
  link,
  onDone,
}: {
  workbook: Workbook;
  view: View;
  link: LinkTarget | null;
  onDone: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const { sheet, selection } = view.state;
  const { row, col } = selection.active;
  const [text, setText] = useState(
    () => workbook.cell(sheet, row, col).display,
  );
  const [place, setPlace] = useState(link !== null && 'sheet' in link);
  const [address, setAddress] = useState(
    link && 'address' in link ? link.address : '',
  );
  const [target, setTarget] = useState(
    link && 'sheet' in link ? link.sheet : sheet,
  );
  const [reference, setReference] = useState(
    link && 'reference' in link ? link.reference : 'A1',
  );
  const [problem, setProblem] = useState<Key | null>(null);
  const kind = (value: boolean, label: Key) => (
    <label className={styles.checkField}>
      <input
        type="radio"
        className={styles.checkbox}
        name={`${id}-kind`}
        checked={place === value}
        onChange={() => {
          setPlace(value);
          setProblem(null);
        }}
      />
      {t(label)}
    </label>
  );
  const field = (
    name: string,
    label: Key,
    value: string,
    set: (value: string) => void,
  ) => (
    <>
      <label className={styles.field} htmlFor={`${id}-${name}`}>
        {t(label)}
      </label>
      <input
        id={`${id}-${name}`}
        className={styles.textInput}
        value={value}
        spellCheck={false}
        autoComplete="off"
        aria-invalid={name !== 'text' && problem !== null}
        onChange={(event) => set(event.target.value)}
      />
    </>
  );
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        const found = view.setLink(
          place ? { sheet: target, reference } : { address },
          text,
        );
        if (found === null) onDone();
        // A rejected entry has its own message.
        else if (found !== 'entry') setProblem(`link.problem.${found}`);
      }}
    >
      <div className={styles.linkKinds}>
        {kind(false, 'link.web')}
        {kind(true, 'link.place')}
      </div>
      {field('text', 'link.text', text, setText)}
      {place ? (
        <>
          <label className={styles.field} htmlFor={`${id}-sheet`}>
            {t('link.sheet')}
          </label>
          <select
            id={`${id}-sheet`}
            className={`${styles.textInput} ${styles.ruleSelect}`}
            value={target}
            onChange={(event) => setTarget(event.target.value)}
          >
            {workbook.sheets.map((each) => (
              <option key={each.id} value={each.id}>
                {each.name}
              </option>
            ))}
          </select>
          {field('reference', 'link.reference', reference, setReference)}
        </>
      ) : (
        field('address', 'link.address', address, setAddress)
      )}
      {problem && (
        <p className={styles.problem} role="alert">
          {t(problem)}
        </p>
      )}
      <Buttons
        onCancel={onDone}
        onRemove={
          link
            ? () => {
                view.removeLink();
                onDone();
              }
            : undefined
        }
      />
    </form>
  );
}

function Buttons({
  onCancel,
  onRemove,
}: {
  onCancel: () => void;
  onRemove?: () => void;
}) {
  const t = useTranslation();
  return (
    <div className={styles.dialogButtons}>
      {onRemove && (
        <Tooltip text={t('link.remove.tooltip')}>
          <button type="button" className={styles.button} onClick={onRemove}>
            {t('link.remove')}
          </button>
        </Tooltip>
      )}
      <button type="button" className={styles.button} onClick={onCancel}>
        {t('cancel')}
      </button>
      <button type="submit" className={`${styles.button} ${styles.primary}`}>
        {t('ok')}
      </button>
    </div>
  );
}
