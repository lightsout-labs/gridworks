import * as Dialog from '@radix-ui/react-dialog';
import { Pencil, Tags, Trash2, type LucideIcon } from 'lucide-react';
import { useCallback, useId, useState, useSyncExternalStore } from 'react';
import type { NamedRange, NameProblem, Workbook } from '../../engine/workbook';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

/**
 * The Name manager button and its dialog, which lists the defined names
 * and edits and deletes them (R27); names are made in the name box.
 */
export function NameControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const [open, setOpen] = useState(false);
  return (
    <Dialog.Root open={open} onOpenChange={setOpen}>
      <Tooltip text={t('names.manager.tooltip')}>
        <Dialog.Trigger className={styles.button}>
          <Tags size={ICON_SIZE} aria-hidden />
          {t('names.manager')}
        </Dialog.Trigger>
      </Tooltip>
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content className={`${styles.dialog} ${styles.wideDialog}`}>
          <Dialog.Title className={styles.dialogTitle}>
            {t('names.manager')}
          </Dialog.Title>
          <NameList workbook={workbook} view={view} />
          <div className={styles.dialogButtons}>
            <Dialog.Close className={styles.button}>
              {t('names.close')}
            </Dialog.Close>
          </div>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

/** The defined names, each with its range, to edit or delete. */
function NameList({ workbook, view }: { workbook: Workbook; view: View }) {
  const t = useTranslation();
  const subscribe = useCallback(
    (listener: () => void) => workbook.subscribe(listener),
    [workbook],
  );
  const names = useSyncExternalStore(subscribe, () => workbook.names);
  const [editing, setEditing] = useState<NamedRange | null>(null);
  if (editing) {
    return (
      <NameForm
        named={editing}
        onSave={(name, reference) => {
          const problem = view.setName(editing.name, name, reference);
          if (!problem) setEditing(null);
          return problem;
        }}
        onCancel={() => setEditing(null)}
      />
    );
  }
  if (names.length === 0) {
    return (
      <Dialog.Description className={styles.dialogText}>
        {t('names.empty')}
      </Dialog.Description>
    );
  }
  return (
    <ul className={styles.ruleList}>
      {names.map((named) => (
        <li key={named.name} className={styles.ruleRow}>
          <span className={styles.ruleText}>
            <span>{named.name}</span>
            <span className={styles.ruleRanges}>{`=${named.refersTo}`}</span>
          </span>
          <IconButton
            icon={Pencil}
            label={t('names.edit')}
            tooltip={t('names.edit.tooltip')}
            onClick={() => setEditing(named)}
          />
          <IconButton
            icon={Trash2}
            label={t('names.delete')}
            tooltip={t('names.delete.tooltip')}
            onClick={() => view.deleteName(named.name)}
          />
        </li>
      ))}
    </ul>
  );
}

/**
 * The name and the range of a defined name, to change; `onSave` gives why
 * it cannot take them, or `null` when it has.
 */
function NameForm({
  named,
  onSave,
  onCancel,
}: {
  named: NamedRange;
  onSave: (
    name: string,
    reference: string,
  ) => NameProblem | 'reference' | 'entry' | null;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [name, setName] = useState(named.name);
  const [reference, setReference] = useState(`=${named.refersTo}`);
  const [problem, setProblem] = useState<
    NameProblem | 'reference' | 'entry' | null
  >(null);
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        setProblem(onSave(name.trim(), reference));
      }}
    >
      <label className={styles.field} htmlFor={`${id}-name`}>
        {t('names.name')}
      </label>
      <input
        id={`${id}-name`}
        className={styles.textInput}
        value={name}
        aria-invalid={problem === 'invalid' || problem === 'taken'}
        onChange={(event) => setName(event.target.value)}
        autoFocus
      />
      <label className={styles.field} htmlFor={`${id}-reference`}>
        {t('names.refersTo')}
      </label>
      <input
        id={`${id}-reference`}
        className={styles.textInput}
        value={reference}
        aria-invalid={problem === 'reference'}
        onChange={(event) => setReference(event.target.value)}
      />
      {/* A rejected entry has its own message. */}
      {problem && problem !== 'entry' && (
        <p className={styles.problem} role="alert">
          {t(`names.problem.${problem}` satisfies Key)}
        </p>
      )}
      <div className={styles.dialogButtons}>
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

/** A button shown by its icon, named for assistive technology. */
function IconButton({
  icon: Icon,
  label,
  tooltip,
  onClick,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  onClick: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.iconButton}
        aria-label={label}
        onClick={onClick}
      >
        <Icon size={ICON_SIZE} aria-hidden />
      </button>
    </Tooltip>
  );
}
