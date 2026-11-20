import * as Dialog from '@radix-ui/react-dialog';
import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import {
  Copy,
  FilePlus,
  FolderOpen,
  MoreHorizontal,
  Pencil,
  Table2,
  Trash2,
  X,
  type LucideIcon,
} from 'lucide-react';
import { useEffect, useId, useRef, useState, type ReactNode } from 'react';
import type { Workbook } from '../../engine/workbook';
import type { Platform, StoredInfo } from '../platform';
import { Tooltip, useTranslation } from './controls';
import type { Language } from './i18n';
import styles from './Shell.module.css';
import { modifiedText, numberedName, renamedSnapshot } from './workbooks';

const ICON_SIZE = 16;

/**
 * The dialog of the workbooks stored in the browser (R34), the one saved
 * last first, to open, rename, duplicate or delete, with the buttons that
 * start a new workbook or open a file. With no workbook open it cannot be
 * closed.
 */
export function WorkbookDialog({
  platform,
  language,
  open,
  current,
  message,
  busy,
  onClose,
  onNew,
  onOpenFile,
  onOpen,
  onFailure,
}: {
  platform: Platform;
  language: Language;
  open: boolean;
  /** The workbook shown behind the dialog, if one is. */
  current: { readonly id: string; readonly workbook: Workbook } | undefined;
  /** The message bar, shown in the dialog while it is open. */
  message: ReactNode;
  /** Whether a file is being opened, when another cannot be. */
  busy: boolean;
  /** Returns to the current workbook. */
  onClose: () => void;
  /** Starts an empty workbook with this name. */
  onNew: (name: string) => void;
  onOpenFile: () => void;
  onOpen: (id: string) => void;
  /** Called when storage cannot be read, or refuses a change. */
  onFailure: (how: 'read' | 'write') => void;
}) {
  const t = useTranslation();
  const headingId = useId();
  const content = useRef<HTMLDivElement>(null);
  const [workbooks, setWorkbooks] = useState<StoredInfo[] | null>(null);
  const [renaming, setRenaming] = useState<StoredInfo | null>(null);
  const [deleting, setDeleting] = useState<StoredInfo | null>(null);
  // Changed to list the workbooks again after a change to them.
  const [listing, setListing] = useState(0);

  useEffect(() => {
    if (!open) return;
    let current = true;
    platform.workbooks
      .list()
      .then((list) => {
        if (current) setWorkbooks(list);
      })
      .catch(() => {
        if (!current) return;
        // New workbook still works, and the user is told.
        setWorkbooks([]);
        onFailure('read');
      });
    return () => {
      current = false;
    };
  }, [platform, open, listing]);

  /** Runs a change to the stored workbooks, then lists them again. */
  const change = (run: () => Promise<void>) => {
    run()
      .catch(() => onFailure('write'))
      .finally(() => setListing((count) => count + 1));
  };
  const saveAs = (id: string, from: string, name: string) =>
    change(async () => {
      const stored = await platform.workbooks.load(from);
      await platform.workbooks.save(id, renamedSnapshot(stored, name));
    });
  const rename = (id: string, name: string) => {
    if (id !== current?.id) {
      saveAs(id, id, name);
      return;
    }
    // The open workbook is renamed as the title box renames it, and saved
    // at once so that the list shows the name.
    const { workbook } = current;
    workbook.execute({ type: 'renameWorkbook', name });
    change(() => platform.workbooks.save(id, workbook.toSnapshot()));
  };

  const closable = current !== undefined;
  const now = Date.now();
  return (
    <Dialog.Root
      open={open}
      onOpenChange={(open) => {
        if (!open) onClose();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content
          className={`${styles.dialog} ${styles.workbooks}`}
          aria-describedby={undefined}
          ref={content}
          // The dialog itself takes the focus: on its first button, the
          // focus would show that button's tooltip, which Esc would close
          // in place of the dialog.
          onOpenAutoFocus={(event) => {
            event.preventDefault();
            content.current?.focus();
          }}
          // With no workbook to return to, the dialog stays.
          onEscapeKeyDown={(event) => {
            if (!closable) event.preventDefault();
          }}
          onInteractOutside={(event) => {
            if (!closable) event.preventDefault();
          }}
        >
          <header className={styles.workbooksHeader}>
            <Dialog.Title className={styles.workbooksTitle}>
              <Table2 size={20} aria-hidden />
              {t('start.title')}
            </Dialog.Title>
            <Tooltip text={t('start.new.tooltip')}>
              <button
                type="button"
                className={`${styles.button} ${styles.primary}`}
                disabled={workbooks === null}
                onClick={() =>
                  onNew(
                    numberedName(
                      t('start.book'),
                      (workbooks ?? []).map((each) => each.name),
                    ),
                  )
                }
              >
                <FilePlus size={ICON_SIZE} aria-hidden />
                {t('start.new')}
              </button>
            </Tooltip>
            <Tooltip text={t('file.open.tooltip')}>
              <button
                type="button"
                className={styles.button}
                disabled={busy}
                onClick={onOpenFile}
              >
                <FolderOpen size={ICON_SIZE} aria-hidden />
                {t('start.open')}
              </button>
            </Tooltip>
            {closable && (
              <Tooltip text={t('start.close.tooltip')}>
                <Dialog.Close
                  className={styles.iconButton}
                  aria-label={t('start.close')}
                >
                  <X size={ICON_SIZE} aria-hidden />
                </Dialog.Close>
              </Tooltip>
            )}
          </header>
          {message}
          <section className={styles.workbooksList} aria-labelledby={headingId}>
            <h2 id={headingId} className={styles.workbooksHeading}>
              {t('start.recent')}
            </h2>
            {workbooks?.length === 0 && (
              <p className={styles.dialogText}>{t('start.empty')}</p>
            )}
            <ul className={styles.workbooksItems} aria-label={t('start.list')}>
              {workbooks?.map((workbook) => (
                <li key={workbook.id} className={styles.workbooksItem}>
                  <Tooltip text={t('start.openWorkbook.tooltip')}>
                    <button
                      type="button"
                      className={`${styles.button} ${styles.workbooksOpen}`}
                      aria-label={t('start.openWorkbook').replace(
                        '{name}',
                        () => workbook.name,
                      )}
                      onClick={() => onOpen(workbook.id)}
                    >
                      <span className={styles.workbooksName}>
                        {workbook.name}
                      </span>
                      <time
                        className={styles.workbooksTime}
                        dateTime={new Date(workbook.modified).toISOString()}
                      >
                        {modifiedText(workbook.modified, now, language)}
                      </time>
                    </button>
                  </Tooltip>
                  <Actions
                    name={workbook.name}
                    // The workbook shown cannot be deleted from under it.
                    deletable={workbook.id !== current?.id}
                    onRename={() => setRenaming(workbook)}
                    onDuplicate={() =>
                      saveAs(
                        crypto.randomUUID(),
                        workbook.id,
                        t('start.copyName').replace(
                          '{name}',
                          () => workbook.name,
                        ),
                      )
                    }
                    onDelete={() => setDeleting(workbook)}
                  />
                </li>
              ))}
            </ul>
          </section>
          <RenameDialog
            workbook={renaming}
            onRename={(name) => {
              if (renaming && name !== renaming.name) {
                rename(renaming.id, name);
              }
              setRenaming(null);
            }}
            onCancel={() => setRenaming(null)}
          />
          <DeleteDialog
            workbook={deleting}
            onDelete={() => {
              const { id } = deleting!;
              change(() => platform.workbooks.delete(id));
              setDeleting(null);
            }}
            onCancel={() => setDeleting(null)}
          />
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

/** The menu of a stored workbook's actions. */
function Actions({
  name,
  deletable,
  onRename,
  onDuplicate,
  onDelete,
}: {
  name: string;
  deletable: boolean;
  onRename: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
}) {
  const t = useTranslation();
  return (
    <DropdownMenu.Root>
      <Tooltip text={t('start.more.tooltip')}>
        <DropdownMenu.Trigger
          className={styles.iconButton}
          aria-label={t('start.more').replace('{name}', () => name)}
        >
          <MoreHorizontal size={ICON_SIZE} aria-hidden />
        </DropdownMenu.Trigger>
      </Tooltip>
      <DropdownMenu.Portal>
        <DropdownMenu.Content className={styles.menu} align="end">
          <MenuItem
            icon={Pencil}
            label={t('start.rename')}
            onSelect={onRename}
          />
          <MenuItem
            icon={Copy}
            label={t('start.duplicate')}
            onSelect={onDuplicate}
          />
          <MenuItem
            icon={Trash2}
            label={t('start.delete')}
            disabled={!deletable}
            onSelect={onDelete}
          />
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}

function MenuItem({
  icon: Icon,
  label,
  disabled,
  onSelect,
}: {
  icon: LucideIcon;
  label: string;
  disabled?: boolean;
  onSelect: () => void;
}) {
  return (
    <DropdownMenu.Item
      className={styles.menuItem}
      disabled={disabled}
      onSelect={onSelect}
    >
      <Icon size={ICON_SIZE} aria-hidden />
      {label}
    </DropdownMenu.Item>
  );
}

/** Asks for a new name of a stored workbook; an empty one is refused. */
function RenameDialog({
  workbook,
  onRename,
  onCancel,
}: {
  workbook: StoredInfo | null;
  onRename: (name: string) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const inputId = useId();
  const [name, setName] = useState('');
  useEffect(() => {
    if (workbook) setName(workbook.name);
  }, [workbook]);
  const trimmed = name.trim();
  return (
    <Dialog.Root
      open={workbook !== null}
      onOpenChange={(open) => {
        if (!open) onCancel();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content className={styles.dialog} aria-describedby={undefined}>
          <Dialog.Title className={styles.dialogTitle}>
            {t('start.renameTitle')}
          </Dialog.Title>
          <form
            onSubmit={(event) => {
              event.preventDefault();
              if (trimmed !== '') onRename(trimmed);
            }}
          >
            <label className={styles.field} htmlFor={inputId}>
              {t('start.name')}
            </label>
            <input
              id={inputId}
              className={styles.textInput}
              value={name}
              autoFocus
              onFocus={(event) => event.target.select()}
              onChange={(event) => setName(event.target.value)}
            />
            <div className={styles.dialogButtons}>
              <Dialog.Close className={styles.button}>
                {t('start.cancel')}
              </Dialog.Close>
              <button
                type="submit"
                className={`${styles.button} ${styles.primary}`}
                disabled={trimmed === ''}
              >
                {t('start.rename')}
              </button>
            </div>
          </form>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

/** Asks whether to delete a stored workbook. */
function DeleteDialog({
  workbook,
  onDelete,
  onCancel,
}: {
  workbook: StoredInfo | null;
  onDelete: () => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  return (
    <Dialog.Root
      open={workbook !== null}
      onOpenChange={(open) => {
        if (!open) onCancel();
      }}
    >
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content className={styles.dialog} role="alertdialog">
          <Dialog.Title className={styles.dialogTitle}>
            {t('start.deleteTitle')}
          </Dialog.Title>
          <Dialog.Description className={styles.dialogText}>
            {t('start.deleteText').replace(
              '{name}',
              () => workbook?.name ?? '',
            )}
          </Dialog.Description>
          <div className={styles.dialogButtons}>
            <Dialog.Close className={styles.button} autoFocus>
              {t('start.cancel')}
            </Dialog.Close>
            <button
              type="button"
              className={`${styles.button} ${styles.danger}`}
              onClick={onDelete}
            >
              {t('start.delete')}
            </button>
          </div>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}
