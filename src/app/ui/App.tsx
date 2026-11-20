import { Provider as TooltipProvider } from '@radix-ui/react-tooltip';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { Dropped } from '../../engine/io';
import { Workbook } from '../../engine/workbook';
import { Grid, MAC, measureText } from '../grid/Grid';
import { View } from '../grid/view';
import type { ClipboardContents, Platform } from '../platform';
import { restoreWorkbook, startAutosave, type Autosave } from './autosave';
import { CellMenu } from './CellMenu';
import { FillOptions } from './FillOptions';
import { FilterButton } from './Filter';
import { ListButton, RejectedEntry } from './Validation';
import { CellDialog, type Editing } from './Notes';
import { FindPanel, type FindMode } from './Find';
import { PrintPanel } from './Print';
import { Translation } from './controls';
import { droppedMessage, isCsv, workbookName } from './files';
import { startingLanguage, translator, type Language } from './i18n';
import { Ribbon } from './Ribbon';
import { FormulaBar, MessageBar, StatusBar } from './Shell';
import { SheetTabs } from './SheetTabs';
import { shortcut } from './shortcuts';
import styles from './Shell.module.css';
import { startingTheme, type Theme } from './theme';
import { WorkbookDialog } from './WorkbookDialog';
import { hashOf, workbookIdOf } from './workbooks';

// The file formats load after the app, but at once: a page that a new
// version has taken over can no longer fetch the files of its own.
const io = import('../../engine/io');
io.catch(() => {});

/** The workbook shown, under its identifier in storage. */
interface Open {
  readonly id: string;
  readonly workbook: Workbook;
  readonly view: View;
  /** Whether it was just made or read from a file, and is not stored yet. */
  readonly fresh?: boolean;
}

/** What the message bar says, about the last file opened or the app. */
type Message =
  | { readonly kind: 'invalid'; readonly name: string }
  | { readonly kind: 'invalidName'; readonly text: string }
  | { readonly kind: 'deletedName'; readonly text: string }
  | { readonly kind: 'saveFailed' }
  | { readonly kind: 'storageFailed' }
  | { readonly kind: 'storageUnreadable' }
  | { readonly kind: 'updated' }
  | {
      readonly kind: 'dropped';
      readonly name: string;
      readonly dropped: Dropped[];
    };

export function App({ platform }: { platform: Platform }) {
  const { preferences } = platform;
  const [language, setLanguage] = useState(() =>
    startingLanguage(preferences.get('language'), navigator.language),
  );
  const [theme, setTheme] = useState(() =>
    startingTheme(
      preferences.get('theme'),
      window.matchMedia('(prefers-color-scheme: dark)').matches,
    ),
  );
  // The workbook the address names after `#`, or none for the workbook
  // dialog alone.
  const [route, setRoute] = useState(() => workbookIdOf(location.hash));
  const [open, setOpen] = useState<Open>();
  // The empty interface shown behind the workbook dialog when no workbook
  // is open: a workbook with no name, which is never stored.
  const [blank] = useState(() => {
    const workbook = Workbook.fromSnapshot(
      { ...restoreWorkbook(null).toSnapshot(), name: '' },
      [],
      measureText,
    );
    return { workbook, view: new View(workbook) };
  });
  // Whether the workbook dialog is open over the workbook shown.
  const [choosing, setChoosing] = useState(false);
  // The autosave of `open`.
  const autosave = useRef<Autosave>(undefined);
  // The identifier of `open`, read where it must not be a dependency.
  const shown = useRef<string | null>(null);
  const [message, setMessage] = useState<Message | null>(null);
  const [busy, setBusy] = useState(false);
  const [finding, setFinding] = useState<FindMode | null>(null);
  const [editing, setEditing] = useState<Editing | null>(null);
  const [printing, setPrinting] = useState(false);

  // On the root element, so that the tooltips and menus drawn outside the
  // app's own element are themed too, and before the grid first paints.
  useLayoutEffect(() => {
    document.documentElement.lang = language === 'zh' ? 'zh-CN' : 'en';
    document.documentElement.dataset.theme = theme;
  }, [language, theme]);

  const show = (next: Open | undefined) => {
    shown.current = next?.id ?? null;
    setOpen(next);
    setChoosing(false);
  };

  useEffect(
    () => platform.listenToUpdates(() => setMessage({ kind: 'updated' })),
    [platform],
  );

  useEffect(() => {
    const onHashChange = () => setRoute(workbookIdOf(location.hash));
    window.addEventListener('hashchange', onHashChange);
    return () => window.removeEventListener('hashchange', onHashChange);
  }, []);

  // The address decides what is shown, so that a reload or the browser's
  // Back button shows it again.
  useEffect(() => {
    if (route === shown.current) return;
    if (route === null) {
      show(undefined);
      return;
    }
    let current = true;
    platform.workbooks
      .load(route)
      .then((stored) => {
        if (!current) return;
        if (!stored) {
          // No longer stored: the workbook dialog instead.
          location.hash = '';
          return;
        }
        const workbook = restoreWorkbook(stored, measureText);
        show({ id: route, workbook, view: new View(workbook) });
      })
      .catch(() => {
        if (!current) return;
        setMessage({ kind: 'storageUnreadable' });
        location.hash = '';
      });
    return () => {
      current = false;
    };
  }, [platform, route]);

  /** Shows a workbook just made, which is stored at once. */
  const showNew = (workbook: Workbook) => {
    const id = crypto.randomUUID();
    show({ id, workbook, view: new View(workbook), fresh: true });
    location.hash = hashOf(id);
  };

  useEffect(() => {
    if (!open) return;
    const { id, workbook, fresh } = open;
    const onFailure = () => setMessage({ kind: 'storageFailed' });
    const saving = startAutosave(
      workbook,
      {
        record: (command) => platform.workbooks.record(id, command),
        save: (snapshot) => platform.workbooks.save(id, snapshot),
      },
      onFailure,
    );
    autosave.current = saving;
    // A workbook made or read from a file is stored at once, so that the
    // workbook dialog lists it.
    if (fresh) {
      platform.workbooks.save(id, workbook.toSnapshot()).catch(onFailure);
    }
    // A hidden tab may be discarded without warning, so it saves at once.
    const onHidden = () => {
      if (document.visibilityState === 'hidden') saving.flush();
    };
    document.addEventListener('visibilitychange', onHidden);
    return () => {
      document.removeEventListener('visibilitychange', onHidden);
      saving.stop();
      autosave.current = undefined;
    };
  }, [platform, open]);

  // Copy, cut and paste work wherever the focus is, with the system
  // clipboard, but an entry being typed copies and pastes its own text.
  // Behind the workbook dialog, the workbook takes none of them.
  useEffect(() => {
    if (!open || choosing) return;
    const { view } = open;
    return platform.listenToClipboard({
      copy: (cut) => (view.state.entry ? null : cut ? view.cut() : view.copy()),
      paste: (text) => {
        if (view.state.entry) return false;
        view.pasteText(text);
        return true;
      },
    });
  }, [platform, open, choosing]);

  // Undo and redo work wherever the focus is, but an entry being typed
  // undoes its own typing.
  useEffect(() => {
    // Behind the workbook dialog, the workbook takes no shortcut.
    if (!open || choosing) return;
    const { view } = open;
    const onKeyDown = (event: KeyboardEvent) => {
      const command = shortcut(
        {
          key: event.key,
          shift: event.shiftKey,
          ctrl: event.ctrlKey,
          meta: event.metaKey,
          alt: event.altKey,
        },
        MAC,
      );
      if (!command) return;
      // The browser would print the app's page; an entry being typed is
      // committed first, unless data validation rejects it.
      if (command === 'print') {
        event.preventDefault();
        if (view.commit()) setPrinting(true);
        return;
      }
      if (view.state.entry) return;
      // Find and replace open from anywhere, a text box too, in place of
      // the browser's own find.
      if (command === 'find' || command === 'replace') {
        event.preventDefault();
        setFinding(command);
        return;
      }
      // A text box, such as a sheet name being typed, undoes its own typing.
      if (isTextBox(event.target)) return;
      event.preventDefault();
      if (
        command === 'bold' ||
        command === 'italic' ||
        command === 'underline'
      ) {
        view.toggle(command);
      } else {
        view[command]();
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [open, choosing]);

  const copyToClipboard = (contents: ClipboardContents | null) => {
    if (contents) platform.writeClipboard(contents);
  };

  const openFile = async () => {
    const file = await platform.openFile();
    if (!file) return;
    setBusy(true);
    try {
      const { readCsv, readXlsx } = await io;
      const name = workbookName(file.name);
      // The decoder drops a byte order mark and replaces bytes that are
      // not UTF-8.
      const { snapshot, dropped } = isCsv(file.name)
        ? readCsv(new TextDecoder().decode(file.bytes), name)
        : await readXlsx(file.bytes, name);
      showNew(Workbook.fromSnapshot(snapshot, [], measureText));
      setMessage(
        dropped.length > 0
          ? { kind: 'dropped', name: file.name, dropped }
          : null,
      );
    } catch {
      // The workbook shown stays as it was.
      setMessage({ kind: 'invalid', name: file.name });
    } finally {
      setBusy(false);
    }
  };

  const saveFile = async (workbook: Workbook) => {
    setBusy(true);
    try {
      const { writeXlsx } = await io;
      const snapshot = workbook.toSnapshot();
      await platform.saveFile(
        `${snapshot.name}.xlsx`,
        await writeXlsx(snapshot),
      );
    } catch {
      setMessage({ kind: 'saveFailed' });
    } finally {
      setBusy(false);
    }
  };

  const exportCsv = async (workbook: Workbook, view: View) => {
    setBusy(true);
    try {
      const { writeCsv } = await io;
      await platform.saveFile(
        `${workbook.name}.csv`,
        new TextEncoder().encode(writeCsv(workbook, view.state.sheet)),
      );
    } catch {
      setMessage({ kind: 'saveFailed' });
    } finally {
      setBusy(false);
    }
  };

  // A workbook being loaded shows nothing yet.
  if (route !== null && !open) return null;
  const t = translator(language);
  const messageBar = message && (
    <MessageBar
      text={
        message.kind === 'invalid'
          ? t('file.invalid').replace('{name}', () => message.name)
          : message.kind === 'invalidName' || message.kind === 'deletedName'
            ? t(
                message.kind === 'invalidName'
                  ? 'names.invalid'
                  : 'names.deleted',
              ).replace('{text}', () => message.text)
            : message.kind === 'saveFailed'
              ? t('file.saveFailed')
              : message.kind === 'storageFailed'
                ? t('storage.failed')
                : message.kind === 'storageUnreadable'
                  ? t('storage.unreadable')
                  : message.kind === 'updated'
                    ? t('app.updated')
                    : droppedMessage(message.name, message.dropped, language)
      }
      error={message.kind !== 'dropped' && message.kind !== 'updated'}
      onClose={() => setMessage(null)}
    />
  );
  const choosingOpen = !open || choosing;
  const shell = open ?? blank;
  // A rejected entry keeps the dialogs closed, so that what they set is
  // not lost.
  const edit = (what: Editing) => {
    if (shell.view.commit()) setEditing(what);
  };
  return (
    <Translation value={t}>
      <TooltipProvider delayDuration={400} disableHoverableContent>
        {/* With no workbook open, the interface does not respond. */}
        <div className={styles.app} inert={!open}>
          <Ribbon
            workbook={shell.workbook}
            view={shell.view}
            language={language}
            onLanguage={(language: Language) => {
              setLanguage(language);
              preferences.set('language', language);
            }}
            theme={theme}
            onTheme={(theme: Theme) => {
              setTheme(theme);
              preferences.set('theme', theme);
            }}
            busy={busy}
            onWorkbooks={() => {
              // Saved first, so that the dialog lists the workbook as it is.
              autosave.current?.flush();
              setChoosing(true);
            }}
            onOpen={() => void openFile()}
            onSave={() => void saveFile(shell.workbook)}
            onExportCsv={() => void exportCsv(shell.workbook, shell.view)}
            onPrint={() => setPrinting(true)}
            onCopy={copyToClipboard}
            onFind={setFinding}
            onEdit={edit}
          />
          {!choosingOpen && messageBar}
          <FormulaBar
            {...shell}
            onInvalid={(text, problem) =>
              setMessage({
                kind: problem === 'invalid' ? 'invalidName' : 'deletedName',
                text,
              })
            }
          />
          <main className={styles.grid}>
            <Grid
              {...shell}
              menu={(lines) => (
                <CellMenu
                  view={shell.view}
                  onCopy={copyToClipboard}
                  onEdit={edit}
                  lines={lines}
                />
              )}
              fillOptions={(filled) => (
                <FillOptions view={shell.view} filled={filled} />
              )}
              filterButton={(col) => (
                <FilterButton
                  workbook={shell.workbook}
                  view={shell.view}
                  col={col}
                />
              )}
              listButton={<ListButton {...shell} />}
            />
          </main>
          <SheetTabs {...shell} />
          <StatusBar {...shell} />
          <FindPanel view={shell.view} mode={finding} onMode={setFinding} />
          <PrintPanel
            {...shell}
            open={printing}
            onOpenChange={setPrinting}
            onPrint={(document) => platform.print(document)}
          />
          <RejectedEntry view={shell.view} />
          <CellDialog
            {...shell}
            editing={editing}
            onClose={() => setEditing(null)}
          />
        </div>
        <WorkbookDialog
          platform={platform}
          language={language}
          open={choosingOpen}
          current={open}
          message={choosingOpen && messageBar}
          busy={busy}
          onClose={() => setChoosing(false)}
          onNew={(name) => {
            const workbook = restoreWorkbook(null, measureText);
            workbook.execute({ type: 'renameWorkbook', name });
            showNew(workbook);
          }}
          onOpenFile={() => void openFile()}
          onOpen={(id) => {
            setMessage(null);
            if (id === shown.current) setChoosing(false);
            else location.hash = hashOf(id);
          }}
          onFailure={(how) =>
            setMessage({
              kind: how === 'read' ? 'storageUnreadable' : 'storageFailed',
            })
          }
        />
      </TooltipProvider>
    </Translation>
  );
}

function isTextBox(target: EventTarget | null): boolean {
  return target instanceof HTMLInputElement && target.type === 'text';
}
