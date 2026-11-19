import type {
  Command,
  NamedRange,
  SheetSnapshot,
  Snapshot,
  Style,
} from '../../../engine/workbook';
import type {
  ChosenFile,
  ClipboardContents,
  ClipboardListener,
  Platform,
  Stored,
  StoredInfo,
} from '..';

// The IndexedDB database `gridworks`. `workbooks` has one record for each
// workbook; `sheets` has one for each sheet, keyed by workbook and sheet;
// `changes` has the changes made since a workbook was last saved, keyed by
// workbook and a sequence number.
const WORKBOOKS = 'workbooks';
const SHEETS = 'sheets';
const CHANGES = 'changes';

interface WorkbookRecord {
  id: string;
  name: string;
  modified: number;
  version: Snapshot['version'];
  sheetIds: string[];
  /** Absent in records saved before cells had styles. */
  styles?: Style[];
  /** The defined names, when there are any. */
  names?: NamedRange[];
}

type SheetRecord = SheetSnapshot & { workbook: string };

interface ChangeRecord {
  workbook: string;
  sequence: number;
  command: Command;
}

/** The sequence number of the last change recorded or loaded. */
let sequence = 0;

let opening: Promise<IDBDatabase> | undefined;

function database(): Promise<IDBDatabase> {
  opening ??= new Promise((resolve, reject) => {
    const request = indexedDB.open('gridworks', 2);
    request.onupgradeneeded = ({ oldVersion }) => {
      const db = request.result;
      if (oldVersion < 1) {
        db.createObjectStore(WORKBOOKS, { keyPath: 'id' });
        db.createObjectStore(SHEETS, { keyPath: ['workbook', 'id'] });
      }
      db.createObjectStore(CHANGES, { keyPath: ['workbook', 'sequence'] });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => {
      opening = undefined; // The next call tries again.
      reject(request.error!);
    };
  });
  return opening;
}

function result<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error!);
  });
}

/** The changes of a workbook up to a sequence number. */
const changesOf = (id: string, last = Infinity) =>
  IDBKeyRange.bound([id, -Infinity], [id, last]);

function completion(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error!);
    transaction.onabort = () => reject(transaction.error!);
  });
}

async function load(id: string): Promise<Stored | null> {
  const transaction = (await database()).transaction([
    WORKBOOKS,
    SHEETS,
    CHANGES,
  ]);
  const sheetStore = transaction.objectStore(SHEETS);
  const [workbook, changes] = await Promise.all([
    result<WorkbookRecord | undefined>(
      transaction.objectStore(WORKBOOKS).get(id),
    ),
    result<ChangeRecord[]>(
      transaction.objectStore(CHANGES).getAll(changesOf(id)),
    ),
  ]);
  sequence = Math.max(sequence, ...changes.map((change) => change.sequence));
  if (!workbook && changes.length === 0) return null;
  const snapshot: Snapshot | null = workbook
    ? {
        version: workbook.version,
        name: workbook.name,
        styles: workbook.styles,
        ...(workbook.names && { names: workbook.names }),
        sheets: await Promise.all(
          workbook.sheetIds.map(async (sheetId) => {
            const record: Partial<SheetRecord> = await result<SheetRecord>(
              sheetStore.get([id, sheetId]),
            );
            // Everything the snapshot of the sheet holds, but the key.
            delete record.workbook;
            return record as SheetSnapshot;
          }),
        ),
      }
    : null;
  return { snapshot, changes: changes.map((change) => change.command) };
}

async function record(id: string, command: Command): Promise<void> {
  // Numbered at once, so that a save called next counts it as saved; its
  // transaction is created first, so it is stored before that save.
  const change: ChangeRecord = { workbook: id, sequence: ++sequence, command };
  const transaction = (await database()).transaction(CHANGES, 'readwrite');
  transaction.objectStore(CHANGES).put(change);
  await completion(transaction);
}

async function save(id: string, snapshot: Snapshot): Promise<void> {
  // The snapshot holds the changes recorded so far, and no later one.
  const last = sequence;
  const transaction = (await database()).transaction(
    [WORKBOOKS, SHEETS, CHANGES],
    'readwrite',
  );
  const workbook: WorkbookRecord = {
    id,
    name: snapshot.name,
    modified: Date.now(),
    version: snapshot.version,
    sheetIds: snapshot.sheets.map((sheet) => sheet.id),
    styles: snapshot.styles,
    ...(snapshot.names && { names: snapshot.names }),
  };
  transaction.objectStore(WORKBOOKS).put(workbook);
  // Deleted sheets go too: every sheet of the workbook is written again.
  // Arrays sort after strings, so `[id, []]` follows every sheet key.
  transaction.objectStore(SHEETS).delete(IDBKeyRange.bound([id], [id, []]));
  for (const sheet of snapshot.sheets) {
    transaction
      .objectStore(SHEETS)
      .put({ ...sheet, workbook: id } satisfies SheetRecord);
  }
  transaction.objectStore(CHANGES).delete(changesOf(id, last));
  await completion(transaction);
}

async function list(): Promise<StoredInfo[]> {
  const transaction = (await database()).transaction(WORKBOOKS);
  const records = await result<WorkbookRecord[]>(
    transaction.objectStore(WORKBOOKS).getAll(),
  );
  return records
    .map(({ id, name, modified }) => ({ id, name, modified }))
    .sort((a, b) => b.modified - a.modified);
}

async function remove(id: string): Promise<void> {
  const transaction = (await database()).transaction(
    [WORKBOOKS, SHEETS, CHANGES],
    'readwrite',
  );
  transaction.objectStore(WORKBOOKS).delete(id);
  transaction.objectStore(SHEETS).delete(IDBKeyRange.bound([id], [id, []]));
  transaction.objectStore(CHANGES).delete(changesOf(id));
  await completion(transaction);
}

// A file input works in every browser; the File System Access API does not.
function openFile(): Promise<ChosenFile | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.csv';
    input.hidden = true;
    let changed = false;
    input.addEventListener('change', () => {
      changed = true;
      input.remove();
      const file = input.files?.[0];
      if (!file) return resolve(null);
      file
        .arrayBuffer()
        .then((buffer) =>
          resolve({ name: file.name, bytes: new Uint8Array(buffer) }),
        )
        .catch(reject);
    });
    // Firefox may fire `cancel` just before the `change` of a file chosen,
    // so a cancellation counts only when no change follows it; a late
    // answer costs nothing, as no file chosen means nothing to do.
    input.addEventListener('cancel', () => {
      setTimeout(() => {
        if (changed) return;
        input.remove();
        resolve(null);
      }, 1000);
    });
    document.body.append(input);
    input.click();
  });
}

// A download: browsers keep it in the downloads folder or ask where.
function saveFile(name: string, bytes: Uint8Array): Promise<void> {
  const url = URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>]));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.hidden = true;
  document.body.append(link);
  link.click();
  link.remove();
  // Revoked later, as some browsers read the address after the click.
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
  return Promise.resolve();
}

/** Whether an event comes from where text is typed, which handles its own. */
function inTextBox(event: Event): boolean {
  const target = event.target;
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

// The clipboard events reach the page on the keyboard shortcuts in every
// browser, with no permission asked, and the focused element gets them.
function listenToClipboard(listener: ClipboardListener): () => void {
  const onCopy = (event: ClipboardEvent) => {
    if (inTextBox(event) || !event.clipboardData) return;
    const contents = listener.copy(event.type === 'cut');
    if (!contents) return;
    event.preventDefault();
    event.clipboardData.setData('text/plain', contents.text);
    event.clipboardData.setData('text/html', contents.html);
  };
  const onPaste = (event: ClipboardEvent) => {
    if (inTextBox(event) || !event.clipboardData) return;
    if (listener.paste(event.clipboardData.getData('text/plain'))) {
      event.preventDefault();
    }
  };
  document.addEventListener('copy', onCopy);
  document.addEventListener('cut', onCopy);
  document.addEventListener('paste', onPaste);
  return () => {
    document.removeEventListener('copy', onCopy);
    document.removeEventListener('cut', onCopy);
    document.removeEventListener('paste', onPaste);
  };
}

// The copy command, unlike the asynchronous clipboard API, needs no
// permission in any browser while the user clicks, and has finished when it
// returns, so a paste right after it reads what it wrote.
function writeClipboard({ text, html }: ClipboardContents): void {
  const onCopy = (event: ClipboardEvent) => {
    // Answered here, not by the app's listener.
    event.stopImmediatePropagation();
    event.preventDefault();
    event.clipboardData?.setData('text/plain', text);
    event.clipboardData?.setData('text/html', html);
  };
  document.addEventListener('copy', onCopy, { capture: true });
  try {
    document.execCommand('copy');
  } finally {
    document.removeEventListener('copy', onCopy, { capture: true });
  }
}

/** The platform for browsers. */
// The service worker that vite-plugin-pwa builds keeps the whole app, so
// that it opens without a network; it is registered as the page loads,
// whoever listens. A new version takes the page over as soon as it is
// installed, while the page still runs the old one.
const workers =
  import.meta.env.PROD && 'serviceWorker' in navigator
    ? navigator.serviceWorker
    : null;
// The worker that served this page, none on the first visit, which the
// version the page runs takes over.
const servedBy = workers?.controller ?? null;
// Without a worker, the app still works online.
workers?.register(`${import.meta.env.BASE_URL}sw.js`).catch(() => {});

function listenToUpdates(listener: () => void): () => void {
  if (!workers) return () => {};
  const onChange = () => {
    if (servedBy && workers.controller !== servedBy) listener();
  };
  // The new version may have taken over before this call.
  onChange();
  workers.addEventListener('controllerchange', onChange);
  return () => workers.removeEventListener('controllerchange', onChange);
}

export const webPlatform: Platform = {
  openFile,
  saveFile,
  workbooks: { list, load, record, save, delete: remove },
  listenToClipboard,
  writeClipboard,
  listenToUpdates,
  // A browser that blocks storage throws; the choice is then not remembered.
  preferences: {
    get(name) {
      try {
        return localStorage.getItem(`gridworks.${name}`);
      } catch {
        return null;
      }
    },
    set(name, value) {
      try {
        localStorage.setItem(`gridworks.${name}`, value);
      } catch {
        // Not remembered.
      }
    },
  },
};
