import {
  createWorkbook,
  Workbook,
  type Command,
  type Snapshot,
  type TextWidth,
} from '../../engine/workbook';
import type { Stored } from '../platform';

const DELAY_MS = 500;

/** Where a workbook is kept. */
export interface Store {
  /** Keeps one change, until the next save includes it. */
  record(command: Command): Promise<void>;
  /** Keeps the whole workbook, in place of the changes recorded so far. */
  save(snapshot: Snapshot): Promise<void>;
}

export interface Autosave {
  /** Saves now if a change is waiting. */
  flush(): void;
  /** Saves a waiting change and stops watching the workbook. */
  stop(): void;
}

/**
 * Keeps every change to the workbook as it is made, so that none is lost
 * when the browser closes, and saves the whole workbook half a second after
 * the last change. `onFailure` is called when the store refuses a change or
 * a save, as when it is full.
 */
export function startAutosave(
  workbook: Workbook,
  store: Store,
  onFailure: (error: unknown) => void,
): Autosave {
  let timer: ReturnType<typeof setTimeout> | undefined;

  const flush = () => {
    if (timer === undefined) return;
    clearTimeout(timer);
    timer = undefined;
    store.save(workbook.toSnapshot()).catch(onFailure);
  };
  const unsubscribe = workbook.subscribe((command) => {
    store.record(command).catch(onFailure);
    clearTimeout(timer);
    timer = setTimeout(flush, DELAY_MS);
  });

  return {
    flush,
    stop() {
      unsubscribe();
      flush();
    },
  };
}

/**
 * The workbook as it was stored: the saved workbook, or a new one, with
 * the changes recorded after it replayed, its rows fitted to text as
 * `measure` measures it.
 */
export function restoreWorkbook(
  stored: Stored | null,
  measure?: TextWidth,
): Workbook {
  return Workbook.fromSnapshot(
    stored?.snapshot ?? createWorkbook().toSnapshot(),
    stored?.changes,
    measure,
  );
}
