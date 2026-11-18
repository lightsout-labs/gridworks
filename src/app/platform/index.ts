import type { Command, Snapshot } from '../../engine/workbook';

/** A workbook as it is stored. */
export interface Stored {
  /** The workbook as it was last saved, or `null` if it never was. */
  readonly snapshot: Snapshot | null;
  /** The changes recorded after that save, in order. */
  readonly changes: Command[];
}

/** What the start screen shows of a stored workbook. */
export interface StoredInfo {
  readonly id: string;
  readonly name: string;
  /** When it was last saved, in milliseconds since 1970. */
  readonly modified: number;
}

/** A file the user chose to open. */
export interface ChosenFile {
  readonly name: string;
  readonly bytes: Uint8Array;
}

/** What a copy puts on the system clipboard. */
export interface ClipboardContents {
  /** Tab-separated text. */
  readonly text: string;
  /** An HTML table. */
  readonly html: string;
}

/**
 * The app's answers to the user's copy, cut and paste outside a text box,
 * which text boxes leave to the system.
 */
export interface ClipboardListener {
  /** What to put on the clipboard, or `null` to leave it to the system. */
  copy(cut: boolean): ClipboardContents | null;
  /** Pastes the text on the clipboard, and gives whether it did. */
  paste(text: string): boolean;
}

/**
 * Everything the app needs from the system it runs on. Only an
 * implementation of this interface may touch files, storage, the clipboard
 * or printing; `web` is the implementation for browsers.
 */
export interface Platform {
  /**
   * Asks the user for an `.xlsx` or `.csv` file and reads it, or gives `null` when
   * the user chooses none.
   */
  openFile(): Promise<ChosenFile | null>;
  /** Gives the user a file, which the system keeps where it keeps files. */
  saveFile(name: string, bytes: Uint8Array): Promise<void>;
  /** The workbooks kept in the system's storage. */
  readonly workbooks: {
    /** The workbooks saved, the one saved last first. */
    list(): Promise<StoredInfo[]>;
    /** The stored workbook, or `null` when there is none with this identifier. */
    load(id: string): Promise<Stored | null>;
    /**
     * Keeps a change to a workbook. It is small and quick to store, so that
     * a change is kept even when the browser closes right after it.
     */
    record(id: string, command: Command): Promise<void>;
    /**
     * Keeps a whole workbook, in place of the changes recorded before this
     * call.
     */
    save(id: string, snapshot: Snapshot): Promise<void>;
    /** Forgets a workbook and the changes recorded to it. */
    delete(id: string): Promise<void>;
  };
  /**
   * Calls `listener` when the user copies, cuts or pastes with the system
   * clipboard, which other applications share, by the keyboard or the
   * browser's menu, until the returned function is called.
   */
  listenToClipboard(listener: ClipboardListener): () => void;
  /**
   * Puts contents on the system clipboard, as a button that copies does.
   */
  writeClipboard(contents: ClipboardContents): void;
  /** The user's choices, such as the language, that outlive the page. */
  readonly preferences: {
    /** The remembered value, or `null` when there is none. */
    get(name: string): string | null;
    set(name: string, value: string): void;
  };
}
