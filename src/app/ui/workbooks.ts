import type { Snapshot } from '../../engine/workbook';
import type { Stored } from '../platform';
import { restoreWorkbook } from './autosave';
import type { Language } from './i18n';

/**
 * The identifier of the workbook an address shows after `#`, or `null` for
 * the start screen.
 */
export function workbookIdOf(hash: string): string | null {
  try {
    const id = decodeURIComponent(hash.replace(/^#/, ''));
    return id === '' ? null : id;
  } catch {
    return null; // Not an address the app made.
  }
}

/** The part of the address after which a workbook is shown. */
export function hashOf(id: string): string {
  return `#${encodeURIComponent(id)}`;
}

/** `stem` and the first number from 1 that no name in `taken` has. */
export function numberedName(stem: string, taken: readonly string[]): string {
  const folded = new Set(taken.map((name) => name.toLowerCase()));
  let number = 1;
  while (folded.has(`${stem}${number}`.toLowerCase())) number++;
  return `${stem}${number}`;
}

/**
 * When a workbook was last modified, at `time`: the time of day alone when
 * it is the same day as `now`, and else the date with it.
 */
export function modifiedText(
  time: number,
  now: number,
  language: Language,
): string {
  const locale = language === 'zh' ? 'zh-CN' : 'en-US';
  const date = new Date(time);
  const today = new Date(now).toDateString() === date.toDateString();
  return date.toLocaleString(locale, {
    ...(!today && { dateStyle: 'medium' }),
    timeStyle: 'short',
  });
}

/**
 * A stored workbook, with its changes replayed, under another name; one no
 * longer stored, as when another tab deleted it, is an error.
 */
export function renamedSnapshot(stored: Stored | null, name: string): Snapshot {
  if (!stored) throw new Error(`The workbook ${name} is no longer stored`);
  const workbook = restoreWorkbook(stored);
  workbook.execute({ type: 'renameWorkbook', name });
  return workbook.toSnapshot();
}
