import type { KeyPress } from '../grid/navigation';

/**
 * A command of the window that a key runs wherever the focus is. Copy, cut
 * and paste come as clipboard events instead, which reach the system
 * clipboard.
 */
export type Shortcut = 'undo' | 'redo' | 'fillDown' | 'fillRight';

/**
 * The shortcut a key runs, as in Excel, with Cmd in place of Ctrl on a
 * Mac, or `null` for any other key.
 */
export function shortcut(press: KeyPress, mac: boolean): Shortcut | null {
  const { key, shift, ctrl, meta, alt } = press;
  const command = mac ? meta && !ctrl : ctrl && !meta;
  if (!command || shift || alt) return null;
  switch (key.toLowerCase()) {
    case 'z':
      return 'undo';
    case 'y':
      return 'redo';
    case 'd':
      return 'fillDown';
    case 'r':
      return 'fillRight';
    default:
      return null;
  }
}
