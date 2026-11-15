import * as ContextMenu from '@radix-ui/react-context-menu';
import {
  ChevronRight,
  ClipboardPaste,
  Copy,
  Eye,
  EyeOff,
  Plus,
  Scissors,
  Trash2,
  type LucideIcon,
} from 'lucide-react';
import { useSyncExternalStore } from 'react';
import { MAC } from '../grid/Grid';
import type { Direction } from '../../engine/workbook';
import type { View } from '../grid/view';
import type { ClipboardContents } from '../platform';
import { useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

/** The kinds of paste special, with their labels. */
export const PASTE_SPECIAL = [
  ['values', 'paste.values'],
  ['formulas', 'paste.formulas'],
  ['formats', 'paste.formats'],
] as const satisfies readonly (readonly [string, Key])[];

/** The cells copied or cut in a view, or `null`. */
export function useCopied(view: View) {
  return useSyncExternalStore(view.subscribe, () => view.state.copied);
}

/**
 * The menu that right-clicking a cell opens, or a row or column header,
 * with the actions on the rows or columns selected.
 */
export function CellMenu({
  view,
  onCopy,
  lines,
}: {
  view: View;
  /** Puts the cells copied or cut on the system clipboard. */
  onCopy: (contents: ClipboardContents) => void;
  /** Whether the menu is that of row or column headers. */
  lines?: Direction;
}) {
  const t = useTranslation();
  const copied = useCopied(view);
  const key = MAC ? '⌘' : 'Ctrl+';
  return (
    <ContextMenu.Portal>
      <ContextMenu.Content className={styles.menu}>
        <MenuItem
          icon={Scissors}
          label={t('cut')}
          keys={`${key}X`}
          onSelect={() => onCopy(view.cut())}
        />
        <MenuItem
          icon={Copy}
          label={t('copy')}
          keys={`${key}C`}
          onSelect={() => onCopy(view.copy())}
        />
        <MenuItem
          icon={ClipboardPaste}
          label={t('paste')}
          keys={`${key}V`}
          disabled={!copied}
          onSelect={() => view.paste()}
        />
        <ContextMenu.Sub>
          <ContextMenu.SubTrigger
            className={styles.menuItem}
            disabled={!copied || copied.cut}
          >
            <ClipboardPaste size={ICON_SIZE} aria-hidden />
            {t('paste.special')}
            <ChevronRight
              className={styles.menuEnd}
              size={ICON_SIZE}
              aria-hidden
            />
          </ContextMenu.SubTrigger>
          <ContextMenu.Portal>
            <ContextMenu.SubContent className={styles.menu} sideOffset={4}>
              {PASTE_SPECIAL.map(([only, label]) => (
                <ContextMenu.Item
                  key={only}
                  className={styles.menuItem}
                  onSelect={() => view.paste(only)}
                >
                  {t(label)}
                </ContextMenu.Item>
              ))}
            </ContextMenu.SubContent>
          </ContextMenu.Portal>
        </ContextMenu.Sub>
        {lines && (
          <>
            <ContextMenu.Separator className={styles.menuSeparator} />
            <MenuItem
              icon={Plus}
              label={t(`${lines}.insert`)}
              onSelect={() => view.insertLines(lines)}
            />
            <MenuItem
              icon={Trash2}
              label={t(`${lines}.delete`)}
              onSelect={() => view.deleteLines(lines)}
            />
            <MenuItem
              icon={EyeOff}
              label={t(`${lines}.hide`)}
              onSelect={() => view.hideLines(lines, true)}
            />
            <MenuItem
              icon={Eye}
              label={t(`${lines}.unhide`)}
              disabled={!view.hasHidden(lines)}
              onSelect={() => view.hideLines(lines, false)}
            />
          </>
        )}
      </ContextMenu.Content>
    </ContextMenu.Portal>
  );
}

function MenuItem({
  icon: Icon,
  label,
  keys,
  disabled,
  onSelect,
}: {
  icon: LucideIcon;
  label: string;
  keys?: string;
  disabled?: boolean;
  onSelect: () => void;
}) {
  return (
    <ContextMenu.Item
      className={styles.menuItem}
      disabled={disabled}
      onSelect={onSelect}
    >
      <Icon size={ICON_SIZE} aria-hidden />
      {label}
      {keys && (
        <span className={styles.menuEnd} aria-hidden>
          {keys}
        </span>
      )}
    </ContextMenu.Item>
  );
}
