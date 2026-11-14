import * as ContextMenu from '@radix-ui/react-context-menu';
import {
  ChevronRight,
  ClipboardPaste,
  Copy,
  Scissors,
  type LucideIcon,
} from 'lucide-react';
import { useSyncExternalStore } from 'react';
import { MAC } from '../grid/Grid';
import type { View } from '../grid/view';
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

/** The menu that right-clicking a cell opens. */
export function CellMenu({ view }: { view: View }) {
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
          onSelect={() => view.cut()}
        />
        <MenuItem
          icon={Copy}
          label={t('copy')}
          keys={`${key}C`}
          onSelect={() => view.copy()}
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
  keys: string;
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
      <span className={styles.menuEnd} aria-hidden>
        {keys}
      </span>
    </ContextMenu.Item>
  );
}
