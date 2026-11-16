import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import { ClipboardList } from 'lucide-react';
import type { Filled, FillHow, View } from '../grid/view';
import { useTranslation } from './controls';
import type { Key } from './i18n';
import styles from './Shell.module.css';

const OPTIONS: readonly [FillHow, Key][] = [
  ['copy', 'fill.copy'],
  ['series', 'fill.series'],
  ['formats', 'fill.formats'],
  ['values', 'fill.values'],
];

/**
 * The Auto Fill Options button of the last fill, and its menu of the ways
 * to redo it; Fill Series only when the cells can form a series.
 */
export function FillOptions({ view, filled }: { view: View; filled: Filled }) {
  const t = useTranslation();
  return (
    <DropdownMenu.Root>
      <DropdownMenu.Trigger asChild>
        <button
          type="button"
          className={styles.fillOptions}
          aria-label={t('fill.options')}
          title={t('fill.options')}
        >
          <ClipboardList size={14} aria-hidden />
        </button>
      </DropdownMenu.Trigger>
      <DropdownMenu.Portal>
        <DropdownMenu.Content
          className={styles.menu}
          align="start"
          sideOffset={4}
        >
          {OPTIONS.filter(([how]) => how !== 'series' || filled.series).map(
            ([how, label]) => (
              <DropdownMenu.Item
                key={how}
                className={styles.menuItem}
                onSelect={() => view.refill(how)}
              >
                {t(label)}
              </DropdownMenu.Item>
            ),
          )}
        </DropdownMenu.Content>
      </DropdownMenu.Portal>
    </DropdownMenu.Root>
  );
}
