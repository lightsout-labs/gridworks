import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import * as Tabs from '@radix-ui/react-tabs';
import {
  ChevronDown,
  ClipboardPaste,
  Copy,
  Download,
  FolderOpen,
  SquareFunction,
  Moon,
  Redo2,
  Scissors,
  Sun,
  Table2,
  Undo2,
  type LucideIcon,
} from 'lucide-react';
import {
  useCallback,
  useId,
  useSyncExternalStore,
  type ReactNode,
} from 'react';
import type { Workbook } from '../../engine/workbook';
import { MAC } from '../grid/Grid';
import type { View } from '../grid/view';
import type { ClipboardContents } from '../platform';
import { PASTE_SPECIAL, useCopied } from './CellMenu';
import { Tooltip, useTranslation } from './controls';
import { AlignmentControls, FontControls } from './Format';
import type { Language } from './i18n';
import styles from './Shell.module.css';
import type { Theme } from './theme';

const ICON_SIZE = 16;

/**
 * The toolbar at the top of the window: tabs, each with the controls of
 * its features in labeled groups. A tab or control is added by the slice
 * that builds its feature.
 */
export function Ribbon({
  workbook,
  view,
  language,
  onLanguage,
  theme,
  onTheme,
  busy,
  onOpen,
  onSave,
  onCopy,
}: {
  workbook: Workbook;
  view: View;
  language: Language;
  onLanguage: (language: Language) => void;
  theme: Theme;
  onTheme: (theme: Theme) => void;
  /** Whether a file is being opened or saved, when neither can start. */
  busy: boolean;
  onOpen: () => void;
  onSave: () => void;
  /** Puts the cells copied or cut on the system clipboard. */
  onCopy: (contents: ClipboardContents) => void;
}) {
  const t = useTranslation();
  return (
    <Tabs.Root className={styles.ribbon} defaultValue="home">
      <div className={styles.ribbonTabs}>
        <span className={styles.brand}>
          <Table2 size={ICON_SIZE} aria-hidden />
          Gridworks
        </span>
        <Tabs.List aria-label={t('ribbon')}>
          <Tooltip text={t('tab.file.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="file">
              {t('tab.file')}
            </Tabs.Trigger>
          </Tooltip>
          <Tooltip text={t('tab.home.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="home">
              {t('tab.home')}
            </Tabs.Trigger>
          </Tooltip>
          <Tooltip text={t('tab.formulas.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="formulas">
              {t('tab.formulas')}
            </Tabs.Trigger>
          </Tooltip>
          <Tooltip text={t('tab.view.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="view">
              {t('tab.view')}
            </Tabs.Trigger>
          </Tooltip>
        </Tabs.List>
      </div>
      <Tabs.Content className={styles.ribbonPanel} value="file">
        <Group label={t('file.group')}>
          <Command
            icon={FolderOpen}
            label={t('file.open')}
            tooltip={t('file.open.tooltip')}
            disabled={busy}
            onRun={onOpen}
            labeled
          />
          <Command
            icon={Download}
            label={t('file.save')}
            tooltip={t('file.save.tooltip')}
            disabled={busy}
            onRun={onSave}
            labeled
          />
        </Group>
      </Tabs.Content>
      <Tabs.Content className={styles.ribbonPanel} value="home">
        <Group label={t('undo')}>
          <History workbook={workbook} view={view} />
        </Group>
        <Group label={t('clipboard')}>
          <Clipboard view={view} onCopy={onCopy} />
        </Group>
        <Group label={t('font')}>
          <FontControls workbook={workbook} view={view} />
        </Group>
        <Group label={t('alignment')}>
          <AlignmentControls workbook={workbook} view={view} />
        </Group>
      </Tabs.Content>
      <Tabs.Content className={styles.ribbonPanel} value="formulas">
        <Group label={t('formulas.library')}>
          <Command
            icon={SquareFunction}
            label={t('formulas.insert')}
            tooltip={t('formulas.insert.tooltip')}
            disabled={false}
            onRun={() => view.insertFunction()}
            keepsEntry
            labeled
          />
        </Group>
      </Tabs.Content>
      <Tabs.Content className={styles.ribbonPanel} value="view">
        <Group label={t('theme')}>
          <Choice
            icon={Sun}
            label={t('theme.light')}
            tooltip={t('theme.light.tooltip')}
            chosen={theme === 'light'}
            onChoose={() => onTheme('light')}
          />
          <Choice
            icon={Moon}
            label={t('theme.dark')}
            tooltip={t('theme.dark.tooltip')}
            chosen={theme === 'dark'}
            onChoose={() => onTheme('dark')}
          />
        </Group>
        <Group label={t('language')}>
          {/* Each language is named in itself, whatever the interface is in. */}
          <Choice
            label="English"
            tooltip={t('language.en.tooltip')}
            chosen={language === 'en'}
            onChoose={() => onLanguage('en')}
          />
          <Choice
            label="中文"
            tooltip={t('language.zh.tooltip')}
            chosen={language === 'zh'}
            onChoose={() => onLanguage('zh')}
          />
        </Group>
      </Tabs.Content>
    </Tabs.Root>
  );
}

function Group({ label, children }: { label: string; children: ReactNode }) {
  const id = useId();
  return (
    <div className={styles.group} role="group" aria-labelledby={id}>
      <span id={id} className={styles.groupLabel}>
        {label}
      </span>
      {children}
    </div>
  );
}

/** The Undo and Redo buttons, enabled when there is something to do. */
function History({ workbook, view }: { workbook: Workbook; view: View }) {
  const t = useTranslation();
  const subscribe = useCallback(
    (listener: () => void) => workbook.subscribe(listener),
    [workbook],
  );
  const canUndo = useSyncExternalStore(subscribe, () => workbook.canUndo);
  const canRedo = useSyncExternalStore(subscribe, () => workbook.canRedo);
  const key = MAC ? '⌘' : 'Ctrl+';
  return (
    <>
      <Command
        icon={Undo2}
        label={t('undo')}
        tooltip={`${t('undo.tooltip')} (${key}Z)`}
        disabled={!canUndo}
        onRun={() => view.undo()}
      />
      <Command
        icon={Redo2}
        label={t('redo')}
        tooltip={`${t('redo.tooltip')} (${key}Y)`}
        disabled={!canRedo}
        onRun={() => view.redo()}
      />
    </>
  );
}

/** Paste, paste special, cut and copy. */
function Clipboard({
  view,
  onCopy,
}: {
  view: View;
  onCopy: (contents: ClipboardContents) => void;
}) {
  const t = useTranslation();
  const copied = useCopied(view);
  const key = MAC ? '⌘' : 'Ctrl+';
  return (
    <>
      <Command
        icon={ClipboardPaste}
        label={t('paste')}
        tooltip={`${t('paste.tooltip')} (${key}V)`}
        disabled={!copied}
        onRun={() => view.paste()}
        labeled
      />
      <DropdownMenu.Root>
        <Tooltip text={t('paste.special.tooltip')}>
          <DropdownMenu.Trigger asChild>
            <button
              type="button"
              className={styles.button}
              aria-label={t('paste.special')}
              disabled={!copied || copied.cut}
            >
              <ChevronDown size={ICON_SIZE} aria-hidden />
            </button>
          </DropdownMenu.Trigger>
        </Tooltip>
        <DropdownMenu.Portal>
          <DropdownMenu.Content
            className={styles.menu}
            align="start"
            sideOffset={4}
          >
            {PASTE_SPECIAL.map(([only, label]) => (
              <DropdownMenu.Item
                key={only}
                className={styles.menuItem}
                onSelect={() => view.paste(only)}
              >
                {t(label)}
              </DropdownMenu.Item>
            ))}
          </DropdownMenu.Content>
        </DropdownMenu.Portal>
      </DropdownMenu.Root>
      <Command
        icon={Scissors}
        label={t('cut')}
        tooltip={`${t('cut.tooltip')} (${key}X)`}
        disabled={false}
        onRun={() => onCopy(view.cut())}
      />
      <Command
        icon={Copy}
        label={t('copy')}
        tooltip={`${t('copy.tooltip')} (${key}C)`}
        disabled={false}
        onRun={() => onCopy(view.copy())}
      />
    </>
  );
}

/**
 * A button that runs a command, shown by its icon, and its label if
 * `labeled`. One that `keepsEntry` works on the entry being typed, which
 * a click on it leaves the keyboard.
 */
function Command({
  icon: Icon,
  label,
  tooltip,
  disabled,
  onRun,
  labeled = false,
  keepsEntry = false,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  disabled: boolean;
  onRun: () => void;
  labeled?: boolean;
  keepsEntry?: boolean;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-label={labeled ? undefined : label}
        disabled={disabled}
        onMouseDown={keepsEntry ? (event) => event.preventDefault() : undefined}
        onClick={onRun}
      >
        <Icon size={ICON_SIZE} aria-hidden />
        {labeled && label}
      </button>
    </Tooltip>
  );
}

/** A button for one of the alternatives of a setting. */
function Choice({
  icon: Icon,
  label,
  tooltip,
  chosen,
  onChoose,
}: {
  icon?: LucideIcon;
  label: string;
  tooltip: string;
  chosen: boolean;
  onChoose: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-pressed={chosen}
        onClick={onChoose}
      >
        {Icon && <Icon size={ICON_SIZE} aria-hidden />}
        {label}
      </button>
    </Tooltip>
  );
}
