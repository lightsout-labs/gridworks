import * as Dialog from '@radix-ui/react-dialog';
import * as DropdownMenu from '@radix-ui/react-dropdown-menu';
import {
  ArrowDown,
  ArrowUp,
  Bold,
  ChevronDown,
  Italic,
  ListFilter,
  Pencil,
  Strikethrough,
  Trash2,
  Underline,
  type LucideIcon,
} from 'lucide-react';
import { useCallback, useId, useState, useSyncExternalStore } from 'react';
import type {
  Condition,
  Range,
  Rule,
  RuleFormat,
  Workbook,
} from '../../engine/workbook';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import { PALETTE } from './Format';
import type { Key } from './i18n';
import { parseRanges, rangesText } from './ranges';
import styles from './Shell.module.css';

const ICON_SIZE = 16;

type RuleType = Condition['type'];

/** The kinds of rule, in the order of Excel's Highlight Cells and Top/Bottom menus. */
const TYPES: readonly RuleType[] = [
  'greater',
  'less',
  'between',
  'equal',
  'contains',
  'duplicate',
  'top',
  'bottom',
];

/** Excel's default format: light red fill with dark red text. */
const DEEP_RED = '#9c0006';
const LIGHT_RED = '#ffc7ce';
const DEFAULT_FORMAT: RuleFormat = { color: DEEP_RED, fill: LIGHT_RED };
const TEXT_COLORS: readonly [string, Key][] = [
  [DEEP_RED, 'color.deepRed'],
  ...PALETTE,
];
const FILL_COLORS: readonly [string, Key][] = [
  [LIGHT_RED, 'color.lightRed'],
  ...PALETTE,
];

/** The font styles a rule can turn on, with their icons. */
const FONT_STYLES: readonly [
  'bold' | 'italic' | 'underline' | 'strike',
  Key,
  LucideIcon,
][] = [
  ['bold', 'bold', Bold],
  ['italic', 'italic', Italic],
  ['underline', 'underline', Underline],
  ['strike', 'strike', Strikethrough],
];

/** What the rules dialog shows: a new rule of a kind, or the sheet's rules. */
type Shown = { kind: 'new'; type: RuleType } | { kind: 'manage' };

/**
 * The Conditional formatting menu: a new rule of each kind for the
 * selected cells, and the rules of the sheet to edit, reorder and delete
 * (R20). Each change is one command, so it is undone on its own.
 */
export function RuleControls({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const [shown, setShown] = useState<Shown | null>(null);
  const close = () => setShown(null);
  return (
    <>
      <DropdownMenu.Root>
        <Tooltip text={t('rules.tooltip')}>
          <DropdownMenu.Trigger asChild>
            <button type="button" className={styles.button}>
              <ListFilter size={ICON_SIZE} aria-hidden />
              {t('rules')}
              <ChevronDown size={12} aria-hidden />
            </button>
          </DropdownMenu.Trigger>
        </Tooltip>
        <DropdownMenu.Portal>
          <DropdownMenu.Content
            className={styles.menu}
            align="start"
            sideOffset={4}
          >
            {TYPES.map((type) => (
              <DropdownMenu.Item
                key={type}
                className={styles.menuItem}
                onSelect={() => setShown({ kind: 'new', type })}
              >
                {t(`rules.new.${type}`)}
              </DropdownMenu.Item>
            ))}
            <DropdownMenu.Separator className={styles.menuSeparator} />
            <DropdownMenu.Item
              className={styles.menuItem}
              onSelect={() => setShown({ kind: 'manage' })}
            >
              {t('rules.manage')}
            </DropdownMenu.Item>
          </DropdownMenu.Content>
        </DropdownMenu.Portal>
      </DropdownMenu.Root>
      <Dialog.Root
        open={shown !== null}
        onOpenChange={(open) => !open && close()}
      >
        <Dialog.Portal>
          <Dialog.Overlay className={styles.overlay} />
          <Dialog.Content className={`${styles.dialog} ${styles.wideDialog}`}>
            {shown?.kind === 'new' && (
              <>
                <Dialog.Title className={styles.dialogTitle}>
                  {t('rule.title')}
                </Dialog.Title>
                <Dialog.Description className={styles.dialogText}>
                  {t('rule.description')}
                </Dialog.Description>
                <RuleForm
                  initial={newRule(shown.type, view.state.selection.ranges)}
                  onSave={(rule) => {
                    // A new rule comes first, as in Excel.
                    view.execute({
                      type: 'setRules',
                      sheet: view.state.sheet,
                      rules: [rule, ...workbook.rules(view.state.sheet)],
                    });
                    close();
                  }}
                  onCancel={close}
                />
              </>
            )}
            {shown?.kind === 'manage' && (
              <ManageRules workbook={workbook} view={view} />
            )}
          </Dialog.Content>
        </Dialog.Portal>
      </Dialog.Root>
    </>
  );
}

/** A rule of a kind for ranges, with Excel's default values and format. */
function newRule(type: RuleType, ranges: readonly Range[]): Rule {
  const condition: Condition =
    type === 'duplicate'
      ? { type }
      : type === 'top' || type === 'bottom'
        ? { type, count: 10 }
        : type === 'between'
          ? { type, value: '', other: '' }
          : { type, value: '' };
  return { ranges, condition, format: DEFAULT_FORMAT };
}

/**
 * The rules of the shown sheet, each with its format, condition and
 * cells, and buttons to edit, move and delete it; or the form of the rule
 * being edited or added.
 */
function ManageRules({ workbook, view }: { workbook: Workbook; view: View }) {
  const t = useTranslation();
  const subscribe = useCallback(
    (listener: () => void) => workbook.subscribe(listener),
    [workbook],
  );
  const { sheet } = view.state;
  const rules = useSyncExternalStore(subscribe, () => workbook.rules(sheet));
  // The rule being edited, by its place, or a new one at `null`.
  const [editing, setEditing] = useState<number | null | undefined>();
  const save = (next: readonly Rule[]) =>
    view.execute({ type: 'setRules', sheet, rules: next });
  const move = (index: number, by: number) => {
    const next = [...rules];
    const [rule] = next.splice(index, 1);
    next.splice(index + by, 0, rule!);
    save(next);
  };
  return (
    <>
      <Dialog.Title className={styles.dialogTitle}>
        {t('rules.title')}
      </Dialog.Title>
      {editing !== undefined ? (
        <>
          <Dialog.Description className={styles.dialogText}>
            {t(editing === null ? 'rule.description' : 'rules.editing')}
          </Dialog.Description>
          <RuleForm
            initial={
              editing === null
                ? newRule('greater', view.state.selection.ranges)
                : rules[editing]!
            }
            onSave={(rule) => {
              save(
                editing === null
                  ? [rule, ...rules]
                  : rules.map((each, index) =>
                      index === editing ? rule : each,
                    ),
              );
              setEditing(undefined);
            }}
            onCancel={() => setEditing(undefined)}
          />
        </>
      ) : (
        <>
          <Dialog.Description className={styles.dialogText}>
            {t('rules.description')}
          </Dialog.Description>
          {rules.length === 0 ? (
            <p className={styles.dialogText}>{t('rules.empty')}</p>
          ) : (
            <ul className={styles.ruleList}>
              {rules.map((rule, index) => (
                <li key={index} className={styles.ruleRow}>
                  <Sample format={rule.format} />
                  <span className={styles.ruleText}>
                    <span>{conditionText(rule.condition, t)}</span>
                    <span className={styles.ruleRanges}>
                      {rangesText(rule.ranges)}
                    </span>
                  </span>
                  <IconButton
                    icon={Pencil}
                    label={t('rules.edit')}
                    tooltip={t('rules.edit.tooltip')}
                    onClick={() => setEditing(index)}
                  />
                  <IconButton
                    icon={ArrowUp}
                    label={t('rules.up')}
                    tooltip={t('rules.up.tooltip')}
                    disabled={index === 0}
                    onClick={() => move(index, -1)}
                  />
                  <IconButton
                    icon={ArrowDown}
                    label={t('rules.down')}
                    tooltip={t('rules.down.tooltip')}
                    disabled={index === rules.length - 1}
                    onClick={() => move(index, 1)}
                  />
                  <IconButton
                    icon={Trash2}
                    label={t('rules.delete')}
                    tooltip={t('rules.delete.tooltip')}
                    onClick={() =>
                      save(rules.filter((_, each) => each !== index))
                    }
                  />
                </li>
              ))}
            </ul>
          )}
          <div className={styles.dialogButtons}>
            <Tooltip text={t('rules.add.tooltip')}>
              <button
                type="button"
                className={styles.button}
                onClick={() => setEditing(null)}
              >
                {t('rules.add')}
              </button>
            </Tooltip>
            <Dialog.Close asChild>
              <button
                type="button"
                className={`${styles.button} ${styles.primary}`}
              >
                {t('rules.close')}
              </button>
            </Dialog.Close>
          </div>
        </>
      )}
    </>
  );
}

/** A rule's condition in words, as the list of rules shows it. */
function conditionText(condition: Condition, t: (key: Key) => string): string {
  const text = t(`rules.text.${condition.type}`);
  return text.replace(/\{(\w+)\}/g, (_, name: string) =>
    String((condition as Record<string, unknown>)[name]),
  );
}

/** Sample text in a rule's format, over the cell's background. */
function Sample({ format }: { format: RuleFormat }) {
  const decoration = [
    format.underline && 'underline',
    format.strike && 'line-through',
  ]
    .filter(Boolean)
    .join(' ');
  return (
    <span
      className={styles.ruleSample}
      aria-hidden
      style={{
        color: format.color,
        background: format.fill,
        fontWeight: format.bold ? 'bold' : undefined,
        fontStyle: format.italic ? 'italic' : undefined,
        textDecoration: decoration || undefined,
      }}
    >
      AaBbCc
    </span>
  );
}

/** A button shown by its icon, named for assistive technology. */
function IconButton({
  icon: Icon,
  label,
  tooltip,
  disabled = false,
  pressed,
  onClick,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  disabled?: boolean;
  pressed?: boolean;
  onClick: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-label={label}
        aria-pressed={pressed}
        disabled={disabled}
        onClick={onClick}
      >
        <Icon size={ICON_SIZE} aria-hidden />
      </button>
    </Tooltip>
  );
}

/**
 * A form for a rule: its kind, its values, the cells it applies to and its
 * format. OK is enabled once every field can be read.
 */
function RuleForm({
  initial,
  onSave,
  onCancel,
}: {
  initial: Rule;
  onSave: (rule: Rule) => void;
  onCancel: () => void;
}) {
  const t = useTranslation();
  const id = useId();
  const start = initial.condition;
  const [type, setType] = useState<RuleType>(start.type);
  const [value, setValue] = useState('value' in start ? start.value : '');
  const [other, setOther] = useState('other' in start ? start.other : '');
  const [count, setCount] = useState(
    'count' in start ? String(start.count) : '10',
  );
  const [applies, setApplies] = useState(rangesText(initial.ranges));
  const [format, setFormat] = useState<RuleFormat>(initial.format);

  const ranges = parseRanges(applies);
  const number = Number(count);
  const condition: Condition | null =
    type === 'duplicate'
      ? { type }
      : type === 'top' || type === 'bottom'
        ? Number.isInteger(number) && number >= 1
          ? { type, count: number }
          : null
        : value === ''
          ? null
          : type === 'between'
            ? other === ''
              ? null
              : { type, value, other }
            : { type, value };
  const change = (property: keyof RuleFormat, setting: string | true | null) =>
    setFormat((old) => {
      const next: Record<string, unknown> = { ...old };
      if (setting === null) delete next[property];
      else next[property] = setting;
      return next as RuleFormat;
    });
  const field = (
    name: string,
    label: string,
    text: string,
    onText: (text: string) => void,
    invalid = false,
  ) => (
    <>
      <label className={styles.field} htmlFor={`${id}-${name}`}>
        {label}
      </label>
      <input
        id={`${id}-${name}`}
        className={styles.textInput}
        value={text}
        spellCheck={false}
        autoComplete="off"
        aria-invalid={invalid}
        aria-describedby={invalid ? `${id}-problem` : undefined}
        onChange={(event) => onText(event.target.value)}
      />
    </>
  );
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        if (!condition || !ranges) return;
        onSave({ ranges, condition, format });
      }}
    >
      <label className={styles.field} htmlFor={`${id}-type`}>
        {t('rule.type')}
      </label>
      <select
        id={`${id}-type`}
        className={`${styles.textInput} ${styles.ruleSelect}`}
        value={type}
        onChange={(event) => setType(event.target.value as RuleType)}
      >
        {TYPES.map((each) => (
          <option key={each} value={each}>
            {t(`rules.type.${each}`)}
          </option>
        ))}
      </select>
      {type === 'top' || type === 'bottom'
        ? field('count', t('rule.count'), count, setCount)
        : type !== 'duplicate' && (
            <>
              {field('value', t('rule.value'), value, setValue)}
              {type === 'between' &&
                field('other', t('rule.other'), other, setOther)}
            </>
          )}
      {field('applies', t('rule.appliesTo'), applies, setApplies, !ranges)}
      {!ranges && (
        <p id={`${id}-problem`} className={styles.problem}>
          {t('rule.appliesTo.invalid')}
        </p>
      )}
      <span className={styles.field}>{t('rule.format')}</span>
      <div className={styles.ruleFormat}>
        {FONT_STYLES.map(([property, label, icon]) => (
          <IconButton
            key={property}
            icon={icon}
            label={t(label)}
            tooltip={t(`${label}.tooltip` as Key)}
            pressed={format[property] === true}
            onClick={() => change(property, format[property] ? null : true)}
          />
        ))}
        <Sample format={format} />
      </div>
      <Swatches
        label={t('color.text')}
        none={t('color.automatic')}
        colors={TEXT_COLORS}
        chosen={format.color}
        onChoose={(color) => change('color', color)}
      />
      <Swatches
        label={t('color.fill')}
        none={t('color.noFill')}
        colors={FILL_COLORS}
        chosen={format.fill}
        onChoose={(color) => change('fill', color)}
      />
      <div className={styles.dialogButtons}>
        <button type="button" className={styles.button} onClick={onCancel}>
          {t('cancel')}
        </button>
        <button
          type="submit"
          className={`${styles.button} ${styles.primary}`}
          disabled={!condition || !ranges}
        >
          {t('ok')}
        </button>
      </div>
    </form>
  );
}

/** A row of colors to choose one from, or none. */
function Swatches({
  label,
  none,
  colors,
  chosen,
  onChoose,
}: {
  label: string;
  none: string;
  colors: readonly [string, Key][];
  chosen: string | undefined;
  onChoose: (color: string | null) => void;
}) {
  const t = useTranslation();
  return (
    <div className={styles.ruleColors} role="group" aria-label={label}>
      <span className={styles.ruleColorLabel}>{label}</span>
      <Tooltip text={none}>
        <button
          type="button"
          className={styles.button}
          aria-label={`${label}: ${none}`}
          aria-pressed={chosen === undefined}
          onClick={() => onChoose(null)}
        >
          {none}
        </button>
      </Tooltip>
      {colors.map(([color, name]) => (
        <Tooltip key={color} text={t(name)}>
          <button
            type="button"
            className={styles.swatch}
            style={{ background: color }}
            aria-label={`${label}: ${t(name)}`}
            aria-pressed={chosen === color}
            data-chosen={chosen === color || undefined}
            onClick={() => onChoose(color)}
          />
        </Tooltip>
      ))}
    </div>
  );
}
