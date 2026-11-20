import * as Dialog from '@radix-ui/react-dialog';
import { useId, useState } from 'react';
import type { Workbook } from '../../engine/workbook';
import { focusGrid, measureText } from '../grid/Grid';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Key } from './i18n';
import {
  DEFAULT_PRINT,
  MARGINS,
  PAPERS,
  printDocument,
  printedRanges,
  type PrintSettings,
} from './print';
import styles from './Shell.module.css';

/**
 * The print panel: what to print and how the pages look, kept until the
 * page reloads; Print gives the document to `onPrint` (R36).
 */
export function PrintPanel({
  workbook,
  view,
  open,
  onOpenChange,
  onPrint,
}: {
  workbook: Workbook;
  view: View;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onPrint: (document: string) => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [settings, setSettings] = useState(DEFAULT_PRINT);
  const change = (changed: Partial<PrintSettings>) =>
    setSettings({ ...settings, ...changed });
  const { sheet, selection } = view.state;
  const ranges = open
    ? printedRanges(workbook, sheet, selection.ranges, settings.what)
    : [];

  const choice = <K extends 'what' | 'orientation' | 'paper' | 'margins'>(
    name: K,
    label: Key,
    values: readonly PrintSettings[K][],
  ) => (
    <>
      <label className={styles.field} htmlFor={`${id}-${name}`}>
        {t(label)}
      </label>
      <select
        id={`${id}-${name}`}
        className={`${styles.textInput} ${styles.ruleSelect}`}
        value={settings[name]}
        onChange={(event) =>
          change({ [name]: event.target.value } as Partial<PrintSettings>)
        }
      >
        {values.map((value) => (
          <option key={value} value={value}>
            {t(`print.${value}` as Key)}
          </option>
        ))}
      </select>
    </>
  );
  const check = (name: 'fitWidth' | 'gridlines' | 'headers') => (
    <label className={styles.checkField}>
      <input
        type="checkbox"
        className={styles.checkbox}
        checked={settings[name]}
        onChange={(event) => change({ [name]: event.target.checked })}
      />
      {t(`print.${name}`)}
    </label>
  );

  return (
    <Dialog.Root open={open} onOpenChange={onOpenChange}>
      <Dialog.Portal>
        <Dialog.Overlay className={styles.overlay} />
        <Dialog.Content
          className={styles.dialog}
          onCloseAutoFocus={(event) => {
            // Opened by a shortcut as well as by the ribbon.
            event.preventDefault();
            focusGrid();
          }}
        >
          <Dialog.Title className={styles.dialogTitle}>
            {t('print')}
          </Dialog.Title>
          <Dialog.Description className={styles.dialogText}>
            {t('print.description')}
          </Dialog.Description>
          <form
            onSubmit={(event) => {
              event.preventDefault();
              onPrint(
                printDocument(workbook, sheet, ranges, settings, measureText),
              );
              onOpenChange(false);
            }}
          >
            {choice('what', 'print.what', ['sheet', 'selection'])}
            {choice('orientation', 'print.orientation', [
              'portrait',
              'landscape',
            ])}
            {choice('paper', 'print.paper', PAPERS)}
            {choice('margins', 'print.margins', MARGINS)}
            <div className={styles.printChecks}>
              {check('fitWidth')}
              {check('gridlines')}
              {check('headers')}
            </div>
            {ranges.length === 0 && (
              <p className={styles.fieldError}>{t('print.empty')}</p>
            )}
            <div className={styles.dialogButtons}>
              <button
                type="button"
                className={styles.button}
                onClick={() => onOpenChange(false)}
              >
                {t('cancel')}
              </button>
              <Tooltip text={t('print.print.tooltip')}>
                <button
                  type="submit"
                  className={`${styles.button} ${styles.primary}`}
                  disabled={ranges.length === 0}
                >
                  {t('print')}
                </button>
              </Tooltip>
            </div>
          </form>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}
