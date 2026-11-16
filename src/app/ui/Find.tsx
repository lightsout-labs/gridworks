import * as Dialog from '@radix-ui/react-dialog';
import { useEffect, useId, useRef, useState } from 'react';
import type { Finding, Found, View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import styles from './Shell.module.css';

/** Which panel is open: find alone, or find and replace. */
export type FindMode = 'find' | 'replace';

/**
 * The find panel, which goes through the cells that match and counts them,
 * and in `replace` mode replaces one match or all of them (R25). It stays
 * open beside the grid, which can still be clicked, until it is closed.
 * What it looks for is kept until the next time it opens.
 */
export function FindPanel({
  view,
  mode,
  onMode,
}: {
  view: View;
  mode: FindMode | null;
  onMode: (mode: FindMode | null) => void;
}) {
  const t = useTranslation();
  const id = useId();
  const [finding, setFinding] = useState<Finding>({ text: '' });
  const [replacement, setReplacement] = useState('');
  const [message, setMessage] = useState('');
  const input = useRef<HTMLInputElement>(null);
  // Where the keyboard was when the panel opened, usually the grid, to give
  // it back on closing: the panel has no trigger for Radix to return to.
  const opener = useRef<Element | null>(null);
  const replacing = mode === 'replace';

  // Switching between find and replace starts from the search again.
  useEffect(() => {
    input.current?.select();
  }, [mode]);

  const change = (next: Partial<Finding>) => {
    setFinding({ ...finding, ...next });
    setMessage('');
  };
  const report = (found: Found | null) =>
    setMessage(
      found
        ? t('find.count')
            .replace('{index}', String(found.index))
            .replace('{count}', String(found.count))
        : t('find.none'),
    );
  const disabled = finding.text === '';

  return (
    <Dialog.Root
      modal={false}
      open={mode !== null}
      onOpenChange={(open) => {
        if (!open) onMode(null);
      }}
    >
      <Dialog.Portal>
        <Dialog.Content
          className={styles.findPanel}
          aria-describedby={undefined}
          onOpenAutoFocus={() => {
            opener.current = document.activeElement;
          }}
          onCloseAutoFocus={(event) => {
            const element = opener.current;
            if (element instanceof HTMLElement && element.isConnected) {
              event.preventDefault();
              element.focus();
            }
          }}
          // The grid stays usable while the panel is open.
          onInteractOutside={(event) => event.preventDefault()}
        >
          <Dialog.Title className={styles.dialogTitle}>
            {t(replacing ? 'replace' : 'find')}
          </Dialog.Title>
          <form
            onSubmit={(event) => {
              event.preventDefault();
              if (!disabled) report(view.findNext(finding));
            }}
          >
            <label className={styles.field} htmlFor={`${id}-what`}>
              {t('find.what')}
            </label>
            <input
              ref={input}
              id={`${id}-what`}
              type="text"
              className={styles.textInput}
              value={finding.text}
              onChange={(event) => change({ text: event.target.value })}
            />
            {replacing && (
              <>
                <label className={styles.field} htmlFor={`${id}-with`}>
                  {t('replace.with')}
                </label>
                <input
                  id={`${id}-with`}
                  type="text"
                  className={styles.textInput}
                  value={replacement}
                  onChange={(event) => setReplacement(event.target.value)}
                />
              </>
            )}
            <div className={styles.findOptions}>
              <label className={styles.field} htmlFor={`${id}-within`}>
                {t('find.within')}
              </label>
              <select
                id={`${id}-within`}
                className={styles.textInput}
                value={finding.workbook ? 'workbook' : 'sheet'}
                onChange={(event) =>
                  change({ workbook: event.target.value === 'workbook' })
                }
              >
                <option value="sheet">{t('find.sheet')}</option>
                <option value="workbook">{t('find.workbook')}</option>
              </select>
              <label className={styles.field} htmlFor={`${id}-in`}>
                {t('find.lookIn')}
              </label>
              {/* Replacing changes what edits a cell, so it looks there. */}
              <select
                id={`${id}-in`}
                className={styles.textInput}
                value={replacing ? 'formulas' : (finding.lookIn ?? 'formulas')}
                disabled={replacing}
                onChange={(event) =>
                  change({
                    lookIn: event.target.value as 'formulas' | 'values',
                  })
                }
              >
                <option value="formulas">{t('find.formulas')}</option>
                <option value="values">{t('find.values')}</option>
              </select>
            </div>
            <label className={styles.checkField}>
              <input
                type="checkbox"
                className={styles.checkbox}
                checked={finding.matchCase ?? false}
                onChange={(event) =>
                  change({ matchCase: event.target.checked })
                }
              />
              {t('find.matchCase')}
            </label>
            <label className={styles.checkField}>
              <input
                type="checkbox"
                className={styles.checkbox}
                checked={finding.entireCell ?? false}
                onChange={(event) =>
                  change({ entireCell: event.target.checked })
                }
              />
              {t('find.entireCell')}
            </label>
            <p className={styles.dialogText} aria-live="polite">
              {message}
            </p>
            <div className={styles.dialogButtons}>
              {replacing && (
                <>
                  <Tooltip text={t('replace.all.tooltip')}>
                    <button
                      type="button"
                      className={styles.button}
                      disabled={disabled}
                      onClick={() => {
                        const count = view.replaceAll(finding, replacement);
                        setMessage(
                          count === 0
                            ? t('find.none')
                            : t('replace.count').replace(
                                '{count}',
                                String(count),
                              ),
                        );
                      }}
                    >
                      {t('replace.all')}
                    </button>
                  </Tooltip>
                  <Tooltip text={t('replace.one.tooltip')}>
                    <button
                      type="button"
                      className={styles.button}
                      disabled={disabled}
                      onClick={() => report(view.replace(finding, replacement))}
                    >
                      {t('replace')}
                    </button>
                  </Tooltip>
                </>
              )}
              <Tooltip text={t('find.next.tooltip')}>
                <button
                  type="submit"
                  className={`${styles.button} ${styles.primary}`}
                  disabled={disabled}
                >
                  {t('find.next')}
                </button>
              </Tooltip>
              <Tooltip text={t('find.close.tooltip')}>
                <button
                  type="button"
                  className={styles.button}
                  onClick={() => onMode(null)}
                >
                  {t('find.close')}
                </button>
              </Tooltip>
            </div>
          </form>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}
