import type { Dropped } from '../../engine/io';
import { translator, type Key, type Language } from './i18n';

/**
 * The name of the workbook in a file: the file's name without `.xlsx` or
 * `.csv`.
 */
export function workbookName(fileName: string): string {
  return fileName.replace(/\.(?:xlsx|csv)$/i, '');
}

/** Whether a file is a `.csv` file, by its name. */
export function isCsv(fileName: string): boolean {
  return /\.csv$/i.test(fileName);
}

/** What the user is told about the content left out of a file opened (R31). */
export function droppedMessage(
  fileName: string,
  dropped: readonly Dropped[],
  language: Language,
): string {
  const t = translator(language);
  const items = dropped.map(({ feature, count }) =>
    t(`dropped.${feature}.${count === 1 ? 'one' : 'other'}` as Key).replace(
      '{count}',
      count.toLocaleString('en-US'),
    ),
  );
  return t('file.dropped')
    .replace('{name}', () => fileName)
    .replace('{list}', () => items.join(language === 'zh' ? '、' : ', '));
}
