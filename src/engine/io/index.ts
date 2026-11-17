import type { Snapshot } from '../workbook';

export { readXlsx } from './read';
export { writeXlsx } from './write';
export { readCsv, writeCsv } from './csv';

/** What a file holds that Gridworks leaves out when it opens it (R31). */
export type DroppedFeature =
  | 'chart'
  | 'pivotTable'
  | 'picture'
  /** An Excel table; its cells are kept. */
  | 'table'
  /** A formula Gridworks cannot calculate; its last result is kept. */
  | 'formula'
  /** A cell beyond the rows and columns of a sheet. */
  | 'cell'
  /** A conditional format of a kind Gridworks does not have. */
  | 'conditionalFormat'
  /** A rule of data validation of a kind Gridworks does not have. */
  | 'validation'
  /** A defined name of a sheet, hidden, or of a constant or a formula. */
  | 'name'
  /** A hyperlink that is neither a web or email address nor a place. */
  | 'link';

export interface Dropped {
  readonly feature: DroppedFeature;
  readonly count: number;
}

/** A workbook read from a file, and what was left out of it. */
export interface Opened {
  readonly snapshot: Snapshot;
  /** Each feature left out, with how many there were, in a fixed order. */
  readonly dropped: Dropped[];
}
