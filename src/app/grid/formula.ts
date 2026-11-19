import {
  argumentHint,
  functions,
  print,
  references as areasOf,
  tokenize,
  type Area,
  type Corner,
  type Token,
} from '../../engine/formula';

/** The number of colors references take in turn. */
export const REFERENCE_COLORS = 6;

/** A reference in the formula being typed, and where it is in the text. */
export interface Highlight {
  readonly start: number;
  readonly end: number;
  /** The same for the same reference, from 0 up to `REFERENCE_COLORS`. */
  readonly color: number;
  /** The sheet named, or none for the formula's own sheet. */
  readonly sheet?: string;
  readonly area: Area;
}

/** What helps while a formula is typed, with the caret at some place. */
export interface FormulaHelp {
  readonly references: readonly Highlight[];
  /**
   * The values Excel offers for the argument at the caret that start with
   * the text typed in it, then the defined names and the functions whose
   * names start with the name typed at the caret: every one, in an
   * argument that offers values, before anything is typed.
   */
  readonly completion: {
    readonly values: readonly string[];
    readonly definedNames: readonly string[];
    readonly names: readonly string[];
  } | null;
  /** The arguments of the function call that holds the caret. */
  readonly hint: {
    readonly name: string;
    readonly arguments: readonly string[];
    /** Which argument the caret is in, or -1 past the last one. */
    readonly current: number;
    /** The values Excel offers for that argument. */
    readonly values: readonly string[];
  } | null;
}

/** A new text of the entry and where the caret goes. */
export interface Edit {
  readonly text: string;
  readonly caret: number;
}

const NOTHING: FormulaHelp = { references: [], completion: null, hint: null };

/**
 * A name that may be the start of a function's or a defined name's, as
 * `SU` or `LOG1`.
 */
const NAME_START = /^[A-Za-z_\u00A1-\uFFFF][\w.\u00A1-\uFFFF]*$/;

/**
 * The help for an entry with the caret at `caret`: none unless it is a
 * formula, which starts with `=`. `definedNames` are the workbook's names.
 */
export function formulaHelp(
  text: string,
  caret: number,
  definedNames: readonly string[] = [],
): FormulaHelp {
  const tokens = formulaTokens(text);
  if (!tokens) return NOTHING;
  return {
    references: highlights(tokens),
    completion: completion(text, tokens, caret, definedNames),
    hint: hint(tokens, caret),
  };
}

/**
 * Whether a reference picked on the grid can go at the caret: where an
 * operand starts, or in place of a reference that ends there.
 */
export function pickable(text: string, caret: number): boolean {
  const tokens = formulaTokens(text);
  if (!tokens) return false;
  const before = tokenBefore(tokens, caret);
  if (!before) return caret > 0;
  if (before.end < caret && text.slice(before.end, caret).trim() !== '') {
    return false; // The caret is inside a token.
  }
  if (before.kind === 'reference') return before.end === caret;
  return startsOperand(text, caret);
}

/**
 * Whether an operand starts at the caret of a formula: after `=`, an
 * operator, `(` or `,`, where a new reference can go.
 */
export function startsOperand(text: string, caret: number): boolean {
  const tokens = formulaTokens(text);
  if (!tokens) return false;
  const before = tokenBefore(tokens, caret);
  if (!before) return caret > 0;
  if (before.end < caret && text.slice(before.end, caret).trim() !== '') {
    return false; // The caret is inside a token.
  }
  return (
    before.kind === 'open' ||
    before.kind === 'comma' ||
    before.kind === 'operator'
  );
}

/**
 * Puts a reference at the caret, in place of a reference that ends there,
 * as a drag on the grid replaces the one it started; with `add`, after
 * that reference and a comma, as Ctrl or Cmd with a click adds one.
 */
export function insertReference(
  text: string,
  caret: number,
  reference: string,
  add = false,
): Edit {
  const before = tokenBefore(formulaTokens(text) ?? [], caret);
  if (before?.kind !== 'reference' || before.end !== caret) {
    return replace(text, caret, caret, reference);
  }
  return add
    ? replace(text, caret, caret, `,${reference}`)
    : replace(text, before.start, caret, reference);
}

/**
 * Cycles the reference at the caret through `$A$1`, `A$1`, `$A1` and `A1`,
 * as F4 does; `null` when the caret is not on a reference.
 */
export function cycleReference(text: string, caret: number): Edit | null {
  const token = formulaTokens(text)?.find(
    (each) =>
      each.kind === 'reference' && each.start <= caret && caret <= each.end,
  );
  if (token?.kind !== 'reference') return null;
  const { reference } = token;
  const [colAbsolute, rowAbsolute] = next(reference.start);
  const absolute = (corner: Corner): Corner => ({
    ...corner,
    rowAbsolute,
    colAbsolute,
  });
  const cycled = print({
    ...reference,
    start: absolute(reference.start),
    ...(reference.end && { end: absolute(reference.end) }),
  });
  return replace(text, token.start, token.end, cycled);
}

/** What a choice of the list is: each goes in its own way. */
export type Completion = 'function' | 'defined' | 'value';

/**
 * Puts a function and its opening parenthesis in place of the name typed,
 * a defined name without one, or an argument's value in place of the text
 * typed in the argument.
 */
export function complete(
  text: string,
  caret: number,
  name: string,
  kind: Completion = 'function',
): Edit {
  const tokens = formulaTokens(text) ?? [];
  const start =
    kind === 'value'
      ? (argumentAt(text, tokens, caret)?.start ?? caret)
      : (nameStart(text, tokens, caret) ?? caret);
  return replace(text, start, caret, kind === 'function' ? `${name}(` : name);
}

/** The tokens of a formula, placed in the whole text, or `null`. */
function formulaTokens(text: string): Token[] | null {
  if (!text.startsWith('=')) return null;
  return tokenize(text.slice(1)).map((token) => ({
    ...token,
    start: token.start + 1,
    end: token.end + 1,
  }));
}

/** The last token that starts before the caret. */
function tokenBefore(tokens: Token[], caret: number): Token | undefined {
  return tokens.findLast((token) => token.start < caret);
}

function highlights(tokens: Token[]): Highlight[] {
  const colors = new Map<string, number>();
  const found: Highlight[] = [];
  for (const token of tokens) {
    if (token.kind !== 'reference') continue;
    const key = print(token.reference).toUpperCase();
    if (!colors.has(key)) colors.set(key, colors.size % REFERENCE_COLORS);
    const [{ sheet, area }] = areasOf(token.reference) as [
      { sheet?: string; area: Area },
    ];
    found.push({
      start: token.start,
      end: token.end,
      color: colors.get(key)!,
      ...(sheet !== undefined && { sheet }),
      area,
    });
  }
  return found;
}

function completion(
  text: string,
  tokens: Token[],
  caret: number,
  definedNames: readonly string[],
): FormulaHelp['completion'] {
  const argument = argumentAt(text, tokens, caret);
  const offered = argument?.values ?? [];
  const typed = argument
    ? text.slice(argument.start, caret).toUpperCase()
    : null;
  // A value typed in full is not offered, so that Enter commits it.
  const values = offered.filter(
    (value) => value.startsWith(typed!) && value !== typed,
  );
  const start = nameStart(text, tokens, caret);
  const prefix =
    start !== null
      ? text.slice(start, caret).toUpperCase()
      : offered.length > 0 && typed === ''
        ? ''
        : null;
  // So is a name typed in full, or a function named as such a value.
  const defined =
    prefix === null
      ? []
      : definedNames.filter((name) => {
          const upper = name.toUpperCase();
          return upper.startsWith(prefix) && upper !== prefix;
        });
  const names =
    prefix === null
      ? []
      : functions.filter(
          (name) =>
            name.startsWith(prefix) &&
            !(name === typed && offered.includes(name)),
        );
  return values.length > 0 || defined.length > 0 || names.length > 0
    ? { values, definedNames: defined, names }
    : null;
}

/** Where the name being typed at the caret starts, or `null`. */
function nameStart(
  text: string,
  tokens: Token[],
  caret: number,
): number | null {
  const token = tokenBefore(tokens, caret);
  if (token?.end !== caret) return null;
  if (token.kind !== 'name' && token.kind !== 'reference') return null;
  const typed = text.slice(token.start, caret);
  // After a `$`, it is the column of a reference.
  if (!NAME_START.test(typed) || text[token.start - 1] === '$') return null;
  return token.start;
}

function hint(tokens: Token[], caret: number): FormulaHelp['hint'] {
  const call = callAt(tokens, caret);
  const found = call && argumentHint(call.name, call.index);
  return found ? { name: call.name.toUpperCase(), ...found } : null;
}

/**
 * The innermost function call whose parentheses hold the caret, which of
 * its arguments the caret is in, and where that argument starts.
 */
function callAt(
  tokens: Token[],
  caret: number,
): { name: string; index: number; start: number } | undefined {
  const calls: { name: string | null; index: number; start: number }[] = [];
  let previous: Token | undefined;
  for (const token of tokens) {
    if (token.end > caret) break;
    if (token.kind === 'open') {
      calls.push({
        name: previous?.kind === 'function' ? previous.name : null,
        index: 0,
        start: token.end,
      });
    } else if (token.kind === 'close') {
      calls.pop();
    } else if (token.kind === 'comma' && calls.length > 0) {
      calls[calls.length - 1]!.index++;
      calls[calls.length - 1]!.start = token.end;
    }
    previous = token;
  }
  return calls.findLast(
    (each): each is { name: string; index: number; start: number } =>
      each.name !== null,
  );
}

/**
 * The argument at the caret, where its text starts after any spaces, and
 * the values Excel offers for it; `undefined` outside a function call.
 */
function argumentAt(
  text: string,
  tokens: Token[],
  caret: number,
): { start: number; values: readonly string[] } | undefined {
  const call = callAt(tokens, caret);
  if (!call) return undefined;
  const spaces = /^\s*/.exec(text.slice(call.start, caret))![0].length;
  const values = argumentHint(call.name, call.index)?.values ?? [];
  return { start: call.start + spaces, values };
}

/** The next form of a reference, by its first corner: `[col, row]` absolute. */
function next({ rowAbsolute, colAbsolute }: Corner): [boolean, boolean] {
  if (!rowAbsolute && !colAbsolute) return [true, true];
  if (rowAbsolute && colAbsolute) return [false, true];
  if (rowAbsolute) return [true, false];
  return [false, false];
}

function replace(
  text: string,
  start: number,
  end: number,
  inserted: string,
): Edit {
  return {
    text: text.slice(0, start) + inserted + text.slice(end),
    caret: start + inserted.length,
  };
}
