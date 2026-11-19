import { describe, expect, it } from 'vitest';
import { argumentHint, functions, parse, tokenize } from './index';

describe('formula tokens', () => {
  it('finds functions, references and separators, with their places', () => {
    const tokens = tokenize('SUM(A1:B2, Sheet2!$C3,"a,b")');
    expect(tokens.map(({ kind, start, end }) => [kind, start, end])).toEqual([
      ['function', 0, 3],
      ['open', 3, 4],
      ['reference', 4, 9],
      ['comma', 9, 10],
      ['reference', 11, 21],
      ['comma', 21, 22],
      ['string', 22, 27],
      ['close', 27, 28],
    ]);
    const reference = tokens[4]!;
    expect(reference.kind === 'reference' && reference.reference).toEqual({
      type: 'range',
      sheet: 'Sheet2',
      start: { row: 2, col: 2, rowAbsolute: false, colAbsolute: true },
    });
  });

  it('reads an unfinished formula as far as it goes', () => {
    const kinds = (text: string) => tokenize(text).map((token) => token.kind);
    expect(kinds("'My sheet'!A1:")).toEqual(['reference', 'operator']);
    expect(kinds('IF(A1>2,"op')).toEqual([
      'function',
      'open',
      'reference',
      'operator',
      'number',
      'comma',
      'string',
    ]);
    expect(kinds('su')).toEqual(['name']);
    expect(kinds('1+#N/A ?')).toEqual(['number', 'operator', 'error', 'other']);
  });
});

describe('argument hints', () => {
  it('names the arguments of a function and the one at an index', () => {
    expect(argumentHint('sum', 0)).toEqual({
      arguments: ['number1', '[number2]', '...'],
      current: 0,
      values: [],
    });
    // Repeated arguments are those after the required ones.
    expect(argumentHint('SUM', 7)?.current).toBe(1);
    expect(argumentHint('SUMIFS', 5)?.current).toBe(3);
    expect(argumentHint('SUMIFS', 6)?.current).toBe(4);
    // One too many is none of them.
    expect(argumentHint('IF', 3)).toEqual({
      arguments: ['logical_test', 'value_if_true', '[value_if_false]'],
      current: -1,
      values: [],
    });
    expect(argumentHint('PI', 0)).toMatchObject({ arguments: [], current: -1 });
    expect(argumentHint('NOPE', 0)).toBeNull();
  });

  it('agrees with the number of arguments each function takes', () => {
    // Whether a call with so many references parses, as the parser refuses
    // the wrong number of arguments.
    const takes = (name: string, count: number) =>
      parse(`${name}(${Array(count).fill('A1').join(',')})`) !== null;
    for (const name of functions) {
      const hint = argumentHint(name, 0)!;
      const names = hint.arguments.filter((each) => each !== '...');
      const optional = names.filter((each) => each.startsWith('['));
      const required = names.length - optional.length;
      const facts = [
        required === 0 || !takes(name, required - 1),
        takes(name, required),
        takes(name, names.length),
      ];
      if (hint.arguments.includes('...')) {
        // The optional arguments repeat.
        facts.push(
          optional.length > 0,
          takes(name, names.length + optional.length),
        );
      } else {
        facts.push(!takes(name, names.length + 1));
      }
      expect([name, facts.every(Boolean)]).toEqual([name, true]);
    }
  });
});

describe('argument values', () => {
  it('gives the values Excel offers for the current argument', () => {
    expect(argumentHint('vlookup', 3)?.values).toEqual(['TRUE', 'FALSE']);
    expect(argumentHint('MATCH', 2)?.values).toEqual(['1', '0', '-1']);
    expect(argumentHint('XLOOKUP', 5)?.values).toEqual(['1', '-1', '2', '-2']);
    expect(argumentHint('ADDRESS', 3)?.values).toEqual(['TRUE', 'FALSE']);
    // Other arguments, and arguments past the last, have none.
    expect(argumentHint('VLOOKUP', 2)?.values).toEqual([]);
    expect(argumentHint('VLOOKUP', 4)?.values).toEqual([]);
    expect(argumentHint('SUM', 0)?.values).toEqual([]);
  });

  it('offers values for every argument for which Excel offers them', () => {
    const offered = functions.flatMap((name) => {
      const { arguments: shown } = argumentHint(name, 0)!;
      return shown.flatMap((argument, index) =>
        argumentHint(name, index)!.values.length > 0
          ? [`${name} ${argument.replace(/[[\]]/g, '')}`]
          : [],
      );
    });
    expect(offered.sort()).toEqual(
      [
        'ADDRESS a1',
        'ADDRESS abs_num',
        'DAYS360 method',
        'HLOOKUP range_lookup',
        'MATCH match_type',
        'NETWORKDAYS.INTL weekend',
        'RANK order',
        'RANK.AVG order',
        'RANK.EQ order',
        'VLOOKUP range_lookup',
        'WEEKDAY return_type',
        'WEEKNUM return_type',
        'WORKDAY.INTL weekend',
        'XLOOKUP match_mode',
        'XLOOKUP search_mode',
        'XMATCH match_mode',
        'XMATCH search_mode',
        'YEARFRAC basis',
      ].sort(),
    );
  });
});
