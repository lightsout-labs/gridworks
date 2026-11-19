import { describe, expect, it } from 'vitest';
import {
  complete,
  cycleReference,
  formulaHelp,
  insertReference,
  pickable,
} from './formula';

describe('function completion', () => {
  it('lists the functions that start with the name typed at the caret', () => {
    const { completion } = formulaHelp('=1+su', 5);
    expect(completion!.values).toEqual([]);
    expect(completion!.names).toContain('SUM');
    expect(completion!.names).toContain('SUMIFS');
    expect(completion!.names.every((name) => name.startsWith('SU'))).toBe(true);
    // A name that looks like a cell, as LOG1, still completes.
    expect(formulaHelp('=LOG1', 5).completion?.names).toEqual(['LOG10']);
  });

  it('offers nothing where no function name is being typed', () => {
    for (const [text, caret] of [
      ['=XYZ', 4],
      ['=SUM(', 5],
      ['=SUM', 2], // The caret is inside the name.
      ['="su', 4],
      ['=A1', 3],
      ['=$A', 3],
      ['su', 2],
    ] as const) {
      expect([text, formulaHelp(text, caret).completion]).toEqual([text, null]);
    }
  });

  it('replaces the name typed with the function chosen and its parenthesis', () => {
    expect(complete('=1+su*2', 5, 'SUM')).toEqual({
      text: '=1+SUM(*2',
      caret: 7,
    });
  });

  it('offers the defined names that start with the name typed, before the functions', () => {
    const names = ['Sales', 'Summary', 'Tax'];
    const { completion } = formulaHelp('=1+su', 5, names);
    expect(completion!.definedNames).toEqual(['Summary']);
    expect(formulaHelp('=SA', 3, names).completion).toMatchObject({
      definedNames: ['Sales'],
      names: [],
    });
    expect(formulaHelp('=ta', 3, names).completion!.definedNames).toEqual([
      'Tax',
    ]);
    // Nothing is offered for a name typed in full.
    expect(formulaHelp('=1+tax', 6, names).completion).toBeNull();
    expect(complete('=1+sa*2', 5, 'Sales', 'defined')).toEqual({
      text: '=1+Sales*2',
      caret: 8,
    });
  });
});

describe('argument value completion', () => {
  const lookup = '=VLOOKUP("Apple",A1:B10,2,';

  it('offers the values of the argument first, then every name', () => {
    const { completion } = formulaHelp(lookup, lookup.length, ['Tax']);
    expect(completion).toMatchObject({
      values: ['TRUE', 'FALSE'],
      definedNames: ['Tax'],
    });
    expect(completion!.names).toContain('SUM');
  });

  it('narrows the values and names to those that start with the text typed', () => {
    const f = formulaHelp(`${lookup}f`, lookup.length + 1).completion!;
    expect(f.values).toEqual(['FALSE']);
    expect(f.names).toContain('FALSE');
    expect(f.names.every((name) => name.startsWith('F'))).toBe(true);
    // A space before the value is not part of it.
    expect(
      formulaHelp(`${lookup} T`, lookup.length + 2).completion!.values,
    ).toEqual(['TRUE']);
    const match = '=MATCH(A1,B1:B9,-';
    expect(formulaHelp(match, match.length).completion).toEqual({
      values: ['-1'],
      definedNames: [],
      names: [],
    });
    // Text that starts no value offers none.
    expect(
      formulaHelp(`${lookup}A1+T`, lookup.length + 4).completion!.values,
    ).toEqual([]);
  });

  it('offers neither the value nor the function for a value typed in full', () => {
    const text = `${lookup}FALSE`;
    expect(formulaHelp(text, text.length).completion).toBeNull();
    // Elsewhere, the function is still offered.
    expect(formulaHelp('=FALSE', 6).completion!.names).toEqual(['FALSE']);
  });

  it('puts the value chosen in place of the text typed, without a parenthesis', () => {
    expect(
      complete(`${lookup} f)`, lookup.length + 2, 'FALSE', 'value'),
    ).toEqual({ text: `${lookup} FALSE)`, caret: lookup.length + 6 });
    expect(complete('=MATCH(A1,B1:B9,-', 17, '-1', 'value')).toEqual({
      text: '=MATCH(A1,B1:B9,-1',
      caret: 18,
    });
    // The function of the same name still takes its parenthesis.
    expect(complete('=FAL', 4, 'FALSE')).toEqual({ text: '=FALSE(', caret: 7 });
  });
});

describe('argument hint', () => {
  it('shows the innermost function whose parentheses hold the caret', () => {
    const hint = (text: string) => formulaHelp(text, text.length).hint;
    expect(hint('=SUM(A1,')).toEqual({
      name: 'SUM',
      arguments: ['number1', '[number2]', '...'],
      current: 1,
      values: [],
    });
    expect(hint('=IF(A1,SUM(1),')).toMatchObject({ name: 'IF', current: 2 });
    expect(hint('=IF(A1,SUM(1,')).toMatchObject({ name: 'SUM', current: 1 });
    expect(hint('=IF(A1,"a,b"')).toMatchObject({ name: 'IF', current: 1 });
    expect(hint('=round((1+2')).toMatchObject({ name: 'ROUND', current: 0 });
    expect(hint('=SUM(1)')).toBeNull();
    expect(hint('=FOO(1')).toBeNull();
    expect(hint('SUM(1')).toBeNull();
    expect(formulaHelp('=SUM(1,2)', 7).hint).toMatchObject({ current: 1 });
  });
});

describe('references', () => {
  it('gives each different reference its own color, and its area and sheet', () => {
    const { references } = formulaHelp("=A1+B2:C3+a1+'My sheet'!A1", 0);
    expect(references).toEqual([
      { start: 1, end: 3, color: 0, area: area(0, 0, 0, 0) },
      { start: 4, end: 9, color: 1, area: area(1, 1, 2, 2) },
      { start: 10, end: 12, color: 0, area: area(0, 0, 0, 0) },
      {
        start: 13,
        end: 26,
        color: 2,
        sheet: 'My sheet',
        area: area(0, 0, 0, 0),
      },
    ]);
    expect(formulaHelp('A1', 0).references).toEqual([]);
  });

  it('can take a reference where an operand starts, or replace the one just before', () => {
    expect(pickable('=', 1)).toBe(true);
    expect(pickable('=A1+ ', 5)).toBe(true);
    expect(pickable('=SUM(', 5)).toBe(true);
    expect(pickable('=SUM(A1,', 8)).toBe(true);
    expect(pickable('=A1:', 4)).toBe(true);
    expect(pickable('=A1', 3)).toBe(true);
    expect(pickable('=A1', 2)).toBe(false);
    expect(pickable('=5', 2)).toBe(false);
    expect(pickable('=SUM(1)', 7)).toBe(false);
    expect(pickable('=SUM', 4)).toBe(false);
    expect(pickable('5+', 2)).toBe(false);
  });

  it('inserts a reference at the caret, or replaces the one just before it', () => {
    expect(insertReference('=SUM()', 5, 'B2')).toEqual({
      text: '=SUM(B2)',
      caret: 7,
    });
    expect(insertReference('=SUM(B2)', 7, 'B2:C4')).toEqual({
      text: '=SUM(B2:C4)',
      caret: 10,
    });
    expect(insertReference('=1+Sheet2!A1', 12, 'Sheet2!B1')).toEqual({
      text: '=1+Sheet2!B1',
      caret: 12,
    });
  });

  it('cycles the reference at the caret through absolute and relative forms', () => {
    let edit = { text: '=A1+B2', caret: 2 };
    const seen: string[] = [];
    for (let step = 0; step < 4; step++) {
      edit = cycleReference(edit.text, edit.caret)!;
      seen.push(edit.text);
    }
    expect(seen).toEqual(['=$A$1+B2', '=A$1+B2', '=$A1+B2', '=A1+B2']);
    expect(edit.caret).toBe(3);
    expect(cycleReference('=SUM(A1:B2)', 10)).toEqual({
      text: '=SUM($A$1:$B$2)',
      caret: 14,
    });
    expect(cycleReference("='My sheet'!a1", 14)).toEqual({
      text: "='My sheet'!$A$1",
      caret: 16,
    });
    expect(cycleReference('=1+2', 2)).toBeNull();
    expect(cycleReference('A1', 2)).toBeNull();
  });
});

function area(top: number, left: number, bottom: number, right: number) {
  return { top, left, bottom, right };
}
