import { describe, expect, it } from 'vitest';
import { readXlsx } from '../../src/engine/io';
import { Workbook } from '../../src/engine/workbook';
import { xlsxFile } from '../xlsx';

describe('the engine', () => {
  it('recalculates a chain of 100,000 formulas within a second (R46)', () => {
    const workbook = Workbook.fromSnapshot({
      version: 1,
      name: 'Chain',
      sheets: [
        {
          id: 's',
          name: 'S',
          columns: [[1]],
          formulas: Array.from(
            { length: 100_000 },
            (_, row): [number, number, string] => [
              row,
              1,
              row === 0 ? '=A1+1' : `=B${row}+1`,
            ],
          ),
        },
      ],
    });
    const start = performance.now();
    workbook.execute({
      type: 'setCell',
      sheet: 's',
      row: 0,
      col: 0,
      input: '2',
    });
    expect(workbook.cell('s', 99_999, 1).value).toBe(100_002);
    expect(performance.now() - start).toBeLessThan(1000);
  });

  it('opens a 10 MB file within 5 seconds (R47)', async () => {
    const bytes = await largeFile();
    expect(bytes.byteLength).toBeGreaterThan(10_000_000);
    const start = performance.now();
    const { snapshot } = await readXlsx(bytes, 'Budget');
    const workbook = Workbook.fromSnapshot(snapshot);
    expect(workbook.cell(snapshot.sheets[0]!.id, 1, 10).value).toBe(
      'text 2-02.5',
    );
    expect(performance.now() - start).toBeLessThan(5000);
  }, 30_000);
});

/**
 * A file of more than 10 MB, as Excel would write a large table: 150,000
 * rows of ten numbers and texts and a formula.
 */
async function largeFile(): Promise<Uint8Array> {
  const strings: string[] = [];
  const rows: string[] = [];
  for (let r = 1; r <= 150_000; r++) {
    const cells: string[] = [];
    for (let c = 0; c < 10; c++) {
      const ref = String.fromCharCode(65 + c) + r;
      if (c % 3 === 0) {
        cells.push(`<c r="${ref}" t="s"><v>${strings.length}</v></c>`);
        strings.push(`<t>text ${r}-${c}</t>`);
      } else {
        cells.push(`<c r="${ref}"><v>${r * c + 0.5}</v></c>`);
      }
    }
    cells.push(
      `<c r="K${r}" t="str"><f>A${r}&amp;B${r}</f><v>text ${r}-0${r + 0.5}</v></c>`,
    );
    rows.push(`<row r="${r}">${cells.join('')}</row>`);
  }
  return xlsxFile({
    strings,
    sheets: [{ name: 'Data', rows: rows.join('') }],
  });
}
