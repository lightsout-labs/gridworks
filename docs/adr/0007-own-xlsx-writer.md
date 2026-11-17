# 0007: Gridworks writes `.xlsx` files with its own writer, and ExcelJS is removed

## Context

ADR 0006 replaced ExcelJS for reading and kept it for writing. Writing
the features of R30 (S28) with ExcelJS 4.4 needs more than a mapping:

- It has no `duplicateValues` conditional format, which Gridworks has.
- Its styles always start with Calibri 11 and its own dxf and font
  tables, while the reader must recognize the file's default font to
  leave it to Gridworks' own.
- Its notes, internal hyperlinks and Excel 2010 data bars each need
  checking against what Excel expects, in a library released rarely.

Meanwhile the reader already holds the knowledge of every part and element
these features use, and JSZip, loaded with it, also writes zip files.

Options considered:

- **Keep ExcelJS** and post-process its output for what it lacks. Two
  models of the file to keep in step, and about 1 MB of library loaded on
  the first save.
- **Our own writer** over JSZip, producing the same parts and elements the
  reader reads.

## Decision

`writeXlsx` is our own writer over JSZip, beside `readXlsx`, in
`engine/io`. ExcelJS is removed from the dependencies. ADR 0003 no
longer applies.

## Consequences

- One module owns the file format in both directions, and the round trip
  of every feature is tested through `readXlsx` and `writeXlsx` alone.
- Saving loads only JSZip, which opening already loads.
- What Excel needs beyond what the reader reads (content types, the VML
  drawing of notes, the extension of data bars) is ours to keep right;
  a unit test checks that every part has a content type and every
  relationship a target, but only Excel itself can confirm a file opens.
- A feature added to files later needs writing code of our own too.
