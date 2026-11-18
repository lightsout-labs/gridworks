# engine/io

The decisions inside the module that reads and writes files, each with its
reason. Its interface, and the decisions that span modules, are in the
[architecture document](../../architecture.md).

- **A file opened is read as far as Gridworks supports it.** A formula
  that does not parse, or calls a function or name that Gridworks does not
  have, is replaced by its last result; a defined name the file keeps is
  calculated. Charts, pivot tables, pictures and Excel tables (whose
  cells stay) are counted from the file's parts and left out, and so are
  cells beyond the 100,000 rows and 100 columns of a sheet, and
  conditional formats, rules of data validation, defined names and
  hyperlinks of kinds Gridworks does not have; `readXlsx` reports what was
  left out (R31), and a file that is not a valid `.xlsx` file leaves the
  workbook shown as it was. Lesser
  differences are made without a word: what Gridworks does not draw of a
  cell (rich text runs, rotation, indents, pattern and gradient fills,
  diagonal borders), filters, which keep their hidden rows, threaded
  comments, read as the notes Excel keeps beside them, and sheet views
  other than frozen panes. A file saved by Gridworks has no formula
  results in it, and asks to be calculated when it is opened.
- **Files map Gridworks' formatting onto Excel's, and back** (R30). A
  file's default font, Calibri 11 in Excel, stands for Gridworks' default
  font and size, so a font keeps its name only when it is one of `FONTS`
  and its size only when it differs from the default's; Gridworks writes
  Calibri 11 as the default, so an explicit size of 11 reads back as the
  default. Theme and indexed colors are read as the colors they show, a
  theme's tint applied in HSL, which is within one step of a channel of
  Excel's swatches; automatic colors, and the theme's text color, are
  left to the theme's text color, and the Hyperlink style is not kept, as
  the grid draws links itself. Excel's other border lines are drawn as
  the nearest of Gridworks' six, and its other alignments as the nearest
  of left, center and right or top and middle. Excel's row and column
  styles are not kept; Excel writes a style on each cell in use as well,
  so only empty cells of a formatted row or column lose it. Sizes keep
  their proportions to the standard ones: Excel's 64 by 20 pixels are
  Gridworks' 96 by 24, a width being Excel's characters of 7 pixels and a
  height its points, and only heights set on rows are written or read, as
  fitted heights are found again. A condition's value is written as a
  number, a boolean or a quoted text, so a date typed in it reads back as
  its serial number. Notes are written with the drawing Excel needs to
  show them, without an author, and an author's name that starts a note
  read is taken out; data bars are written with Excel 2010's extension,
  so that they start from zero as Gridworks draws them.
- **A `.csv` file is UTF-8 text** (R32). The app decodes a file opened,
  dropping a byte order mark, and `readCsv` reads it as one sheet named
  after the file, or `Sheet1` when the name cannot name a sheet. Its fields
  are read as RFC 4180 and Excel have them, with lines ended by CRLF, LF
  or CR, and each is recognized as if it were typed, taking the format it
  implies; a field starting with `=` stays text, so that a file runs no
  formula, and an error such as `#N/A` stays text, as in `.xlsx` files.
  Exporting writes the cells in use with a byte order mark, so that Excel
  reads it as UTF-8, and CRLF line ends, quoting a field that holds a
  comma, a quote or a line break. A number is written in full whatever
  its format, and one in a date or time format as `yyyy-mm-dd hh:mm:ss`,
  or its date or time alone, unless it has a part of a second, which
  would not read back, so that reopening gives the same values. Text is
  written as it is, as Excel writes it, so a text that reads as a number,
  a date or a boolean, such as `007`, comes back as one, an error as its
  text, and a text starting with `=` is not escaped and is a formula to a
  program that runs them.
