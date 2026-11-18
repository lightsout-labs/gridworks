# engine/workbook

The decisions inside the workbook module, each with its reason. Its interface,
and the decisions that span modules, are in the
[architecture document](../../architecture.md).

## Storage and recalculation

- **Cells are stored by column**, as arrays of values and of style
  identifiers, with formulas in a sparse map. Ten million populated cells
  (R45) do not fit in memory as one object each.
- **Recalculation follows a dependency graph.** Each sheet keeps, for
  each cell, the formulas that read it alone, and for each column, the
  formulas that read a range over it (a whole row is in every column). A
  change marks the formulas that read the changed cells, and theirs in
  turn, as needing calculation; a formula is calculated when it is read,
  after the ones it reads, with an explicit stack
  ([ADR 0004](../../adr/0004-main-thread-lazy-recalculation.md)). Nothing
  else is calculated, and `calculations` counts what is, so tests check it
  (R14). Cells that nothing reads stay uncalculated until read.
- **Circular references give 0 and are reported**, as in Excel without
  iteration. The workbook calculates every formula a change sets, and the
  formulas that were in a circular reference, before it tells its
  listeners, so `circularReference()` knows every one (R15). A stored workbook calculates all its
  formulas when it opens.
- **References to other sheets are by name**, as written. Renaming a
  sheet rewrites the formulas that name it; a reference to a sheet that
  no longer exists gives `#REF!` but keeps its text, so undoing the
  deletion brings it back (Excel writes `#REF!` into the text). Any change
  to the sheets other than a color or the order finds again what every
  formula reads and recalculates every formula. A rename also rewrites
  references left from a deleted sheet of the same name, which an undo of
  the rename then does not put back.

## Commands and undo

- **Undo restores what a command changed.** Before it runs a command, the
  workbook records the contents and styles of the cells the command will
  change as a `restore` command, which puts them back. A change to the
  sheets (add or duplicate, delete, rename, recolor, move) records its own
  reverse: deleting a sheet records a `restoreSheet` command holding a
  snapshot of it, and the others their opposite. Undo runs that command and
  records its own reverse for redo. The workbook keeps the last 100
  changes, Excel's number, and a new change forgets what could be redone.
- **Typing sets a cell's format.** A typed date, time, percentage or
  number with separators sets the number format of its cell's style, and
  so does a formula that is a call to a function whose result Excel shows
  as a date or time (`DATE`, `TODAY`, `NOW`, `TIME`, `EDATE`, `EOMONTH`,
  `WORKDAY` and `WORKDAY.INTL`) in a cell with the General format; unlike
  Excel, arithmetic on such a call (`=TODAY()+7`) leaves the format
  alone. A plain number keeps the format the cell has, and Delete clears values
  and formulas but not formats, as in Excel. The formula bar and the
  editor start from `editText`, which reads back as the same value.
- **A paste works out its cells.** `paste` names the source range and the
  target, and the workbook works out the cells: a target made of whole
  copies of the source is filled with them, any other gets one copy at its
  top left. A copy moves the relative references of its formulas; a cut
  moves the cells, and every reference in the workbook to a range wholly
  inside them follows them, absolute or not, as in Excel. Listeners get,
  and undo keeps, a `group` of `restore` commands of different cells.
  Paste special pastes values (formula results, errors as their text),
  formulas and values, or formats, and keeps the rest of the target.
  Unlike Excel, a copy at the edge of the sheet keeps what fits and a cut
  moves back from the edge, where Excel refuses both; and references to
  the cells a cut lands on keep reading them, where Excel turns them into
  `#REF!`. A restore lists every cell it fills, so pasting one cell over a
  whole column is 100,000 cells in one recorded change; over a whole sheet
  it would be too many, which nothing prevents yet.
- **A fill extends series or repeats.** `fill` names the source range and
  the target, which extends it on one side, and the workbook fills each
  row or column of the source along the way the fill goes: as a series
  when its cells start one (numbers by their step, a least-squares line
  for more than two; dates by days, or by months when they fall on the
  same day of the month; English weekday and month names, in the case of
  the first; texts that differ only in the number they end in, by its
  step, keeping its digits, so `A001` gives `A002`), and otherwise by
  repeating them, formulas moved and formats kept, aligned to the source
  so that a fill upward ends on the source's last cell. One number is
  repeated, and one date, name or numbered text counts on by one, as in
  Excel. Its `how` asks for another way: `copy` repeats every row or
  column, `series` counts one number on too, and `formats` and `values`
  fill only the formats or only the contents and keep the rest of the
  target; `fillsSeries(fill)` tells whether a fill would extend any
  series. Like a paste, listeners and undo get a `group` of one `restore`.
- **A sort orders whole rows.** `sort` names a range, the columns to sort
  by, each ascending or descending, and whether its first row holds
  headers, which stay in place. The workbook orders the rows by the values
  of those columns, formula results included, as Excel does: numbers, then
  text in any case, then `FALSE` and `TRUE`, then errors, and empty cells
  last in either direction; rows equal on every column keep their order.
  Whole rows of the range move with their formats, and the relative
  references of their formulas move with them, as a copy moves them;
  references from elsewhere keep their cells, and conditional formats stay
  where they are, as in Excel. Like a paste, listeners and undo get a
  `group` of one `restore`. Ranges of whole rows or columns sort the cells
  in use only. A range with merged cells is refused, where Excel refuses
  only merges of different sizes. Under a filter, the rows it hides keep
  their places and the rows shown are sorted among themselves, as in
  Excel; rows hidden by hand outside the filter's range are sorted as
  before.
- **Inserting or deleting rows or columns rebuilds the sheets it
  changes.** `insertLines` and `deleteLines` name spans of rows or columns;
  an insertion puts as many empty lines before each span as it has, as
  Excel does before each selected range. The workbook makes the sheet anew
  from its snapshot with the cells, formulas and line sizes moved, and
  every sheet whose formulas refer to the lines with those formulas
  rewritten by the `lines` move of `shift`: a reference follows its cells,
  absolute or not, a range grows or shrinks with the lines inside it, and a
  reference wholly to deleted cells becomes `#REF!` in the text, as in
  Excel. It then finds again what every formula reads. Undo keeps a
  `restoreSheets` command holding the snapshots of the sheets as they
  were, as deleting a sheet does, so that the `#REF!` it wrote comes back
  as the references they were; what an insertion pushes off the sheet,
  which Excel refuses, comes back too. Listeners get the insertion or
  deletion itself.

## Layout

- **Row heights and column widths are CSS pixels at 100%**, kept in each
  sheet's `layout` with the lines hidden, and only for the lines that
  differ from the standard 24 by 96. `setLines` sets the size, or the
  standard one, and the visibility of spans, and records the lines as they
  were for undo. The layout is replaced, never changed, so the grid can
  build what it derives from it once per layout. `.xlsx` files keep sizes
  and hidden lines.
- **Frozen panes are a count of rows and columns** kept in the sheet's
  `layout` and its snapshot: `freeze` sets the rows above and columns
  left of a cell, the top row or the first column, and 0 and 0 unfreeze,
  recording what was frozen for undo. They are always the first lines of
  the sheet: unlike Excel, which freezes from the row and column scrolled
  to the top left, Gridworks freezes from A1, so a sheet scrolled down
  shows its first rows again when frozen, and freezing at a cell further
  down or right than the window reaches leaves nothing to scroll in view
  until it is unfrozen. Inserting or deleting lines keeps the counts.
- **Rows fit their text.** The workbook keeps, beside the sizes set on
  rows, the `fitted` height of each row whose wrapped text or font larger
  than the default needs more than the standard 24 pixels, and a size set
  on a row overrides it, as Excel's custom heights do. It fits the rows a
  change of cells touches, every row of a sheet when column widths change
  or the sheet is rebuilt, and none when no style can grow a row. Text is
  measured by a `TextWidth` function the app gives `fromSnapshot`, which
  measures with the canvas in the fonts the grid draws with, so the
  engine's wrapping (`wrapText`, also used by the grid) and the drawn lines
  agree; without one, as in Node, it guesses. Fitted heights are not
  stored: a sheet is fitted when it is opened. A formula whose result grows
  after a change elsewhere does not refit its row until its own row
  changes.
- **Merged cells are ranges kept by their sheet.** `merge` merges each
  range, grown over the merges and the other ranges it overlaps, keeps the
  value of its top-left cell and empties the others, as Excel does after
  its warning; `unmerge` removes the merges the ranges overlap. Undo is a
  `restore` that also carries the merges of its ranges as they were.
  Inserting and deleting lines moves a merge, grows it with lines inserted
  inside it, shrinks it with lines deleted, and drops it when one cell is
  left. Pasting formats, and so the format painter, gives the target the
  merges inside the cells copied in place of its own, as in Excel; other
  pastes and fills neither make nor remove merges. Merged cells are not
  fitted to their text, as in Excel.

## Formats

- **A format is a style change.** `setStyle` gives the cells of ranges
  properties of their style (font, size, bold, italic, underline,
  strikethrough, text and fill colors, horizontal and vertical alignment,
  wrap), `null` taking one back to the default, and keeps the rest; undo
  restores the cells like any other change. Styles are kept with their
  properties in one order, so that cells that look the same share one.
  Ranges of whole rows or columns change only the cells in use there, as
  formatting the lines themselves would need row and column styles, which
  Gridworks does not have: a cell typed later below a bold column is not
  bold.
- **A number format is a style property.** `numberFormat` holds an Excel
  format code, set by `setStyle` like any other property, so undo needs
  nothing of its own. A cell whose format is text keeps what is typed as
  text, a formula too, as in Excel; giving the format to a number does not
  turn it into text.
- **Borders are style properties.** A cell's style has a border for each
  side (`borderTop`, `borderRight`, `borderBottom`, `borderLeft`), each a
  line (thin, medium, thick, dashed, dotted or double) and a color, the
  theme's text color when it has none, as Excel's Automatic is black on
  white. `setBorders` puts one on a side of the selected ranges (their
  top, right, bottom or left edge), on their outside edges, or on every
  side of every cell, and `null` takes them off; like `setStyle`, ranges
  of whole lines change only the cells in use there, so the bottom border
  of a whole column goes on its last cell in use. Each cell keeps its
  own borders, as in Excel.
- **Conditional formats are rules kept by their sheet** (R20), in order of
  priority: `setRules` replaces the list, so adding, editing, reordering
  and deleting a rule are each one command whose reverse is the list as it
  was. A rule has the ranges it applies to, a condition (greater than,
  less than, between or equal to a value typed as in a cell, text that
  contains a string in any case, duplicate values, or the top or bottom N
  numbers, ties included) and a format of text and fill colors and bold,
  italic, underline and strikethrough. A color scale and a data bar are
  conditions too, with an empty format: a scale of two or three colors
  fills each number with the color between those of the lowest number,
  the median (Excel's 50th percentile) for three, and the highest, mixed
  in RGB; a data bar's length is the number's share of the way from zero,
  or the lowest number below it, to the highest, as Excel's automatic
  bounds. Text, empty cells and errors get no color and no bar, and
  Excel's other bounds (numbers, percents, formulas) are not offered.
  Values compare as Excel sorts them: numbers before text before booleans,
  text in any case; empty cells and errors match nothing, where Excel
  takes an empty cell as 0. Rules are evaluated when a cell is read, so
  they follow every change of a value or formula result; `cell` gives what
  they change as `conditional` (a scale's color as its fill) and the first
  data bar as `bar`, apart from the style, so the ribbon still shows a
  cell's own format. What a duplicate, top/bottom, scale or bar rule needs
  of all its values is worked out once and forgotten after every change.
  Where rules differ on a property, the first wins, and a new rule goes
  first, as in Excel. Inserting and deleting lines moves, grows and
  shrinks a rule's ranges and drops a rule whose cells are all deleted;
  pasting and cutting cells neither copy nor move rules, unlike Excel.

## Data features

- **A filter is kept by its sheet and hides rows in its layout** (R24).
  A sheet has one filter, as in Excel: a range whose first row holds the
  headers, and for some of its columns what the rows below must have to be
  shown: display texts checked in a list (`''` for empty cells), a text
  that the display text equals, contains, begins or ends with in any case,
  or a number that the value equals, is greater or less than, or lies
  between. Filters of several columns combine. The `filter` command
  replaces the sheet's filter, or removes it with `null`, and sets the
  `hidden` rows of the layout below the headers of the old and new ranges:
  hidden where the new filter does not show them, shown otherwise; rows
  hidden by hand there are shown, and rows elsewhere keep theirs. So the
  grid, hit testing, the row headers' gap, navigation and snapshots need
  nothing of their own, and rows stay as they were filtered until the
  filter changes, as in Excel. Undo is a `filter` command with the old
  filter and the rows that were hidden (`hidden`), which it sets in place
  of what a filter finds. Inserting and deleting lines moves and shrinks
  the range as a rule's, drops the filter of a column deleted, and drops
  the whole filter with its header row, leaving its rows as they are.
  `filterValues` lists a column's display texts in the rows the other
  columns' filters show, as values sort. A filter allows merged cells, and
  hides the rows of a merge like any other, as Excel does.
- **Find is a query, and a replace one command** (R25). `find` lists the
  cells of the sheets given whose text holds the search, sheet by sheet
  and row by row: the text that edits the cell (Excel's Look in Formulas,
  the default) or the text it shows (Values), in any case unless Match
  case, and the whole text for Match entire cell contents. Excel's
  wildcards (`*`, `?`, `~`) are not built: the search is plain text.
  `only` searches one cell, which the view uses to check the active cell
  without reading the sheet. `find` reads only the text it needs of each
  cell, as a whole sheet of R45's size is ten million cells. `replace`
  replaces every match in the editing text of each cell found, and sets
  the cell to the new text as if it were typed, so `1,000` becomes a
  number and a formula is parsed again; like a paste, listeners and undo
  get a `group` of a `restore` for each sheet, so replacing across the
  workbook is one undo. Replacing always looks in formulas, as in Excel,
  whose Replace offers no other choice.
- **Data validation is rules kept by their sheet** (R26), each the
  ranges it applies to and a criterion: a list, of items typed with commas
  between them or of the shown texts of a range written after `=` (on any
  sheet, read when asked for, empty cells left out; the reference is
  kept as typed, so unlike Excel it does not follow a renamed sheet or
  lines inserted or deleted there); or a number, a date or a text length
  between bounds, both included, either of which may be absent, which
  gives Excel's between, greater or equal and less or equal. The
  `validate` command gives the cells of ranges a criterion, taking them
  out of the rules they were in, or takes them out with `null`; undo is
  the same command with the sheet's rules as they were (`validations`),
  which it puts back. A cell's rule is the first that covers it. `allows`
  reads typed input as `setCell` would, a formula by its result evaluated
  in place without storing it, and compares a list in any case by the text
  the value would show, as Excel does; empty input and cells without a
  rule are always allowed (Excel's Ignore blank), and errors never are.
  Excel's whole-number and decimal kinds, its other operators, input
  messages, the Warning and Information styles, custom error messages and
  the formula kind are not built. A list's source may be a defined name
  (`=Colors`), read as its range. Inserting and deleting lines moves,
  grows and shrinks a rule's ranges as a conditional format's; pasting and
  cutting cells neither copy nor move rules, unlike Excel.
- **Defined names are a list kept by the workbook** (R27), each a name and
  the range it refers to, written as Excel writes it with its sheet and
  absolute (`Sheet1!$A$1:$B$2`), or `#REF!` once its cells are deleted.
  Names are workbook-wide: Excel's names scoped to one sheet, and names of
  constants or formulas, are not built. A name follows Excel's rules
  (`nameProblem`): letters, digits, `_` and `.`, starting with a letter or
  `_`, at most 255 characters, and not a reference in either style or a
  boolean, so `TAX1` is a cell and `R1C1` is refused; names differ in more
  than case. `setNames` replaces the list, kept in alphabetical order, as
  `setRules` replaces a sheet's rules, so defining, editing and deleting
  a name are each one command whose reverse is the list as it was; its
  `renamed` rewrites the formulas that use the old name, as Excel does;
  listeners get, and undo keeps, the formulas rewritten as `formulas`, so
  undo puts back only those, as they were typed, and a formula that
  already used the new name keeps it. A formula keeps a name as written;
  the evaluator asks the context what it refers to, and a name not
  defined gives `#NAME?`, so deleting a name, or a formula typed before
  its name, needs nothing of its own. The dependency graph enters the
  range of each name a formula uses, and any change to the names finds
  again what every formula reads, as a change to the sheets does.
  Inserting and deleting lines moves names by the same `lines` move as
  references, and their undo, `restoreSheets`, carries the names as they
  were; renaming a sheet renames it in names; a name on a deleted sheet
  gives `#REF!` until the deletion is undone. Pasting, filling and
  sorting leave names alone. `.xlsx` files keep names.
- **Notes and hyperlinks belong to cells** (R28, R29), kept by their
  sheet by the cell's key, apart from values and styles, and are part of
  what `contents` reads and a `restore` puts back, as a style is. So undo
  needs nothing of its own: `setNote` and `setLink`, each for one cell,
  change cells like `setCell`, and their reverse is a `restore`. A paste
  copies them and a cut and a sort move them with their cells, as in
  Excel; pasting only values, formulas or formats, and a fill, keep the
  target's own; typing, Delete and merging keep them, and rows and columns
  inserted and deleted move and drop them as cells. A note is plain text;
  Excel's author, threaded comments and a note's own size are not built.
  A link is a web address, given `https://` when typed without a scheme,
  or `#` and a reference in the workbook, as Excel's `HYPERLINK` writes
  one (`#'My data'!C3`; a defined name typed is kept as its range). Only
  `http:`, `https:` and `mailto:` addresses are taken, and only they are
  opened, so that no link runs script in the app. `setLink` may also set
  the cell's text as if typed. A reference in a link is kept as written,
  so it does not follow a renamed sheet or lines inserted, as in Excel.
  Excel gives a link the Hyperlink cell style; Gridworks stores none, and
  the grid draws links itself. `.xlsx` files keep notes and links.
