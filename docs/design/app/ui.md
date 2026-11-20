# app/ui

The decisions inside the module that holds everything around the grid, each
with its reason. Its interface, and the decisions that span modules, are in
the [architecture document](../../architecture.md).

## Shell

- **The address names the workbook shown** (R33, R34): `#` and its
  identifier, or nothing for the start screen, so that a reload, a link
  kept or the browser's Back button shows the same again; an identifier
  no longer stored gives the start screen. The start screen lists the
  stored workbooks, the one saved last first, with the time of that save
  (the time alone for today), and holds New workbook, which starts an
  empty workbook named `Book` and the first number no stored workbook has,
  Open file, and a menu on each workbook to rename, duplicate (as `Name
  (copy)`) and delete it, deleting after a confirmation. A rename or
  duplicate there loads the stored workbook, replays its changes and saves
  it under the name, so the platform needs no operation of its own for
  them. A workbook made new or opened from a file gets a new random
  identifier and is saved at once, so that the start screen lists it.
  The File tab's Start screen button goes back to it.
- **The workbook's name is a text box** beside the brand in the ribbon, as
  in Excel for the web; Enter or leaving it renames the workbook with
  `renameWorkbook`, an empty name keeps the one before, and the page's
  title is the name and ` - Gridworks`. The rename is recorded and saved
  like any change, but not undone, as the name is the file's rather than
  its contents. Renaming the workbook is no change to its cells, so the
  view keeps what was copied and the last fill.
- **Storage problems go to the message bar.** When storage refuses a
  change or a save, as when it is full, autosave tells the app, and the
  message bar under the ribbon asks the user to save the workbook to their
  computer; storage that cannot be read gives the start screen with a
  message. The same bar lists what a file opened left out, says that a
  file is not a valid `.xlsx` file, that a save failed, or that a new
  version of Gridworks is ready and a reload shows it.
- **The Open button takes `.xlsx` and `.csv` files**, told apart by their
  names; Export CSV saves the active sheet.
- **Window shortcuts are handled on the window**, not by the grid, so that
  Ctrl/Cmd+Z, Ctrl/Cmd+Y, Ctrl/Cmd+D, Ctrl/Cmd+R, Ctrl/Cmd+B, I and U, and
  Ctrl/Cmd+F and H work after a click on the ribbon too. During an
  entry they are left to the editor, which undoes its own typing.
- **Printing builds a plain HTML table** of the sheet or selection with
  print CSS for page size, orientation and margins, and rows that do not
  break across pages (R36). The canvas itself cannot be paginated.
  `print.ts` builds the document and has no DOM, so it is unit tested; the
  print panel (`Print.tsx`) only chooses its settings, which it keeps
  until the page reloads. The sheet prints its cells in use from A1, as
  `extent` gives them, and the selection each of its ranges within them,
  on a page of its own, as Excel prints the areas of a selection; with
  nothing in them, Print is disabled. Excel's defaults apply: portrait,
  Letter whatever the language (as number formats are), Normal margins
  (0.75 and 0.7 inch), no gridlines and no headings. Fit to width sets
  CSS `zoom` on a table wider than the page, from the paper's width less
  the margins, and never enlarges one. Column headings are a `thead`,
  which browsers repeat on each page. Cells are drawn as the grid draws
  them (`rowTexts` and `cellText`, so wrapping, overflow and `####`
  match), on white paper: black text and borders where the theme's would
  be, Excel's blue for links. A border is given to both cells beside it,
  since the collapsed borders of a table would otherwise let a gridline
  win over the border of the cell to its left or above. Hidden rows and
  columns, those of a filter too, are left out, and merged cells span
  those printed. Ctrl/Cmd+P opens the panel from anywhere, an entry
  being typed committed first, in place of the browser's print of the
  app's page.

## Ribbon

- **The ribbon grows with the features.** A tab or control is added by the
  slice that builds its feature, so none is ever a placeholder (R44).
- **The ribbon's panel scrolls sideways** when its controls are wider than
  the window, so that the page and the grid never do.
- **The ribbon shows the style of the active cell**, and Bold, Italic,
  Underline, Strikethrough and Wrap text turn their property off for the
  selection when the active cell has it and on otherwise, as in Excel.
- **The Number format menu** offers Excel's General, Number, Currency,
  Accounting, Percentage, Scientific, Date, Time and Text, with the US
  dollar and dates as in the United States whatever the interface
  language, and names the format of the active cell, or Custom; its last
  item opens a dialog for any code, previewing the active cell's value in
  it and refusing a code `numfmt` cannot read. Increase decimal, Decrease
  decimal and the thousands separator change the format of the active cell
  (`changedFormat`) and give the result to the whole selection, as Excel
  does.
- **The Borders menu keeps the line and color last chosen** for the next
  border.
- **Conditional formatting** in the Home tab opens a dialog for a new rule
  of each kind on the selection, with Excel's default colors (red, yellow
  and green scales, blue bars), and Manage rules a dialog that lists the
  sheet's rules.
- **The View tab's Freeze panes menu** holds the four commands, Freeze
  panes disabled at A1 and Unfreeze panes when nothing is frozen.
- **Rows and columns are inserted, deleted, hidden and shown** from the
  context menu of the row and column headers, for the selected lines; the
  ribbon has no row and column commands.
- **Sort** in the Data tab holds Sort A to Z and Sort Z to A by the active
  cell's column, and Custom sort, a dialog of levels, each a column and an
  order, whose columns the header row names when the first row holds
  headers.
- **Filter** in the Data tab turns the filter on and off, and Clear clears
  every column. A filter's button opens a panel under it (a Radix dialog
  without an overlay, placed under the button) with Clear filter, a choice
  of the values or a condition, a search box over a checklist whose search
  keeps only the values it finds, as Excel's, and the condition's value.
  Excel's sort items in the dropdown, its second condition joined by and
  or or, and its blue row numbers are not built.
- **Data validation** in the Data tab opens a dialog with the rule of the
  active cell, to give the selected cells, change or remove (Clear all, or
  Any value); dates in it are typed and shown as `m/d/yyyy`.
- **Insert function**, on the Formulas tab, opens the function list with
  every function, starting a formula in the formula bar if none is being
  typed.
- **Notes and links**: the Review tab holds New note (Edit note when the
  active cell has one) and Delete note, and the Insert tab Link, which
  inserts or edits the link of the active cell in a dialog of its text, a
  web address or a sheet and cell reference, with Remove link when it has
  one; the dialog shows the address or reference as the text of an empty
  cell left without one. The cell menu holds all of these, Remove link
  included.

## Find, names and the status bar

- **The find panel** opens with Ctrl/Cmd+F and Ctrl/Cmd+H, and Find and
  Replace in the Home tab's Editing group: a Radix dialog without an
  overlay at the top right that stays open while the grid is used, with
  Find what, Replace with, Within (sheet or workbook), Look in, Match case
  and Match entire cell contents; it keeps them until the page reloads.
  The shortcuts replace the browser's own find, even in a text box, but
  not while an entry is typed. macOS hides the browser on Cmd+H, so Ctrl+H
  replaces there too, as in Excel for Mac, and the tooltip names it. The
  panel has no Radix trigger, so on closing it gives the keyboard back to
  the element that had it when it opened, usually the grid. Excel's Find
  all list and its search by columns are not built.
- **The name box** shows the active cell, or the name whose range is
  exactly the one range selected; Enter there goes to a cell address, a
  range (on another sheet too) or a name, revealing its top-left cell and
  giving the keyboard back to the grid, or names the last range selected
  with a new name, and anything else, or a name whose cells were deleted,
  gives a message in the message bar. Its list, beside it, goes to a name
  chosen. The Formulas tab's Name manager lists the names with their
  ranges, edits a name and its range (typed with or without `=`, on the
  sheet shown unless it names one) and deletes one; Excel's New, Filter
  and comment of a name, and Define name, are not built, since the name
  box defines names.
- **The status bar shows a circular reference** as `Circular References:
  B1`, with its sheet when another sheet is shown (R15).
