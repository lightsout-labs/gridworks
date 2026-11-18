# app/grid

The decisions inside the grid module, the canvas and the `View` it edits,
each with its reason. Its interface, and the decisions that span modules, are
in the [architecture document](../../architecture.md).

## The view

- **The view follows the sheets.** `View` subscribes to the workbook when
  it is made, right after the workbook and before anything else, so that
  when the sheet it shows is deleted, by a command or an undo, it shows the
  sheet that took its place before any other listener reads the deleted
  one. Switching to another sheet selects A1 there; each sheet's selection
  is not remembered.
- **Selection is a list of ranges** with an active cell, and an anchor and
  focus for the last range, which Shift moves. A whole row, column or sheet
  is a range that spans the sheet, so the engine's `aggregate` sums any
  selection the same way. The view grows every selection over the merges
  it overlaps and makes the top-left cell active, so a click inside a
  merged cell selects all of it, and an arrow leaves a merged cell from its
  far edge. Keyboard navigation passes over hidden rows and columns, and
  Ctrl+arrow treats them as absent.
- **What was copied lives in the view**, not the workbook: a range and
  whether it was cut. Pasting reads the source as it is then. Escape,
  undo and any change to the workbook other than a paste forget it, as
  Excel does, and a cut is pasted once.
- **The fill handle turns a drag into a `fill`.** It is a square on the
  bottom right corner of a single selected range; the view turns a drag of
  it into the target, extending the selection whichever way the pointer is
  furthest beyond it, and a double-click fills down as far as the column
  on its left has data, or else the one on its right. Ctrl/Cmd+D and
  Ctrl/Cmd+R copy without series, as in Excel, so they are a `paste` of
  the first row or column of the selection into the rest of it, or of the
  row above or column left of a selection one cell deep. A drag released
  with Ctrl (Option on macOS, where Ctrl clicks open the menu) swaps
  copying and counting on: `copy` when the fill would extend a series, and
  `series` otherwise.
- **The last fill can be redone another way.** After a drag or
  double-click of the fill handle, the view keeps the fill as `filled`
  (with whether `series` could extend it) until any other change to the
  workbook, as it does what was copied; the grid shows the Auto Fill
  Options button that the app gives it at the bottom-right corner of the
  target. Choosing an option calls `refill(how)`, which undoes the fill
  and runs it again with that `how`, so one undo still restores the cells
  as they were before the fill, and the button stays.
- **The format painter is a paste of formats.** Its button keeps the
  range selected in the view; the next range a click or drag selects, on
  the cells or the headers, gets its formats and merges through `paste`
  with `only: 'formats'`, so a target is tiled with whole copies, or gets
  one copy at its top left, as a paste is. Escape, or the button again,
  puts it away. It paints once: Excel's double-click to keep painting is
  not built.
- **Sorting and filtering start from the selection.** The view sorts the
  one range selected, cut down to the cells in use (`sortRange`), and none
  when that is one cell, so a single cell is not grown to the table around
  it as Excel does. It turns a filter on over the one range selected, or
  the data around the active cell (Ctrl+A's region) when one cell is
  selected, cut down to the cells in use; the grid places the button the
  app gives it at the right of each header cell of the filter.
- **Find goes from the active cell.** The view's `findNext` goes to the
  first match after the active cell, or round to the first, on the shown
  sheet or through the sheets in tab order, and gives which match it is
  and how many there are; it sets `reveal`, which the grid scrolls into
  view without taking the keyboard from the panel. `replace` replaces in
  the active cell when it holds a match, then goes to the next, and
  `replaceAll` gives how many cells it replaced.

## Entry

- **An entry has two modes**, as in Excel. Typing starts Enter mode, where
  an arrow commits and moves; F2, a double-click and the formula bar start
  Edit mode, where the arrows move the caret. Enter and Tab commit and
  move in both, and Tab therefore stays in the grid.
- **Only typing is validated.** The view's `commit` asks the workbook's
  `allows` first, and a rejected entry stays, marked `rejected`, and the
  selection does not move, until the user tries again (Retry, in Edit
  mode) or abandons it (Cancel or Escape), as Excel's Stop message; every
  command that commits first stops there too, Copy and Cut included, and
  the Data validation dialog does not open. Pastes, fills, text pasted
  from other applications, replace (which types each new text without
  asking) and choosing from a list are not checked, as in Excel. The grid
  shows the button the app gives it to the right of the active cell when
  its rule is a list, and it opens a menu of the items, which sets the
  cell to the one chosen.
- **An entry knows its cell's sheet and its caret.** While a formula is
  typed, a click or drag on the grid inserts a reference instead of
  selecting when a reference can go at the caret: after `=`, an operator,
  `(` or `,`, or in place of a reference that ends there, which is how a
  drag grows the reference it started. Clicking another sheet's tab then
  keeps the entry going in the formula bar, references picked there name
  the sheet, the grid there shows no selection, and committing or
  cancelling shows the entry's sheet again. Elsewhere a click commits, as
  before. The entry keeps the cells of the reference just picked as a
  selection of its own, `point`, which Shift with a click and a drag
  stretch from where it started; Ctrl/Cmd with a click adds a comma and a
  new reference after it. In Enter mode the arrows point too, where an
  operand starts or after the reference they or a click put there, but
  not after one typed (`canPoint`): `keyAction` gives `point` in place of
  `navigate`, and the view moves `point` as navigation moves the
  selection, starting from the cell edited. Any other change to the text
  or the caret forgets `point`, so the next arrow after an operator starts
  from that cell again.
- **The editors draw a formula's text in colors** by laying a copy of the
  text, with each reference in a span of its color, over the input box,
  whose own text is transparent and which still draws the caret and the
  selection. The same reference has the same color, from six theme
  variables `--reference-0` to `--reference-5` taken in turn, and the
  canvas outlines each referenced range of the shown sheet in it. Defined
  names are not colored as references.
- **Function help works on tokens, not the parser**, since a formula
  being typed is rarely complete. The list offers the defined names that
  start with the text typed, before the functions, and puts a name without
  a parenthesis; a name typed in full is not offered, so that Enter
  commits it. The arrows move in the list, Tab or Enter or a click
  chooses, and Escape closes it until the entry changes. The hint is that
  of the innermost function call whose parentheses hold the caret.
  Argument names are Excel's English ones in both languages, as function
  names are.

## Drawing

- **The grid scrolls natively.** An element that scrolls holds an element
  the size of the whole sheet at the zoom, and the canvas inside it sticks
  to the visible area and draws the cells under it on every scroll event.
  Wheels, trackpads, scroll bars and momentum then behave as everywhere
  else in the browser. The scroll position is that element's, not the
  `View`'s. It caps the sheet height: browsers limit an element to between
  about 17 and 33 million pixels, so a sheet has 100,000 rows of 24 pixels
  (4.8 million at 200%) and 100 columns, the numbers R1 states. More rows
  would need scroll bars of our own.
- **Lines are found by binary search.** The grid builds an `Axis` for each
  direction once per layout: it finds a line's place, or the line at a
  place, by binary search over the lines that differ from the standard
  size, so that drawing a frame stays fast with thousands of them. A
  hidden line is a line of size 0, which hit testing passes over. Dragging
  a header border resizes its line, or every whole line selected with it,
  and dragged to nothing hides them, as in Excel; a double-click fits a
  column to its widest text, and gives a row back its fitted height.
- **Frozen panes are drawn as up to four panes.** The page still scrolls
  natively over the whole sheet; the grid places a frozen line without the
  scroll and the others with it, so the lines just past the frozen ones
  slide under them, and draws the canvas in up to four panes (`panes`),
  each clipped and scrolled its own way, with a line where the frozen
  lines end. A selection, a merged cell or text that runs over the
  boundary is drawn in each pane it reaches, so ranges need nothing of
  their own; hit testing and scrolling to show a cell leave frozen lines
  in place, and the in-cell editor of a frozen cell follows the scroll.
- **Cells align by the type of their value** unless their style aligns
  them, as Excel's General alignment does: numbers right, text left,
  booleans and errors centered. Text that is too long runs over the empty
  cells to its right; text aligned center or right, and wrapped text, do
  not run over into other cells.
- **A number too wide for its column shows `####`**, as many as fit, as
  in Excel, wrapped or not; text runs over or is cut instead.
- **Colors are drawn in layers.** Explicit colors are the same in both
  themes; text with no color of its own on a fill is black or white,
  whichever shows on the fill, and otherwise the theme's. A number
  format's color is drawn over the style's, and a conditional format's
  `conditional` over both; a data bar is drawn in a solid color over the
  fill, inside the cell's lines and under its text, and merged cells show
  no bar.
- **Borders are drawn from each cell's own**, so the line between two
  cells is drawn when either has it there.
- **A merged cell is drawn as one box** with the fill and text of its
  top-left cell, and a cell's borders on the sides on the edge of its
  merge only.
- **Notes and links are drawn by the grid.** A link's text is underlined,
  in the theme's `--link` color unless the cell has a color of its own, so
  it shows in both themes. A note is marked with a triangle in the cell's
  top-right corner, in `--note`, and shown beside the cell once the
  pointer has rested on it for 300 ms. A click that selects a cell with a
  link and does not drag follows it when released: the view selects the
  cells of a reference, scrolling them into view, or gives a web address,
  which the grid opens in a new tab with `noopener`; Excel's
  click-and-hold to select a link without following it is not built, but
  the keyboard selects it.
