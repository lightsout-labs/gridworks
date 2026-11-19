# engine/formula

The decisions inside the formula module, each with its reason. Its interface,
and the decisions that span modules, are in the
[architecture document](../../architecture.md).

- **Formulas are stored as Excel writes them**: function names and
  references in capitals, sheet names as the sheets have them, without
  spaces other than the intersection operator.
- **A range where one value is expected gives the cell in the formula's
  row or column**, Excel's implicit intersection, and `#VALUE!` when there
  is none; dynamic arrays are out of scope.
- **Functions are specifications in a library**, one table for each
  family (math, statistics, logic, conditional totals, text, dates,
  lookup and reference, information), each giving its number of
  arguments, whether files prefix its name with `_xlfn.`, the format
  Excel gives a cell for its result if it is a date or time (which
  `resultFormat` reads), and its implementation. The parser refuses a call with the wrong number of
  arguments, or with a constant where it needs a range (as the first
  argument of `SUMIF`), as Excel does, so it stays text. An implementation
  receives its arguments unevaluated and asks for each as a value, a
  number or a reference, so that `IF`, `IFS` and `IFERROR` evaluate only
  what they need; most read each argument as one type through `typed`,
  where a read that fails gives its error as the result. A formula reads a
  range by visiting its cells that are not empty, with their places, so
  that a whole column costs only its cells and `SUMIF` can pair its ranges
  cell by cell. A lookup reads only the row or column it searches.
  `INDEX`, `CHOOSE` and `XLOOKUP` give references, as in Excel.
  `ISFORMULA` and `FORMULATEXT` read the formula of a cell, as printed for
  the formula bar, from the context of the evaluation. Their argument is a
  dependency like any other, so one that names its own cell is a circular
  reference; the graph does not tell a read of a formula from a read of a
  value, and nothing requires the difference.
- **No function is volatile.** The dependency graph only knows the cells
  a formula names, so `TODAY` and `NOW` read the clock when their formula
  is calculated (on entry, on a change to what it reads, and when the
  workbook opens), and `RAND`, `OFFSET` and `INDIRECT`, whose results or
  references change without an edit, are not built.
- **Argument values are listed beside the signatures**, by function and
  argument, as Excel offers them (`VLOOKUP`'s `range_lookup`: `TRUE`,
  `FALSE`; `MATCH`'s `match_type`: `1`, `0`, `-1`), and `argumentHint`
  gives those of the current argument. Their descriptions are interface
  text and live in the app's dictionaries.
