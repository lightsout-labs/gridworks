# engine/format

The decisions inside the format module, each with its reason. Its interface,
and the decisions that span modules, are in the
[architecture document](../../architecture.md).

- **A change of decimals starts from what the value shows.**
  `changedFormat` gives a format one decimal more or fewer, or turns the
  thousands separator over; a General cell starts from the decimals its
  value shows, and a date, time, fraction or text format changes nothing.
- **A color in a code (`[Red]`) is given apart from the text**, by
  `formatColor`, so that the grid draws it over the style's color.
- **The fill character of a code (`* `, as in the accounting format) is
  left out** rather than repeated across the cell.
- **General is `numfmt`'s**, which already keeps a number to 11
  characters, so a General number is not rounded again to fit its column.
