# app/platform

The decisions inside the platform module, each with its reason. Its
interface, the browser storage it keeps, and the decisions that span modules
are in the [architecture document](../../architecture.md).

- **Files are opened with a file input and saved as a download.** It works
  in all four browsers; the File System Access API does not.
- **The clipboard is heard through the browser's events.** The platform
  listens for copy, cut and paste on the whole document, outside text
  boxes, which works in every browser without a permission; the
  asynchronous clipboard API would ask one in Firefox and Safari to read.
  `writeClipboard` runs the browser's copy command and has finished when
  it returns.
- **A save cannot start while the page unloads**: browsers discard an
  IndexedDB transaction that begins then, so the workbook is saved when
  the page is hidden instead.
