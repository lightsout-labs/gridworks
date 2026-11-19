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
- **The app is kept offline by a service worker that serves it first from
  its cache** (R35), built by `vite-plugin-pwa` over the whole build and
  registered as the page loads. Any address under `/gridworks/` gets
  `index.html`, so a workbook's `#` address opens offline too. A visit
  after a deploy is served the version cached, while the browser fetches
  the new worker; that worker takes the page over at once (`skipWaiting`,
  `clientsClaim`), and the change of controller of a page that already had
  one tells the app that a reload shows the new version. The controller is
  noted as the page loads, so a takeover before the app listens is
  reported when it does. The new worker drops the old version's files,
  so the app imports its lazily loaded file formats as it starts. Serving the page
  from the network first would show the new version without a reload, but
  would wait for a slow network on every visit.
