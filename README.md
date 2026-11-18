# Gridworks

Gridworks is a spreadsheet application that provides the main features of
Microsoft Excel in the browser, including opening and saving real `.xlsx`
files and `.csv` files, with no installation and no account. Its interface is in English and
Simplified Chinese, with light and dark themes. It is written in TypeScript,
runs entirely client-side, and is deployed to GitHub Pages.

**Use it at <https://lightsout-labs.github.io/gridworks/>.**

The requirements are in [docs/requirements.md](docs/requirements.md), the
design in [docs/architecture.md](docs/architecture.md) and, inside each
module, in [docs/design/](docs/design/), and progress in
[docs/backlog.md](docs/backlog.md).

## Development

Gridworks needs the Node.js version in [.nvmrc](.nvmrc) and `make`.

| Task | Command |
|---|---|
| Install | `npm ci`, then `npx playwright install` once for the end-to-end tests |
| Run locally | `npm run dev`, then open <http://localhost:5173/gridworks/> |
| Test | `make check` (format, lint, types, unit tests with coverage, documents) `make e2e` (the built app in Chromium, Firefox and WebKit), and `make perf` (the time limits for scrolling, recalculation and opening files) |
| Build | `npm run build`, which writes the static site to `dist/` |
| Upgrade Playwright | Change the version of `@playwright/test`, the tag of both `mcr.microsoft.com/playwright` images in [.github/workflows/ci.yml](.github/workflows/ci.yml) to the same version, and run `npx playwright install` again |

Every merge to `main` is deployed to the address above by
[a GitHub Actions workflow](.github/workflows/deploy.yml), which then runs a
smoke test against it.

## License

[MIT](LICENSE)
