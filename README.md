# HROne Counter Builder

HROne-themed validator and counter generator for monthly HROne-to-SAP journal
extracts. It is available as both a hosted browser application and an offline
Windows desktop application.

Posting keys `40` and `29` are debit; `50`, `34`, and `39` are credit. The app
rejects a source file unless total debit equals total credit. It then creates
balanced counters containing no more than 999 rows. When a sequential boundary
is unavailable, it finds exact combinations and reorders complete source rows.
Every original row must be used exactly once.

The app accepts `.xls`, `.xlsx`, and `.csv`, processes the file locally, and
exports an `.xlsx` that preserves the source values, dates, and row count while
appending `DR/CR` and `COUNTER BUILDER`.

## Windows desktop build

The Electron desktop edition uses the same validated React interface and
counter engine. Network requests are blocked in the desktop shell, so workbook
processing works without an internet connection or browser login.

Build the Windows installer and portable executable with:

```powershell
pnpm run desktop:build
```

Artifacts are written to `outputs/windows`.

## GitHub Pages

The browser application can also be published as a static GitHub Pages site.
Build the deployable `docs` directory with:

```powershell
pnpm run pages:build
```

Pushes that change `docs` on the default branch are deployed automatically by
the GitHub Pages workflow.
