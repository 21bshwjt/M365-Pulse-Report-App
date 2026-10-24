# Pulse Report Builder

Turns an M365 Pulse Report `.xlsx` workbook into a colorful, standalone HTML security report.
No AI, no install, no network. The workbook never leaves your computer.

## Use it

**Drag and drop:** open `index.html` in Edge or Chrome, drop the workbook on the page, then click **Save HTML report**.

**Command line** (Node 18+), for scheduled or batch runs:

```
node cli.js report.xlsx                     # writes report.html next to the workbook
node cli.js report.xlsx -o C:\Reports\out.html
node cli.js *.xlsx -d C:\Reports            # many workbooks at once
```

Keep `pulse-engine.js` in the same folder as `index.html` and `cli.js`.

## How it reads the workbook

Sheet names are ignored. Each sheet's header row is compared with the column signatures in
`SIGNATURES` at the top of `pulse-engine.js`. For example, any sheet with a UPN column and an
MFA-registration column is treated as the user list. Column names are compared without case,
spaces or punctuation, so `MFA Registered`, `MFARegistered` and `mfa_registered` all match.

Any other small sheet with a `Status`, `Result` or `Compliant` column is picked up as a set of
tenant setting checks automatically. Values such as `COMPLIANT`, `NOT-COMPLIANT`, `Yes`,
`No - Expected …` and `Risk Detected` are all understood.

Sections whose data is missing are left out, and the report footer lists what was not found.

## Tuning

- **Thresholds** (stale days, expiry warnings, MFA targets): `CONFIG` in `pulse-engine.js`.
- **A column was renamed and is no longer recognised:** add the new name, lower-case with
  spaces and punctuation removed, to the matching alias list in `SIGNATURES` or in the
  section's `col(...)` call.
- **New findings:** add a rule inside `analyse()`. Each rule calls
  `add(severity, area, title, detail, sheetName)`, with severity `crit`, `high`, `med` or `info`.

Fonts load from Google Fonts when online. Offline, the report falls back to system fonts.
