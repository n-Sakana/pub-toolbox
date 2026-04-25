# pub-vba-toolbox

Small and medium VBA tools collected behind a single Excel add-in.

## Current Tool

- `outlook-draft`: create an empty draft table in the active workbook and generate Outlook drafts from it.

## Design

- FE: worksheet-based tool UI in the user's workbook
- BE: VBA logic called from the add-in
- Default mode: inline execution in the current Excel process
- Future mode: optional background execution with the same job model

## Layout

```text
src/
  addin/
  common/
  tools/
    outlook-draft/
scripts/
docs/
dist/
sample/
```

## Build

- `build-addin.bat`: build `dist/vba-toolbox.xlam`
- `build-sample.bat`: build `sample/vba-toolbox-sample.xlsx`

## Trial Flow

1. Build the add-in.
2. Install `dist/vba-toolbox.xlam` in Excel.
3. Build and open the sample workbook.
4. Open the `vba-tools` ribbon tab and click `Create Draft Sheet`.
5. Copy columns `A:G` from `source` into `outlook_draft`, or fill `outlook_draft` with formulas.
6. `from` is optional. Leave it blank to use the default Outlook account.
7. Use `;` for multiple mail addresses and attachment paths.
8. Click `Run Drafts` on the same ribbon tab.
