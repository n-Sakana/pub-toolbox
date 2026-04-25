# Architecture

## Host

The add-in is the host. It installs a small menu and exposes tool entry points.

## FE / BE

- FE is expressed as sheets inside the user's workbook.
- BE is VBA code in the add-in.
- The first implementation runs inline.
- Background mode is a later adapter, not a different business model.

## Outlook Draft Tool

- Input source: current table or current region on the active sheet
- FE output: `outlook_draft` sheet with a generated table
- Run action: create Outlook drafts row by row and write status back to the sheet
