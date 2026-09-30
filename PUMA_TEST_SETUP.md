# PUMA TEST REVIEW SETUP

This branch is for the isolated PUMA review environment only.

## Test workbook

PUMA TEST - DO NOT USE FOR PRODUCTION - 2026-09-30

Spreadsheet ID:
`1wjGB4dUTbBUiWVC7Kh0huVw8wM_ppmjTgktLw6ElVf4`

URL:
https://docs.google.com/spreadsheets/d/1wjGB4dUTbBUiWVC7Kh0huVw8wM_ppmjTgktLw6ElVf4/edit

The workbook was copied from live PUMA and verified against representative ranges in:
- Tracker config
- RAW_PO_IMPORT
- Kinney - Project Tracker
- Harrison Smith Residence - Project Tracker

The copied workbook also contains a visible `PUMA TEST CONTROL` tab.

## Bound Apps Script

Google copies bound Apps Script code with a copied Google Sheet. The copied
script is a separate Apps Script project.

For safety, this branch's `.clasp.json` does NOT point at production. It is
intentionally set to:

`REPLACE_WITH_PUMA_TEST_SCRIPT_ID`

Before any CLASP push:
1. Open the PUMA TEST spreadsheet.
2. Extensions -> Apps Script.
3. Project Settings.
4. Copy the Script ID.
5. Replace the placeholder in this branch only.
6. Confirm the Script ID is NOT the live production ID:
   `1ddj6ChwgU9bnUpO-by2EtkWFuQ2ZIno4lRwfLJOqt7ce6XQi-0kbWmMs`

## Timezone

This test branch uses:
`America/New_York`

The live repository previously used `America/La_Paz`.

## Trigger stabilization

This branch removes duplicate global simple-trigger definitions:
- one project-level `onEdit(e)` dispatcher
- one project-level `onOpen(e)` menu definition

Individual ESD handlers now have unique function names.

## First test order

1. Open PUMA TEST.
2. Confirm the PUMA TEST CONTROL tab appears first.
3. Push this branch to the copied Apps Script project only.
4. Reload PUMA TEST.
5. Run `PUMA -> Purchase Orders -> Audit RAW PO Project Resolution (Read Only)`.
6. Do not run `Apply RAW PO Import` until the audit output has been reviewed.
7. Validate Kinney, Harrison Smith, Maddux, Gough, Stony/Stoney Shore, blank
   project overrides, repeated part numbers, freight, and ambiguous cases.
8. Only after read-only results are clean, perform write tests inside PUMA TEST.
