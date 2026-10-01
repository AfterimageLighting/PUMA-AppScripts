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

The verified bound Script IDs are stored separately:

- TEST: `.clasp.test.json`
  `1TBhtqb5-3L3z34ozM87-mlJmxNxkcWofyeScPnjlHkCRBIxKdXBpiDpp`
- Production: `.clasp.production.json`
  `1ddj6ChwgU9bnUpO-by2EtkWFuQ2ZIno4lRwfLJOqt7ce6XQi-0kbWmMs`

The repository intentionally does **not** track an active `.clasp.json`.
Before any Apps Script push, explicitly select the intended target:

TEST:
`python select_clasp_target.py test`

Production:
`python select_clasp_target.py production`

That command generates a local, git-ignored `.clasp.json`. Confirm the target
shown by the selector before running `clasp push`. CI verifies that the active
file is not committed and that both stored target configs contain the exact
verified Script IDs.

## Timezone

This test branch uses:
`America/New_York`

The live repository previously used `America/La_Paz`.

## Trigger stabilization

This branch removes duplicate global simple-trigger definitions:
- one project-level `onEdit(e)` dispatcher
- one project-level `onOpen(e)` menu definition

Individual ESD handlers now have unique function names.

## Verified TEST integrations

Gmail account:
`adrian@afterimagelighting.com`

Verified TEST labels:
- PUMA TEST/OKD = Label_1
- PUMA TEST/RFA = Label_2
- PUMA TEST/RFA Processed = Label_8
- PUMA TEST/RFPO = Label_3
- PUMA TEST/PO = Label_4
- PUMA TEST/RR = Label_5
- PUMA TEST/RFPS = Label_6
- PUMA TEST/DR = Label_7

Verified TEST Drive destinations:
- PO PDFs = `1EeexxItqTX896tq64lf1-8lBz2bqCBmH`
- Receiving Reports = `1WMpQwDwGGVq8R1AYxvJPAP8gOrgmP9Bg`
- Delivery Reports = `12hNy09LDDFAjyDl4q97GTTsCysyvSgJf`

These are isolated TEST targets. Production targets remain unchanged.

## First test order

1. Run `python select_clasp_target.py test`, verify the TEST target, then push to the copied PUMA TEST Apps Script project only.
2. Reload PUMA TEST and confirm the PUMA menu loads with no authorization/runtime errors.
3. Run `PUMA -> Purchase Orders -> Audit RAW PO Project Resolution (Read Only)`.
4. Run the Open Projects manifest audit, config read/SafeSync dry test, missing-quotes report, and the separated quote/general audit outputs.
5. Validate Kinney, Harrison Smith, Maddux, Gough, Stony/Stoney Shore, blank
   project overrides, repeated part numbers, freight, shared POs, and ambiguous cases.
6. Do not run `Apply RAW PO Import` until all read-only results are reviewed.
7. Only after read-only results are clean, perform write tests inside PUMA TEST.
8. Configure/deploy an isolated Cloud TEST service with `PUMA_TEST_MODE=1` and the verified TEST labels/folders, then replay TEST-labeled messages.
9. Do not merge or deploy production until Apps Script and Cloud end-to-end TEST results pass.
