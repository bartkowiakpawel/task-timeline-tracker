# VBA source

The workbook contains two main pieces of VBA logic:

- `modTimeline.bas` - validates source data and generates the weekly timeline.
- `TrackerSheet.bas` - worksheet event used to filter `UPDATES` when a task is selected in `TRACKER`.

## Expected Excel table names

The worksheet-selection event expects:

- `tblTracker` on the `TRACKER` sheet
- `tblUpdates` on the `UPDATES` sheet

Both tables must contain a column named `Task_name`.

## Installation notes

`modTimeline.bas` belongs in a standard VBA module.

`TrackerSheet.bas` contains `Worksheet_SelectionChange`, so its code belongs in the worksheet module for `TRACKER` rather than in a standard module.

The `.xlsm` workbook in the repository is the ready-to-use version; these source files are provided so the macro logic can be reviewed directly on GitHub.
