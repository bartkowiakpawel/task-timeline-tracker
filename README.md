# Task Timeline Tracker

A lightweight Excel/VBA task tracking tool with weekly timelines, update history, task grouping, hover notes, quick filtering, and built-in data validation.

> **Track what happened, not just how long it took.**

## Why this exists

Sometimes a full project-management platform is unnecessary, but a simple task list is not enough. This workbook is designed for lightweight BA / PM / transformation / operations workflows where you want to understand:

- when a task was created,
- when work actually started,
- how long it remained active,
- what happened along the way,
- and why a task may have stayed open for several weeks.

The workbook keeps source data simple and generates a visual weekly timeline with update markers and hover notes.

## Features

- Weekly task timeline generated with one macro
- Separate `TRACKER`, `UPDATES`, and `TIMELINE` sheets
- Optional task grouping
- Empty groups automatically shown as **Unassigned**
- Waiting period visualization: `Creation Date -> Start Date`
- Active period visualization: `Start Date -> End Date / Today`
- Update history displayed as `!` markers on the timeline
- Multiple updates in the same week displayed as `!!`, `!!!`, etc.
- Hover notes with date, update type, reporter, and comment
- Current week highlighting
- Expand/collapse groups using Excel outline controls
- Click a task in `TRACKER` to filter `UPDATES`
- Source sheets are read-only from the timeline-generation macro
- Built-in validation before timeline generation

## Workbook structure

### TRACKER

| Column | Field |
|---|---|
| A | `Task_name` |
| B | `Creation_date` |
| C | `Start_date` |
| D | `End_date` |
| E | `Status` |
| F | `Owner` |
| G | `Group` |

`Group` is optional. Blank values are shown under **Unassigned** in the generated timeline.

### UPDATES

| Column | Field |
|---|---|
| A | `Task_name` |
| B | `Update_date` |
| C | `Update_type` |
| D | `Reporter` |
| E | `Comment` |

Example update types: `Progress`, `Issue`, `Shift`, `Hold`.

### TIMELINE

The timeline is generated in weekly intervals (Monday-Sunday).

- **Grey** - task existed but work had not started yet
- **Blue** - active/completed working period
- **Yellow header** - current week
- **!** - one update exists in that week
- **!! / !!!** - multiple updates exist in that week

Hover over a cell containing `!` to see update details.

## Validation rules

Timeline generation stops if:

- duplicate task names exist in `TRACKER`,
- an update refers to a task that does not exist in `TRACKER`,
- an update date is missing or invalid,
- an update date is later than the task's `End_date`,
- a populated source row has no task name.

This intentionally favors explicit data-quality errors over silently generating an inconsistent timeline.

## Recommended workspace setup

For day-to-day use, open `TRACKER` and `UPDATES` side by side:

**View -> New Window -> Arrange All -> Vertical**

Recommended layout:

- left window: `TRACKER`
- right window: `UPDATES`

Selecting a task in `TRACKER` automatically filters `UPDATES` to related entries. Clicking outside the task list removes the filter.

## Getting started

1. Download `task_timeline_tracker.xlsm`.
2. Open it in desktop Microsoft Excel.
3. Enable macros if you trust the workbook and your Excel security policy allows it.
4. Add or replace sample data in `TRACKER` and `UPDATES`.
5. Open `TIMELINE` and run **Generate Timeline**.

> Excel for the web does not run VBA macros. Use the desktop version of Excel.

## Source code

The VBA source is also included separately in [`src/`](src/) so it can be reviewed without opening the macro-enabled workbook.

## Screenshots

Screenshots will be added in [`screenshots/`](screenshots/) as the public template is refined.

## Security note

This repository includes a macro-enabled Excel file (`.xlsm`). Review the VBA source in `src/` before enabling macros if you downloaded the workbook from an untrusted source or modified fork.

## Author

Developed by **Paweł Bartkowiak**.

AI assistance: **ChatGPT by OpenAI**.

## License

MIT License. See [`LICENSE`](LICENSE).
