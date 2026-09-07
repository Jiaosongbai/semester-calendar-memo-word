---
name: semester-calendar-memo-word
description: Convert university academic calendars supplied as PDFs, images, or text into minimalist, printable and editable Word weekly planners. Use for 校历转记事本、学期规划表、周历本、学年记事本 requests. Generate DOCX only; exclude class timetables, daily planners, reminders, subscriptions, apps, and spreadsheet output.
---

# Semester Calendar Memo Word

Convert the requested portion of an academic calendar into a minimalist `.docx` weekly planner.

## Determine the requested range

Use the range explicitly requested by the user when the combined information from supplied materials and permitted authoritative research is sufficient for that range.

- Generate a full academic year when the user requests it and both semesters are sufficiently documented.
- Generate only the requested semester when only that semester is requested or supported.
- If required information remains missing, list the missing fields and wait. Never invent dates or week counts.

Each semester requires its name, the Monday of teaching week 1, and the total included weeks. Include examination weeks when documented.

## Research and source precedence

By default, verify the applicable national statutory-holiday schedule online using authoritative government sources. Research a specific local, ethnic, religious, commemorative, or school-created holiday only when the user asks for it.

Apply sources in this order:

1. The user's academic calendar for teaching, examinations, registration, campus events, and school vacations.
2. The institution's official publication when additional school information is needed.
3. Authoritative government notices for statutory-holiday dates.

Do not let general web results override the institution's teaching or vacation schedule. Preserve source notes for researched dates.

## Classify dates before rendering

- `statutory_holiday`: an officially published holiday interval. Mark every confirmed day in the interval.
- `calendar_day_marker`: a festival name shown on the academic calendar without a confirmed leave interval. Label it as a calendar marker; do not infer consecutive leave or make-up workdays.
- `school_break`: an institution-defined winter, summer, anniversary, or other school vacation. Keep it distinct from statutory holidays.
- `activity`: a campus event with day, month, or range precision.

## Prepare and generate

Read [references/spec-schema.md](references/spec-schema.md) before structuring extracted calendar data. Use [references/spec-example.json](references/spec-example.json) only when a concrete example helps; do not copy its institution-specific facts into another calendar.

Create a validated JSON file, then run the bundled generator using paths relative to this skill folder:

```bash
python scripts/generate_notebook.py <spec.json> --output <result.docx>
```

The script requires `python-docx` and `lxml`. It validates required dates, builds automatically expanding Word tables, and creates true Word footnotes. Never edit the script to embed a user's institution or an environment-specific absolute path.

## Word output requirements

- Generate DOCX only.
- Use one complete, non-splitting weekly block with four columns and two segments: Monday–Wednesday above, Thursday–Sunday below.
- Show every date as `M/D`; include the year on each Monday as `YYYY/M/D`.
- Keep each body row at a minimum height that expands automatically when the user types or presses Enter. Never use an exact fixed row height for writable cells.
- Keep approximately 10 pt of unbordered whitespace between weekly blocks.
- Fit three to four blank weekly blocks per page under normal conditions. Allow Word to repaginate after the user adds content.
- Start each semester on a new page and identify examination weeks.

## Activities and footnotes

- Anchor an event with a confirmed start date to that date.
- Distribute month-only events across pages within that month; the layout anchor is not an asserted start date.
- Put shorter activities before longer-span activities on the same page, and place longer-span items later when dates are imprecise.
- Use true Word footnotes with matching sequential references.
- Display both the date-cell reference and the footnote number as a single gray `#777777`, 14 pt superscript. Each footnote body must begin with exactly one real footnote-reference element; never add a duplicate literal number.
- Style footnote text as gray `#777777`, 8 pt. Keep content inside page margins.
- Treat 16 rendered footnote lines as the per-page upper bound under the validated blank-page layout. Continue or redistribute overflow without shrinking below 8 pt.

## Validate before delivery

Render and inspect every page. Confirm:

- requested range and week boundaries;
- statutory holidays, calendar markers, and school breaks are not conflated;
- weekly blocks stay intact and whitespace is balanced;
- writable rows expand when populated with multiple paragraphs;
- footnote references and page-bottom numbers are sequential, unique, matched, and visibly superscript;
- no content crosses the page margins.

Deliver only the final `.docx` unless the user asks for supporting material.
