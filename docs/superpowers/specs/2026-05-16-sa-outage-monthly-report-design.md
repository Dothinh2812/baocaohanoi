# SA Outage Monthly Report Design

## Goal

Add a monthly report below the existing `/su_co_sa` outage sections. The report lists all SA outage incidents for the selected month and shows start time, end time, SA name/code, affected subscriber count, and date/month context.

## Data Source

Use the existing SQLite database at `/home/vtst/1bss/runtime/default/sqlite/sa_outage.db`, table `sa_outage_incidents`.

The affected subscriber count is `max_off_count`, because it represents the highest number of subscribers that lost contact during the incident.

## Filtering

The report is filtered by month using `started_at`.

The API accepts an optional `month` query parameter in `YYYY-MM` format. If omitted or invalid, it uses the current local month.

## UI

The new report appears below:

1. Dang ton dong
2. Da clear hom nay
3. Thong ke tat ca su co theo thang

The report includes a month picker, total incident count, total affected subscriber count, and a table with incident code, SA, team, status, start time, end time, date, affected subscribers, duration, recovery reason, and user handling notes.

## User Notes

Each monthly report row includes a separate user-editable handling note. Notes are saved with a per-row `Luu` button.

Do not use the existing `notes` column for user input because the `/home/vtst/1bss` SA outage pipeline owns that column and rewrites it during incident open/update/close operations. Dashboard user notes are stored in separate nullable columns:

- `user_notes`
- `user_notes_updated_at`
- `user_notes_updated_by`

The dashboard performs an idempotent SQLite migration before reading or writing these columns. Writes use a short SQLite connection with WAL mode and `busy_timeout=5000` to reduce contention with the pipeline writer.

## Testing

Add route-level tests for the monthly report query and note-save behavior using a temporary SQLite database and monkeypatched DB path/cache.
