# Task 1 Report: Backend Helpers For OFF-In-Day And Sleeping Rows

## Status: DONE

## Commits
- None (pending user request to commit)

## Test Results
- New test `test_quangchudong_helpers_split_off_today_and_sleeping_rows`: PASSED
- Full suite: 121 passed, 0 failed

## What Was Done
Added four helper functions to `blueprints/quangchudong_routes.py`:

- `_get_event_time(row)` — extracts datetime from `first_off_time` or falls back to `alert_time`
- `_filter_off_today_rows(rows, cutoff_time)` — filters rows with event time >= cutoff, sorted newest first
- `_with_sleep_days(row, now)` — enriches a row with `thoi_gian_ngu` (integer days since event, min 1)
- `_filter_sleeping_rows(rows, cutoff_time, now)` — filters rows with event time < cutoff, sorted by sleep days descending

Added one test in `tests/test_quangchudong_routes.py` covering all four helpers with edge cases (empty time, unparseable time, fallback to alert_time).

## Concerns
None.
