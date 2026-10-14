# V7 — conflict detection for web schedule edits

- The web schedule editor now performs PostgreSQL row-locked compare-and-swap against the record loaded by the editor. Concurrent edits from two website sessions return HTTP 409 instead of silently overwriting one another.
- On conflict, the editor refreshes and asks the user to review changes before saving again.
- Added GET /api/hub/dayoff-requests for viewing the existing XLSX day-off queue.

## Limitations
- Telegram's legacy schedule callbacks still use full-document read-modify-write and can race with the web editor. Migrating all Telegram mutations to the same atomic operation is still required for full concurrency guarantees.
- Day-off approvals are **not** enabled in the website. Existing XLSX is a local filesystem snapshot and may not persist across Railway redeploys or multiple replicas. Approval workflows should first migrate to PostgreSQL and share one backend with the bot.
- No live Railway or browser integration test was performed.
