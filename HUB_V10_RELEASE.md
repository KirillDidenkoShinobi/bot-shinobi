# Shinobi Team Hub V10 — incremental safety update

Implemented:
- Shared PostgreSQL storage for Telegram user timezone profiles, with one-time import of JSON data when the table is empty. Local JSON fallback when DATABASE_URL is unset.
- Server-side audit of authenticated mutating API requests (actor, path, HTTP method, status), including errors. No request bodies or secrets are logged.
- Referral delete route now agrees with the owner/deputy authorization policy. When a Telegram bot token is present, a failed Telegram revocation prevents deleting the local link record.

Limitations and pending work:
- XLSX work-hours schedules are still not migrated to PostgreSQL. The staff-duty schedule is a separate data model.
- Referral link creation and editing through the new UI are not complete.
- AI moderation settings editing remains incomplete.
- Auditing is API-level; actions in Telegram commands are not all included in the unified log.
- No live browser/Railway/PostgreSQL/Telegram end-to-end test has been run. Deploy to staging, back up database, and validate before production.
- Multiple concurrent writes to data.json still require broader consistency work.

Tests performed: Python syntax compilation, ZIP integrity verification.
