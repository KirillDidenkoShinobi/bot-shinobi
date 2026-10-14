# Shinobi Team Hub — Control Center V2

## New
- Glass UI overview with live summaries from existing report/chat/referral/status endpoints.
- Staff section with server-side owner-only Telegram ID role assignments: deputy, admin, moderator, disabled.
- Persistent staff role overrides in PostgreSQL; existing Railway role variables remain fallback.
- Immediate revocation of active sessions whenever an employee role changes.
- Existing reporting, referrals, moderation log, chat analytics, settings and bot management routes retained.
- Responsive UI, motion preferences and confirmation before role changes.

## Deploy
Upload ZIP as existing Railway project (keep existing DATABASE_URL and Telegram OIDC variables). Keep HUB_OWNER_IDS=1197818066. Staff role changes persist in PostgreSQL and take priority over legacy role variables. Owner IDs are configured only via Railway, not editable in UI.

## Notes
This release integrates old functional pages into the new shell and adds live overview + staff management. It does not claim to add a new independent frontend framework or implement unsupported bot features. Test on a staging deployment before replacing production. Do not distribute data.json or existing backup exports publicly.
