# Shinobi Team Hub — Telegram-confirmed login

## Railway variables
- `HUB_PUBLIC_URL=https://your-public-domain.example` — REQUIRED, exact HTTPS site origin, without trailing slash.
- `HUB_OWNER_IDS=1197818066` — owner allowlist (1197818066 is also built-in).
- `HUB_DEPUTY_IDS=...` — comma-separated Telegram IDs.
- `HUB_ADMIN_IDS=...` — comma-separated Telegram IDs.
- `HUB_MODERATOR_IDS=...` — comma-separated Telegram IDs.
- `DATABASE_URL` — recommended/required for persistence across Railway redeploys. Auth records stored in PostgreSQL when configured. Without it, fallback SQLite is local and ephemeral.

## Flow
1. Authorized user opens bot DM and sends `/start`.
2. User clicks **Подтвердить вход на сайт** in Telegram DM. This is the Telegram confirmation step.
3. Bot generates a single-use URL valid for 5 minutes.
4. Opening the URL atomically consumes the token and creates an opaque 12-hour HttpOnly Secure cookie session.
5. Server rechecks current role on every request; role removal in env invalidates access.

## Roles
- owner: full access, can revoke sessions, see login audit.
- deputy: view all ordinary reports; edit links, schedules and operational settings.
- admin: view reports, create reports and assign tags.
- moderator: read-only operational dashboards.

Mutating APIs default to **owner-only** unless explicitly listed in `WRITE_ROLES` in `stats_api.py`. Backend checks are authoritative, even if a button is visible in UI.

## Admin APIs
- `GET /api/login-audit` owner only: login events.
- `POST /api/revoke-user` owner only: JSON `{ "uid": 123 }` revokes all sessions for the user.
- `GET /logout`: revoke current session.

**Security:** Treat a login URL as a bearer secret; Telegram confirmation proves the requesting Telegram account but cannot prevent another person from using a leaked URL first. Never share the URL. Restrict access to `/auth/telegram` in access logs if possible (URL contains the token). For stronger protection, add an additional device-bound confirmation step.
