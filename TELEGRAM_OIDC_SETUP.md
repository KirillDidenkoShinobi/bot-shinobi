# Telegram Login — Shinobi Team Hub

BotFather → Login Widget → Redirect URIs:
`https://shinobi-bot-production.up.railway.app/auth/telegram/callback`

Railway Variables:
- `HUB_PUBLIC_URL=https://shinobi-bot-production.up.railway.app`
- `TELEGRAM_CLIENT_ID=8784220957`
- `TELEGRAM_CLIENT_SECRET=` (новый секрет после Revoke, **не помещать в код или чат**)
- `DATABASE_URL` (постоянная PostgreSQL)
- `HUB_OWNER_IDS`, `HUB_DEPUTY_IDS`, `HUB_ADMIN_IDS`, `HUB_MODERATOR_IDS` — Telegram ID через запятую.

Теперь `/start` у администраторов даёт кнопку перехода на сайт. На сайте нажимают «Войти через Telegram»; используется официальный OIDC authorization code + PKCE S256, проверка state, RS256 подписи JWT через Telegram JWKS, issuer/audience/expiry, проверка роли, сессия 12 часов, отзыв сессий и журнал входов. Одноразовый старый endpoint `/auth/telegram?token=` сохранён только для обратной совместимости, но бот больше не выдаёт такие ссылки.

Для проверки: разверните на Railway, откройте страницу, войдите, затем `/logout`. Реальная авторизация требует доступного Telegram OIDC и правильных настроек BotFather.
