# Проверка уведомлений о заявках

- `drop_pending_updates=False`: Telegram updates не удаляются при запуске.
- Логи `[JOIN REQUEST RECEIVED]`, `[JOIN REQUEST IGNORED]`, `[JOIN REQUEST DUPLICATE]`, `[JOIN REQUEST DELIVERED]`.
- Необходимо оставить один polling-процесс на один BOT_TOKEN, иначе TelegramConflictError продолжится.
- `/delivery_test` проверяет только отправку сообщений, не получение chat_join_request.
- При новой заявке смотреть Railway logs на `[JOIN REQUEST RECEIVED]`.
- Старые ссылки и статистика не менялись.
- Не загружайте старый data.json поверх актуальной БД.
