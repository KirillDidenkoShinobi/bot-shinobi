# Shinobi Team Hub V8 — заявки на выходные

- PostgreSQL: общая таблица hub_dayoff_requests для Telegram и сайта.
- Первичный импорт из существующего day_off_requests.xlsx с ON CONFLICT DO NOTHING; существующие решения в БД не перезаписываются.
- Новые заявки через Telegram сохраняются в PostgreSQL при DATABASE_URL.
- Одобрение/отклонение с сайта: owner/deputy, атомарный переход только из Pending, журнал действий, попытка уведомить пользователя через Bot API.
- Telegram-кнопки тоже используют общий PostgreSQL-статус; повторные решения отклоняются.
- Если DATABASE_URL отсутствует, старый Excel-режим Telegram остаётся, но решения через сайт недоступны.
- XLSX-снимки в разделе отчётов не являются живой базой.

Проверить в тестовом Railway с резервной копией PostgreSQL. Реальный Telegram и Railway не тестировались.
