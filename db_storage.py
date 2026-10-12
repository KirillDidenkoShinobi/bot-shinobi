import asyncio
import hashlib
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()
BASE_DIR = Path(__file__).resolve().parent
DATABASE_URL = os.getenv('DATABASE_URL', '').strip()
_ALLOWED_TOP_EXT = {'.json', '.xlsx'}


def enabled():
    return bool(DATABASE_URL)


def _connect():
    import psycopg2
    return psycopg2.connect(DATABASE_URL, connect_timeout=10)


def _safe_path(rel: str) -> Path:
    p = (BASE_DIR / rel).resolve()
    if BASE_DIR.resolve() not in p.parents and p != BASE_DIR.resolve():
        raise ValueError('unsafe path')
    return p


def _managed_files():
    files = []
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.suffix.lower() in _ALLOWED_TOP_EXT:
            files.append(p)
    media = BASE_DIR / 'site_media'
    if media.exists():
        files.extend(p for p in media.rglob('*') if p.is_file() and p.suffix.lower() in {'.png','.jpg','.webp'})
    reports = BASE_DIR / 'stats_reports'
    if reports.exists():
        files.extend(p for p in reports.rglob('*') if p.is_file())
    return files


def _rel(p: Path):
    return p.relative_to(BASE_DIR).as_posix()


def init_schema():
    if not enabled():
        print('[DB] DATABASE_URL не задан — PostgreSQL persistence отключён')
        return False
    with _connect() as conn, conn.cursor() as cur:
        cur.execute('''
            CREATE TABLE IF NOT EXISTS shinobi_storage (
                path TEXT PRIMARY KEY,
                content BYTEA NOT NULL,
                sha256 TEXT NOT NULL,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        ''')
    return True


def restore_or_seed():
    """На старте: если БД уже содержит данные — восстанавливаем их; если пустая — загружаем текущие файлы как начальную базу."""
    if not init_schema():
        return
    with _connect() as conn, conn.cursor() as cur:
        cur.execute('SELECT path, content FROM shinobi_storage ORDER BY path')
        rows = cur.fetchall()
        if rows:
            for rel, content in rows:
                try:
                    dst = _safe_path(rel)
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    dst.write_bytes(bytes(content))
                except Exception as e:
                    print(f'[DB] Не удалось восстановить {rel}: {e}')
            print(f'[DB] Восстановлено файлов из PostgreSQL: {len(rows)}')
            return
    sync_to_db()
    print('[DB] База была пустой — текущие данные сохранены как начальный снимок')


def sync_to_db():
    if not enabled():
        return 0
    files = _managed_files()
    local = {}
    for p in files:
        try:
            blob = p.read_bytes()
            local[_rel(p)] = (blob, hashlib.sha256(blob).hexdigest())
        except OSError:
            pass
    with _connect() as conn, conn.cursor() as cur:
        for rel, (blob, sha) in local.items():
            cur.execute('''
                INSERT INTO shinobi_storage(path, content, sha256, updated_at)
                VALUES (%s, %s, %s, NOW())
                ON CONFLICT(path) DO UPDATE SET
                    content = EXCLUDED.content,
                    sha256 = EXCLUDED.sha256,
                    updated_at = NOW()
                WHERE shinobi_storage.sha256 IS DISTINCT FROM EXCLUDED.sha256
            ''', (rel, blob, sha))
    return len(local)



def _init_settings_schema():
    if not enabled():
        return False
    with _connect() as conn, conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS shinobi_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
    return True

def set_setting(key: str, value):
    """Shared runtime setting for bot + web, persisted in PostgreSQL."""
    if not enabled():
        return False
    _init_settings_schema()
    text = 'true' if value is True else 'false' if value is False else str(value)
    with _connect() as conn, conn.cursor() as cur:
        cur.execute("""
            INSERT INTO shinobi_settings(key, value, updated_at) VALUES (%s,%s,NOW())
            ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value, updated_at=NOW()
        """, (key, text))
    return True

def get_setting(key: str, default=None):
    if not enabled():
        return default
    _init_settings_schema()
    with _connect() as conn, conn.cursor() as cur:
        cur.execute('SELECT value FROM shinobi_settings WHERE key=%s', (key,))
        row = cur.fetchone()
    if not row:
        return default
    value = row[0]
    if value == 'true': return True
    if value == 'false': return False
    return value

async def persistence_loop(interval=3):
    if not enabled():
        return
    print('[DB] Автосохранение PostgreSQL включено')
    while True:
        try:
            await asyncio.to_thread(sync_to_db)
        except Exception as e:
            print(f'[DB] Ошибка синхронизации: {e}')
        await asyncio.sleep(interval)
