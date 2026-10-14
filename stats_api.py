import base64
import hashlib
import hmac
import json
import os
import shutil
import secrets
import time
import io
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from aiohttp import web
import hub_auth
import aiohttp
import jwt
from urllib.parse import urlencode
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from chat_summary import snapshot as chat_summary_snapshot, uptime_seconds

BASE_DIR = Path(__file__).resolve().parent
REPORTS_DIR = BASE_DIR / "stats_reports"
REPORTS_DIR.mkdir(exist_ok=True)
SOURCE_FILES = [
    "shinobi_links_log.xlsx",
    "work_schedules.xlsx",
    "day_off_requests.xlsx",
    "data.json",
    "admin_assignments.json",
    "user_timezones.json",
    "project_settings.json",
    "shinobi_server_tags.txt",
    "changelog.json",
]
ADVERTISING_LINK = "https://t.me/+xXGHYqokIiMzNjYy"
MEDIA_DIR = BASE_DIR / "site_media"
MEDIA_DIR.mkdir(exist_ok=True)
MAX_IMAGE_BYTES = 8 * 1024 * 1024



def _build_staff_schedule_xlsx(dst: Path):
    """Build website staff schedule from admin duties only; no work-hour columns."""
    try:
        assignments = json.loads((BASE_DIR / "admin_assignments.json").read_text(encoding="utf-8"))
    except Exception:
        assignments = {}
    try:
        users = json.loads((BASE_DIR / "data.json").read_text(encoding="utf-8")).get("users", {})
    except Exception:
        users = {}

    weekdays = [("Monday","Понедельник"),("Tuesday","Вторник"),("Wednesday","Среда"),("Thursday","Четверг"),("Friday","Пятница"),("Saturday","Суббота"),("Sunday","Воскресенье")]
    duty_names = {"search":"Искать людей", "chat":"Следить за чатом", "tags":"Выдавать теги", "owner_tech":"Владелец / Тех админ"}
    wb = Workbook(); ws = wb.active; ws.title = "График сотрудников"
    headers = ["Сотрудник", "День", "Обязанность"]
    for c, value in enumerate(headers, 1):
        cell=ws.cell(1,c,value); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E79"); cell.alignment=Alignment(horizontal="center")
    row=2
    for uid, week in assignments.items():
        u = users.get(str(uid), {})
        name = u.get("name") or u.get("username") or str(uid)
        if u.get("username") and u.get("name"):
            name += f" (@{u['username']})"
        for day_en, day_ru in weekdays:
            rec = (week or {}).get(day_en, {}) or {}
            if rec.get("off"):
                duty = "Выходной"
            elif rec.get("always"):
                duty = "Владелец / Тех админ"
            else:
                duty = ", ".join(duty_names.get(x,x) for x in rec.get("duties", [])) or "Без обязанностей"
            ws.cell(row,1,name); ws.cell(row,2,day_ru); ws.cell(row,3,duty)
            for c in range(1,4): ws.cell(row,c).alignment=Alignment(vertical="top",wrap_text=True)
            row += 1
    ws.column_dimensions["A"].width=30; ws.column_dimensions["B"].width=18; ws.column_dimensions["C"].width=55
    ws.freeze_panes="A2"
    wb.save(dst)

def _manifest_path(report_id):
    return REPORTS_DIR / report_id / "manifest.json"


def create_report_snapshot(requested_by=None):
    now = datetime.now()
    report_id = now.strftime("%Y%m%d-%H%M%S") + "-" + secrets.token_hex(2)
    folder = REPORTS_DIR / report_id
    folder.mkdir(parents=True, exist_ok=False)
    files = []
    for name in SOURCE_FILES:
        src = BASE_DIR / name
        if name == "work_schedules.xlsx":
            _build_staff_schedule_xlsx(folder / name)
            files.append(name)
        elif src.exists():
            shutil.copy2(src, folder / name)
            files.append(name)

    # Отдельный сводный снимок нужен сайту: рекламная ссылка всегда отображается
    # отдельным блоком и хранит значение именно на момент создания отчёта.
    bot_data = {}
    data_path = BASE_DIR / "data.json"
    if data_path.exists():
        try:
            bot_data = json.loads(data_path.read_text(encoding="utf-8"))
        except Exception:
            bot_data = {}
    ad_link = bot_data.get("advertising_link") or ADVERTISING_LINK
    ad_info = bot_data.get("links", {}).get(ad_link, {})
    all_links = bot_data.get("links", {})
    users = bot_data.get("users", {})
    overview = {
        "advertising": {
            "label": "Рекламная ссылка",
            "link": ad_link,
            "invited": int(ad_info.get("invited", 0) or 0),
            "tracked_users": len(ad_info.get("users", []) or []),
        },
        "totals": {
            "users": len(users),
            "links": len(all_links),
            "referral_links": sum(1 for x in all_links.values() if x.get("type") != "advertising"),
            "invited": sum(int(x.get("invited", 0) or 0) for x in all_links.values()),
            "manual_adjustments": len(bot_data.get("invite_adjustments", []) or []),
        },
        "links": [
            {
                "link": link,
                "type": info.get("type", "referral"),
                "label": info.get("label", ""),
                "owner_id": info.get("owner_id"),
                "invited": int(info.get("invited", 0) or 0),
                "created": info.get("created", ""),
                "join_request": bool(info.get("join_request")),
            }
            for link, info in all_links.items()
        ],
    }
    (folder / "overview.json").write_text(json.dumps(overview, ensure_ascii=False, indent=2), encoding="utf-8")
    files.insert(0, "overview.json")

    manifest = {"id": report_id, "created_at": now.isoformat(timespec="seconds"), "requested_by": requested_by, "files": files}
    _manifest_path(report_id).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    # Не ждём фонового цикла: новый отчёт сразу зеркалируется в PostgreSQL.
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc:
        print(f"[DB] Не удалось немедленно сохранить отчёт: {exc}")
    return manifest


def list_reports():
    reports = []
    for folder in REPORTS_DIR.iterdir():
        if not folder.is_dir():
            continue
        mp = folder / "manifest.json"
        if mp.exists():
            try:
                reports.append(json.loads(mp.read_text(encoding="utf-8")))
            except Exception:
                pass
    return sorted(reports, key=lambda x: x.get("created_at", ""), reverse=True)


def _api_authorized(request):
    expected = os.getenv("SHINOBI_STATS_API_KEY", "").strip()
    if not expected:
        return False
    supplied = request.headers.get("X-Shinobi-Key", "") or request.query.get("key", "")
    return bool(supplied) and secrets.compare_digest(expected, supplied)


def _session_secret():
    return os.getenv("SESSION_SECRET", os.getenv("SHINOBI_STATS_API_KEY", "change-me-session-secret")).encode()


def _read_session(request):
    return hub_auth.session(request.cookies.get("shinobi_session", ""))

# GET/read operations available to staff; state-changing endpoints require elevated role.
WRITE_ROLES = {
    "/api/reports": {"owner", "deputy", "admin"},
    "/api/referrals/delete": {"owner", "deputy"},
    "/api/server-tags/assign": {"owner", "deputy", "admin"},
    "/api/control/group-access": {"owner", "deputy", "tech_admin"},
    "/api/control/server-tags": {"owner", "deputy", "tech_admin"},
    "/api/control/auto-reactions": {"owner", "deputy", "tech_admin"},
    "/api/market/wallet": {"owner"},
    "/api/backup/full": {"owner"},
    "/api/system/version": {"owner"},
    "/api/october-plan": {"owner", "tech_admin"},
    "/api/hub/schedule-live": {"owner", "deputy"},
    "/api/hub/dayoff-review": {"owner", "deputy"},
    "/api/tech/join-review": {"owner", "deputy", "tech_admin"},
    "/api/tech/invites": {"owner", "deputy"},
}

@web.middleware
async def auth_middleware(request, handler):
    public = {"/", "/api/status", "/media/banner", "/media/favicon", "/auth/telegram", "/auth/telegram/start", "/auth/telegram/callback"}
    if request.path in public:
        return await handler(request)
    session = _read_session(request)
    # API key is retained only for read-only server-to-server API calls.
    api_read = request.method == 'GET' and request.path.startswith('/api/') and _api_authorized(request)
    if not session and not api_read:
        if request.path.startswith('/api/'):
            return web.json_response({"error":"unauthorized"},status=401)
        raise web.HTTPFound('/')
    if request.method not in ('GET','HEAD','OPTIONS'):
        if not session or session['role'] not in WRITE_ROLES.get(request.path, {'owner'}):
            return web.json_response({"error":"forbidden: insufficient role"},status=403)
    if request.path.startswith('/api/backup') or request.path.startswith('/api/market/wallets'):
        if not session or session['role'] != 'owner':
            return web.json_response({"error":"forbidden"},status=403)
    if request.path in ('/api/login-audit','/api/revoke-user','/api/staff','/api/staff/role','/api/sessions','/api/sessions/revoke-others'):
        if not session or session['role'] != 'owner':
            return web.json_response({"error":"forbidden"},status=403)
    if request.method in ('POST', 'PUT', 'PATCH', 'DELETE') and request.path.startswith('/api/') and session:
        try:
            response = await handler(request)
            status = getattr(response, 'status', 200)
            hub_auth.audit(session['uid'], 'api_' + request.method.lower() + '_' + request.path.replace('/', '_') + '_status_' + str(status))
            return response
        except Exception:
            hub_auth.audit(session['uid'], 'api_' + request.method.lower() + '_' + request.path.replace('/', '_') + '_error')
            raise
    return await handler(request)

SETTINGS_FILE = BASE_DIR / "project_settings.json"
DEFAULT_SETTINGS = {
    "bot_version": "v1.0.0",
    "project_name": "Shinobi Team",
    "bot_name": "Shinobi Bot",
    "environment": "PRODUCTION",
    "platform": "MLBB",
}

def _load_settings():
    data = dict(DEFAULT_SETTINGS)
    try:
        if SETTINGS_FILE.exists():
            loaded = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                data.update({k: v for k, v in loaded.items() if isinstance(v, (str, int, float, bool))})
    except Exception as exc:
        print(f"[SETTINGS] Ошибка чтения: {exc}")
    # Платформа проекта фиксирована как MLBB. Это также перекрывает старое
    # значение Railway, которое могло восстановиться из PostgreSQL.
    data["platform"] = "MLBB"
    return data

def _save_settings(data):
    SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc:
        print(f"[DB] Не удалось сохранить настройки: {exc}")

async def status(request):
    st = _load_settings()
    return web.json_response({"ok": True, "service": "Shinobi Stats", "version": st.get("bot_version", "v1.0.0")})

async def system_info(request):
    st = _load_settings()
    db_connected = bool(os.getenv("DATABASE_URL", "").strip())
    return web.json_response({
        "settings": st,
        "database": "PostgreSQL" if db_connected else "Not configured",
        "database_connected": db_connected,
        "bot_status": "ONLINE",
        "role": (_read_session(request) or {}).get("role", ""),
    }, dumps=lambda x: json.dumps(x, ensure_ascii=False))

async def update_version(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "tech_admin"):
        return web.json_response({"error": "tech_only"}, status=403)
    try:
        payload = await request.json()
    except Exception:
        return web.json_response({"error": "invalid_json"}, status=400)
    version = str(payload.get("version", "")).strip()
    if not version or len(version) > 64:
        return web.json_response({"error": "invalid_version"}, status=400)
    st = _load_settings()
    st["bot_version"] = version
    st["version_updated_at"] = datetime.now().isoformat(timespec="seconds")
    _save_settings(st)
    return web.json_response({"ok": True, "version": version}, dumps=lambda x: json.dumps(x, ensure_ascii=False))


def _detect_image_ext(blob):
    if blob.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if blob.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if blob.startswith(b"RIFF") and blob[8:12] == b"WEBP":
        return ".webp"
    return None

def _media_file(kind):
    for ext in (".png", ".jpg", ".webp"):
        p = MEDIA_DIR / (kind + ext)
        if p.exists():
            return p
    return None

def _sync_media():
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc:
        print(f"[DB] Не удалось сохранить оформление: {exc}")

async def media_file(request):
    kind = request.match_info["kind"]
    if kind not in ("banner", "favicon"):
        raise web.HTTPNotFound()
    p = _media_file(kind)
    if not p:
        raise web.HTTPNotFound()
    return web.FileResponse(p)

async def upload_media(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "tech_admin"):
        return web.json_response({"error": "tech_only"}, status=403)
    kind = request.match_info["kind"]
    if kind not in ("banner", "favicon"):
        return web.json_response({"error": "invalid_kind"}, status=400)
    try:
        reader = await request.multipart()
        field = await reader.next()
        if not field or field.name != "image":
            return web.json_response({"error": "image_required"}, status=400)
        chunks, total = [], 0
        while True:
            chunk = await field.read_chunk()
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_IMAGE_BYTES:
                return web.json_response({"error": "too_large"}, status=413)
            chunks.append(chunk)
        blob = b"".join(chunks)
        ext = _detect_image_ext(blob)
        if not ext:
            return web.json_response({"error": "unsupported_image"}, status=415)
        for old_ext in (".png", ".jpg", ".webp"):
            old = MEDIA_DIR / (kind + old_ext)
            if old.exists():
                old.unlink()
        dst = MEDIA_DIR / (kind + ext)
        dst.write_bytes(blob)
        _sync_media()
        return web.json_response({"ok": True, "url": f"/media/{kind}?v={int(time.time())}"})
    except Exception as exc:
        print(f"[MEDIA] Ошибка загрузки: {exc}")
        return web.json_response({"error": "upload_failed"}, status=500)



OCTOBER_PLAN_FILE = BASE_DIR / "october_content_plan.json"

def _read_october_plan():
    try:
        return json.loads(OCTOBER_PLAN_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"month":"2026-10","items":[]}

def _write_october_plan(data):
    tmp = BASE_DIR / "october_content_plan.web.tmp"
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(OCTOBER_PLAN_FILE)

async def october_plan_api(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "tech_admin"):
        return web.json_response({"error":"tech_only"}, status=403)
    return web.json_response(_read_october_plan(), dumps=lambda x: json.dumps(x, ensure_ascii=False))

async def october_plan_update(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "tech_admin"):
        return web.json_response({"error":"tech_only"}, status=403)
    try:
        payload = await request.json(); date = str(payload.get("date", "")); done = bool(payload.get("done"))
    except Exception:
        return web.json_response({"error":"invalid_data"}, status=400)
    data = _read_october_plan(); found = False
    for item in data.get("items", []):
        if item.get("date") == date:
            item["done"] = done; found = True; break
    if not found: return web.json_response({"error":"not_found"}, status=404)
    _write_october_plan(data)
    return web.json_response({"ok":True,"date":date,"done":done})

LOGIN_HTML = r'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shinobi Stats • Authorization</title><link rel="icon" href="/media/favicon"><style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(circle at 50% 0,#d6edff 0,#eaf6ff 38%,#f1f9ff 100%);font-family:Arial,sans-serif;color:#ffffff;display:grid;place-items:center}.card{width:min(440px,92vw);background:rgba(17,17,19,.94);border:1px solid #d4e7f5;border-radius:22px;padding:34px;box-shadow:0 30px 90px #000,0 0 60px #e5091422}.mark{width:64px;height:64px;border-radius:18px;background:linear-gradient(145deg,#71c9f5,#248dc9);display:grid;place-items:center;font-weight:900;font-size:28px;margin:0 auto 18px;box-shadow:0 0 35px #e5091444}h1{text-align:center;margin:0;font-size:28px}.sub{text-align:center;color:#66859b;margin:8px 0 28px;font-size:12px;letter-spacing:1.5px}.roles{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:18px}.role{border:1px solid #d4e7f5;background:#171719;color:#597a94;padding:13px;border-radius:11px;cursor:pointer}.role.active{border-color:#2499dc;color:#ffffff;background:#e5091418}.label{font-size:12px;color:#64849b;margin:0 0 7px 2px}input{width:100%;background:#ffffff;border:1px solid #d4e7f5;color:#ffffff;padding:14px;border-radius:10px;outline:none;font-size:15px}input:focus{border-color:#2499dc}.btn{width:100%;border:0;border-radius:10px;padding:14px;margin-top:15px;background:linear-gradient(135deg,#55b8ee,#258bc7);color:#ffffff;font-weight:800;cursor:pointer}.err{background:#e5091415;border:1px solid #e5091444;color:#d14f61;padding:10px;border-radius:9px;margin-bottom:15px;font-size:13px}.foot{text-align:center;color:#444;font-size:11px;margin-top:22px}
/* SHINOBI TEAM HUB - ICE BLUE / WHITE THEME */
:root{color-scheme:light;--ice:#f1f9ff;--sky:#2499dc;--ink:#173b59;--line:#d7e8f5}
body{background:radial-gradient(ellipse at 15% 0%,#d7efff 0%,#f4fbff 45%,#ffffff 100%)!important;color:var(--ink)!important;font-family:Inter,Arial,sans-serif!important}
header,.panel,.card{background:rgba(255,255,255,.93)!important;border-color:#d7e8f5!important;box-shadow:0 10px 38px rgba(37,132,190,.09)}
.card{box-shadow:0 24px 65px rgba(37,132,190,.15)!important}
.brand span,.eyebrow{color:#218fd0!important}
input,select,textarea{background:#fff!important;color:#173b59!important;border-color:#c8e2f2!important}
input:focus,select:focus,textarea:focus{border-color:#2499dc!important;outline-color:#2499dc!important}
button,.btn,.tab.active,.navbtn.active{transition:background .2s,box-shadow .2s,transform .2s}
.btn,.tab.active,.navbtn.active{background:linear-gradient(135deg,#57b9ef,#218dd0)!important;color:#fff!important;border-color:#218dd0!important}
.btn:hover{box-shadow:0 6px 18px rgba(36,153,220,.22);transform:translateY(-1px)}
.role,.tab,.navbtn{background:#eef8ff!important;color:#315b79!important;border-color:#d2e7f6!important}
.role.active{background:#d9efff!important;color:#166eaa!important;border-color:#2499dc!important}
th{background:#dff1fc!important;color:#173b59!important;border-color:#cbe3f3!important}
td{color:#244c68!important;border-color:#e0edf6!important}
tr:hover td,.report:hover,.report.active{background:#edf8ff!important}
.ph{background:#f7fcff!important;color:#173b59!important}
.logout{color:#315b79!important;border-color:#cce4f4!important}
@media(max-width:700px){header{padding:12px 16px!important;height:auto;min-height:64px}.container{padding:18px 12px!important}}
</style></head><body><form class="card" method="post" action="/login"><div class="mark">S</div><h1>SHINOBI <span style="color:#2499dc">HUB</span></h1><div class="sub">INTERNAL ACCESS • AUTHORIZATION</div>__ERROR__<div class="roles"><button type="button" class="role active" data-role="owner">👑 Владелец</button><button type="button" class="role" data-role="tech">🛠 Тех. админ</button></div><input type="hidden" name="role" id="role" value="owner"><div class="label">ПАРОЛЬ ДОСТУПА</div><input type="password" name="password" placeholder="Введите пароль" autocomplete="current-password" required><button class="btn">ВОЙТИ В СИСТЕМУ</button><div class="foot">SHINOBI TEAM © 2026 • PROTECTED SYSTEM</div></form><script>document.querySelectorAll('.role').forEach(b=>b.onclick=()=>{document.querySelectorAll('.role').forEach(x=>x.classList.remove('active'));b.classList.add('active');document.getElementById('role').value=b.dataset.role})</script></body></html>'''


DASHBOARD_HTML = r'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shinobi Stats</title><link rel="icon" href="/media/favicon"><script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script><script src="https://cdn.jsdelivr.net/npm/chart.js"></script><style>
*{box-sizing:border-box}body{margin:0;background:#f3faff;color:#163c59;font-family:Arial,sans-serif}header{height:72px;padding:0 28px;display:flex;align-items:center;justify-content:space-between;background:#ffffff;border-bottom:1px solid #d4e7f5;position:sticky;top:0;z-index:10}.brand{font-weight:900;font-size:21px}.brand span{color:#2499dc}.who{font-size:12px;color:#64849b}.who b{color:#258bc7}.logout{color:#597a94;text-decoration:none;border:1px solid #c5deef;padding:8px 11px;border-radius:8px;margin-left:12px}main{max-width:1450px;margin:auto;padding:30px 22px}.eyebrow{color:#2499dc;font-size:11px;letter-spacing:2px;font-weight:bold}.title{font-size:32px;margin:7px 0 5px}.muted{color:#6b879d}.toolbar{display:flex;gap:12px;align-items:center;margin:25px 0}.btn{border:0;background:#2499dc;color:#ffffff;padding:11px 16px;border-radius:8px;font-weight:700;cursor:pointer}.grid{display:grid;grid-template-columns:300px 1fr;gap:18px}.panel{background:#ffffff;border:1px solid #d7e8f5;border-radius:14px;overflow:hidden}.ph{padding:15px 17px;border-bottom:1px solid #d7e8f5;font-weight:700}.reports{max-height:650px;overflow:auto}.report{padding:14px 16px;border-bottom:1px solid #d7e8f5;cursor:pointer}.report:hover,.report.active{background:#e5091410;border-left:3px solid #2499dc}.rid{font-size:12px;font-weight:bold}.rdate{font-size:11px;color:#6b879d;margin-top:5px}.content{min-height:500px}.tabs{display:flex;gap:7px;padding:12px;border-bottom:1px solid #d7e8f5;overflow:auto}.tab{background:#eaf6ff;color:#597a94;border:1px solid #d4e7f5;border-radius:8px;padding:9px 12px;cursor:pointer;white-space:nowrap}.tab.active{background:#2499dc;color:white;border-color:#2499dc}.tablewrap{overflow:auto;max-height:650px}table{border-collapse:collapse;width:100%;min-width:720px;font-size:13px}td,th{padding:11px 14px;border-bottom:1px solid #d7e8f5;border-right:1px solid #d7e8f5;white-space:nowrap;text-align:left}th{background:#258bc7;color:white;position:sticky;top:0}tr:hover td{background:#eaf6ff}.empty{padding:70px 20px;text-align:center;color:#6b879d}.status{font-size:12px;color:#66859b}.topnav{display:flex;gap:8px;margin:20px 0}.navbtn{background:#edf7ff;border:1px solid #d4e7f5;color:#597a94;padding:10px 14px;border-radius:9px;cursor:pointer;font-weight:700}.navbtn.active{background:#2499dc;color:#ffffff;border-color:#2499dc}.infoGrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px}.infoCard{background:#ffffff;border:1px solid #29292e;border-radius:14px;padding:18px}.infoLabel{font-size:11px;color:#66859b;letter-spacing:1px}.infoValue{font-size:22px;font-weight:900;margin-top:8px}.hero{background:linear-gradient(0deg,rgba(5,5,6,.70),rgba(5,5,6,.22)),url('/media/banner') center/cover no-repeat,linear-gradient(135deg,#24080b,#ffffff 55%,#190306);border:1px solid #4c1116;border-radius:18px;padding:28px;margin-bottom:18px;min-height:170px;display:flex;flex-direction:column;justify-content:flex-end}.hero h2{font-size:34px;margin:0 0 7px}.editrow{display:flex;gap:8px;margin-top:14px}.editrow input{flex:1;background:#0b0b0d;border:1px solid #c5deef;color:#ffffff;border-radius:8px;padding:11px}.mediaTools{margin-top:18px;padding:16px;background:#ffffff;border:1px solid #29292e;border-radius:14px}.mediaTools h3{margin:0 0 10px}.mediaRow{display:flex;gap:10px;flex-wrap:wrap;align-items:center}.mediaRow input[type=file]{max-width:320px;color:#597a94}.mediaPreview{width:100%;max-width:520px;max-height:180px;object-fit:cover;border-radius:12px;border:1px solid #c5deef;margin-top:12px}.page{display:none}.page.active{display:block}@media(max-width:800px){header{padding:0 14px}.grid{grid-template-columns:1fr}.reports{max-height:230px}.title{font-size:26px}.who{font-size:0}.who b,.logout{font-size:11px}main{padding:22px 12px}}
/* SHINOBI TEAM HUB - ICE BLUE / WHITE THEME */
:root{color-scheme:light;--ice:#f1f9ff;--sky:#2499dc;--ink:#173b59;--line:#d7e8f5}
body{background:radial-gradient(ellipse at 15% 0%,#d7efff 0%,#f4fbff 45%,#ffffff 100%)!important;color:var(--ink)!important;font-family:Inter,Arial,sans-serif!important}
header,.panel,.card{background:rgba(255,255,255,.93)!important;border-color:#d7e8f5!important;box-shadow:0 10px 38px rgba(37,132,190,.09)}
.card{box-shadow:0 24px 65px rgba(37,132,190,.15)!important}
.brand span,.eyebrow{color:#218fd0!important}
input,select,textarea{background:#fff!important;color:#173b59!important;border-color:#c8e2f2!important}
input:focus,select:focus,textarea:focus{border-color:#2499dc!important;outline-color:#2499dc!important}
button,.btn,.tab.active,.navbtn.active{transition:background .2s,box-shadow .2s,transform .2s}
.btn,.tab.active,.navbtn.active{background:linear-gradient(135deg,#57b9ef,#218dd0)!important;color:#fff!important;border-color:#218dd0!important}
.btn:hover{box-shadow:0 6px 18px rgba(36,153,220,.22);transform:translateY(-1px)}
.role,.tab,.navbtn{background:#eef8ff!important;color:#315b79!important;border-color:#d2e7f6!important}
.role.active{background:#d9efff!important;color:#166eaa!important;border-color:#2499dc!important}
th{background:#dff1fc!important;color:#173b59!important;border-color:#cbe3f3!important}
td{color:#244c68!important;border-color:#e0edf6!important}
tr:hover td,.report:hover,.report.active{background:#edf8ff!important}
.ph{background:#f7fcff!important;color:#173b59!important}
.logout{color:#315b79!important;border-color:#cce4f4!important}
@media(max-width:700px){header{padding:12px 16px!important;height:auto;min-height:64px}.container{padding:18px 12px!important}}

/* SHINOBI HUB SIDEBAR + CONTAINERS */
:root{--hub-side:252px;--hub-head:72px}
body{background:linear-gradient(125deg,#e8f5ff,#f9fcff 48%,#edf8ff)!important}
header{z-index:60!important;box-shadow:0 3px 20px #4c9bcc14!important}
main{max-width:none!important;width:auto!important;margin:0 0 0 var(--hub-side)!important;padding:30px clamp(18px,3vw,42px) 65px!important;min-height:calc(100vh - var(--hub-head))}
.topnav{position:fixed!important;left:0!important;top:var(--hub-head)!important;bottom:0!important;width:var(--hub-side)!important;display:flex!important;flex-direction:column!important;flex-wrap:nowrap!important;align-items:stretch!important;gap:5px!important;overflow-y:auto!important;overflow-x:hidden!important;margin:0!important;padding:64px 12px 22px!important;background:linear-gradient(#fff,#f0f8ff)!important;border-right:1px solid #d4e7f5!important;z-index:40!important}
.topnav:before{content:'РАЗДЕЛЫ СИСТЕМЫ';position:absolute;top:23px;left:22px;font-size:10px;letter-spacing:1.7px;font-weight:800;color:#84a5ba}
.topnav .navbtn{width:100%!important;min-height:44px!important;flex:0 0 auto!important;display:flex!important;align-items:center!important;justify-content:flex-start!important;text-align:left!important;white-space:normal!important;line-height:1.4!important;padding:11px 13px!important;border:1px solid transparent!important;border-radius:11px!important;background:transparent!important;color:#466d87!important;font-size:13px!important;box-shadow:none!important}
.topnav .navbtn:hover{background:#e8f5ff!important;color:#126fae!important}
.topnav .navbtn.active{background:#d9efff!important;color:#136ca7!important;border-color:#b5ddf6!important;box-shadow:inset 3px 0 0 #319cda!important}
.page{min-width:0}
.hero,.panel,.infoCard,.mediaTools{background:#ffffffee!important;border:1px solid #dcecf7!important;border-radius:17px!important;box-shadow:0 8px 26px #2c76a90e!important}
.hero{padding:24px!important;margin-bottom:20px!important}
.infoGrid{gap:14px!important}.infoCard{padding:19px!important}.mediaTools{padding:21px!important;margin-bottom:17px!important}.panel{overflow:hidden}.grid{gap:17px!important}.tablewrap,.mediaTools{max-width:100%;overflow-x:auto}
.hub-menu-toggle{display:none;background:#e9f6ff;color:#176fa8;border:1px solid #c8e5f6;border-radius:10px;font-size:20px;width:40px;height:40px;cursor:pointer}.hub-overlay{display:none}
@media(max-width:900px){header{padding:0 15px!important}.hub-menu-toggle{display:inline-flex;align-items:center;justify-content:center}header .brand{margin-right:auto;margin-left:10px}main{margin-left:0!important;padding:20px 14px 55px!important}.topnav{transform:translateX(-105%);transition:transform .2s ease;z-index:55!important}body.hub-menu-open .topnav{transform:translateX(0)}.hub-overlay{position:fixed;inset:var(--hub-head) 0 0;background:#183b5947;z-index:45}body.hub-menu-open .hub-overlay{display:block}}
@media(max-width:550px){.infoGrid{grid-template-columns:repeat(2,minmax(0,1fr))!important}.infoValue{font-size:clamp(17px,5vw,26px)!important;overflow-wrap:anywhere}.grid{grid-template-columns:1fr!important}.hero{padding:18px!important}.mediaTools{padding:15px!important}.page .mediaRow{flex-wrap:wrap}}

/* SHINOBI TEAM HUB 2026 — GLASS DESIGN SYSTEM */
:root{--glass-blue:#2a9ddd;--glass-ink:#173d5a;--glass-line:rgba(130,190,226,.34);--glass-shadow:0 14px 44px rgba(38,111,163,.10);--glass-radius:22px}
html{scroll-behavior:smooth}body{font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif!important;background:radial-gradient(ellipse at 78% -12%,#d7edff 0,transparent 44%),radial-gradient(ellipse at 15% 40%,#edf8ff 0,transparent 60%),#f6fbff!important;color:var(--glass-ink)!important}
header{height:76px!important;background:rgba(255,255,255,.79)!important;backdrop-filter:blur(22px) saturate(155%);-webkit-backdrop-filter:blur(22px) saturate(155%);border-bottom:1px solid rgba(255,255,255,.92)!important;padding-inline:30px!important}
.brand{letter-spacing:.035em;font-size:20px!important}.brand span{background:linear-gradient(125deg,#1494da,#66b9ee);background-clip:text;-webkit-background-clip:text;color:transparent!important}
.topnav{top:76px!important;background:rgba(255,255,255,.65)!important;backdrop-filter:blur(24px) saturate(145%);-webkit-backdrop-filter:blur(24px) saturate(145%);border-right:1px solid rgba(160,206,232,.5)!important;padding:60px 13px 24px!important;box-shadow:8px 0 35px rgba(74,149,192,.04)}
.topnav .navbtn{border:1px solid transparent!important;border-radius:13px!important;background:transparent!important;font-size:13px!important;font-weight:650!important;color:#52758e!important;transition:transform .18s ease,background .18s ease,box-shadow .18s ease,color .18s ease!important;gap:12px!important}
.topnav .navbtn:hover{transform:translateX(3px);background:rgba(220,242,255,.9)!important;color:#137db9!important}
.topnav .navbtn.active{background:linear-gradient(120deg,#e1f3ff,#d2eaff)!important;color:#126eaa!important;border-color:#c3e5fb!important;box-shadow:0 5px 18px rgba(44,145,204,.11)!important}
.hub-nav-icon{display:inline-flex;flex:0 0 19px;width:19px;height:19px;align-items:center;justify-content:center}.hub-nav-icon svg{width:18px;height:18px;stroke:currentColor;stroke-width:1.8;fill:none;stroke-linecap:round;stroke-linejoin:round}
main{padding-top:34px!important}.eyebrow{font-weight:800!important}.title{font-size:clamp(27px,3vw,38px)!important;letter-spacing:-.045em;font-weight:800!important}.muted,.status{color:#7896aa!important}
.hero{background:linear-gradient(100deg,rgba(16,68,106,.74),rgba(29,107,163,.22)),url('/media/banner') center/cover no-repeat!important;color:white!important;border:1px solid #ffffff99!important;border-radius:26px!important;min-height:190px!important;box-shadow:var(--glass-shadow)!important;overflow:hidden}
.hero h2{color:white!important;letter-spacing:-.035em}
.panel,.infoCard,.mediaTools{background:rgba(255,255,255,.78)!important;backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,.98)!important;outline:1px solid rgba(162,209,235,.22);border-radius:var(--glass-radius)!important;box-shadow:var(--glass-shadow)!important;transition:box-shadow .2s ease,transform .2s ease!important}
.infoCard:hover{transform:translateY(-3px);box-shadow:0 18px 38px rgba(38,111,163,.15)!important}.infoValue{color:#1a80ba!important;letter-spacing:-.035em}.infoLabel{color:#7899ae!important;font-weight:750}
.ph{background:rgba(235,247,255,.55)!important;border-bottom:1px solid #e2f0f8!important;padding:18px 20px!important}
.btn,.tab,.logout,.hub-menu-toggle{border-radius:12px!important;transition:transform .2s ease,box-shadow .2s ease,background .2s ease!important}.btn{padding:12px 18px!important;background:linear-gradient(135deg,#4eb5ed,#1c8fcd)!important;box-shadow:0 6px 16px #319ad329!important}.btn:hover{transform:translateY(-2px)!important;box-shadow:0 10px 24px #319ad33b!important}.btn:active,.tab:active,.navbtn:active{transform:scale(.975)!important}
button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible{outline:3px solid #65c4f3!important;outline-offset:3px!important}
input,select,textarea{border-radius:11px!important;min-height:38px}.tab{font-weight:650}.tab:hover{background:#d8efff!important}.tablewrap{border-radius:0 0 20px 20px}th{background:#e8f5fd!important;color:#28678e!important;font-weight:750!important}td{color:#375e76!important}tr:hover td{background:#f1faff!important}
.report{transition:background .2s ease,padding-left .2s ease}.report:hover{padding-left:20px;background:#eaf7ff!important}.logout:hover{background:#e8f5ff!important;color:#167cb5!important}
#hub-toast-stack{position:fixed;right:22px;bottom:22px;z-index:9999;display:flex;flex-direction:column;gap:10px;pointer-events:none}.hub-toast{padding:13px 17px;background:rgba(255,255,255,.95);border:1px solid #c8e6f6;border-radius:14px;color:#235576;box-shadow:0 14px 36px #286c9a2c;backdrop-filter:blur(15px);font-size:13px;font-weight:650;max-width:340px;animation:hub-toast-in .25s ease both}.hub-toast.error{border-color:#ffd3d3;color:#a34343}@keyframes hub-toast-in{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:translateY(0)}}
@media(max-width:900px){.topnav{top:76px!important}header{padding:0 14px!important}main{padding:20px 14px 60px!important}}@media(max-width:550px){.title{font-size:28px!important}.hero{min-height:160px!important}.infoCard{padding:15px!important}.topnav{width:min(290px,86vw)!important}}
@media(prefers-reduced-motion:reduce){*,*:before,*:after{animation-duration:.01ms!important;transition-duration:.01ms!important;scroll-behavior:auto!important}}

/* V4 — searchable glass data workspaces */
.hub-table-toolbar{display:flex;flex-wrap:wrap;align-items:end;gap:12px;margin:14px 0 18px;padding:15px;background:rgba(232,247,255,.65);border:1px solid rgba(162,208,235,.5);border-radius:17px}
.hub-search-label{flex:1 1 230px;display:flex;flex-direction:column;gap:6px;color:#497995;font-size:12px;font-weight:750}
.hub-search-label input{width:100%;background:#fff!important;border:1px solid #cde5f5!important;padding:12px 14px;min-height:43px}
.hub-secondary{background:rgba(255,255,255,.9)!important;color:#1d80b8!important;border:1px solid #b9e0f5!important;box-shadow:none!important}
.hub-secondary:hover{background:#e4f4ff!important}
.hub-filter-hidden{display:none!important}
#growthPage .mediaTools,#aiModPage .mediaTools{overflow-x:auto}
#growthPage td,#aiModPage td{vertical-align:middle}
#growthPage tbody tr,#aiModPage tbody tr{transition:background .16s ease}
@media(max-width:650px){.hub-table-toolbar{padding:12px}.hub-table-toolbar button{flex:1 1 120px}}
</style></head><body><header><button type="button" class="hub-menu-toggle" aria-label="Открыть разделы" aria-expanded="false" onclick="toggleHubMenu()">☰</button><div class="brand">SHINOBI <span>STATS</span></div><div class="who"><b>__ROLE__</b> • AUTHORIZED <a class="logout" href="/logout">Выйти</a></div></header><div class="hub-overlay" onclick="closeHubMenu()"></div><main><div class="eyebrow">SHINOBI TEAM • INTERNAL SYSTEM</div><div class="title">Контроль сотрудников</div><div class="muted">Отчёты и архив администрации проекта • BUILD 2026.09.27.2</div><div class="topnav"><button class="navbtn active" onclick="showPage('hubOverview',this);hubOverviewLoad()">Обзор</button><button class="navbtn" onclick="showPage('hubStaff',this);hubStaffLoad()">Сотрудники</button><button class="navbtn" onclick="showPage('hubSecurity',this);hubSecurityLoad()">Безопасность</button><button class="navbtn" onclick="showPage('hubSchedule',this);hubScheduleLoad();hubLiveLoad();hubDayoffLoad()">График и выходные</button><button class="navbtn active" onclick="showPage('reportsPage',this)">📊 Отчёты</button><button class="navbtn" onclick="showPage('projectPage',this)">🥷 О проекте</button><button class="navbtn" onclick="showPage('botPage',this)">🤖 О боте</button><button class="navbtn" onclick="showPage('controlPage',this);loadControl()">🎛 Управление</button><button class="navbtn" onclick="showPage('changesPage',this);loadChanges()">🗓 История изменений</button><button class="navbtn" onclick="showPage('chatPage',this);loadChatSummary()">💬 Сводка чатов</button><button class="navbtn" onclick="showPage('growthPage',this);loadGrowth()">📈 Приглашения и рост</button>__TECH_NAV__<button class="navbtn" onclick="showPage('aiModPage',this);loadAiModeration()">🧪 AI-модерация <b>BETA</b></button></div><div id="aiModPage" class="page"><div class="hero"><div class="eyebrow">AI MODERATION • BETA</div><h2>🧪 Журнал модерации <span style="font-size:15px;border:1px solid #e50914;padding:4px 8px;border-radius:8px">BETA</span></h2><div class="muted">AI и ручные наказания Staff в единой системе Shinobi Moderation.</div></div><div class="infoGrid"><div class="infoCard"><div class="infoLabel">СРАБАТЫВАНИЙ В ЖУРНАЛЕ</div><div class="infoValue" id="aiModCount">—</div></div><div class="infoCard"><div class="infoLabel">СТАТУС</div><div class="infoValue">🧪 BETA</div></div></div><div class="mediaTools"><h3>🛡 Последние нарушения • BETA</h3><div class="status" style="margin-bottom:12px">Показываются последние 500 применённых решений AI-модерации.</div><div class="hub-table-toolbar"><label class="hub-search-label">Поиск по нарушениям <input id="hubModFilter" type="search" placeholder="Участник, правило, причина" oninput="hubFilterTable('aiModLog',this.value)"></label><button class="btn hub-secondary" onclick="hubExportTable('aiModLog','shinobi-moderation.csv')">Экспорт CSV</button><button class="btn hub-secondary" onclick="loadAiModeration()">Обновить</button></div><div id="aiModLog"><div class="empty">Загрузка…</div></div></div></div><div id="growthPage" class="page"><div class="hero"><div class="eyebrow">SHINOBI ANALYTICS</div><h2>📈 Приглашения и рост чата</h2><div class="muted">Рейтинг реферальных ссылок, дневной прирост и история входов в основной чат.</div></div><div class="infoGrid"><div class="infoCard"><div class="infoLabel">ВОШЛО СЕГОДНЯ</div><div class="infoValue" id="joinsToday">—</div></div><div class="infoCard"><div class="infoLabel">ВОШЛО ВЧЕРА</div><div class="infoValue" id="joinsYesterday">—</div></div><div class="infoCard"><div class="infoLabel">ПРИРОСТ К ВЧЕРА</div><div class="infoValue" id="joinsDelta">—</div></div><div class="infoCard"><div class="infoLabel">ЗА 30 ДНЕЙ</div><div class="infoValue" id="joinsPeriod">—</div></div></div><div class="mediaTools"><h3>📊 График входов за 30 дней</h3><div style="height:340px"><canvas id="joinsChart"></canvas></div></div><div class="mediaTools"><h3>🏆 Ссылки по количеству приглашённых</h3><div class="status" style="margin-bottom:12px">Сортировка от большего к меньшему. Персональную ссылку можно удалить прямо здесь.</div><div class="hub-table-toolbar"><label class="hub-search-label">Поиск по ссылкам <input id="hubReferralFilter" type="search" placeholder="Имя, @username или ссылка" oninput="hubFilterTable('referralRanking',this.value)"></label><button class="btn hub-secondary" onclick="hubExportTable('referralRanking','shinobi-referrals.csv')">Экспорт CSV</button><button class="btn hub-secondary" onclick="loadGrowth()">Обновить</button></div><div id="referralRanking"><div class="empty">Загрузка…</div></div></div></div><div id="octoberPage" class="page"><div class="hero"><div class="eyebrow">TECH ADMIN • OCTOBER 2026</div><h2>📅 План на октябрь</h2><div class="muted">31 день контента MLBB. Отмечай выполненные публикации галочкой.</div></div><div class="panel"><div class="ph">Контент-план • Октябрь 2026</div><div id="octoberPlanBox"><div class="empty">Загрузка…</div></div></div></div><section class="page" id="techOpsPage"><div class="eyebrow">SHINOBI TEAM / TECHNICAL OPERATIONS</div><h1 class="title">🛠 Техническая панель</h1><p class="muted">Диагностика, группы и управление доступными настройками. Права проверяются на сервере.</p><div class="infoGrid"><div class="infoCard"><div class="infoLabel">ПОЛЬЗОВАТЕЛЕЙ</div><div class="infoValue" id="techUsers">—</div></div><div class="infoCard"><div class="infoLabel">ГРУПП</div><div class="infoValue" id="techGroups">—</div></div><div class="infoCard"><div class="infoLabel">ТЕГОВ</div><div class="infoValue" id="techTags">—</div></div></div><div class="panel"><div class="ph">Управление ботом</div><p class="muted">Настройки групп, тегов и реакций доступны в существующем разделе «Управление». Сохраняются через API, с проверкой роли.</p><div class="hub-quick"><button class="btn" onclick="hubGo('controlPage');loadControl()">Открыть управление</button><button class="btn hub-secondary" onclick="techOpsLoad()">Обновить показатели</button><button class="btn hub-secondary" onclick="hubGo('chatPage');loadChatSummary()">Сводка чатов</button><button class="btn hub-secondary" onclick="hubGo('aiModPage');loadAiModeration()">Журнал модерации</button></div><p class="status" id="techOpsStatus"></p></div>
<div class="panel"><div class="ph">📩 Заявки на вступление</div><p class="muted">Одобрение и отклонение выполняются через Telegram API. Уже обработанные заявки защищены от повторного решения.</p><button class="btn hub-secondary" onclick="techJoinLoad()">Обновить заявки</button><div id="techJoinList" style="margin-top:12px"></div></div>
<div class="panel"><div class="ph">👥 Пользователи и приглашения</div><p class="muted">Рейтинг пользователей и корректировка числа приглашений (только владелец и заместитель).</p><input id="techUserFilter" placeholder="Поиск по имени, ID, username" oninput="techUserRender()" style="padding:10px;width:100%;max-width:420px"><button class="btn hub-secondary" onclick="techUsersLoad()">Обновить пользователей</button><div id="techUserList" style="margin-top:12px"></div></div>
<div class="panel"><div class="ph">📁 Excel и отчёты</div><div class="hub-quick"><a class="btn hub-secondary" href="/api/tech/export/referrals">Ссылки Excel</a><a class="btn hub-secondary" href="/api/tech/export/schedule">Графики Excel</a><a class="btn hub-secondary" href="/api/tech/export/dayoffs">Выходные Excel</a><button class="btn" onclick="hubGo('reportsPage')">Все отчёты</button><button class="btn hub-secondary" onclick="hubGo('growthPage');loadGrowth()">Реферальные ссылки</button></div></div></section><section class="page active" id="hubOverview"><div class="eyebrow">SHINOBI TEAM / CONTROL CENTER</div><h1 class="title">Панель управления</h1><p class="muted">Живые данные из существующих модулей Shinobi Team Hub</p><div class="infoGrid" id="hubOverviewStats"><div class="infoCard">Загружаем показатели…</div></div><div class="hub-feature-grid"><div class="panel"><div class="ph">Быстрые действия</div><div class="hub-quick"><button class="btn" onclick="hubGo('growthPage')">Приглашения</button><button class="btn" onclick="hubGo('controlPage')">Управление ботом</button><button class="btn" onclick="hubGo('chatPage')">Сводка чатов</button><button class="btn" onclick="hubGo('reportsPage')">Отчёты и графики</button></div></div><div class="panel"><div class="ph">Безопасность</div><div class="hub-quick"><p class="muted">Авторизация через Telegram, роли на сервере, выход после 30 минут бездействия.</p><a class="btn" href="/logout" style="text-decoration:none;display:inline-block">Завершить сессию</a></div></div></div></section>
<section class="page" id="hubSecurity"><div class="eyebrow">SECURITY CENTER</div><h1 class="title">Безопасность и сессии</h1><p class="muted">Только владелец. История авторизации и активные сессии сотрудников.</p><div class="infoGrid"><div class="infoCard"><div class="infoLabel">СЕССИИ</div><div class="infoValue" id="hubSessionCount">—</div></div><div class="infoCard"><div class="infoLabel">СОБЫТИЯ АВТОРИЗАЦИИ</div><div class="infoValue" id="hubAuditCount">—</div></div></div><div class="mediaTools"><h3>Активные сессии</h3><p class="status">Сессии автоматически истекают после 30 минут бездействия, максимум через 12 часов. Завершение сессии не выходит из Telegram.</p><div class="mediaRow"><button class="btn" onclick="hubSecurityLoad()">Обновить</button><button class="navbtn" onclick="hubRevokeAll()">Завершить все чужие сессии</button></div><div class="tablewrap" id="hubSessionRows" style="margin-top:16px">Загрузка…</div></div><div class="mediaTools"><h3>Журнал входов</h3><div class="mediaRow"><input id="hubAuditSearch" placeholder="Поиск по ID или событию" oninput="hubRenderAudit()" aria-label="Поиск в журнале"></div><div class="tablewrap" id="hubAuditRows" style="margin-top:16px">Загрузка…</div></div></section><section class="page" id="hubStaff"><div class="eyebrow">КОМАНДА / ДОСТУП</div><h1 class="title">Сотрудники</h1><p class="muted">Владелец может назначать роли, отключать доступ и завершать активные сессии. Изменения сохраняются в PostgreSQL.</p><div class="panel"><div class="ph">Добавить или изменить сотрудника</div><div class="hub-quick"><div class="mediaRow"><input id="hubStaffUid" inputmode="numeric" placeholder="Telegram ID сотрудника"><select id="hubStaffRole"><option value="moderator">Модератор</option><option value="admin">Администратор</option><option value="tech_admin">Технический администратор</option><option value="deputy">Заместитель</option><option value="disabled">Отключить доступ</option></select><button class="btn" onclick="hubStaffSave()">Сохранить</button></div><div id="hubStaffMsg" class="status"></div></div></div><div class="panel" style="margin-top:18px"><div class="ph">Список сотрудников</div><div id="hubStaffRows" class="tablewrap"><div class="empty">Загрузка…</div></div></div></section>
<section class="page" id="hubSchedule"><div class="eyebrow">STAFF OPERATIONS</div><h1 class="title">График и выходные</h1><p class="muted">Живой график и заявки из общей PostgreSQL. Ниже также доступен архив XLSX-снимков для отчётности.</p><div class="infoGrid"><div class="infoCard"><div class="infoLabel">ЗАПИСЕЙ ГРАФИКА</div><div class="infoValue" id="hubScheduleCount">—</div></div><div class="infoCard"><div class="infoLabel">ЗАЯВОК НА ВЫХОДНОЙ</div><div class="infoValue" id="hubDaysOffCount">—</div></div></div><div class="panel" style="margin-top:18px"><div class="ph">Живой график обязанностей · PostgreSQL</div><p class="muted">Изменения сохраняются в общей базе и используются Telegram-ботом. Редактирование доступно владельцу и заместителю.</p><div class="mediaRow"><select id="hubLiveUser" onchange="hubLiveRender()" aria-label="Сотрудник"></select><select id="hubLiveDay" onchange="hubLiveRender()" aria-label="День недели"></select><label><input type="checkbox" id="hubLiveOff"> Выходной</label></div><div id="hubLiveDuties" class="hub-feature-grid" style="margin:12px 0"></div><div class="mediaRow"><button class="btn" onclick="hubLiveSave()">Сохранить график</button><button class="btn hub-secondary" onclick="hubLiveLoad()">Обновить данные</button></div><p class="status" id="hubLiveStatus"></p></div><div class="panel" style="margin-top:18px"><div class="ph">График сотрудников</div><div class="hub-table-toolbar"><label class="hub-search-label">Поиск по сотруднику или дню<input id="hubScheduleFilter" placeholder="Имя, ID, день недели…" oninput="hubFilterTable('hubScheduleRows',this.value)"></label><button class="btn hub-secondary" onclick="hubScheduleLoad()">Обновить</button><button class="btn hub-secondary" onclick="hubExportTable('hubScheduleRows','shinobi-schedule.csv')">Экспорт CSV</button></div><div id="hubScheduleRows" class="tablewrap"><div class="empty">Откройте раздел для загрузки</div></div></div><div class="panel" style="margin-top:18px"><div class="ph">Рассмотрение заявок · PostgreSQL</div><p class="muted">Одобрение или отклонение доступно владельцу и заместителю. Повторное решение по одной заявке блокируется.</p><div class="mediaRow"><button class="btn hub-secondary" onclick="hubDayoffLoad()">Обновить заявки</button></div><div id="hubDayoffReviewRows" class="tablewrap"><div class="empty">Нажмите «Обновить заявки»</div></div><p id="hubDayoffReviewStatus" class="status"></p></div><div class="panel" style="margin-top:18px"><div class="ph">Заявки на выходные</div><div class="hub-table-toolbar"><label class="hub-search-label">Поиск по заявкам<input id="hubDaysOffFilter" placeholder="Имя, статус, дата…" oninput="hubFilterTable('hubDaysOffRows',this.value)"></label><button class="btn hub-secondary" onclick="hubExportTable('hubDaysOffRows','shinobi-days-off.csv')">Экспорт CSV</button></div><div id="hubDaysOffRows" class="tablewrap"><div class="empty">Откройте раздел для загрузки</div></div></div><p class="status" id="hubScheduleUpdated"></p></section><div id="reportsPage" class="page active"><div class="toolbar"><button class="btn" id="newReportBtn" onclick="createNewReport()">＋ Новый отчёт</button><button class="btn" onclick="loadReports()">↻ Обновить архив</button><span class="status" id="status">Готово</span></div><div class="grid"><section class="panel"><div class="ph">Архив отчётов</div><div class="reports" id="reports"><div class="empty">Загрузка…</div></div></section><section class="panel content"><div class="ph" id="reportTitle">Выберите отчёт</div><div class="tabs" id="tabs"></div><div class="tablewrap" id="table"><div class="empty">Выберите отчёт слева</div></div></section></div></div><div id="projectPage" class="page"><div class="hero"><div class="eyebrow">SHINOBI TEAM</div><h2>О проекте</h2><div class="muted">Внутренняя система управления, отчётности и контроля Shinobi Team.</div></div><div class="infoGrid"><div class="infoCard"><div class="infoLabel">ПРОЕКТ</div><div class="infoValue" id="projectName">Shinobi Team</div></div><div class="infoCard"><div class="infoLabel">ПЛАТФОРМА</div><div class="infoValue" id="platform">MLBB</div></div><div class="infoCard"><div class="infoLabel">СРЕДА</div><div class="infoValue" id="environment">PRODUCTION</div></div><div class="infoCard"><div class="infoLabel">БАЗА ДАННЫХ</div><div class="infoValue" id="database">—</div></div></div><div id="mediaEditor"></div></div><div id="botPage" class="page"><div class="hero"><div class="eyebrow">SHINOBI SYSTEM</div><h2>О боте</h2><div class="muted">Техническая информация о текущей версии Telegram-бота.</div></div><div class="infoGrid"><div class="infoCard"><div class="infoLabel">BOT STATUS</div><div class="infoValue" style="color:#43d17a" id="botStatus">● ONLINE</div></div><div class="infoCard"><div class="infoLabel">BOT VERSION</div><div class="infoValue" id="botVersion">—</div><div id="versionEditor"></div></div><div class="infoCard"><div class="infoLabel">DATABASE</div><div class="infoValue" id="botDatabase">—</div></div><div class="infoCard"><div class="infoLabel">ACCESS ROLE</div><div class="infoValue">__ROLE__</div></div></div></div><div id="controlPage" class="page"><div class="hero"><div class="eyebrow">TELEGRAM WEB APP</div><h2>Управление ботом</h2><div class="muted">Состояние Telegram-функций, группы, пользователи и серверные теги.</div></div><div class="infoGrid"><div class="infoCard"><div class="infoLabel">ПОЛЬЗОВАТЕЛЕЙ</div><div class="infoValue" id="ctlUsers">—</div></div><div class="infoCard"><div class="infoLabel">ГРУПП</div><div class="infoValue" id="ctlGroups">—</div></div><div class="infoCard"><div class="infoLabel">ВЫДАНО ТЕГОВ</div><div class="infoValue" id="ctlTags">—</div></div><div class="infoCard"><div class="infoLabel">ТЕХ. РАБОТЫ</div><div class="infoValue" id="ctlMaint">—</div></div></div><div class="mediaTools"><h3>🎲 Автоматические реакции</h3><div class="status">Мгновенные случайные реакции на новые сообщения в основном чате. Переключение применяется без перезапуска бота (до 2 секунд).</div><div class="mediaRow" style="margin-top:12px"><button class="btn" id="autoReactionsToggle" onclick="toggleAutoReactions()">Загрузка…</button><span class="status" id="autoReactionsState"></span></div></div><div class="mediaTools"><h3>⚡ Автовыдача серверных тегов</h3><div class="status">Если выключено, бот не спрашивает сервер у новых участников, не реагирует на ответы и реакции для выдачи тегов.</div><div class="mediaRow" style="margin-top:12px"><button class="btn" id="serverTagsToggle" onclick="toggleServerTags()">Загрузка…</button><span class="status" id="serverTagsState"></span></div></div><div class="mediaTools"><h3>🏷 Серверные теги</h3><div class="status">Ручная выдача использует Telegram setChatMemberTag. Боту нужно право can_manage_tags.</div><div class="mediaRow" style="margin-top:12px"><input id="tagChat" placeholder="Chat ID"><input id="tagUser" placeholder="User ID"><select id="tagValue" style="background:#0d0d0f;color:#fff;border:1px solid #333;padding:11px;border-radius:8px"></select><button class="btn" onclick="assignTagWeb()">Выдать тег</button></div><div class="status" id="tagStatus" style="margin-top:10px">Готово</div><div id="tagHistory" style="margin-top:14px;overflow:auto"></div></div><div class="mediaTools"><h3>👤 Участники без тега</h3><div class="status">Проверяет участников ОСНОВНОГО чата, которых бот уже знает. Администраторы исключаются автоматически — их специальные теги не учитываются. Telegram Bot API не отдаёт боту полный список всех участников группы.</div><div class="mediaRow" style="margin-top:12px"><button class="btn" onclick="loadUntagged()">Обновить список</button><button class="navbtn" onclick="copyUntagged()">Копировать</button><span class="status" id="untaggedStatus">Нажмите «Обновить список»</span></div><textarea id="untaggedList" readonly style="width:100%;min-height:180px;margin-top:12px;background:#0b0b0d;color:#eee;border:1px solid #333;border-radius:8px;padding:12px;resize:vertical"></textarea></div><div class="mediaTools"><h3>💾 Полный бэкап бота</h3><div class="status">Собирает ZIP в момент запроса: актуальные файлы бота, локальные данные, отчёты, медиа и полный экспорт PostgreSQL. Для скачивания требуется отдельный резервный код.</div><div class="mediaRow" style="margin-top:12px"><button class="btn" id="backupBtn" onclick="downloadFullBackup()">Скачать актуальную версию</button><span class="status" id="backupStatus">Готово</span></div></div><div class="mediaTools" id="marketWalletPanel"><h3>💎 Монеты Чёрного рынка</h3><div class="status">Доступно владельцу и тех. админу. Выбери игрока и добавь ему виртуальные монеты либо установи точный баланс.</div><div class="mediaRow" style="margin-top:12px"><select id="marketPlayer" style="min-width:320px;background:#0d0d0f;color:#fff;border:1px solid #333;padding:11px;border-radius:8px"></select><input id="marketAmount" type="number" min="0" placeholder="Количество монет"><select id="marketMode" style="background:#0d0d0f;color:#fff;border:1px solid #333;padding:11px;border-radius:8px"><option value="add">Добавить к балансу</option><option value="set">Установить баланс</option></select><button class="btn" onclick="updateMarketWallet()">Применить</button></div><div class="status" id="marketWalletStatus" style="margin-top:10px">Загрузка игроков…</div></div><div class="mediaTools"><h3>🏘 Группы и доступ к командам</h3><div class="status" style="margin-bottom:10px">Переключатель действует на все команды бота, включая игровые. Новые группы по умолчанию получают доступ.</div><div id="ctlGroupList" class="status">Загрузка…</div></div><div class="mediaTools"><h3>👥 Пользователи бота</h3><div id="ctlUserList" class="status">Загрузка…</div></div></div><div id="changesPage" class="page"><div class="hero"><div class="eyebrow">SHINOBI CHANGELOG</div><h2>История изменений</h2><div class="muted">Что менялось в боте и веб-панели — по датам.</div></div><div id="changesList"><div class="empty">Загрузка…</div></div></div><div id="chatPage" class="page"><div class="hero"><div class="eyebrow">LIVE CHAT ANALYTICS</div><h2>Сводка чатов</h2><div class="muted">Статистика собирается ботом отдельно для каждого установленного чата.</div></div><div id="chatSummaryBox" class="infoGrid"><div class="empty">Загрузка…</div></div><div class="status" style="margin-top:14px">🟢 «Активны» = уникальные пользователи, от которых бот видел сообщения за последние 15 минут. Telegram Bot API не предоставляет точный список online-пользователей группы.</div></div></main><script>
let reports=[],current=null;const labels={'overview.json':'Обзор','shinobi_links_log.xlsx':'Результаты и приглашения','work_schedules.xlsx':'График сотрудников','day_off_requests.xlsx':'Заявки на выходные','data.json':'Полная база бота','admin_assignments.json':'Данные обязанностей','user_timezones.json':'Часовые пояса'};

let techUsersCache=[];
function techEsc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
async function techUsersLoad(){const box=document.getElementById('techUserList');box.textContent='Загрузка…';try{const r=await fetch('/api/tech/users');if(!r.ok)throw Error('HTTP '+r.status);const j=await r.json();techUsersCache=j.users||[];techUserRender();}catch(e){box.textContent='Ошибка: '+e.message}}
function techUserRender(){
 const box=document.getElementById('techUserList');if(!box)return;
 const q=(document.getElementById('techUserFilter')?.value||'').toLowerCase();
 const users=techUsersCache.filter(u=>[u.name,u.username,u.id].some(x=>String(x||'').toLowerCase().includes(q))).slice(0,100);
 if(!users.length){box.textContent='Пользователи не найдены';return;}
 const table=document.createElement('table');table.innerHTML='<thead><tr><th>Участник</th><th>ID</th><th>Приглашено</th><th>Действие</th></tr></thead>';
 const tbody=document.createElement('tbody');
 users.forEach(u=>{
  const tr=document.createElement('tr');
  for(const value of [String(u.name||'')+' @'+String(u.username||''),String(u.id),String(u.invited)]){
   const td=document.createElement('td');td.textContent=value;tr.append(td);
  }
  const td=document.createElement('td');const b=document.createElement('button');b.className='btn hub-secondary';b.textContent='Изменить';
  b.addEventListener('click',()=>techSetInvites(u.id,u.invited));td.append(b);tr.append(td);tbody.append(tr);
 });table.append(tbody);box.replaceChildren(table);
}
async function techSetInvites(id,current){const value=prompt('Количество приглашённых для ID '+id,current);if(value===null)return;if(!/^\d+$/.test(value))return alert('Введите целое неотрицательное число');try{const r=await fetch('/api/tech/invites',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({user_id:id,invited:Number(value)})});const j=await r.json();if(!r.ok)throw Error(j.error||'HTTP '+r.status);await techUsersLoad();window.hubToast?.('Количество приглашений сохранено')}catch(e){alert('Не удалось сохранить: '+e.message)}}
async function techJoinLoad(){
 const box=document.getElementById('techJoinList');if(!box)return;
 box.textContent='Загрузка…';
 try{
  const r=await fetch('/api/tech/join-requests');if(!r.ok)throw Error('HTTP '+r.status);
  const j=await r.json();const items=j.requests||[];
  if(!items.length){box.textContent='Активных заявок нет';return;}
  const table=document.createElement('table');
  table.innerHTML='<thead><tr><th>Участник</th><th>Группа</th><th>Действия</th></tr></thead>';
  const tbody=document.createElement('tbody');
  items.forEach(x=>{
   const tr=document.createElement('tr');
   const person=document.createElement('td');person.textContent=String(x.name||'')+' ('+String(x.user_id)+')';
   const group=document.createElement('td');group.textContent=String(x.chat_id);
   const actions=document.createElement('td');
   for(const [decision,title,cls] of [['approve','Принять','btn'],['decline','Отклонить','btn hub-secondary']]){
    const b=document.createElement('button');b.className=cls;b.textContent=title;
    b.addEventListener('click',()=>techReview(x.chat_id,x.user_id,x.request_id,decision));actions.append(b);
   }
   tr.append(person,group,actions);tbody.append(tr);
  });table.append(tbody);box.replaceChildren(table);
 }catch(e){box.textContent='Ошибка: '+e.message}
}
async function techReview(chat_id,user_id,request_id,decision){if(!confirm(decision==='approve'?'Принять участника?':'Отклонить заявку?'))return;try{const r=await fetch('/api/tech/join-review',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({chat_id,user_id,request_id,decision})});const j=await r.json();if(!r.ok)throw Error(j.error||'HTTP '+r.status);window.hubToast?.('Заявка обработана');await techJoinLoad()}catch(e){alert('Ошибка: '+e.message)}}

async function techOpsLoad(){const status=document.getElementById('techOpsStatus');try{const r=await fetch('/api/control');if(!r.ok)throw Error('HTTP '+r.status);const j=await r.json();for(const [id,key] of [['techUsers','users'],['techGroups','groups'],['techTags','tags']])document.getElementById(id).textContent=(j.counts||{})[key]??'—';status.textContent='Данные обновлены';techUsersLoad();techJoinLoad();}catch(e){status.textContent='Ошибка загрузки: '+e.message}}
let sysInfo=null;function showPage(id,btn){document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));document.getElementById(id).classList.add('active');document.querySelectorAll('.navbtn').forEach(x=>x.classList.remove('active'));btn.classList.add('active');if(id!=='reportsPage')loadSystemInfo()}
async function loadSystemInfo(){let r=await fetch('/api/system-info');if(!r.ok)return;sysInfo=await r.json();let s=sysInfo.settings||{};document.getElementById('projectName').textContent=s.project_name||'Shinobi Team';document.getElementById('platform').textContent=s.platform||'MLBB';document.getElementById('environment').textContent=s.environment||'PRODUCTION';document.getElementById('database').textContent=sysInfo.database_connected?'● PostgreSQL Connected':'Not configured';document.getElementById('botStatus').textContent='● '+(sysInfo.bot_status||'ONLINE');document.getElementById('botVersion').textContent=s.bot_version||'v1.0.0';document.getElementById('botDatabase').textContent=sysInfo.database||'—';let e=document.getElementById('versionEditor');e.innerHTML=['owner','tech_admin'].includes(sysInfo.role)?`<div class="editrow"><input id="newVersion" maxlength="64" value="${escapeHtml(s.bot_version||'v1.0.0')}" placeholder="Например v1920.20wsW"><button class="btn" onclick="saveVersion()">Сохранить</button></div><div class="status" id="versionStatus" style="margin-top:8px">Доступно владельцу и техническому администратору.</div>`:'<div class="status" style="margin-top:10px">Доступно владельцу и техническому администратору.</div>';renderMediaEditor(sysInfo.role)}
async function saveVersion(){let v=document.getElementById('newVersion').value.trim(),st=document.getElementById('versionStatus');if(!v){st.textContent='Введите версию';return}let r=await fetch('/api/system/version',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({version:v})});let j=await r.json();if(!r.ok){st.textContent='Ошибка сохранения';return}st.textContent='✓ Версия сохранена в PostgreSQL';document.getElementById('botVersion').textContent=j.version;loadSystemInfo()}
function renderMediaEditor(role){let e=document.getElementById('mediaEditor');if(!e)return;if(!['owner','tech_admin'].includes(role)){e.innerHTML='<div class="mediaTools"><h3>🎨 Оформление проекта</h3><div class="status">Доступно владельцу и техническому администратору.</div></div>';return}e.innerHTML=`<div class="mediaTools"><h3>🎨 Оформление проекта</h3><div class="status" style="margin-bottom:12px">PNG / JPG / WEBP, до 8 МБ. Изображения сохраняются в PostgreSQL.</div><div class="mediaRow"><input id="bannerFile" type="file" accept="image/png,image/jpeg,image/webp"><button class="btn" onclick="uploadMedia('banner','bannerFile','mediaStatus')">Загрузить баннер</button><input id="faviconFile" type="file" accept="image/png,image/jpeg,image/webp"><button class="btn" onclick="uploadMedia('favicon','faviconFile','mediaStatus')">Загрузить иконку</button></div><div class="status" id="mediaStatus" style="margin-top:10px">Готово</div><img class="mediaPreview" src="/media/banner?ts=${Date.now()}" onerror="this.style.display='none'" alt="Баннер"></div>`}
async function uploadMedia(kind,inputId,statusId){let input=document.getElementById(inputId),st=document.getElementById(statusId);if(!input||!input.files||!input.files[0]){st.textContent='Сначала выберите изображение';return}let fd=new FormData();fd.append('image',input.files[0]);st.textContent='Загрузка…';let r=await fetch('/api/media/'+kind,{method:'POST',body:fd});let j={};try{j=await r.json()}catch(e){}if(!r.ok){st.textContent='Ошибка загрузки: '+(j.error||r.status);return}st.textContent='✓ Изображение сохранено';if(kind==='favicon'){let link=document.querySelector('link[rel="icon"]');if(link)link.href='/media/favicon?ts='+Date.now()}else{document.querySelectorAll('.hero').forEach(x=>x.style.backgroundImage=`linear-gradient(0deg,rgba(5,5,6,.70),rgba(5,5,6,.22)),url('/media/banner?ts=${Date.now()}')`);let img=document.querySelector('.mediaPreview');if(img){img.style.display='block';img.src='/media/banner?ts='+Date.now()}}}
async function createNewReport(){let btn=document.getElementById('newReportBtn');btn.disabled=true;status.textContent='Формирование нового отчёта…';try{let r=await fetch('/api/reports',{method:'POST'});let j={};try{j=await r.json()}catch(e){}if(!r.ok){status.textContent='Ошибка создания: '+(j.error||r.status);return}current=j.report.id;await loadReports();openReport(j.report.id);status.textContent='✓ Новый отчёт создан: '+j.report.id}catch(e){status.textContent='Ошибка соединения при создании отчёта'}finally{btn.disabled=false}}
async function loadReports(){status.textContent='Обновление…';let r=await fetch('/api/reports');if(!r.ok){status.textContent='Ошибка доступа';return}let j=await r.json();reports=j.reports||[];renderReports();status.textContent='Отчётов: '+reports.length;if(reports.length&&!current)openReport(reports[0].id)}
function renderReports(){let el=document.getElementById('reports');if(!reports.length){el.innerHTML='<div class="empty">Архив пока пуст</div>';return}el.innerHTML=reports.map(x=>`<div class="report ${current===x.id?'active':''}" onclick="openReport('${x.id}')"><div class="rid">REPORT ${x.id}</div><div class="rdate">${new Date(x.created_at).toLocaleString('ru-RU')}</div></div>`).join('')}
function openReport(id){current=id;renderReports();let r=reports.find(x=>x.id===id);document.getElementById('reportTitle').textContent='Отчёт • '+new Date(r.created_at).toLocaleString('ru-RU');let files=r.files||[];document.getElementById('tabs').innerHTML=files.map((f,i)=>`<button class="tab ${i===0?'active':''}" onclick="openFile('${id}','${f}',this)">${labels[f]||f}</button>`).join('');if(files[0])openFile(id,files[0],document.querySelector('.tab'))}
async function openFile(id,file,btn){document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));if(btn)btn.classList.add('active');let box=document.getElementById('table');box.innerHTML='<div class="empty">Загрузка данных…</div>';let r=await fetch('/api/reports/'+encodeURIComponent(id)+'/'+encodeURIComponent(file));if(!r.ok){box.innerHTML='<div class="empty">Не удалось загрузить данные</div>';return}if(file==='overview.json'){let j=await r.json();let a=j.advertising||{},t=j.totals||{};box.innerHTML=`<div style="padding:18px"><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-bottom:18px"><div style="padding:18px;background:#180b0d;border:1px solid #5b1117;border-radius:12px"><div style="color:#ff3942;font-size:12px;font-weight:800">📣 РЕКЛАМНАЯ ССЫЛКА</div><div style="font-size:32px;font-weight:900;margin:9px 0">${a.invited||0}</div><div style="color:#888;font-size:12px">приглашённых</div><div style="margin-top:12px;font-size:12px;word-break:break-all">${a.link||''}</div></div><div style="padding:18px;background:#151517;border:1px solid #2b2b30;border-radius:12px"><div style="color:#888;font-size:12px">ВСЕГО ССЫЛОК</div><div style="font-size:30px;font-weight:900;margin-top:9px">${t.links||0}</div></div><div style="padding:18px;background:#151517;border:1px solid #2b2b30;border-radius:12px"><div style="color:#888;font-size:12px">ВСЕГО ПРИГЛАШЕНО</div><div style="font-size:30px;font-weight:900;margin-top:9px">${t.invited||0}</div></div><div style="padding:18px;background:#151517;border:1px solid #2b2b30;border-radius:12px"><div style="color:#888;font-size:12px">ПОЛЬЗОВАТЕЛЕЙ</div><div style="font-size:30px;font-weight:900;margin-top:9px">${t.users||0}</div></div></div><h3>Все ссылки</h3><div style="overflow:auto"><table><tr><th>Тип</th><th>Ссылка</th><th>Приглашено</th><th>Создатель ID</th><th>Создана</th></tr>${(j.links||[]).map(x=>`<tr><td>${x.type==='advertising'?'📣 Рекламная':'Реферальная'}</td><td>${x.link||''}</td><td><b>${x.invited||0}</b></td><td>${x.owner_id??''}</td><td>${x.created||''}</td></tr>`).join('')}</table></div></div>`;return}if(file.endsWith('.json')){let j=await r.json();box.innerHTML='<pre style="padding:18px;white-space:pre-wrap;word-break:break-word;color:#bbb">'+escapeHtml(JSON.stringify(j,null,2))+'</pre>';return}let buf=await r.arrayBuffer();let wb=XLSX.read(buf,{type:'array'});let s=wb.Sheets[wb.SheetNames[0]];let html=XLSX.utils.sheet_to_html(s);box.innerHTML=html;let first=document.querySelector('#table table tr');if(first)first.querySelectorAll('td').forEach(td=>{let th=document.createElement('th');th.innerHTML=td.innerHTML;td.replaceWith(th)})}
function escapeHtml(s){return s.replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}

let joinsChartInstance=null;
async function loadGrowth(){
 let [ar,jr]=await Promise.all([fetch('/api/referrals'),fetch('/api/joins-analytics?days=30')]);
 if(jr.ok){let j=await jr.json();document.getElementById('joinsToday').textContent=j.today||0;document.getElementById('joinsYesterday').textContent=j.yesterday||0;let d=Number(j.delta||0);document.getElementById('joinsDelta').textContent=(d>0?'+':'')+d;document.getElementById('joinsPeriod').textContent=j.period_total||0;let ctx=document.getElementById('joinsChart');if(joinsChartInstance)joinsChartInstance.destroy();joinsChartInstance=new Chart(ctx,{type:'line',data:{labels:(j.series||[]).map(x=>x.date.slice(5).split('-').reverse().join('.')),datasets:[{label:'Входы',data:(j.series||[]).map(x=>x.joins),borderWidth:2,tension:.28,fill:true}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#aaa'}}},scales:{x:{ticks:{color:'#777'},grid:{color:'#1d1d20'}},y:{beginAtZero:true,ticks:{color:'#777',precision:0},grid:{color:'#1d1d20'}}}}});}
 let box=document.getElementById('referralRanking');if(!ar.ok){box.innerHTML='<div class="empty">Ошибка загрузки ссылок</div>';return}let a=await ar.json(),rows=a.links||[];box.innerHTML='<div style="overflow:auto"><table><tr><th>#</th><th>Владелец</th><th>Ссылка</th><th>Приглашено</th><th>Действие</th></tr>'+rows.map((x,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(x.username?'@'+x.username:(x.name||'Без имени'))}</td><td style="max-width:420px;overflow:hidden;text-overflow:ellipsis">${escapeHtml(x.link||'')}</td><td><b>${x.invited||0}</b></td><td>${x.type==='advertising'?'<span class="status">Системная</span>':`<button class="navbtn" onclick='deleteReferral(${JSON.stringify(x.link)})'>Удалить</button>`}</td></tr>`).join('')+'</table></div>';hubFilterTable('referralRanking',document.getElementById('hubReferralFilter')?.value||'');
}
async function deleteReferral(link){if(!confirm('Удалить эту персональную ссылку? Она также будет отозвана в Telegram, если Bot API разрешит.'))return;let r=await fetch('/api/referrals/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({link})});let j={};try{j=await r.json()}catch(e){};if(!r.ok){alert('Ошибка: '+(j.error||r.status));return}loadGrowth()}

async function loadControl(){let r=await fetch('/api/control');if(!r.ok)return;let j=await r.json();document.getElementById('ctlUsers').textContent=j.counts.users;document.getElementById('ctlGroups').textContent=j.counts.groups;document.getElementById('ctlTags').textContent=j.counts.tags;document.getElementById('ctlMaint').textContent=j.maintenance?'ВКЛ':'ВЫКЛ';let ar=j.auto_reactions_enabled!==false;let arb=document.getElementById('autoReactionsToggle');arb.dataset.enabled=ar?'1':'0';arb.textContent=ar?'Выключить реакции':'Включить реакции';document.getElementById('autoReactionsState').textContent=ar?'🟢 Включены':'🔴 Выключены';let tg=!!j.server_tags_enabled;document.getElementById('serverTagsState').textContent=tg?'🟢 Функция включена':'🔴 Функция выключена';let tgb=document.getElementById('serverTagsToggle');tgb.textContent=tg?'Выключить выдачу тегов':'Включить выдачу тегов';tgb.dataset.enabled=tg?'1':'0';let sel=document.getElementById('tagValue');sel.innerHTML=(j.server_tags||[]).map(x=>`<option>${escapeHtml(x)}</option>`).join('');document.getElementById('tagHistory').innerHTML='<table><tr><th>Игрок</th><th>User ID</th><th>Тег</th><th>Дата</th></tr>'+Object.values(j.assignments||{}).slice(-50).reverse().map(x=>`<tr><td>${escapeHtml(x.name||'')}</td><td>${x.user_id||''}</td><td>${escapeHtml(x.tag||'')}</td><td>${escapeHtml(x.assigned_at||'')}</td></tr>`).join('')+'</table>';document.getElementById('ctlGroupList').innerHTML=(j.groups||[]).map(g=>`<div style="padding:9px;border-bottom:1px solid #252529"><b>${escapeHtml(g.title||String(g.id))}</b> · <code>${g.id}</code> · ${g.access==='deny'?'🔴 запрещены':'🟢 разрешены'} <button class="navbtn" onclick="toggleGroup(${g.id},'${g.access==='deny'?'allow':'deny'}')">${g.access==='deny'?'Разрешить':'Запретить'}</button></div>`).join('')||'Нет групп';document.getElementById('ctlUserList').innerHTML=(j.users||[]).slice(0,100).map(u=>`<div style="padding:7px;border-bottom:1px solid #252529">${escapeHtml(u.name||u.username||'Без имени')} · <code>${u.id}</code> · приглашено ${u.total_invited||0}</div>`).join('')||'Нет пользователей';loadMarketWallets()}
async function loadMarketWallets(){let sel=document.getElementById('marketPlayer'),st=document.getElementById('marketWalletStatus'),panel=document.getElementById('marketWalletPanel');if(!sel)return;let r=await fetch('/api/market/wallets');if(r.status===403){panel.style.display='none';return}let j={};try{j=await r.json()}catch(e){}if(!r.ok){st.textContent='Ошибка загрузки: '+(j.error||r.status);return}let rows=j.players||[];sel.innerHTML=rows.map(x=>`<option value="${x.chat_id}:${x.user_id}">${escapeHtml((x.username?'@'+x.username:x.name)||x.user_id)} · ${x.balance} 💎 · chat ${x.chat_id}</option>`).join('');st.textContent=rows.length?'Игроков на рынке: '+rows.length:'Пока нет игроков в базе рынка'}
async function updateMarketWallet(){let sel=document.getElementById('marketPlayer'),amount=Number(document.getElementById('marketAmount').value),mode=document.getElementById('marketMode').value,st=document.getElementById('marketWalletStatus');if(!sel.value||!Number.isInteger(amount)||amount<0){st.textContent='Укажи игрока и целое количество монет';return}let [chat_id,user_id]=sel.value.split(':');let r=await fetch('/api/market/wallet',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({chat_id:Number(chat_id),user_id:Number(user_id),amount,mode})});let j={};try{j=await r.json()}catch(e){};st.textContent=r.ok?'✓ Новый баланс: '+j.balance+' 💎':'Ошибка: '+(j.error||r.status);if(r.ok)loadMarketWallets()}
async function loadChanges(){let box=document.getElementById('changesList');box.innerHTML='<div class="empty">Загрузка…</div>';let r=await fetch('/api/changelog');if(!r.ok){box.innerHTML='<div class="empty">Не удалось загрузить историю</div>';return}let j=await r.json();box.innerHTML=(j.items||[]).map(c=>`<div class="mediaTools"><div class="eyebrow">${escapeHtml(c.date||'')}</div><h3 style="margin:7px 0 12px">${escapeHtml(c.title||'Обновление')}</h3><div>${(c.items||[]).map(i=>`<div style="padding:7px 0;border-bottom:1px solid #242428">✓ ${escapeHtml(i)}</div>`).join('')}</div></div>`).join('')||'<div class="empty">История пока пустая</div>'}
async function loadUntagged(){let st=document.getElementById('untaggedStatus'),box=document.getElementById('untaggedList');st.textContent='Проверяю участников…';box.value='';let r=await fetch('/api/control/untagged');let j={};try{j=await r.json()}catch(e){}if(!r.ok){st.textContent='Ошибка: '+(j.error||r.status);return}let rows=(j.untagged||[]).map(u=>{let who=u.username?'@'+u.username:(u.name||'Без имени');return who+' | ID: '+u.id});box.value=rows.join('\n');st.textContent='В основном чате: '+(j.total_members||0)+' · админов исключено: '+(j.admins_excluded||0)+' · проверено известных: '+(j.checked||0)+' · без тега: '+rows.length}
async function copyUntagged(){let box=document.getElementById('untaggedList'),st=document.getElementById('untaggedStatus');try{await navigator.clipboard.writeText(box.value);st.textContent='✓ Список скопирован'}catch(e){box.select();document.execCommand('copy');st.textContent='✓ Список скопирован'}}
async function assignTagWeb(){let st=document.getElementById('tagStatus');st.textContent='Выдаю…';let body={chat_id:Number(document.getElementById('tagChat').value),user_id:Number(document.getElementById('tagUser').value),tag:document.getElementById('tagValue').value};let r=await fetch('/api/server-tags/assign',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});let j=await r.json();st.textContent=r.ok?'✓ Тег выдан: '+j.tag:'Ошибка: '+(j.error||r.status);if(r.ok)loadControl()}
async function toggleAutoReactions(){let b=document.getElementById('autoReactionsToggle');let enabled=b.dataset.enabled!=='1';b.disabled=true;try{let r=await fetch('/api/control/auto-reactions',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({enabled})});let j=await r.json();if(!r.ok)throw Error(j.error||r.status);await loadControl()}catch(e){document.getElementById('autoReactionsState').textContent='Ошибка: '+e.message}finally{b.disabled=false}}
async function toggleServerTags(){let b=document.getElementById('serverTagsToggle');let enabled=b.dataset.enabled!=='1';b.disabled=true;let r=await fetch('/api/control/server-tags',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({enabled})});b.disabled=false;if(r.ok)loadControl();else{let j=await r.json().catch(()=>({}));document.getElementById('serverTagsState').textContent='Ошибка: '+(j.error||r.status)}}
async function toggleGroup(chat_id,access){let r=await fetch('/api/control/group-access',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({chat_id,access})});if(r.ok)loadControl()}
async function downloadFullBackup(){let code=prompt('Введите код полного бэкапа:');if(code===null)return;let st=document.getElementById('backupStatus'),btn=document.getElementById('backupBtn');st.textContent='Собираю актуальный ZIP…';btn.disabled=true;try{let r=await fetch('/api/backup/full',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({code})});if(!r.ok){let j={};try{j=await r.json()}catch(e){};st.textContent=r.status===401?'Неверный код':'Ошибка: '+(j.error||r.status);return}let blob=await r.blob();let cd=r.headers.get('Content-Disposition')||'';let m=cd.match(/filename=\"?([^\";]+)\"?/i);let name=m?m[1]:'shinobi-full-backup.zip';let u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(u);st.textContent='✓ ZIP скачан'}catch(e){st.textContent='Ошибка соединения'}finally{btn.disabled=false}}

async function loadAiModeration(){let box=document.getElementById('aiModLog');box.innerHTML='<div class="empty">Загрузка…</div>';let r=await fetch('/api/ai-moderation-log');if(!r.ok){box.innerHTML='<div class="empty">Ошибка загрузки</div>';return}let j=await r.json(),rows=j.log||[];document.getElementById('aiModCount').textContent=rows.length;if(!rows.length){box.innerHTML='<div class="empty">Нарушений пока нет</div>';return}box.innerHTML='<div class="tablewrap"><table><thead><tr><th>Время</th><th>Пользователь</th><th>Чат</th><th>Сообщение</th><th>Правило</th><th>Наказание</th><th>Причина</th><th>Источник</th></tr></thead><tbody>'+rows.map(x=>`<tr><td>${escapeHtml(x.time||'—')}</td><td>${escapeHtml(x.user_name||String(x.user_id||''))}<br><span class="status">${x.username?'@'+escapeHtml(x.username):'ID '+x.user_id}</span></td><td>${escapeHtml(x.chat_title||String(x.chat_id||''))}</td><td style="max-width:320px;white-space:normal">${escapeHtml(x.message_text||'—')}</td><td>${x.rule ? 'Правило '+x.rule+', пункт '+(x.subpoint||1) : '—'}</td><td>${escapeHtml(x.action_text||x.action||'—')}</td><td style="max-width:280px;white-space:normal">${escapeHtml(x.reason||'—')}</td><td>${escapeHtml(x.source_text||'AI')}</td></tr>`).join('')+'</tbody></table></div>'}
async function loadChatSummary(){let box=document.getElementById('chatSummaryBox');box.innerHTML='<div class="empty">Загрузка…</div>';let r=await fetch('/api/chat-summary');if(!r.ok){box.innerHTML='<div class="empty">Ошибка загрузки</div>';return}let j=await r.json();let chats=j.chats||[];if(!chats.length){box.innerHTML='<div class="empty">Пока нет собранной статистики. Она появится после новых сообщений в чатах.</div>';return}box.innerHTML=chats.map(c=>`<div class="infoCard" style="grid-column:span 2"><div class="infoLabel">${escapeHtml(c.title||String(c.chat_id))}</div><div style="font-size:12px;color:#666;margin:6px 0 14px">ID ${c.chat_id}</div><div class="infoGrid"><div><div class="infoLabel">УЧАСТНИКОВ</div><div class="infoValue">${c.members??'—'}</div></div><div><div class="infoLabel">АКТИВНЫ ЗА 15 МИН</div><div class="infoValue">${c.active_15m||0}</div></div><div><div class="infoLabel">СООБЩЕНИЙ СЕГОДНЯ</div><div class="infoValue">${c.messages_today||0}</div></div><div><div class="infoLabel">ПИК ЗА ЧАС</div><div class="infoValue">${c.peak_messages||0}</div><div class="status">${c.peak_hour||'—'}</div></div><div><div class="infoLabel">PING API</div><div class="infoValue">${c.ping_ms??'—'} мс</div></div><div><div class="infoLabel">ВСЕГО ОТСЛЕЖЕНО СООБЩЕНИЙ</div><div class="infoValue">${c.tracked_total_messages||0}</div></div></div></div>`).join('')}


async function loadOctoberPlan(){let box=document.getElementById('octoberPlanBox');if(!box)return;box.innerHTML='<div class="empty">Загрузка…</div>';let r=await fetch('/api/october-plan');if(!r.ok){box.innerHTML='<div class="empty">Доступ только тех. администратору</div>';return}let j=await r.json(),items=j.items||[];let done=items.filter(x=>x.done).length;box.innerHTML=`<div style="padding:16px;border-bottom:1px solid #252529"><b>${done} / ${items.length}</b> выполнено • ${items.length-done} осталось</div><div style="overflow:auto"><table><tr><th>Готово</th><th>Дата</th><th>Активность</th><th>Формат</th><th>Текст поста</th></tr>${items.map(x=>`<tr style="${x.done?'opacity:.55':''}"><td><input type="checkbox" ${x.done?'checked':''} onchange="setOctoberDone('${x.date}',this.checked)"></td><td><b>${escapeHtml(x.date.split('-').reverse().join('.'))}</b></td><td>${escapeHtml(x.title||'')}</td><td>${escapeHtml(x.format||'')}</td><td style="white-space:normal;min-width:360px">${escapeHtml(x.post_text||'')}</td></tr>`).join('')}</table></div>`}
async function setOctoberDone(date,done){let r=await fetch('/api/october-plan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({date,done})});if(!r.ok){alert('Не удалось сохранить отметку');loadOctoberPlan();return}loadOctoberPlan()}

loadReports();</script><script>
// V4: client-side filtering of already-authorized API results; no new permissions.
function hubFilterTable(containerId,query){
 const box=document.getElementById(containerId);if(!box)return;
 const term=String(query||'').trim().toLocaleLowerCase('ru');
 box.querySelectorAll('tbody tr').forEach(tr=>tr.classList.toggle('hub-filter-hidden',!!term&&!tr.textContent.toLocaleLowerCase('ru').includes(term)));
}
function hubExportTable(containerId,filename){
 const box=document.getElementById(containerId),table=box?.querySelector('table');
 if(!table){window.hubToast?.('Таблица пока не загружена',true);return}
 const rows=[...table.querySelectorAll('tr')].filter(tr=>!tr.classList.contains('hub-filter-hidden'));
 const csv=rows.map(tr=>[...tr.querySelectorAll('th,td')].filter((cell,i)=>!(/действие/i.test(table.querySelectorAll('th')[i]?.textContent||''))).map(cell=>'"'+cell.innerText.replace(/\s+/g,' ').trim().replace(/"/g,'""')+'"').join(';')).join('\r\n');
 const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
 window.hubToast?.('Таблица экспортирована');
}
</script><script>
function hubScheduleTable(target, payload){
 const el=document.getElementById(target);
 if(payload.missing){el.innerHTML='<div class="empty">Файл графика отсутствует</div>';return}
 if(!payload.rows.length){el.innerHTML='<div class="empty">Записей пока нет</div>';return}
 const table=document.createElement('table');const head=document.createElement('thead');const hr=document.createElement('tr');
 for(const col of payload.columns){const th=document.createElement('th');th.textContent=col;hr.append(th)}head.append(hr);table.append(head);
 const body=document.createElement('tbody');for(const row of payload.rows){const tr=document.createElement('tr');for(const val of row){const td=document.createElement('td');td.textContent=val;tr.append(td)}body.append(tr)}table.append(body);el.replaceChildren(table);
}
let hubLiveData=null;
async function hubLiveLoad(){
 const status=document.getElementById('hubLiveStatus');
 try{const r=await fetch('/api/hub/schedule-live',{credentials:'same-origin'});if(!r.ok)throw Error('HTTP '+r.status);
 hubLiveData=await r.json();const u=document.getElementById('hubLiveUser'),d=document.getElementById('hubLiveDay');
 const prev=u.value;u.replaceChildren();(hubLiveData.users||[]).forEach(x=>{const o=new Option((x.name||x.id)+' · '+x.id,x.id);u.add(o)});
 if(prev&&[...u.options].some(o=>o.value===prev))u.value=prev;
 d.replaceChildren();const days=['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];(hubLiveData.weekdays||[]).forEach((x,i)=>d.add(new Option(days[i]||x,x)));
 hubLiveRender();status.textContent='Загружены актуальные данные графика';
 }catch(e){status.textContent='Не удалось загрузить живой график: '+e.message}
}
function hubLiveRender(){if(!hubLiveData)return;
 const uid=document.getElementById('hubLiveUser').value,day=document.getElementById('hubLiveDay').value;
 const rec=(((hubLiveData.assignments||{})[uid]||{})[day]||{});
 document.getElementById('hubLiveOff').checked=!!rec.off;
 const box=document.getElementById('hubLiveDuties');box.replaceChildren();
 Object.entries(hubLiveData.duties||{}).forEach(([key,label])=>{const wrap=document.createElement('label');wrap.className='hub-duty-choice';const input=document.createElement('input');input.type='checkbox';input.value=key;input.checked=(rec.duties||[]).includes(key);wrap.append(input,document.createTextNode(' '+label));box.append(wrap)});
}
async function hubLiveSave(){const status=document.getElementById('hubLiveStatus');const uid=document.getElementById('hubLiveUser').value,day=document.getElementById('hubLiveDay').value;
 if(!uid||!day){status.textContent='Выбери сотрудника и день';return}
 const body={uid,day,expected:(((hubLiveData.assignments||{})[uid]||{})[day]||{}),off:document.getElementById('hubLiveOff').checked,duties:[...document.querySelectorAll('#hubLiveDuties input:checked')].map(x=>x.value)};
 try{const r=await fetch('/api/hub/schedule-live',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});const data=await r.json();if(r.status===409){await hubLiveLoad();throw Error('График изменён другим администратором. Данные обновлены, проверьте их и сохраните снова.')}if(!r.ok)throw Error(data.error||'HTTP '+r.status);
 hubLiveData.assignments[uid]??={};hubLiveData.assignments[uid][day]=data.record;hubLiveRender();status.textContent='Сохранено в общей базе';window.hubToast?.('График сохранён');
 }catch(e){status.textContent='Ошибка сохранения: '+e.message;window.hubToast?.('Ошибка сохранения',true)}
}
async function hubDayoffLoad(){
 const el=document.getElementById('hubDayoffReviewRows'),st=document.getElementById('hubDayoffReviewStatus');
 try{const r=await fetch('/api/hub/dayoff-requests',{credentials:'same-origin'});if(!r.ok)throw Error('HTTP '+r.status);const data=await r.json();
 el.replaceChildren();const table=document.createElement('table');const head=document.createElement('tr');
 for(const label of ['Заявка','Сотрудник','Дата','Причина','Статус','Решение']){const th=document.createElement('th');th.textContent=label;head.append(th)}table.append(head);
 for(const item of data.requests||[]){const tr=document.createElement('tr');for(const val of [item.id,item.name+' ('+item.user_id+')',item.date,item.reason,item.status]){const td=document.createElement('td');td.textContent=val;tr.append(td)}
 const action=document.createElement('td');if(item.status==='Pending')for(const [label,decision] of [['Одобрить','Approved'],['Отклонить','Rejected']]){const btn=document.createElement('button');btn.className='btn hub-secondary';btn.textContent=label;btn.onclick=()=>hubDayoffReview(item.id,decision);action.append(btn)}tr.append(action);table.append(tr)}
 el.append(table);st.textContent='Источник: '+data.source+' · '+(data.requests||[]).length+' заявок';
 }catch(e){st.textContent='Не удалось загрузить заявки: '+e.message}
}
async function hubDayoffReview(id,decision){
 if(!confirm((decision==='Approved'?'Одобрить':'Отклонить')+' заявку '+id+'?'))return;
 const st=document.getElementById('hubDayoffReviewStatus');
 try{const r=await fetch('/api/hub/dayoff-review',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify({id,decision})});const data=await r.json();if(!r.ok)throw Error(data.error||'HTTP '+r.status);window.hubToast?.('Решение сохранено');await hubDayoffLoad()}
 catch(e){st.textContent='Ошибка: '+e.message;await hubDayoffLoad()}
}
async function hubScheduleLoad(){
 const status=document.getElementById('hubScheduleUpdated');status.textContent='Загрузка данных…';
 try{const r=await fetch('/api/hub/schedule-snapshot',{credentials:'same-origin'});if(!r.ok)throw Error('HTTP '+r.status);const data=await r.json();
 hubScheduleTable('hubScheduleRows',data.schedule);hubScheduleTable('hubDaysOffRows',data.days_off);
 document.getElementById('hubScheduleCount').textContent=data.schedule.rows.length;
 document.getElementById('hubDaysOffCount').textContent=data.days_off.rows.length;
 hubFilterTable('hubScheduleRows',document.getElementById('hubScheduleFilter').value);
 hubFilterTable('hubDaysOffRows',document.getElementById('hubDaysOffFilter').value);
 status.textContent='Источник: XLSX-снимки · График обновлён: '+(data.schedule.updated_at||'нет данных')+' · Заявки: '+(data.days_off.updated_at||'нет данных');
 }catch(e){status.textContent='Не удалось загрузить график: '+e.message;window.hubToast?.('Ошибка загрузки графика',true)}
}
</script><script>function toggleHubMenu(){const o=document.body.classList.toggle('hub-menu-open');document.querySelector('.hub-menu-toggle')?.setAttribute('aria-expanded',String(o))}function closeHubMenu(){document.body.classList.remove('hub-menu-open');document.querySelector('.hub-menu-toggle')?.setAttribute('aria-expanded','false')}document.querySelectorAll('.topnav .navbtn').forEach(b=>b.addEventListener('click',closeHubMenu));document.addEventListener('keydown',e=>{if(e.key==='Escape')closeHubMenu()});</script><script>
(function(){
 const icons={overview:'<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',staff:'<circle cx="9" cy="8" r="3"/><path d="M3 20v-2a6 6 0 0 1 12 0v2"/><path d="M17 5a3 3 0 0 1 0 6M18 14a5 5 0 0 1 3 5"/>',link:'<path d="M10 13a5 5 0 0 0 7.1 0l2-2a5 5 0 0 0-7.1-7.1l-1.1 1.1"/><path d="M14 11a5 5 0 0 0-7.1 0l-2 2a5 5 0 0 0 7.1 7.1l1.1-1.1"/>',calendar:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 10h18"/>',shield:'<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',chart:'<path d="M4 20V10M10 20V4M16 20v-8M22 20v-6"/>',settings:'<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M19 5l-2 2M7 17l-2 2"/>',file:'<path d="M6 2h9l5 5v15H6z"/><path d="M15 2v6h5M9 12h8M9 16h8"/>',bot:'<rect x="4" y="7" width="16" height="13" rx="3"/><path d="M12 3v4M9 3h6M8 13h.01M16 13h.01M9 17h6"/>'};
 document.querySelectorAll('.topnav .navbtn').forEach((b,i)=>{if(b.querySelector('.hub-nav-icon'))return;const t=(b.textContent||'').toLowerCase();const k=/сотруд|админ|персонал/.test(t)?'staff':/ссыл|реферал|приглаш/.test(t)?'link':/график|выходн|календар|план/.test(t)?'calendar':/модерац|наруш|безопас/.test(t)?'shield':/аналит|статист|отчёт/.test(t)?'chart':/настрой|управлен/.test(t)?'settings':/бот|игр/.test(t)?'bot':/истори|лог|изменен/.test(t)?'file':'overview';const e=document.createElement('span');e.className='hub-nav-icon';e.setAttribute('aria-hidden','true');e.innerHTML='<svg viewBox="0 0 24 24">'+icons[k]+'</svg>';b.prepend(e)});
 const stack=document.createElement('div');stack.id='hub-toast-stack';stack.setAttribute('aria-live','polite');document.body.appendChild(stack);
 window.hubToast=function(message,isError){const n=document.createElement('div');n.className='hub-toast'+(isError?' error':'');n.textContent=message;stack.appendChild(n);setTimeout(()=>n.remove(),4200)};
})();
</script><style>
/* Shinobi Hub: unified application shell */
.hub-feature-grid{display:grid;grid-template-columns:1.4fr 1fr;gap:18px;margin-top:22px}.hub-quick{padding:20px;display:flex;gap:12px;flex-wrap:wrap;align-items:center}.hub-quick input,.hub-quick select{padding:12px;border-radius:12px;border:1px solid #cce3f3;min-width:170px}.infoCard{background:rgba(255,255,255,.77)!important;backdrop-filter:blur(14px);transition:transform .2s,box-shadow .2s}.infoCard:hover{transform:translateY(-3px);box-shadow:0 14px 30px #3d8fba21!important}.hub-quick .btn{border-radius:12px}.page.active{animation:hubReveal .3s ease both}@keyframes hubReveal{from{opacity:.5;transform:translateY(8px)}to{opacity:1;transform:none}}@media(max-width:800px){.hub-feature-grid{grid-template-columns:1fr}}
@media(prefers-reduced-motion:reduce){.page.active,.infoCard{animation:none!important;transition:none!important}}
</style><script>
function hubGo(id){const b=[...document.querySelectorAll('.topnav>.navbtn')].find(x=>(x.getAttribute('onclick')||'').includes("'"+id+"'"));if(b)b.click()}
function hubMetric(label,value){const d=document.createElement('div');d.className='infoCard';const l=document.createElement('div');l.className='infoLabel';l.textContent=label;const v=document.createElement('div');v.className='infoValue';v.textContent=value;d.append(l,v);return d}
async function hubOverviewLoad(){const box=document.getElementById('hubOverviewStats');if(!box)return;box.replaceChildren(hubMetric('СТАТУС','Загрузка…'));const endpoints=[['/api/status','БОТ / API'],['/api/chat-summary','ЧАТОВ В СВОДКЕ'],['/api/reports','СОХРАНЁННЫХ ОТЧЁТОВ'],['/api/referrals','РЕФЕРАЛЬНЫЕ ДАННЫЕ']];const res=await Promise.allSettled(endpoints.map(async ([url,label])=>{const r=await fetch(url,{credentials:'same-origin'});if(!r.ok)throw Error(String(r.status));return {label,data:await r.json()}}));box.replaceChildren();res.forEach((r,i)=>{if(r.status!=='fulfilled'){box.append(hubMetric(endpoints[i][1],'Недоступно'));return}const {label,data}=r.value;let value='Доступно';if(label==='ЧАТОВ В СВОДКЕ')value=(data.chats||[]).length;if(label==='СОХРАНЁННЫХ ОТЧЁТОВ')value=(data.reports||[]).length;if(label==='БОТ / API')value=data.status||data.ok===true&&'Работает'||'На связи';if(label==='РЕФЕРАЛЬНЫЕ ДАННЫЕ')value=Array.isArray(data)?data.length:(Array.isArray(data.links)?data.links.length:'Доступно');box.append(hubMetric(label,value))})}
async function hubStaffLoad(){const box=document.getElementById('hubStaffRows');if(!box)return;box.textContent='Загрузка…';try{const r=await fetch('/api/staff');if(r.status===403){box.textContent='Только владелец может управлять сотрудниками.';document.getElementById('hubStaffUid').disabled=true;document.getElementById('hubStaffRole').disabled=true;return}if(!r.ok)throw Error(r.status);const j=await r.json();const t=document.createElement('table');t.innerHTML='<thead><tr><th>Telegram ID</th><th>Роль</th><th>Действия</th></tr></thead>';const tb=document.createElement('tbody');(j.staff||[]).forEach(u=>{const tr=document.createElement('tr');const a=document.createElement('td');a.textContent=u.uid;const b=document.createElement('td');b.textContent=({owner:'Владелец',deputy:'Заместитель',admin:'Администратор',tech_admin:'Технический администратор',moderator:'Модератор',disabled:'Нет доступа'})[u.role]||u.role;const c=document.createElement('td');if(u.role!=='owner'){const btn=document.createElement('button');btn.className='navbtn';btn.textContent='Изменить';btn.onclick=()=>{document.getElementById('hubStaffUid').value=u.uid;document.getElementById('hubStaffRole').value=u.role==='disabled'?'moderator':u.role;window.scrollTo({top:0,behavior:'smooth'})};c.append(btn)}tr.append(a,b,c);tb.append(tr)});t.append(tb);box.replaceChildren(t)}catch(e){box.textContent='Ошибка загрузки: '+e.message}}
async function hubStaffSave(){const uid=Number(document.getElementById('hubStaffUid').value),role=document.getElementById('hubStaffRole').value,msg=document.getElementById('hubStaffMsg');if(!Number.isSafeInteger(uid)||uid<=0){msg.textContent='Укажи корректный Telegram ID';return}if(!confirm('Изменить роль сотрудника '+uid+' на '+role+'? Все его активные сессии будут завершены.'))return;try{const r=await fetch('/api/staff/role',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid,role})});const j=await r.json();if(!r.ok)throw Error(j.error||r.status);msg.textContent='Роль сохранена. Сессии сотрудника отозваны.';if(window.hubToast)hubToast('Права обновлены');hubStaffLoad()}catch(e){msg.textContent='Ошибка: '+e.message;if(window.hubToast)hubToast(e.message,true)}}
let hubAuditEvents=[];
async function hubSecurityLoad(){const sb=document.getElementById('hubSessionRows'),ab=document.getElementById('hubAuditRows');sb.textContent='Загрузка…';ab.textContent='Загрузка…';try{const [sr,ar]=await Promise.all([fetch('/api/sessions'),fetch('/api/login-audit')]);if(sr.status===403||ar.status===403){sb.textContent='Доступ только владельцу';ab.textContent='Доступ только владельцу';return}if(!sr.ok||!ar.ok)throw Error('Ошибка сервера');const sessions=(await sr.json()).sessions||[];hubAuditEvents=(await ar.json()).events||[];document.getElementById('hubSessionCount').textContent=sessions.length;document.getElementById('hubAuditCount').textContent=hubAuditEvents.length;const t=document.createElement('table');t.innerHTML='<thead><tr><th>Telegram ID</th><th>Роль</th><th>Последняя активность</th><th>Истекает</th><th>Действие</th></tr></thead>';const body=document.createElement('tbody');for(const x of sessions){const tr=document.createElement('tr');for(const value of [x.uid,x.role,new Date(x.last_seen*1000).toLocaleString('ru-RU'),new Date(x.expires*1000).toLocaleString('ru-RU')]){const td=document.createElement('td');td.textContent=value;tr.append(td)}const td=document.createElement('td'),btn=document.createElement('button');btn.className='navbtn';btn.textContent='Завершить';btn.onclick=()=>hubRevokeUser(x.uid);td.append(btn);tr.append(td);body.append(tr)}t.append(body);sb.replaceChildren(t);hubRenderAudit()}catch(e){sb.textContent='Ошибка: '+e.message;ab.textContent='Ошибка: '+e.message}}
function hubRenderAudit(){const box=document.getElementById('hubAuditRows');if(!box)return;const q=(document.getElementById('hubAuditSearch')?.value||'').trim().toLowerCase();const events=hubAuditEvents.filter(e=>String(e.uid).includes(q)||String(e.event).toLowerCase().includes(q));const t=document.createElement('table');t.innerHTML='<thead><tr><th>Дата</th><th>Telegram ID</th><th>Событие</th></tr></thead>';const body=document.createElement('tbody');for(const e of events){const tr=document.createElement('tr');for(const v of [new Date(e.at*1000).toLocaleString('ru-RU'),e.uid,e.event]){const td=document.createElement('td');td.textContent=v;tr.append(td)}body.append(tr)}t.append(body);box.replaceChildren(t)}
async function hubRevokeUser(uid){if(!confirm('Завершить все сессии Telegram ID '+uid+'?'))return;try{const r=await fetch('/api/revoke-user',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid})});if(!r.ok)throw Error('Не удалось завершить сессии');window.hubToast?.('Сессии завершены');hubSecurityLoad()}catch(e){window.hubToast?.(e.message,true)}}
async function hubRevokeAll(){if(!confirm('Завершить все сессии сотрудников, кроме своей?'))return;try{const r=await fetch('/api/sessions/revoke-others',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});if(!r.ok)throw Error('Ошибка завершения сессий');window.hubToast?.('Чужие сессии завершены');hubSecurityLoad()}catch(e){window.hubToast?.(e.message,true)}}
window.addEventListener('DOMContentLoaded',()=>{const first=document.querySelector('.topnav>.navbtn');if(first){document.querySelectorAll('.topnav>.navbtn').forEach(b=>b.classList.remove('active'));first.classList.add('active')}document.getElementById('reportsPage')?.classList.remove('active');document.getElementById('hubOverview')?.classList.add('active');hubOverviewLoad()});
</script></body></html>'''


OIDC_DISCOVERY = 'https://oauth.telegram.org/.well-known/openid-configuration'

def _oidc_settings():
    client_id=os.getenv('TELEGRAM_CLIENT_ID','').strip()
    client_secret=os.getenv('TELEGRAM_CLIENT_SECRET','').strip()
    base=os.getenv('HUB_PUBLIC_URL','').strip().rstrip('/')
    if not (client_id and client_secret and base.startswith('https://')):
        return None
    return client_id,client_secret,base+'/auth/telegram/callback'

async def _oidc_discovery():
    async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=12)) as client:
        async with client.get(OIDC_DISCOVERY) as response:
            response.raise_for_status()
            data=await response.json()
    if not all(data.get(k,'').startswith('https://') for k in ('authorization_endpoint','token_endpoint','jwks_uri')):
        raise ValueError('Invalid Telegram OIDC discovery')
    return data

async def login_page(request):
    if _read_session(request):raise web.HTTPFound('/dashboard')
    configured=_oidc_settings() is not None
    action='<a class="btn" href="/auth/telegram/start">✈ Войти через Telegram</a>' if configured else '<div class="tip">Авторизация ещё не настроена. Укажите TELEGRAM_CLIENT_ID, TELEGRAM_CLIENT_SECRET и HUB_PUBLIC_URL в Railway.</div>'
    html = """<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Shinobi Team Hub — Telegram Login</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:radial-gradient(ellipse at 20% 10%,#bde5ff,#f8fcff 65%);font:16px system-ui;color:#24506e}.card{max-width:390px;margin:20px;padding:38px;text-align:center;border:1px solid #ffffffaa;border-radius:28px;background:#ffffffa8;box-shadow:0 18px 50px #5490b22e;backdrop-filter:blur(25px)}h1{color:#157fbe}.btn{display:block;margin:25px 0 8px;padding:16px;border-radius:14px;color:white;background:#229ed9;text-decoration:none;font-weight:700;box-shadow:0 10px 24px #229ed944}.btn:hover{background:#1689c3}.tip{padding:16px;border-radius:14px;background:#e5f5ff;margin-top:20px;line-height:1.65}.note{font-size:12px;opacity:.7}</style><div class="card"><h1>SHINOBI TEAM HUB</h1><p>Защищённый вход для администрации</p>"""+action+"""<p class="note">Доступ только сотрудникам Shinobi Team. Подтвердите вход через Telegram.</p></div></html>"""
    return web.Response(text=html,content_type='text/html',headers={'Cache-Control':'no-store'})

async def oidc_start(request):
    cfg=_oidc_settings()
    if not cfg:return web.Response(text='Telegram Login не настроен в Railway.',status=503)
    client_id,secret,redirect=cfg
    try: discovery=await _oidc_discovery()
    except Exception as exc:
        print('[OIDC DISCOVERY ERROR]',repr(exc))
        return web.Response(text='Не удалось подключиться к Telegram Login.',status=502)
    state,verifier=hub_auth.oidc_state_create()
    challenge=base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b'=').decode()
    params={'response_type':'code','client_id':client_id,'redirect_uri':redirect,'scope':'openid profile','state':state,'code_challenge':challenge,'code_challenge_method':'S256'}
    response=web.HTTPFound(discovery['authorization_endpoint']+'?'+urlencode(params))
    response.set_cookie('hub_oidc_state',state,max_age=300,secure=True,httponly=True,samesite='Lax',path='/auth/telegram')
    raise response

async def oidc_callback(request):
    cfg=_oidc_settings()
    if not cfg:return web.Response(text='Авторизация не настроена',status=503)
    client_id,secret,redirect=cfg
    state=request.query.get('state','')
    if not state or not secrets.compare_digest(state,request.cookies.get('hub_oidc_state','')):
        return web.Response(text='Недействительное состояние авторизации. Начните вход заново.',status=403)
    verifier=hub_auth.oidc_state_consume(state)
    if not verifier:return web.Response(text='Запрос входа уже использован или устарел.',status=403)
    code=request.query.get('code','')
    if not code:return web.Response(text='Telegram не подтвердил вход.',status=403)
    try:
        discovery=await _oidc_discovery()
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=15)) as client:
            async with client.post(discovery['token_endpoint'],data={'grant_type':'authorization_code','code':code,'redirect_uri':redirect,'client_id':client_id,'client_secret':secret,'code_verifier':verifier}) as r:
                if r.status!=200:raise ValueError('Token exchange failed')
                tokens=await r.json()
            async with client.get(discovery['jwks_uri']) as r:
                r.raise_for_status();jwks=await r.json()
        encoded=tokens.get('id_token','')
        header=jwt.get_unverified_header(encoded)
        if header.get('alg')!='RS256':raise ValueError('Unexpected signing algorithm')
        key_data=next((k for k in jwks['keys'] if k.get('kid')==header.get('kid') and k.get('kty')=='RSA'),None)
        if not key_data:raise ValueError('Signing key not found')
        key=jwt.algorithms.RSAAlgorithm.from_jwk(key_data)
        claims=jwt.decode(encoded,key,algorithms=['RS256'],audience=client_id,issuer=discovery['issuer'],options={'require':['exp','iat','iss','aud','sub']},leeway=30)
        # Telegram OIDC `sub` is an OIDC subject, not the Bot API user ID.
        # The verified `profile` claim `id` is the Telegram numeric user ID.
        raw_uid=claims.get('id')
        if isinstance(raw_uid,bool) or not isinstance(raw_uid,(int,str)) or not str(raw_uid).isascii() or not str(raw_uid).isdecimal():
            raise ValueError('Verified Telegram ID token has no valid numeric id claim (profile scope required)')
        uid=int(raw_uid)
        if not 1 <= uid <= 1099511627775:
            raise ValueError('Verified Telegram user ID outside Bot API user ID range')
        session=hub_auth.oidc_session_create(uid,request.remote or '')
        if not session:return web.Response(text='Доступ запрещён: этот Telegram-аккаунт не назначен сотрудником Shinobi Team. Проверьте Telegram ID аккаунта и HUB_OWNER_IDS в Railway. Подробности — в Railway Logs ([HUB OIDC ACCESS DENIED]).',status=403,content_type='text/plain',charset='utf-8')
    except Exception as exc:
        print('[TELEGRAM OIDC LOGIN ERROR]',type(exc).__name__,str(exc)[:150])
        return web.Response(text='Не удалось подтвердить вход через Telegram. Попробуйте ещё раз.',status=403)
    response=web.HTTPFound('/dashboard')
    response.del_cookie('hub_oidc_state',path='/auth/telegram')
    response.set_cookie('shinobi_session',session,httponly=True,secure=True,samesite='Lax',max_age=43200,path='/')
    raise response

async def telegram_login(request):
    token=request.query.get('token','')
    session=hub_auth.redeem(token,request.remote or '')
    if not session:return web.Response(text='Ссылка недействительна, использована или срок истёк. Запросите новую через /start.',status=403)
    response=web.HTTPFound('/dashboard')
    response.set_cookie('shinobi_session',session,httponly=True,secure=True,samesite='Lax',max_age=43200,path='/')
    raise response

async def logout(request):
    hub_auth.revoke(request.cookies.get('shinobi_session',''),request.remote or '')
    response=web.HTTPFound('/')
    response.del_cookie('shinobi_session',path='/')
    raise response

async def dashboard(request):
    session=_read_session(request)
    if not session:raise web.HTTPFound('/')
    role=session['role']
    html=DASHBOARD_HTML.replace('__ROLE__',f'{role.upper()} · ID {session["uid"]}')
    tech_nav='<button class="navbtn" onclick="showPage(\'octoberPage\',this);loadOctoberPlan()">📅 План на октябрь</button>' if role in ('owner','deputy') else ''
    tech_nav += '<button class="navbtn" onclick="showPage(\'techOpsPage\',this);techOpsLoad()">🛠 Тех. панель</button>' if role in ('owner','deputy','tech_admin') else ''
    return web.Response(text=html.replace('__TECH_NAV__',tech_nav),content_type='text/html')

async def login_audit(request):
    return web.json_response({'events':hub_auth.recent()})

async def revoke_user_api(request):
    body=await request.json()
    try:uid=int(body['uid'])
    except (KeyError,TypeError,ValueError):return web.json_response({'error':'uid required'},status=400)
    hub_auth.revoke_user(uid)
    return web.json_response({'ok':True,'uid':uid})

async def sessions_api(request):
    return web.json_response({'sessions':hub_auth.active_sessions()})

async def revoke_others_api(request):
    session=_read_session(request)
    if not session:return web.json_response({'error':'unauthorized'},status=401)
    hub_auth.revoke_other_sessions(session['uid'],request.cookies.get('shinobi_session',''))
    return web.json_response({'ok':True})

async def staff_api(request):
    if (_read_session(request) or {}).get('role') != 'owner':
        return web.json_response({'error':'forbidden'},status=403)
    return web.json_response({'staff':hub_auth.staff_list()})

async def staff_role_api(request):
    data=await request.json()
    try:
        uid=int(data['uid']);new_role=str(data['role'])
        if uid<=0:raise ValueError('invalid uid')
        hub_auth.set_staff_role(uid,new_role)
    except (KeyError,TypeError,ValueError) as e:
        return web.json_response({'error':str(e)},status=400)
    return web.json_response({'ok':True,'uid':uid,'role':new_role})

async def hub_schedule_live_api(request):
    """Current duties, shared with the Telegram bot via PostgreSQL."""
    from schedule_system import _load_assignments, _known_users, WEEKDAYS, DUTIES
    try:
        return web.json_response({"assignments": _load_assignments(), "users": _known_users(), "weekdays": WEEKDAYS, "duties": DUTIES}, dumps=lambda v: json.dumps(v, ensure_ascii=False))
    except Exception as exc:
        print("[HUB SCHEDULE LIVE]", type(exc).__name__)
        return web.json_response({"error": "schedule_unavailable"}, status=503)


async def hub_schedule_save_api(request):
    from schedule_system import update_schedule_record_atomic, WEEKDAYS, DUTIES
    session = _read_session(request)
    if not session or session["role"] not in ("owner", "deputy"):
        return web.json_response({"error": "forbidden"}, status=403)
    try:
        body = await request.json()
        uid = str(int(body["uid"]))
        day = body["day"]
        if int(uid) <= 0 or day not in WEEKDAYS:
            raise ValueError("invalid user/day")
        duties = body.get("duties", [])
        if not isinstance(duties, list) or len(duties) > len(DUTIES) or any(x not in DUTIES for x in duties):
            raise ValueError("invalid duties")
        off = body.get("off", False)
        if not isinstance(off, bool):
            raise ValueError("invalid off")
        expected = body.get("expected")
        if not isinstance(expected, dict):
            raise ValueError("missing expected record")
        rec = update_schedule_record_atomic(uid, day, expected, {"off": off, "duties": [] if off else list(dict.fromkeys(duties)), "always": False})
        if rec is None:
            return web.json_response({"error": "schedule_conflict", "message": "График изменился. Обновите данные перед сохранением."}, status=409)
        hub_auth.audit(session["uid"], "schedule_updated_" + uid + "_" + day)
        return web.json_response({"ok": True, "record": rec})
    except (ValueError, KeyError, TypeError, json.JSONDecodeError):
        return web.json_response({"error": "invalid_schedule_input"}, status=400)
    except Exception as exc:
        print("[HUB SCHEDULE SAVE]", type(exc).__name__)
        return web.json_response({"error": "schedule_save_failed"}, status=503)


async def hub_dayoff_pending_api(request):
    from schedule_system import list_dayoff_requests
    try:
        return web.json_response({'requests':list_dayoff_requests(300), 'source':'postgresql' if os.getenv('DATABASE_URL') else 'local_xlsx_snapshot'})
    except Exception as exc:
        print('[HUB DAYOFF]',type(exc).__name__)
        return web.json_response({'error':'dayoff_unavailable'},status=503)


async def hub_dayoff_review_api(request):
    session = _read_session(request)
    if not session or session['role'] not in ('owner','deputy'):
        return web.json_response({'error':'forbidden'},status=403)
    try:
        body = await request.json()
        request_id = str(body['id'])
        decision = str(body['decision'])
        from schedule_system import review_dayoff_request
        found = review_dayoff_request(request_id,decision,session['uid'])
        if not found:
            return web.json_response({'error':'already_reviewed_or_not_found'},status=409)
        hub_auth.audit(session['uid'],'dayoff_'+decision.lower()+'_'+request_id)
        # Telegram delivery is best effort; database decision is already committed.
        try:
            import aiohttp
            token=os.getenv('BOT_TOKEN','')
            if token:
                msg=('✅' if decision=='Approved' else '❌')+' Заявка '+request_id+(' одобрена' if decision=='Approved' else ' отклонена')+'\nДата: '+str(found['date'])
                async with aiohttp.ClientSession() as http:
                    async with http.post('https://api.telegram.org/bot'+token+'/sendMessage',json={'chat_id':int(found['user_id']),'text':msg},timeout=aiohttp.ClientTimeout(total=5)) as response:
                        if response.status!=200:print('[HUB DAYOFF NOTIFY] Telegram HTTP',response.status)
        except Exception as exc:
            print('[HUB DAYOFF NOTIFY]',type(exc).__name__)
        return web.json_response({'ok':True,'request':found})
    except (ValueError,TypeError,KeyError,json.JSONDecodeError):
        return web.json_response({'error':'invalid_request'},status=400)
    except Exception as exc:
        print('[HUB DAYOFF REVIEW]',type(exc).__name__)
        return web.json_response({'error':'review_failed'},status=503)


async def hub_schedule_snapshot_api(request):
    """Read-only display of the same XLSX snapshots used by report exports."""
    from openpyxl import load_workbook
    def read_sheet(filename, limit=1500):
        path = BASE_DIR / filename
        if not path.is_file():
            return {"columns": [], "rows": [], "updated_at": None, "missing": True}
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = book.active
            it = sheet.iter_rows(values_only=True)
            columns = [str(v or '') for v in next(it, ())]
            rows = [[str(v) if v is not None else '' for v in row] for _, row in zip(range(limit), it)]
        finally:
            book.close()
        return {"columns": columns, "rows": rows, "updated_at": datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec='seconds'), "missing": False}
    try:
        return web.json_response({"schedule": read_sheet('work_schedules.xlsx'), "days_off": read_sheet('day_off_requests.xlsx'), "source": "local_xlsx_snapshot"}, dumps=lambda v: json.dumps(v, ensure_ascii=False))
    except Exception:
        return web.json_response({"error": "schedule_snapshot_unavailable"}, status=500)


async def reports(request):
    return web.json_response({"reports": list_reports()}, dumps=lambda x: json.dumps(x, ensure_ascii=False))


async def create_report_from_web(request):
    session = _read_session(request)
    # Веб-генерация доступна обеим внутренним ролям: владельцу и тех. админу.
    # API-key также остаётся допустимым для служебной интеграции.
    if not session and not _api_authorized(request):
        return web.json_response({"error": "forbidden"}, status=403)
    if session and session.get("role") not in ("owner", "deputy"):
        return web.json_response({"error": "forbidden"}, status=403)
    try:
        requested_by = "web:" + (session.get("role") if session else "api")
        report = create_report_snapshot(requested_by=requested_by)
        return web.json_response({"ok": True, "report": report}, dumps=lambda x: json.dumps(x, ensure_ascii=False))
    except Exception as exc:
        return web.json_response({"error": str(exc)}, status=500)


async def report_file(request):
    report_id = request.match_info["report_id"]
    filename = request.match_info["filename"]
    if Path(filename).name != filename or Path(report_id).name != report_id:
        raise web.HTTPBadRequest()
    path = REPORTS_DIR / report_id / filename
    if not path.exists() or not path.is_file():
        raise web.HTTPNotFound()
    return web.FileResponse(path)



SERVER_TAGS = ["Ру серв","Крд серв","Мск серв","Еу серв","Узб серв","Екб серв","Укр серв","Сиб серв","Дв серв","Вн серв","Еу/Ру серв","Ру/Крд серв","Екб/Серв Ру/Мск серв","Сиб/Дв серв","Ру/Еу серв","Ру/Вн серв","Ру/Екб серв","Ру/Сиб/Мск серв","Унив серв"]

def _bot_data_read():
    p=BASE_DIR/'data.json'
    try: return json.loads(p.read_text(encoding='utf-8')) if p.exists() else {}
    except Exception: return {}

def _bot_data_write(data):
    p=BASE_DIR/'data.json'; tmp=BASE_DIR/'data.web.tmp'
    tmp.write_text(json.dumps(data,ensure_ascii=False,indent=4),encoding='utf-8'); tmp.replace(p)
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc: print(f'[WEB DB] {exc}')

async def changelog_api(request):
    path = BASE_DIR / "changelog.json"
    try:
        items = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    except Exception:
        items = []
    items = sorted(items, key=lambda item: str(item.get("date", "")), reverse=True)
    return web.json_response({"items": items}, dumps=lambda obj: json.dumps(obj, ensure_ascii=False))

async def control_snapshot(request):
    d=_bot_data_read(); installed={str(x) for x in d.get('installed_groups',[])}
    groups=[]
    for cid,g in d.get('known_groups',{}).items():
        if cid in installed:
            groups.append({'id':int(cid),'title':g.get('title',''),'status':g.get('bot_status',''),'access':d.get('group_command_access',{}).get(cid,'allow')})
    users=[]
    for uid,u in d.get('users',{}).items():
        users.append({'id':int(uid),'name':u.get('name') or u.get('full_name',''),'username':u.get('username',''),'total_invited':int(u.get('total_invited',0) or 0)})
    assignments=d.get('server_tag_assignments',{})
    
    tag_enabled = bool(d.get('server_tags_enabled', True))
    try:
        from db_storage import get_setting
        v = get_setting('server_tags_enabled', None)
        if isinstance(v, bool): tag_enabled = v
    except Exception: pass
    try:
        from db_storage import get_setting
        auto_reactions_enabled = get_setting('auto_reactions_enabled', True) is not False
    except Exception:
        auto_reactions_enabled = True
    return web.json_response({'counts':{'users':len(users),'groups':len(groups),'tags':len(assignments)},'maintenance':bool(d.get('maintenance_mode')),'groups':groups,'users':users,'assignments':assignments,'server_tags':SERVER_TAGS,'server_tags_enabled':tag_enabled,'auto_reactions_enabled':auto_reactions_enabled},dumps=lambda x:json.dumps(x,ensure_ascii=False))


async def referrals_api(request):
    d = _bot_data_read()
    users = d.get("users", {}) or {}
    links = d.get("links", {}) or {}
    rows = []
    for link, info in links.items():
        owner_id = info.get("owner_id")
        owner = users.get(str(owner_id), {}) if owner_id is not None else {}
        rows.append({
            "link": link,
            "type": info.get("type", "referral"),
            "owner_id": owner_id,
            "name": owner.get("name") or info.get("label") or "Без имени",
            "username": owner.get("username", ""),
            "invited": int(info.get("invited", 0) or 0),
            "created": info.get("created", ""),
        })
    rows.sort(key=lambda x: (-x["invited"], str(x["name"]).lower()))
    return web.json_response({"links": rows}, dumps=lambda x: json.dumps(x, ensure_ascii=False))

async def referral_delete(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "deputy"):
        return web.json_response({"error": "owner_or_tech_only"}, status=403)
    try:
        payload = await request.json()
        link = str(payload.get("link", "")).strip()
    except Exception:
        return web.json_response({"error": "invalid_data"}, status=400)
    d = _bot_data_read(); links = d.get("links", {})
    info = links.get(link)
    if not info:
        return web.json_response({"error": "link_not_found"}, status=404)
    if info.get("type") == "advertising" or link == d.get("advertising_link") or link == ADVERTISING_LINK:
        return web.json_response({"error": "advertising_link_protected"}, status=400)
    # Сначала пробуем отозвать ссылку в Telegram, затем удаляем её из базы сайта/бота.
    revoked = False
    token = os.getenv("BOT_TOKEN", "").strip()
    chat_id = os.getenv("CHAT_ID", "-1004368701081").strip()
    if token:
        try:
            async with aiohttp.ClientSession() as session_http:
                async with session_http.post(f"https://api.telegram.org/bot{token}/revokeChatInviteLink", json={"chat_id": chat_id, "invite_link": link}, timeout=10) as resp:
                    result = await resp.json()
                    revoked = bool(result.get("ok"))
        except Exception:
            revoked = False
    if token and not revoked:
        return web.json_response({"error": "telegram_revoke_failed", "message": "Telegram не подтвердил отзыв ссылки; запись сохранена."}, status=502)
    owner_id = info.get("owner_id")
    del links[link]
    if owner_id is not None:
        u = (d.get("users") or {}).get(str(owner_id), {})
        if isinstance(u.get("links"), list): u["links"] = [x for x in u["links"] if x != link]
    _bot_data_write(d)
    # Web API работает в том же процессе, что и bot.py: синхронизируем и RAM-копию,
    # чтобы следующий save_data() не вернул удалённую ссылку обратно.
    try:
        import sys
        bot_module = sys.modules.get("__main__")
        runtime_data = getattr(bot_module, "data", None)
        if isinstance(runtime_data, dict):
            runtime_data.get("links", {}).pop(link, None)
            if owner_id is not None:
                ru = runtime_data.get("users", {}).get(str(owner_id), {})
                if isinstance(ru.get("links"), list): ru["links"] = [x for x in ru["links"] if x != link]
    except Exception:
        pass
    return web.json_response({"ok": True, "revoked_in_telegram": revoked})

async def joins_analytics_api(request):
    d = _bot_data_read(); history = d.get("join_history", []) or []
    today = datetime.now().date(); days = 30
    try: days = max(7, min(365, int(request.query.get("days", 30))))
    except Exception: pass
    start = today - timedelta(days=days-1)
    counts = {(start + timedelta(days=i)).isoformat(): 0 for i in range(days)}
    for row in history:
        raw = row.get("joined_at") or row.get("joined") or ""
        try: day = datetime.fromisoformat(raw).date()
        except Exception: continue
        key = day.isoformat()
        if key in counts: counts[key] += 1
    today_count = counts.get(today.isoformat(), 0)
    yesterday = today - timedelta(days=1); yesterday_count = counts.get(yesterday.isoformat(), 0)
    delta = today_count - yesterday_count
    total = sum(counts.values())
    series = [{"date": k, "joins": v} for k,v in counts.items()]
    return web.json_response({"today": today_count, "yesterday": yesterday_count, "delta": delta, "period_total": total, "days": days, "series": series}, dumps=lambda x: json.dumps(x, ensure_ascii=False))

MARKET_WALLETS_FILE = BASE_DIR / "market_wallets.json"

def _market_wallets_read():
    try:
        return json.loads(MARKET_WALLETS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _market_wallets_write(data):
    tmp = BASE_DIR / "market_wallets.web.tmp"
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(MARKET_WALLETS_FILE)
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc:
        print(f"[DB] market wallets sync failed: {exc}")

async def market_wallets_api(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "deputy"):
        return web.json_response({"error":"owner_or_tech_only"}, status=403)
    d = _bot_data_read(); wallets = _market_wallets_read(); rows=[]
    for chat_id, chat in (d.get("game_hub") or {}).items():
        for uid, player in (chat.get("players") or {}).items():
            balance = int(wallets.get(str(chat_id), {}).get(str(uid), player.get("balance", 0)))
            try:
                from db_storage import get_setting
                persisted = get_setting(f"market_wallet:{int(chat_id)}:{str(uid)}", None)
                if persisted is not None:
                    balance = max(0, int(persisted))
            except Exception:
                pass
            rows.append({"chat_id":str(chat_id),"user_id":str(uid),"username":player.get("username", ""),"name":player.get("name", ""),"balance":balance})
    return web.json_response({"players":rows}, dumps=lambda x:json.dumps(x,ensure_ascii=False))

async def market_wallet_update(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "deputy"):
        return web.json_response({"error":"owner_or_tech_only"}, status=403)
    try:
        payload=await request.json(); chat_id=str(int(payload.get("chat_id"))); uid=str(int(payload.get("user_id"))); amount=int(payload.get("amount")); mode=str(payload.get("mode","add"))
        if amount < 0 or amount > 1_000_000_000: raise ValueError()
        if mode not in {"add","set"}: raise ValueError()
    except Exception:
        return web.json_response({"error":"invalid_data"}, status=400)
    d=_bot_data_read(); player=((d.get("game_hub") or {}).get(chat_id,{}).get("players") or {}).get(uid)
    if not player: return web.json_response({"error":"player_not_found"}, status=404)
    wallets=_market_wallets_read(); current=int(wallets.get(chat_id,{}).get(uid,player.get("balance",0)))
    new_balance = current + amount if mode == "add" else amount
    wallets.setdefault(chat_id,{})[uid]=new_balance
    _market_wallets_write(wallets)
    # Save the wallet independently in PostgreSQL too. This prevents a bot/web
    # process race from restoring an older JSON snapshot after a site edit.
    try:
        from db_storage import set_setting
        set_setting(f"market_wallet:{int(chat_id)}:{uid}", new_balance)
    except Exception as exc:
        print(f"[DB] market wallet setting save failed: {exc}")
    # Keep game_hub's cached balance consistent for exports/fallbacks.
    try:
        player["balance"] = new_balance
        _bot_data_write(d)
    except Exception as exc:
        print(f"[WEB] game_hub balance cache save failed: {exc}")
    return web.json_response({"ok":True,"balance":new_balance})

async def web_assign_server_tag(request):
    try: payload=await request.json(); chat_id=int(payload.get('chat_id')); user_id=int(payload.get('user_id')); tag=str(payload.get('tag','')).strip()
    except Exception: return web.json_response({'error':'invalid_data'},status=400)
    if tag not in SERVER_TAGS: return web.json_response({'error':'unknown_tag'},status=400)
    token=os.getenv('BOT_TOKEN','').strip()
    if not token: return web.json_response({'error':'BOT_TOKEN_not_configured'},status=500)
    try:
        async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=15)) as session:
            async with session.post(f'https://api.telegram.org/bot{token}/setChatMemberTag',json={'chat_id':chat_id,'user_id':user_id,'tag':tag}) as resp:
                result=await resp.json(content_type=None)
                if not resp.ok or not result.get('ok'): return web.json_response({'error':result.get('description','telegram_error')},status=400)
        d=_bot_data_read(); d.setdefault('server_tag_assignments',{})[f'{chat_id}:{user_id}']={'chat_id':chat_id,'user_id':user_id,'name':'','username':'','tag':tag,'assigned_at':datetime.now().isoformat()}; _bot_data_write(d)
        return web.json_response({'ok':True,'tag':tag})
    except Exception as exc: return web.json_response({'error':str(exc)},status=500)

async def web_auto_reactions_toggle(request):
    try:
        payload = await request.json()
        enabled = payload.get('enabled')
        if not isinstance(enabled, bool):
            return web.json_response({'error':'enabled_must_be_boolean'}, status=400)
        from db_storage import set_setting
        set_setting('auto_reactions_enabled', enabled)
        return web.json_response({'ok':True,'enabled':enabled})
    except Exception as exc:
        return web.json_response({'error':'settings_db_error','detail':str(exc)},status=500)

async def web_server_tags_toggle(request):
    try:
        payload = await request.json()
        enabled = payload.get('enabled')
        if not isinstance(enabled, bool):
            raise ValueError('enabled must be boolean')
    except Exception:
        return web.json_response({'error':'invalid_data'}, status=400)
    # PostgreSQL — общий источник истины между веб-процессом и ботом (даже если это разные контейнеры Railway).
    try:
        from db_storage import set_setting
        set_setting('server_tags_enabled', enabled)
    except Exception as exc:
        return web.json_response({'error':'settings_db_error','detail':str(exc)}, status=500)
    d = _bot_data_read()
    d['server_tags_enabled'] = enabled
    _bot_data_write(d)
    return web.json_response({'ok':True,'enabled':enabled})

async def web_untagged_users(request):
    """Проверяет известных боту пользователей через Telegram и возвращает участников CHAT_ID без тега."""
    d = _bot_data_read()
    token = os.getenv('BOT_TOKEN','').strip()
    try:
        chat_id = int(os.getenv('CHAT_ID') or d.get('main_chat_id') or -1004368701081)
    except Exception:
        chat_id = -1004368701081
    if not token:
        return web.json_response({'error':'BOT_TOKEN_not_configured'}, status=500)

    candidates = {}
    for uid, u in d.get('users', {}).items():
        try: candidates[int(uid)] = {'name':u.get('name') or u.get('full_name',''), 'username':u.get('username','')}
        except Exception: pass
    for rec in d.get('server_tag_assignments', {}).values():
        try: candidates[int(rec.get('user_id'))] = {'name':rec.get('name',''), 'username':rec.get('username','')}
        except Exception: pass
    # Пользователи из заявок тоже считаются известными боту.
    for rec in d.get('pending_join_requests', {}).values():
        try: candidates[int(rec.get('user_id'))] = {'name':rec.get('name',''), 'username':rec.get('username','')}
        except Exception: pass

    untagged=[]; blocked=[]; checked=0
    total_members=0; admin_ids=set()
    timeout=aiohttp.ClientTimeout(total=20)
    async with aiohttp.ClientSession(timeout=timeout) as session:
        # Всегда работаем только с основным чатом. Staff здесь не используется.
        try:
            async with session.post(f'https://api.telegram.org/bot{token}/getChatMemberCount', json={'chat_id':chat_id}) as resp:
                count_result=await resp.json(content_type=None)
            if count_result.get('ok'):
                total_members=int(count_result.get('result') or 0)
        except Exception:
            pass
        # Администраторов основного чата не считаем: у них специальные админские теги.
        try:
            async with session.post(f'https://api.telegram.org/bot{token}/getChatAdministrators', json={'chat_id':chat_id}) as resp:
                admins_result=await resp.json(content_type=None)
            if admins_result.get('ok'):
                admin_ids={int((x.get('user') or {}).get('id')) for x in (admins_result.get('result') or []) if (x.get('user') or {}).get('id')}
        except Exception:
            pass
        for uid, meta in candidates.items():
            if uid in admin_ids:
                continue
            try:
                async with session.post(f'https://api.telegram.org/bot{token}/getChatMember', json={'chat_id':chat_id,'user_id':uid}) as resp:
                    result=await resp.json(content_type=None)
                if not result.get('ok'): continue
                member=result.get('result') or {}; status=str(member.get('status','')); checked += 1
                if status == 'kicked':
                    blocked.append({'id':uid,'name':meta.get('name',''),'username':meta.get('username',''),'reason':'ЧС / заблокирован в чате'})
                    continue
                if status in ('left',''): continue
                tag=str(member.get('tag') or '').strip()
                if not tag:
                    user=member.get('user') or {}
                    untagged.append({'id':uid,'name':user.get('first_name') or meta.get('name',''),'username':user.get('username') or meta.get('username','')})
            except Exception:
                continue
    return web.json_response({'chat_id':chat_id,'total_members':total_members,'admins_excluded':len(admin_ids),'known_candidates':len(candidates),'checked':checked,'untagged':untagged,'blocked':blocked}, dumps=lambda x:json.dumps(x,ensure_ascii=False))


async def web_group_access(request):
    try: payload=await request.json(); chat_id=int(payload.get('chat_id')); access=str(payload.get('access'))
    except Exception: return web.json_response({'error':'invalid_data'},status=400)
    if access not in ('allow','deny'): return web.json_response({'error':'invalid_access'},status=400)
    d=_bot_data_read(); d.setdefault('group_command_access',{})[str(chat_id)]=access; _bot_data_write(d)
    # Сайт и Telegram-бот работают в одном процессе. Обновляем RAM-копию сразу,
    # чтобы переключатель применялся без рестарта и следующий save_data() не откатил настройку.
    try:
        import sys
        bot_module = sys.modules.get('__main__')
        runtime_data = getattr(bot_module, 'data', None)
        if isinstance(runtime_data, dict):
            runtime_data.setdefault('group_command_access', {})[str(chat_id)] = access
    except Exception:
        pass
    return web.json_response({'ok':True,'access':access})


def _json_safe(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {"__type__": "bytes", "base64": base64.b64encode(bytes(value)).decode("ascii")}
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass
    return str(value)


def _export_postgres_full():
    database_url = os.getenv("DATABASE_URL", "").strip()
    if not database_url:
        return {"configured": False, "tables": {}}
    import psycopg2
    result = {"configured": True, "exported_at": datetime.now().isoformat(timespec="seconds"), "tables": {}}
    with psycopg2.connect(database_url, connect_timeout=15) as conn, conn.cursor() as cur:
        cur.execute("""SELECT tablename FROM pg_catalog.pg_tables WHERE schemaname='public' ORDER BY tablename""")
        tables = [row[0] for row in cur.fetchall()]
        from psycopg2 import sql
        for table in tables:
            cur.execute(sql.SQL("SELECT * FROM {} ORDER BY 1").format(sql.Identifier(table)))
            cols = [d.name for d in cur.description]
            rows = [[_json_safe(v) for v in row] for row in cur.fetchall()]
            result["tables"][table] = {"columns": cols, "rows": rows, "row_count": len(rows)}
    return result


def _build_full_backup_zip():
    # Сначала принудительно синхронизируем локальные данные с PostgreSQL.
    try:
        from db_storage import sync_to_db
        sync_to_db()
    except Exception as exc:
        print(f"[BACKUP] sync_to_db: {exc}")
    db_export = _export_postgres_full()
    out = io.BytesIO()
    skip_dirs = {"__pycache__", ".git", ".idea", ".vscode", ".venv", "venv"}
    skip_names = {".env"}
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for src in BASE_DIR.rglob("*"):
            if not src.is_file():
                continue
            rel = src.relative_to(BASE_DIR)
            if any(part in skip_dirs for part in rel.parts):
                continue
            if src.name in skip_names or src.suffix.lower() in {".pyc", ".pyo"}:
                continue
            if src.name.startswith("shinobi-full-backup-") and src.suffix.lower()==".zip":
                continue
            zf.write(src, (Path("shinobi-bot-main") / rel).as_posix())
        zf.writestr("BACKUP/database_full_export.json", json.dumps(db_export, ensure_ascii=False, indent=2))
        zf.writestr("BACKUP/README.txt", "SHINOBI TEAM FULL BACKUP\nСоздан: " + datetime.now().isoformat(timespec="seconds") + "\nСодержит актуальные файлы бота и полный экспорт таблиц PostgreSQL.\n.env и секреты окружения намеренно не включены.\n")
    return out.getvalue()


async def download_full_backup(request):
    session = _read_session(request)
    if not session or session.get("role") not in ("owner", "deputy"):
        return web.json_response({"error": "forbidden"}, status=403)
    expected = os.getenv("BACKUP_EXPORT_CODE", "").strip()
    if not expected:
        return web.json_response({"error": "BACKUP_EXPORT_CODE_not_configured"}, status=503)
    try:
        payload = await request.json()
        code = str(payload.get("code", ""))
    except Exception:
        return web.json_response({"error": "invalid_data"}, status=400)
    if not secrets.compare_digest(code, expected):
        return web.json_response({"error": "invalid_backup_code"}, status=401)
    try:
        blob = await __import__('asyncio').to_thread(_build_full_backup_zip)
    except Exception as exc:
        print(f"[BACKUP] Ошибка: {exc}")
        return web.json_response({"error": "backup_failed"}, status=500)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"shinobi-full-backup-{stamp}.zip"
    return web.Response(body=blob, content_type="application/zip", headers={"Content-Disposition": f'attachment; filename="{filename}"', "Cache-Control": "no-store"})



async def ai_moderation_log_api(request):
    try:
        payload=json.loads((BASE_DIR / "data.json").read_text(encoding="utf-8"))
    except Exception:
        payload={}
    records=list((payload.get("shinobi_moderation",{}) or {}).get("records",[]) or [])[-500:]
    out=[]
    for x in reversed(records):
        y=dict(x)
        try: y["time"]=datetime.fromtimestamp(int(y.get("ts",0))).strftime("%d.%m.%Y %H:%M:%S")
        except Exception: y["time"]="—"
        kind=y.get("kind",""); dur=int(y.get("duration",0) or 0)
        if kind in {"pred","warn"}: action={"pred":"ПРЕД • 7 дней","warn":"WARN • 7 дней"}[kind]
        elif kind=="mute": action=f"Мут {max(1,dur//60)} мин."
        elif kind=="ban": action="Бан навсегда" if not dur else f"Бан {max(1,dur//3600)} ч."
        else: action=kind
        y["action_text"]=action + (" • снято" if y.get("revoked") else "")
        y["user_name"]=y.get("user_name") or str(y.get("user_id", ""))
        y["message_text"]=y.get("message_text") or "—"
        y["chat_title"]=y.get("chat_title") or str(y.get("chat_id", ""))
        y["confidence"]=y.get("confidence") or 0
        y["source_text"]="AI" if y.get("source")=="ai" else ("SYSTEM" if y.get("source")=="system" else y.get("actor_name","ADMIN"))
        out.append(y)
    return web.json_response({"beta":True,"log":out},dumps=lambda x:json.dumps(x,ensure_ascii=False))

async def chat_summary_api(request):
    rows_by_id = {int(x["chat_id"]): x for x in chat_summary_snapshot()}
    try:
        bot_data = json.loads((BASE_DIR / "data.json").read_text(encoding="utf-8"))
    except Exception:
        bot_data = {}
    for raw_id in bot_data.get("installed_groups", []):
        cid = int(raw_id); meta = (bot_data.get("known_groups", {}) or {}).get(str(cid), {})
        rows_by_id.setdefault(cid, {"chat_id": cid, "title": meta.get("title", str(cid)), "messages_today": 0, "peak_hour": "—", "peak_messages": 0, "active_15m": 0, "active_60m": 0, "tracked_total_messages": 0})
    rows = list(rows_by_id.values()); token = os.getenv("BOT_TOKEN", "").strip()
    async with aiohttp.ClientSession() as session:
        for row in rows:
            row["members"] = None; row["ping_ms"] = None
            if not token: continue
            try:
                started = time.perf_counter()
                async with session.get(f"https://api.telegram.org/bot{token}/getChatMemberCount", params={"chat_id": row["chat_id"]}, timeout=8) as resp: payload = await resp.json()
                row["ping_ms"] = round((time.perf_counter() - started) * 1000)
                if payload.get("ok"): row["members"] = payload.get("result")
            except Exception: pass
    rows.sort(key=lambda x: str(x.get("title", "")))
    return web.json_response({"chats": rows, "uptime_seconds": uptime_seconds(), "updated_at": datetime.now().isoformat(timespec="seconds")}, dumps=lambda x: json.dumps(x, ensure_ascii=False))


# V13: operational functions previously accessible only via Telegram inline buttons.
def _v13_runtime():
    import sys
    runtime = sys.modules.get('__main__')
    if runtime is None or not isinstance(getattr(runtime, 'data', None), dict):
        raise RuntimeError('Telegram bot runtime unavailable')
    return runtime

async def tech_users_api(request):
    d = _bot_data_read()
    users = []
    for uid, u in (d.get('users') or {}).items():
        if not isinstance(u, dict): continue
        users.append({'id':str(uid),'name':str(u.get('name') or u.get('full_name') or ''),
                      'username':str(u.get('username') or ''),'invited':int(u.get('total_invited',0) or 0),
                      'bot_admin':str(uid) in {str(x) for x in d.get('bot_admins',[])}})
    users.sort(key=lambda x:(-x['invited'],x['name']))
    return web.json_response({'users':users[:3000],'total':len(users)},dumps=lambda v:json.dumps(v,ensure_ascii=False))

async def tech_join_requests_api(request):
    d = _bot_data_read()
    pending = []
    for key, req in (d.get('pending_join_requests') or {}).items():
        if not isinstance(req,dict) or req.get('status') != 'pending': continue
        pending.append({'key':str(key),'chat_id':req.get('chat_id'),'user_id':req.get('user_id'),
                        'name':str(req.get('name') or ''),'request_id':str(req.get('request_id') or ''),
                        'created_at':str(req.get('created_at') or '')})
    return web.json_response({'requests':pending[-300:][::-1],'total':len(pending)},dumps=lambda v:json.dumps(v,ensure_ascii=False))

async def tech_join_review_api(request):
    try:
        payload=await request.json(); chat_id=int(payload['chat_id']); user_id=int(payload['user_id'])
        decision=str(payload['decision']); expected=str(payload['request_id'])
        if decision not in ('approve','decline'): raise ValueError('invalid decision')
    except (ValueError,TypeError,KeyError):
        return web.json_response({'error':'invalid_request'},status=400)
    try: runtime=_v13_runtime()
    except RuntimeError: return web.json_response({'error':'bot_not_running'},status=503)
    key=f'{chat_id}:{user_id}'
    req=runtime.data.get('pending_join_requests',{}).get(key)
    if not req or req.get('status')!='pending' or str(req.get('request_id') or '')!=expected:
        return web.json_response({'error':'request_changed_or_processed'},status=409)
    try:
        if decision=='approve': await runtime.bot.approve_chat_join_request(chat_id,user_id)
        else: await runtime.bot.decline_chat_join_request(chat_id,user_id)
    except Exception as exc:
        return web.json_response({'error':'telegram_rejected','detail':str(exc)[:250]},status=502)
    from datetime import datetime as _dt
    session=_read_session(request)
    req.update(status='approved' if decision=='approve' else 'declined',decided_by=session['uid'],
               decided_by_name='Web Hub',decided_at=_dt.now().isoformat(),decided_in='Web Hub')
    runtime.save_data()
    sync=getattr(runtime,'_sync_join_request_messages',None)
    if sync:
        try: await sync(req)
        except Exception as exc: print('[HUB JOIN SYNC]',exc)
    return web.json_response({'ok':True,'status':req['status']})

async def tech_invites_update_api(request):
    # Owner/deputy only: set a participant's referral count while preserving audit trail.
    try:
        body=await request.json(); uid=str(int(body['user_id'])); new=int(body['invited'])
        if new<0 or new>10000000: raise ValueError('invalid count')
    except (KeyError,TypeError,ValueError):return web.json_response({'error':'invalid_count_or_user'},status=400)
    try: runtime=_v13_runtime()
    except RuntimeError:return web.json_response({'error':'bot_not_running'},status=503)
    user=runtime.data.get('users',{}).get(uid)
    if not isinstance(user,dict):return web.json_response({'error':'user_not_found'},status=404)
    old=int(user.get('total_invited',0) or 0); delta=new-old
    user['total_invited']=new
    links=runtime.data.get('links',{})
    primary=next((link for link in user.get('links',[]) if link in links),None)
    if primary: links[primary]['invited']=max(0,int(links[primary].get('invited',0) or 0)+delta)
    runtime.data.setdefault('invite_adjustments',[]).append({'user_id':int(uid),'old':old,'new':new,
       'delta':delta,'source':'web_hub','admin_id':_read_session(request)['uid'],
       'at':datetime.now().isoformat()})
    runtime.save_data()
    return web.json_response({'ok':True,'old':old,'new':new})

async def tech_export_api(request):
    from io import BytesIO
    from openpyxl import Workbook
    kind=request.match_info['kind']
    files={'referrals':'shinobi_links_log.xlsx','schedule':'work_schedules.xlsx','dayoffs':'day_off_requests.xlsx'}
    if kind not in files: return web.json_response({'error':'unknown_export'},status=404)
    path=BASE_DIR/files[kind]
    if not path.exists():return web.json_response({'error':'export_unavailable'},status=404)
    return web.FileResponse(path,headers={'Content-Disposition':f'attachment; filename="{files[kind]}"'})

async def start_stats_api():
    app = web.Application(middlewares=[auth_middleware])
    app.router.add_get("/", login_page)
    app.router.add_get("/auth/telegram", telegram_login)
    app.router.add_get("/auth/telegram/start", oidc_start)
    app.router.add_get("/auth/telegram/callback", oidc_callback)
    app.router.add_get("/api/login-audit", login_audit)
    app.router.add_get("/api/sessions", sessions_api)
    app.router.add_post("/api/sessions/revoke-others", revoke_others_api)
    app.router.add_get("/api/staff", staff_api)
    app.router.add_post("/api/staff/role", staff_role_api)
    app.router.add_post("/api/revoke-user", revoke_user_api)
    app.router.add_get("/logout", logout)
    app.router.add_get("/dashboard", dashboard)
    app.router.add_get("/api/status", status)
    app.router.add_get("/api/tech/users", tech_users_api)
    app.router.add_get("/api/tech/join-requests", tech_join_requests_api)
    app.router.add_post("/api/tech/join-review", tech_join_review_api)
    app.router.add_post("/api/tech/invites", tech_invites_update_api)
    app.router.add_get("/api/tech/export/{kind}", tech_export_api)
    app.router.add_get("/api/system-info", system_info)
    app.router.add_get("/api/chat-summary", chat_summary_api)
    app.router.add_get("/api/ai-moderation-log", ai_moderation_log_api)
    app.router.add_get("/api/october-plan", october_plan_api)
    app.router.add_post("/api/october-plan", october_plan_update)
    app.router.add_post("/api/system/version", update_version)
    app.router.add_get("/media/{kind}", media_file)
    app.router.add_post("/api/media/{kind}", upload_media)
    app.router.add_get("/api/hub/schedule-snapshot", hub_schedule_snapshot_api)
    app.router.add_get("/api/hub/dayoff-requests", hub_dayoff_pending_api)
    app.router.add_post("/api/hub/dayoff-review", hub_dayoff_review_api)
    app.router.add_get("/api/hub/schedule-live", hub_schedule_live_api)
    app.router.add_post("/api/hub/schedule-live", hub_schedule_save_api)
    app.router.add_get("/api/reports", reports)
    app.router.add_get("/api/control", control_snapshot)
    app.router.add_get("/api/referrals", referrals_api)
    app.router.add_post("/api/referrals/delete", referral_delete)
    app.router.add_get("/api/joins-analytics", joins_analytics_api)
    app.router.add_get("/api/changelog", changelog_api)
    app.router.add_post("/api/server-tags/assign", web_assign_server_tag)
    app.router.add_post("/api/control/server-tags", web_server_tags_toggle)
    app.router.add_post("/api/control/auto-reactions", web_auto_reactions_toggle)
    app.router.add_get("/api/control/untagged", web_untagged_users)
    app.router.add_post("/api/control/group-access", web_group_access)
    app.router.add_get("/api/market/wallets", market_wallets_api)
    app.router.add_post("/api/market/wallet", market_wallet_update)
    app.router.add_post("/api/backup/full", download_full_backup)
    app.router.add_post("/api/reports", create_report_from_web)
    app.router.add_get("/api/reports/{report_id}/{filename}", report_file)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.getenv("PORT", os.getenv("SHINOBI_STATS_PORT", "8080")))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    print(f"SHINOBI TEAM HUB WEB + API запущен на порту {port}")
    return runner
