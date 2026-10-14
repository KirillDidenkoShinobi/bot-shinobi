"""One-time Telegram-confirmed web login and server-side RBAC."""
import os, time, secrets, hashlib, sqlite3, threading
from pathlib import Path

DB_PATH = Path(__file__).with_name('hub_auth.sqlite3')
LOCK = threading.RLock()
ROLES = ('owner','deputy','admin','moderator')
IDLE_TIMEOUT = 30 * 60
SESSION_MAX_AGE = 12 * 60 * 60

def _db():
    url=os.getenv('DATABASE_URL','').strip()
    if url:
        import psycopg2
        c=psycopg2.connect(url,connect_timeout=8)
        return c, '%s'
    c=sqlite3.connect(str(DB_PATH),timeout=15)
    return c,'?'

def _execute(sql,args=(),fetch=False):
    with LOCK:
        c,placeholder=_db()
        try:
            cur=c.cursor()
            if placeholder=='%s': sql=sql.replace('?', '%s')
            cur.execute(sql,args)
            result=cur.fetchone() if fetch else None
            c.commit()
            return result
        finally: c.close()

def init():
    _execute('CREATE TABLE IF NOT EXISTS hub_login_tokens (digest TEXT PRIMARY KEY, uid BIGINT NOT NULL, expires BIGINT NOT NULL, consumed BIGINT)')
    _execute('CREATE TABLE IF NOT EXISTS hub_sessions (digest TEXT PRIMARY KEY, uid BIGINT NOT NULL, expires BIGINT NOT NULL, revoked BIGINT)')
    # Safe migration for existing PostgreSQL and SQLite installations.
    try:_execute('ALTER TABLE hub_sessions ADD COLUMN last_seen BIGINT')
    except Exception:pass
    _execute('CREATE TABLE IF NOT EXISTS hub_oidc_states (digest TEXT PRIMARY KEY, verifier TEXT NOT NULL, expires BIGINT NOT NULL, consumed BIGINT)')
    _execute('CREATE TABLE IF NOT EXISTS hub_login_audit (id BIGINT PRIMARY KEY, uid BIGINT NOT NULL, event TEXT NOT NULL, at BIGINT NOT NULL, ip TEXT)')

def role(uid):
    uid=str(uid)
    if uid in ('1197818066', *os.getenv('HUB_OWNER_IDS','').replace(' ','').split(',')):return 'owner'
    for role_name, env in [('deputy','HUB_DEPUTY_IDS'),('admin','HUB_ADMIN_IDS'),('moderator','HUB_MODERATOR_IDS')]:
        if uid in os.getenv(env,'').replace(' ','').split(','):return role_name
    return None

def audit(uid,event,ip=''):
    try:_execute('INSERT INTO hub_login_audit(id,uid,event,at,ip) VALUES(?,?,?,?,?)',(secrets.randbits(62),int(uid),event,int(time.time()),str(ip)[:100]))
    except Exception as exc:print('[HUB AUTH AUDIT]',exc)

def issue(uid):
    init()
    if not role(uid):return None
    token=secrets.token_urlsafe(36);digest=hashlib.sha256(token.encode()).hexdigest()
    _execute('INSERT INTO hub_login_tokens(digest,uid,expires,consumed) VALUES(?,?,?,NULL)',(digest,int(uid),int(time.time())+300))
    audit(uid,'link_issued')
    return token

def redeem(token,ip=''):
    init()
    if not token or len(token)>256:return None
    digest=hashlib.sha256(token.encode()).hexdigest()
    with LOCK:
        c,ph=_db()
        try:
            cur=c.cursor()
            sql='UPDATE hub_login_tokens SET consumed=? WHERE digest=? AND consumed IS NULL AND expires>=? RETURNING uid'
            if ph=='%s':sql=sql.replace('?','%s')
            cur.execute(sql,(int(time.time()),digest,int(time.time())))
            row=cur.fetchone()
            if not row or not role(row[0]):c.rollback();return None
            session=secrets.token_urlsafe(40);session_hash=hashlib.sha256(session.encode()).hexdigest()
            sql='INSERT INTO hub_sessions(digest,uid,expires,revoked,last_seen) VALUES(?,?,?,NULL,?)'
            if ph=='%s':sql=sql.replace('?','%s')
            cur.execute(sql,(session_hash,int(row[0]),int(time.time())+SESSION_MAX_AGE,int(time.time())))
            c.commit()
            audit(row[0],'login',ip)
            return session
        finally:c.close()

def session(token):
    if not token or len(token)>256:return None
    init()
    digest=hashlib.sha256(token.encode()).hexdigest()
    now=int(time.time())
    with LOCK:
        c,ph=_db()
        try:
            cur=c.cursor()
            sql='SELECT uid,expires,revoked,last_seen FROM hub_sessions WHERE digest=?'
            cur.execute(sql.replace('?',ph),(digest,))
            row=cur.fetchone()
            if not row or row[2] is not None or int(row[1])<=now:return None
            last_seen=int(row[3]) if row[3] is not None else max(now-1,int(row[1])-SESSION_MAX_AGE)
            if now-last_seen>=IDLE_TIMEOUT:
                cur.execute('UPDATE hub_sessions SET revoked=? WHERE digest=?'.replace('?',ph),(now,digest))
                c.commit()
                audit(row[0],'session_idle_expired')
                return None
            current=role(row[0])
            if not current:return None
            if now-last_seen>=60:
                cur.execute('UPDATE hub_sessions SET last_seen=? WHERE digest=?'.replace('?',ph),(now,digest))
                c.commit()
            return {'uid':int(row[0]),'role':current}
        finally:c.close()

def revoke(token,ip=''):
    s=session(token)
    if s:
        _execute('UPDATE hub_sessions SET revoked=? WHERE digest=?',(int(time.time()),hashlib.sha256(token.encode()).hexdigest()))
        audit(s['uid'],'logout',ip)

def revoke_user(uid):
    init();_execute('UPDATE hub_sessions SET revoked=? WHERE uid=? AND revoked IS NULL',(int(time.time()),int(uid)))
    audit(uid,'sessions_revoked')

def recent(limit=100):
    init()
    with LOCK:
        c,ph=_db()
        try:
            cur=c.cursor();cur.execute('SELECT uid,event,at,ip FROM hub_login_audit ORDER BY at DESC LIMIT '+str(max(1,min(int(limit),500))))
            return [{'uid':r[0],'event':r[1],'at':r[2],'ip':r[3]} for r in cur.fetchall()]
        finally:c.close()


def oidc_state_create():
    init()
    state=secrets.token_urlsafe(32)
    verifier=secrets.token_urlsafe(48)
    _execute('INSERT INTO hub_oidc_states(digest,verifier,expires,consumed) VALUES(?,?,?,NULL)',(hashlib.sha256(state.encode()).hexdigest(),verifier,int(time.time())+300))
    return state,verifier

def oidc_state_consume(state):
    init()
    if not state or len(state)>256:return None
    with LOCK:
        c,ph=_db()
        try:
            cur=c.cursor()
            sql='UPDATE hub_oidc_states SET consumed=? WHERE digest=? AND consumed IS NULL AND expires>=? RETURNING verifier'
            if ph=='%s':sql=sql.replace('?','%s')
            cur.execute(sql,(int(time.time()),hashlib.sha256(state.encode()).hexdigest(),int(time.time())))
            row=cur.fetchone();c.commit()
            return row[0] if row else None
        finally:c.close()

def oidc_session_create(uid,ip=''):
    init()
    if not role(uid):
        audit(uid,'oidc_access_denied',ip)
        print('[HUB OIDC ACCESS DENIED] uid=%s role=none; verify HUB_OWNER_IDS / signed Telegram account' % uid)
        return None
    session=secrets.token_urlsafe(40)
    _execute('INSERT INTO hub_sessions(digest,uid,expires,revoked,last_seen) VALUES(?,?,?,NULL,?)',(hashlib.sha256(session.encode()).hexdigest(),int(uid),int(time.time())+SESSION_MAX_AGE,int(time.time())))
    audit(uid,'oidc_login',ip)
    return session
