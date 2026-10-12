import asyncio, json, os, random, time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
STATS_FILE = BASE_DIR / 'chat_summary.json'
BOT_STARTED_AT = time.time()
_lock = asyncio.Lock()

MLBB_FACTS = [
 'Mobile Legends: Bang Bang вышла в 2016 году.',
 'Классический матч MLBB строится вокруг уничтожения вражеской базы, а не только количества убийств.',
 'Черепаха и Лорд — ключевые нейтральные объекты, которые могут сильно изменить темп матча.',
 'Карта MLBB разделена на три основные линии: EXP, Mid и Gold, плюс лес.',
 'Роумер обычно помогает создавать численное преимущество и контролировать карту.',
 'Gold Lane даёт дополнительное золото на ранней стадии и особенно важна для многих стрелков.',
 'EXP Lane ускоряет получение опыта и часто используется бойцами и танковыми героями.',
 'В MLBB кусты скрывают героя от обычного обзора противника, пока он не будет обнаружен.',
 'Контроль Лорда часто важнее случайной погони за убийствами: он помогает давить линии.',
 'Правильная ротация после зачистки линии позволяет команде быстрее создавать преимущество на карте.',
 'Battle Spells подбираются под героя, роль и конкретный матч — универсального выбора для всех нет.',
 'Эмблемы позволяют дополнительно адаптировать характеристики и таланты героя под стиль игры.',
 'Система рангов MLBB использует звёзды/очки и сезонные сбросы, поэтому ранг меняется между сезонами.',
 'Контроль зрения через позиционирование и проверку кустов помогает избегать неожиданных гангов.',
 'Даже при преимуществе по убийствам команда может проиграть, если теряет башни, объекты и базу.',
]

def _load():
    try:
        d=json.loads(STATS_FILE.read_text(encoding='utf-8')) if STATS_FILE.exists() else {}
    except Exception: d={}
    return d if isinstance(d,dict) else {}

def _save(d):
    tmp=STATS_FILE.with_suffix('.tmp')
    tmp.write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf-8')
    tmp.replace(STATS_FILE)

def _day(): return datetime.now().strftime('%Y-%m-%d')
def _hour(): return datetime.now().strftime('%H:00')

def record_message(chat_id:int,title:str,user_id:int|None):
    d=_load(); cid=str(chat_id); day=_day(); now=time.time()
    c=d.setdefault(cid,{'title':title,'days':{},'recent_users':{},'total_seen_messages':0})
    c['title']=title or c.get('title',''); c['total_seen_messages']=int(c.get('total_seen_messages',0))+1
    dd=c.setdefault('days',{}).setdefault(day,{'messages':0,'hours':{}}); dd['messages']=int(dd.get('messages',0))+1
    h=_hour(); dd.setdefault('hours',{})[h]=int(dd.get('hours',{}).get(h,0))+1
    if user_id: c.setdefault('recent_users',{})[str(user_id)]=now
    # keep storage compact
    for old in sorted(c.get('days',{}))[:-31]: c['days'].pop(old,None)
    cutoff=now-86400
    c['recent_users']={k:v for k,v in c.get('recent_users',{}).items() if float(v)>=cutoff}
    _save(d)

def snapshot():
    d=_load(); now=time.time(); out=[]
    for cid,c in d.items():
        dd=c.get('days',{}).get(_day(),{}); hours=dd.get('hours',{}) or {}
        peak=max(hours.items(),key=lambda x:x[1]) if hours else ('—',0)
        out.append({'chat_id':int(cid),'title':c.get('title',''),'messages_today':int(dd.get('messages',0)),'peak_hour':peak[0],'peak_messages':int(peak[1]),'active_15m':sum(1 for v in c.get('recent_users',{}).values() if now-float(v)<=900),'active_60m':sum(1 for v in c.get('recent_users',{}).values() if now-float(v)<=3600),'tracked_total_messages':int(c.get('total_seen_messages',0))})
    return out

def ping_ms(): return max(0,int((time.time()-time.time())*1000))
def uptime_seconds(): return int(time.time()-BOT_STARTED_AT)
def random_fact(): return random.choice(MLBB_FACTS)
