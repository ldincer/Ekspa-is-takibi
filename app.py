#!/usr/bin/env python3
"""Ekspa İş Takibi — Satış (fırsat/proje) + Teknik Servis (çağrı karşılama) tek uygulamada.

Yalnızca Python 3.8+ standart kütüphanesini kullanır; ek kurulum gerekmez.
Veriler data/ekspa.db (SQLite) dosyasında tutulur. İlk açılışta, aynı klasörde
data/firsatlar.db (Ekspa Fırsat Takibi) ve/veya data/cagri.db (Çağrı Karşılama)
varsa verileri ve kullanıcıları otomatik taşınır; eski dosyalara dokunulmaz.

    python app.py                                   # http://0.0.0.0:8080
    python app.py --port 9000
    python app.py --import Ekspa-Proje.xlsx         # Odoo CRM fırsat dışa aktarımı (openpyxl gerekir)
    python app.py --import-envanter Envanter.xlsx   # Odoo envanter dışa aktarımı (openpyxl gerekir)
    python app.py --import-talepler Talepler.xlsx   # Odoo yardım masası dışa aktarımı (openpyxl gerekir)
    python app.py --kullanici-ekle ali --satis duzenleyici --servis uzman
    python app.py --kullanici-ekle levent --yonetici
    python app.py --sifre-sifirla ali
    python app.py --kullanicilar
"""
import argparse
import getpass
import hashlib
import hmac
import json
import os
import queue
import re
import secrets
import signal
import smtplib
import sqlite3
import ssl
import sys
import threading
import time
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(BASE, "static")
# Railway (railway.com) üzerinde çalışıyor muyuz? Railway bu değişkenleri her servise kendisi ekler.
ON_RAILWAY = bool(os.environ.get("RAILWAY_ENVIRONMENT_ID") or os.environ.get("RAILWAY_ENVIRONMENT"))
# Railway'de servise bir Volume bağlıysa veriler oraya yazılır (bağlama yolu ne olursa olsun).
DATA_DIR = (os.environ.get("RAILWAY_VOLUME_MOUNT_PATH") or os.environ.get("EKSPA_DATA_DIR")
            or os.path.join(BASE, "data"))
DB_PATH = os.path.join(DATA_DIR, "ekspa.db")
LEGACY_CRM = os.path.join(DATA_DIR, "firsatlar.db")     # Ekspa Fırsat Takibi sürüm 1-3
LEGACY_SRV = os.path.join(DATA_DIR, "cagri.db")         # Çağrı Karşılama sürüm 1
SEED_CRM = os.path.join(BASE, "data", "baslangic_firsatlar.json")
SEED_INV = os.path.join(BASE, "data", "baslangic_envanter.json")
SEED_TIC = os.path.join(BASE, "data", "baslangic_talepler.json")
# HTTPS arkasında çalışırken oturum çerezi "Secure" işaretlenir. Railway TLS'i kendisi sonlandırdığı için orada varsayılan açıktır.
HTTPS = (os.environ.get("EKSPA_HTTPS") == "1" or os.environ.get("CAGRI_HTTPS") == "1"
         or (ON_RAILWAY and os.environ.get("EKSPA_HTTPS") != "0"))


def norm_code(v):
    return re.sub(r"[^0-9A-Za-z]", "", str(v or "")).upper()


# İlk yönetici hesabı için kurulum kodu. İnternete açık kurulumda (Railway) adresi ilk bulan kişinin
# yönetici olmasını engeller. EKSPA_KURULUM_KODU verilmemişse Railway'de açılışta rastgele üretilip loglara yazılır.
SETUP_CODE_ENV = norm_code(os.environ.get("EKSPA_KURULUM_KODU"))
SETUP_CODE = SETUP_CODE_ENV or (secrets.token_hex(6).upper() if ON_RAILWAY else "")


def public_base(st):
    """E-postalardaki bağlantılar için uygulama adresi: Ayarlar'daki adres, yoksa Railway'in verdiği alan adı."""
    base = st.get("base_url") or ""
    if not base and os.environ.get("RAILWAY_PUBLIC_DOMAIN"):
        base = f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}/servis"
    return base.rstrip("/")

# Modül bazlı yetkiler. Yönetici her iki modülde tam yetkilidir.
CRM_LV = {"": 0, "goruntuleyici": 1, "duzenleyici": 2}
SRV_LV = {"": 0, "goruntuleyici": 1, "uzman": 2}
CRM_T = {"": "Erişim yok", "goruntuleyici": "Görüntüleyici", "duzenleyici": "Düzenleyici"}
SRV_T = {"": "Erişim yok", "goruntuleyici": "Görüntüleyici", "uzman": "Uzman"}
STAGES = [
    ("acik", "Çağrı açıldı"),
    ("inceleniyor", "İnceleniyor"),
    ("parca_siparis", "Parça sipariş edildi"),
    ("parca_geldi", "Parça teslim alındı"),
    ("mudahale", "Müdahale / onarım"),
    ("musteri_bekleniyor", "Müşteri yanıtı bekleniyor"),
    ("cozuldu", "Çözüldü"),
    ("iade", "Cihaz müşteriye iade edildi"),
    ("kapandi", "Kapandı"),
    ("iptal", "İptal edildi"),
]
STAGE_T = dict(STAGES)
CLOSED = {"cozuldu", "iade", "kapandi", "iptal"}
NEEDS_SOLUTION = {"cozuldu", "iade", "kapandi"}
PRIORITIES = {"acil": "Acil", "yuksek": "Yüksek", "normal": "Normal", "dusuk": "Düşük"}
# Değiştiğinde müşteriye bildirilen alanlar ve görünen adları
TRACKED = {"stage": "Aşama", "priority": "Öncelik", "assignee": "Sorumlu uzman", "sla_due": "Hedef çözüm tarihi",
           "subject": "Konu", "description": "Arıza tanımı", "part_no": "Parça no", "order_no": "Sipariş no",
           "rma_no": "RMA / iade no", "cargo_no": "Kargo takip no", "solution": "Çözüm açıklaması",
           "contact_name": "İletişim kişisi", "contact_phone": "Telefon", "location": "Konum"}
EDITABLE = set(TRACKED) | {"contact_email"}
SESSION_SHORT, SESSION_LONG = 12 * 3600, 30 * 24 * 3600
PBKDF2_ITER = 240_000
COOKIE = "ekspa_oturum"
LOCK = threading.Lock()
FAILS = {}
MAILQ = queue.Queue()
USERNAME_RE = re.compile(r"^[a-z0-9._-]{3,32}$")
EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def today():
    return date.today().isoformat()


def add_business_days(d, n):
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def norm_serial(s):
    s = str(s or "").strip().upper()
    s = re.sub(r"^[:\-\s]+", "", s)
    return re.sub(r"\s+", "", s)


def warranty(start, end, on=None):
    """Garanti durumunu döndürür: kod, açıklama."""
    on = on or today()
    if not start and not end:
        return "bilgi_yok", "Garanti tarihleri girilmemiş"
    if not end:
        return "bilgi_yok", "Garanti bitiş tarihi girilmemiş"
    if start and on < start:
        return "baslamadi", f"Garanti {fmt_d(start)} tarihinde başlıyor"
    if on > end:
        return "bitti", f"Garanti {fmt_d(end)} tarihinde sona ermiş"
    left = (date.fromisoformat(end) - date.fromisoformat(on)).days
    return "kapsamda", f"Garanti kapsamında · bitiş {fmt_d(end)} ({left} gün kaldı)"


def fmt_d(v):
    return f"{v[8:10]}.{v[5:7]}.{v[0:4]}" if v else ""


# ---------------------------------------------------------------- veritabanı
def connect():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA foreign_keys=ON")
    return con


SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS settings (k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT NOT NULL, email TEXT, role TEXT NOT NULL,
  pw TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1, must_change INTEGER NOT NULL DEFAULT 0, created_at TEXT, last_login TEXT,
  crm_role TEXT NOT NULL DEFAULT '', servis_role TEXT NOT NULL DEFAULT '');
CREATE TABLE IF NOT EXISTS deals (id TEXT PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sessions (token_hash TEXT PRIMARY KEY, username TEXT NOT NULL, expires REAL NOT NULL, created_at TEXT);
CREATE TABLE IF NOT EXISTS inventory (id INTEGER PRIMARY KEY, firma TEXT, serial TEXT, serial_norm TEXT, brand TEXT, part_no TEXT,
  description TEXT, service TEXT, w_start TEXT, w_end TEXT, w_info TEXT, created_at TEXT, updated_at TEXT);
CREATE INDEX IF NOT EXISTS inv_serial ON inventory(serial_norm);
CREATE TABLE IF NOT EXISTS firms (name TEXT PRIMARY KEY, emails TEXT, phone TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS tickets (id INTEGER PRIMARY KEY, no TEXT UNIQUE NOT NULL, subject TEXT NOT NULL, description TEXT,
  priority TEXT, stage TEXT NOT NULL, firma TEXT, contact_name TEXT, contact_email TEXT, contact_phone TEXT, location TEXT,
  serial TEXT, inventory_id INTEGER, product TEXT, warranty_code TEXT, warranty_text TEXT, out_of_warranty INTEGER DEFAULT 0,
  override_reason TEXT, assignee TEXT, assignee_name TEXT, sla_due TEXT, part_no TEXT, order_no TEXT, rma_no TEXT, cargo_no TEXT,
  solution TEXT, source TEXT, created_at TEXT, created_by TEXT, updated_at TEXT, closed_at TEXT);
CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY, ticket_id INTEGER NOT NULL, at TEXT, user TEXT, kind TEXT,
  text TEXT, public INTEGER DEFAULT 1);
CREATE INDEX IF NOT EXISTS ev_ticket ON events(ticket_id);
CREATE TABLE IF NOT EXISTS emails (id INTEGER PRIMARY KEY, ticket_id INTEGER, kind TEXT, to_addr TEXT, subject TEXT, body TEXT,
  status TEXT, error TEXT, created_at TEXT, sent_at TEXT);
"""

DEFAULT_SETTINGS = {
    "company": "Ekspa", "base_url": "", "sla_days": "5",
    "smtp_host": "", "smtp_port": "587", "smtp_security": "starttls", "smtp_user": "", "smtp_password": "",
    "from_addr": "", "from_name": "Ekspa Teknik Servis", "reply_to": "", "signature": "Ekspa Teknik Servis",
    "notify_expert": "1", "notify_customer_open": "1", "notify_customer_update": "1",
}


def add_user(con, un, name, email, admin, crm, srv, pw_hash, must_change=0, created=None):
    con.execute("INSERT INTO users (username, name, email, role, pw, active, must_change, created_at, crm_role, servis_role) "
                "VALUES (?,?,?,?,?,1,?,?,?,?)", (un, name, email or "", "yonetici" if admin else "kullanici", pw_hash,
                                                 1 if must_change else 0, created or now(), crm or "", srv or ""))


def migrate_users_table(con):
    """Eski (Ekspa Fırsat Takibi) kullanıcı tablosunu modül bazlı rollere çevirir."""
    cols = {r[1] for r in con.execute("PRAGMA table_info(users)")}
    if not cols or "crm_role" in cols:
        return
    con.execute("ALTER TABLE users RENAME TO users_eski")
    con.executescript(SCHEMA)
    for r in con.execute("SELECT * FROM users_eski").fetchall():
        old = r["role"]
        admin = old == "yonetici"
        crm = "duzenleyici" if old in ("yonetici", "duzenleyici") else "goruntuleyici"
        con.execute("INSERT INTO users (username, name, email, role, pw, active, must_change, created_at, last_login, crm_role, servis_role) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?)", (r["username"], r["name"], r["email"] if "email" in cols else "",
                                                       "yonetici" if admin else "kullanici", r["pw"], r["active"], r["must_change"],
                                                       r["created_at"], r["last_login"], crm, ""))
    con.execute("DROP TABLE users_eski")
    con.execute("DELETE FROM sessions")
    con.commit()
    print("Fırsat Takibi kullanıcıları yeni yetki yapısına taşındı.")


def import_legacy_service(con):
    """Ayrı kurulmuş Çağrı Karşılama veritabanını (data/cagri.db) bir kez içeri alır."""
    if not os.path.exists(LEGACY_SRV) or con.execute("SELECT v FROM meta WHERE k='servis_tasindi'").fetchone():
        return False
    if con.execute("SELECT COUNT(*) FROM tickets").fetchone()[0] or con.execute("SELECT COUNT(*) FROM inventory").fetchone()[0]:
        print("Uyarı: data/cagri.db bulundu ama servis tabloları dolu olduğu için taşınmadı.")
        return False
    con.execute("ATTACH DATABASE ? AS eski", (LEGACY_SRV,))
    for tbl in ("inventory", "firms", "tickets", "events", "emails"):
        con.execute(f"INSERT INTO {tbl} SELECT * FROM eski.{tbl}")
    con.execute("INSERT OR REPLACE INTO settings SELECT * FROM eski.settings")
    for r in con.execute("SELECT * FROM eski.users").fetchall():
        srv = {"yonetici": "uzman", "uzman": "uzman", "goruntuleyici": "goruntuleyici"}.get(r["role"], "")
        ex = con.execute("SELECT * FROM users WHERE username=?", (r["username"],)).fetchone()
        if ex:
            con.execute("UPDATE users SET servis_role=?, email=CASE WHEN COALESCE(email,'')='' THEN ? ELSE email END, "
                        "role=CASE WHEN ?='yonetici' THEN 'yonetici' ELSE role END WHERE username=?",
                        (srv, r["email"] or "", r["role"], r["username"]))
        else:
            con.execute("INSERT INTO users (username, name, email, role, pw, active, must_change, created_at, last_login, crm_role, servis_role) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?)", (r["username"], r["name"], r["email"] or "", "yonetici" if r["role"] == "yonetici" else "kullanici",
                                                           r["pw"], r["active"], r["must_change"], r["created_at"], r["last_login"], "", srv))
    con.commit()
    con.execute("DETACH DATABASE eski")
    con.execute("INSERT OR REPLACE INTO meta VALUES ('servis_tasindi', ?)", (now(),))
    con.commit()
    print("Çağrı Karşılama verileri ve kullanıcıları taşındı.")
    return True


def init_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(DB_PATH) and os.path.exists(LEGACY_CRM):
        src, dst = sqlite3.connect(LEGACY_CRM), sqlite3.connect(DB_PATH)
        src.backup(dst)
        src.close(); dst.close()
        print("Ekspa Fırsat Takibi veritabanı (firsatlar.db) kopyalandı; eski dosya yedek olarak duruyor.")
    con = connect()
    con.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")
    con.execute("CREATE TABLE IF NOT EXISTS sessions (token_hash TEXT PRIMARY KEY, username TEXT NOT NULL, expires REAL NOT NULL, created_at TEXT)")
    migrate_users_table(con)
    con.executescript(SCHEMA)
    con.execute("INSERT OR IGNORE INTO meta VALUES ('rev','0')")
    for k, v in DEFAULT_SETTINGS.items():
        con.execute("INSERT OR IGNORE INTO settings VALUES (?,?)", (k, v))
    con.execute("DELETE FROM sessions WHERE expires < ?", (time.time(),))
    con.commit()
    import_legacy_service(con)
    if con.execute("SELECT COUNT(*) FROM deals").fetchone()[0] == 0 and os.path.exists(SEED_CRM):
        with open(SEED_CRM, encoding="utf-8") as f:
            rows = json.load(f)
        for d in rows:
            did = d.pop("id")
            con.execute("INSERT INTO deals VALUES (?,?,?)", (did, json.dumps(d, ensure_ascii=False), now()))
        bump(con)
        print(f"Başlangıç fırsat verisi yüklendi: {len(rows)} fırsat")
    if con.execute("SELECT COUNT(*) FROM inventory").fetchone()[0] == 0 and os.path.exists(SEED_INV):
        with open(SEED_INV, encoding="utf-8") as f:
            n = insert_inventory(con, json.load(f))
        print(f"Başlangıç envanteri yüklendi: {n} ürün")
    if con.execute("SELECT COUNT(*) FROM tickets").fetchone()[0] == 0 and os.path.exists(SEED_TIC):
        with open(SEED_TIC, encoding="utf-8") as f:
            n = insert_tickets(con, json.load(f))
        print(f"Başlangıç talepleri yüklendi: {n} çağrı")
    link_assignees(con)
    con.commit()
    return con


def insert_inventory(con, rows):
    t = now()
    for r in rows:
        con.execute("INSERT INTO inventory (firma, serial, serial_norm, brand, part_no, description, service, w_start, w_end, w_info, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (r.get("firma") or "", r.get("serial") or "", norm_serial(r.get("serial")), r.get("brand") or "", r.get("part_no") or "",
                     r.get("description") or "", r.get("service") or "", r.get("w_start") or "", r.get("w_end") or "", r.get("w_info") or "", t, t))
    bump(con)
    return len(rows)


def insert_tickets(con, rows):
    n = 0
    for r in rows:
        if con.execute("SELECT 1 FROM tickets WHERE no=?", (r["no"],)).fetchone():
            continue
        cur = con.execute(
            "INSERT INTO tickets (no, subject, description, priority, stage, firma, contact_name, assignee, assignee_name, source, created_at, created_by, updated_at, closed_at, warranty_code, warranty_text) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (r["no"], r["subject"], r.get("description", ""), r.get("priority", "normal"), r.get("stage", "kapandi"), r.get("firma", ""),
             r.get("contact_name", ""), r.get("assignee", ""), r.get("assignee_name", ""), "Odoo aktarımı", now(), "Odoo aktarımı", now(),
             None, "", "Odoo'dan aktarıldı; garanti kontrolü yapılmadı"))
        con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,0)",
                    (cur.lastrowid, now(), "Sistem", "olusturuldu", f"Odoo yardım masasından aktarıldı (Aşama: {r.get('odoo_stage', '')})"))
        n += 1
    bump(con)
    return n


from contextlib import contextmanager


@contextmanager
def tx(con):
    """Kilit + hata olursa yarım kalan değişiklikleri geri al."""
    with LOCK:
        try:
            yield
        except BaseException:
            con.rollback()
            raise


def bump(con):
    con.execute("UPDATE meta SET v = CAST(v AS INTEGER) + 1 WHERE k='rev'")


def rev(con):
    return int(con.execute("SELECT v FROM meta WHERE k='rev'").fetchone()[0])


def settings(con):
    return {r["k"]: r["v"] for r in con.execute("SELECT k, v FROM settings")}


# ---------------------------------------------------------------- şifreler
def hash_pw(pw):
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt, PBKDF2_ITER)
    return f"pbkdf2_sha256${PBKDF2_ITER}${salt.hex()}${dk.hex()}"


def check_pw(pw, stored):
    try:
        _, it, salt, dk = stored.split("$")
        return hmac.compare_digest(hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), int(it)).hex(), dk)
    except Exception:
        return False


def pw_problem(pw):
    if len(pw) < 8:
        return "Şifre en az 8 karakter olmalı."
    if pw.isdigit() or pw.isalpha():
        return "Şifre hem harf hem rakam içermeli."
    return None


def norm_username(u):
    return str(u or "").strip().translate(str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosucgiosu")).lower()


def tok_hash(t):
    return hashlib.sha256(t.encode()).hexdigest()


def emails_in(s):
    out = [e.strip() for e in re.split(r"[;,\s]+", str(s or "")) if e.strip()]
    bad = [e for e in out if not EMAIL_RE.match(e)]
    if bad:
        raise ValueError(f"Geçersiz e-posta adresi: {bad[0]}")
    return out


# ---------------------------------------------------------------- e-posta
def queue_mail(con, ticket_id, kind, to_list, subject, body):
    to_list = [t for t in to_list if t]
    if not to_list:
        return
    st = settings(con)
    status = "bekliyor" if st.get("smtp_host") and st.get("from_addr") else "ayar_yok"
    cur = con.execute("INSERT INTO emails (ticket_id, kind, to_addr, subject, body, status, created_at) VALUES (?,?,?,?,?,?,?)",
                      (ticket_id, kind, ", ".join(to_list), subject, body, status, now()))
    if status == "bekliyor":
        MAILQ.put(cur.lastrowid)


def smtp_send(st, to_addr, subject, body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr((st.get("from_name") or "", st["from_addr"]))
    msg["To"] = to_addr
    if st.get("reply_to"):
        msg["Reply-To"] = st["reply_to"]
    msg["Message-ID"] = make_msgid(domain=st["from_addr"].split("@")[-1])
    msg.set_content(body)
    port = int(st.get("smtp_port") or 587)
    sec = st.get("smtp_security") or "starttls"
    ctx = ssl.create_default_context()
    if sec == "ssl":
        srv = smtplib.SMTP_SSL(st["smtp_host"], port, context=ctx, timeout=30)
    else:
        srv = smtplib.SMTP(st["smtp_host"], port, timeout=30)
        if sec == "starttls":
            srv.starttls(context=ctx)
    try:
        if st.get("smtp_user"):
            srv.login(st["smtp_user"], st.get("smtp_password") or "")
        srv.send_message(msg)
    finally:
        try:
            srv.quit()
        except Exception:
            pass


def mail_worker(con):
    while True:
        mid = MAILQ.get()
        with LOCK:
            row = con.execute("SELECT * FROM emails WHERE id=?", (mid,)).fetchone()
            st = settings(con)
        if not row or row["status"] not in ("bekliyor",):
            continue
        try:
            smtp_send(st, row["to_addr"], row["subject"], row["body"])
            status, err = "gonderildi", None
        except Exception as e:
            status, err = "hata", f"{type(e).__name__}: {e}"[:500]
        with LOCK:
            con.execute("UPDATE emails SET status=?, error=?, sent_at=? WHERE id=?", (status, err, now() if status == "gonderildi" else None, mid))
            bump(con)
            con.commit()


def ticket_lines(t, st):
    p = json.loads(t["product"] or "{}")
    lines = [f"Çağrı no      : {t['no']}", f"Konu          : {t['subject']}", f"Aşama         : {STAGE_T.get(t['stage'], t['stage'])}",
             f"Öncelik       : {PRIORITIES.get(t['priority'], t['priority'] or '')}", f"Firma         : {t['firma'] or '-'}"]
    if t["serial"]:
        lines.append(f"Ürün          : {' '.join(x for x in [p.get('brand'), p.get('part_no')] if x)} {('- ' + p['description'][:80]) if p.get('description') else ''}".rstrip())
        lines.append(f"Seri no       : {t['serial']}")
    if t["warranty_text"]:
        lines.append(f"Garanti       : {t['warranty_text']}")
    if t["assignee_name"]:
        lines.append(f"Sorumlu uzman : {t['assignee_name']}")
    if t["sla_due"]:
        lines.append(f"Hedef çözüm   : {fmt_d(t['sla_due'])}")
    return lines


def mail_open(con, t, user):
    st = settings(con)
    tag = f"[Çağrı {t['no']}]"
    base = public_base(st)
    if st.get("notify_expert") == "1" and t["assignee"]:
        u = con.execute("SELECT email, name FROM users WHERE username=?", (t["assignee"],)).fetchone()
        if u and u["email"]:
            body = "\n".join([f"Merhaba {u['name']},", "", "Size yeni bir teknik servis çağrısı atandı.", ""] + ticket_lines(t, st) +
                             [f"İletişim       : {t['contact_name'] or '-'} {t['contact_email'] or ''} {t['contact_phone'] or ''}".rstrip(),
                              f"Konum          : {t['location'] or '-'}", "", "Arıza tanımı:", t["description"] or "-", "",
                              f"Çağrıyı açan: {user['name']}"] + ([f"Çağrıya git: {base}/#cagri-{t['id']}"] if base else []) +
                             ["", st.get("signature") or ""])
            queue_mail(con, t["id"], "uzman", [u["email"]], f"{tag} Size yeni çağrı atandı: {t['subject']}", body)
    if st.get("notify_customer_open") == "1" and t["contact_email"]:
        body = "\n".join([f"Sayın {t['contact_name'] or 'yetkili'},", "",
                          f"Teknik servis talebiniz {st.get('company') or ''} tarafından kayda alınmıştır.", ""] + ticket_lines(t, st) +
                         ["", "Talebinizle ilgili gelişmeleri bu e-posta adresine bildireceğiz.",
                          "Bize yazarken konu satırındaki çağrı numarasını korumanızı rica ederiz.", "", st.get("signature") or ""])
        queue_mail(con, t["id"], "musteri", emails_in(t["contact_email"]), f"{tag} Teknik servis kaydınız açıldı: {t['subject']}", body)


def mail_update(con, t, changes, comment):
    st = settings(con)
    if st.get("notify_customer_update") != "1" or not t["contact_email"] or (not changes and not comment):
        return
    tag = f"[Çağrı {t['no']}]"
    stage_changed = any(c[0] == "Aşama" for c in changes)
    head = f"Çağrınızın durumu: {STAGE_T.get(t['stage'])}" if stage_changed else "Çağrınızda güncelleme var"
    lines = [f"Sayın {t['contact_name'] or 'yetkili'},", "", f"{t['no']} numaralı teknik servis çağrınızda aşağıdaki değişiklikler yapıldı:", ""]
    for label, old, new in changes:
        lines.append(f"• {label}: {new or '(boş)'}" + (f"  (önceki: {old})" if old else ""))
    if comment:
        lines += ["", "Açıklama:", comment]
    lines += ["", "Güncel çağrı bilgileri:"] + ticket_lines(t, st) + ["", st.get("signature") or ""]
    subj = f"{tag} {head}" if stage_changed else f"{tag} {head}: {t['subject']}"
    queue_mail(con, t["id"], "musteri", emails_in(t["contact_email"]), subj, "\n".join(lines))


# ---------------------------------------------------------------- HTTP
class HTTPError(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


def s(v, n=2000):
    return str(v or "").strip()[:n]


class Handler(BaseHTTPRequestHandler):
    server_version = "EkspaIsTakibi/4.0"

    def log_message(self, fmt, *args):
        sys.stderr.write("%s %s\n" % (self.client_ip(), fmt % args))

    @property
    def con(self):
        return self.server.con

    def send_json(self, obj, code=200, cookie=None):
        raw = json.dumps(obj, ensure_ascii=False, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        if cookie is not None:
            self.send_header("Set-Cookie", cookie)
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def redirect(self, to):
        self.send_response(303)
        self.send_header("Location", to)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 2_000_000:
            raise HTTPError(413, "İstek çok büyük")
        try:
            b = json.loads(self.rfile.read(n) or b"{}")
        except json.JSONDecodeError:
            raise HTTPError(400, "Geçersiz istek")
        if not isinstance(b, dict):
            raise HTTPError(400, "Geçersiz istek")
        return b

    def client_ip(self):
        ip = self.client_address[0]
        # Railway'in kenar sunucusu gerçek istemci adresini X-Real-IP'ye yazar (istemcinin gönderdiğini ezer).
        # Bu olmadan tüm kullanıcılar aynı proxy adresinden geliyor görünür ve hatalı giriş kilidi herkesi etkiler.
        if ON_RAILWAY and self.headers.get("X-Real-IP"):
            return self.headers["X-Real-IP"].strip()[:64]
        # Cloudflare Tunnel arkasında gerçek istemci adresi bu başlıkta gelir
        if ip in ("127.0.0.1", "::1") and self.headers.get("CF-Connecting-IP"):
            return self.headers["CF-Connecting-IP"].strip()[:64]
        return ip

    def make_cookie(self, token, max_age):
        parts = [f"{COOKIE}={token}", "Path=/", "HttpOnly", "SameSite=Strict"]
        if max_age is not None:
            parts.append(f"Max-Age={max_age}")
        if HTTPS:
            parts.append("Secure")
        return "; ".join(parts)

    def current_user(self):
        c = SimpleCookie(self.headers.get("Cookie", ""))
        if COOKIE not in c:
            return None
        th = tok_hash(c[COOKIE].value)
        with tx(self.con):
            row = self.con.execute("SELECT u.username, u.name, u.email, u.role, u.crm_role, u.servis_role, u.active, u.must_change, s.expires FROM sessions s "
                                   "JOIN users u ON u.username=s.username WHERE s.token_hash=?", (th,)).fetchone()
            if not row or row["expires"] < time.time() or not row["active"]:
                if row:
                    self.con.execute("DELETE FROM sessions WHERE token_hash=?", (th,)); self.con.commit()
                return None
        admin = row["role"] == "yonetici"
        return {"username": row["username"], "name": row["name"], "email": row["email"] or "", "admin": admin,
                "role": "yonetici" if admin else "kullanici", "crm": row["crm_role"] or "", "servis": row["servis_role"] or "",
                "crmLevel": 3 if admin else CRM_LV.get(row["crm_role"] or "", 0),
                "servisLevel": 3 if admin else SRV_LV.get(row["servis_role"] or "", 0),
                "mustChange": bool(row["must_change"]), "_th": th}

    def require(self, module=None, level=1, admin=False, allow_must_change=False):
        """module: 'crm' (Satış) veya 'servis'; level: 1 görüntüleme, 2 düzenleme."""
        u = self.current_user()
        if not u:
            raise HTTPError(401, "Oturum açmanız gerekiyor")
        if u["mustChange"] and not allow_must_change:
            raise HTTPError(403, "Devam etmeden önce şifrenizi değiştirmeniz gerekiyor")
        if admin and not u["admin"]:
            raise HTTPError(403, "Bu işlem yalnızca yönetici içindir")
        if module and u[module + "Level"] < level:
            raise HTTPError(403, "Bu bölüme erişim yetkiniz yok" if level == 1 else "Bu işlem için yetkiniz yok")
        return u

    def dispatch(self, m):
        url = urlparse(self.path)
        path = url.path.rstrip("/") or "/"
        self.q = {k: v[0] for k, v in parse_qs(url.query).items()}
        try:
            if m != "GET" and self.headers.get("X-Requested-With") != "fetch":
                raise HTTPError(403, "Geçersiz istek kaynağı")
            if path.startswith("/api/"):
                return self.api(m, path.split("/")[2:])
            if m != "GET":
                raise HTTPError(404, "Bulunamadı")
            return self.static(path)
        except HTTPError as e:
            if not path.startswith("/api/") and e.code == 401:
                return self.redirect("/giris")
            return self.send_json({"error": str(e)}, e.code)
        except ValueError as e:
            return self.send_json({"error": str(e)}, 400)

    def do_GET(self): self.dispatch("GET")
    def do_POST(self): self.dispatch("POST")
    def do_PUT(self): self.dispatch("PUT")
    def do_DELETE(self): self.dispatch("DELETE")

    # ------------------------------------------------------------ yönlendirme
    def api(self, m, p):
        r = (m, p[0] if p else "")
        if r == ("GET", "kurulum"):
            with tx(self.con):
                return self.send_json({"kurulumGerekli": self.con.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0,
                                       "kodGerekli": bool(SETUP_CODE)})
        if r == ("POST", "kurulum"):
            return self.first_admin()
        if r == ("POST", "login"):
            return self.login()
        if r == ("POST", "logout"):
            u = self.current_user()
            if u:
                with tx(self.con):
                    self.con.execute("DELETE FROM sessions WHERE token_hash=?", (u["_th"],)); self.con.commit()
            return self.send_json({"ok": True}, cookie=self.make_cookie("", 0))
        if p == ["me"] and m == "GET":
            u = self.require(allow_must_change=True)
            return self.send_json({**{k: v for k, v in u.items() if not k.startswith("_")},
                                   "roleName": "Yönetici" if u["admin"] else " · ".join(x for x in [
                                       ("Satış: " + CRM_T[u["crm"]]) if u["crm"] else "", ("Servis: " + SRV_T[u["servis"]]) if u["servis"] else ""] if x) or "Erişim yok"})
        if p == ["me", "sifre"] and m == "POST":
            return self.change_own_pw()
        if p and p[0] == "kullanicilar":
            return self.users_api(m, p)
        if r == ("GET", "uzmanlar"):
            self.require("servis")
            with tx(self.con):
                rows = self.con.execute("SELECT username, name, email FROM users WHERE active=1 AND (servis_role='uzman' OR role='yonetici') ORDER BY name").fetchall()
            return self.send_json({"uzmanlar": [dict(x) for x in rows]})
        if r == ("GET", "rev"):
            self.require()
            with tx(self.con):
                return self.send_json({"rev": rev(self.con)})
        if r == ("GET", "meta"):
            self.require("servis")
            return self.send_json({"stages": STAGES, "closed": sorted(CLOSED), "priorities": PRIORITIES, "today": today()})
        if p and p[0] == "deals":
            return self.deals_api(m, p)
        if p and p[0] == "envanter":
            return self.inventory_api(m, p)
        if p and p[0] == "firmalar":
            return self.firms_api(m, p)
        if p and p[0] == "cagrilar":
            return self.tickets_api(m, p)
        if p and p[0] == "eposta":
            return self.mail_api(m, p)
        if p and p[0] == "ayarlar":
            return self.settings_api(m, p)
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ kimlik
    def start_session(self, username, remember):
        token = secrets.token_urlsafe(32)
        life = SESSION_LONG if remember else SESSION_SHORT
        self.con.execute("INSERT INTO sessions VALUES (?,?,?,?)", (tok_hash(token), username, time.time() + life, now()))
        self.con.execute("UPDATE users SET last_login=? WHERE username=?", (now(), username))
        self.con.commit()
        return self.make_cookie(token, life if remember else None)

    def login(self):
        b = self.body()
        un, pw = norm_username(b.get("username")), str(b.get("password") or "")
        key, t = (self.client_ip(), un), time.time()
        with tx(self.con):
            FAILS[key] = [x for x in FAILS.get(key, []) if t - x < 300]
            if len(FAILS[key]) >= 5:
                raise HTTPError(429, "Çok fazla hatalı deneme. 5 dakika sonra tekrar deneyin.")
            row = self.con.execute("SELECT pw, active FROM users WHERE username=?", (un,)).fetchone()
            if not row or not check_pw(pw, row["pw"]):
                FAILS[key].append(t)
                raise HTTPError(401, "Kullanıcı adı veya şifre hatalı.")
            if not row["active"]:
                raise HTTPError(403, "Bu hesap devre dışı bırakılmış. Yöneticinize başvurun.")
            FAILS.pop(key, None)
            cookie = self.start_session(un, bool(b.get("remember")))
        return self.send_json({"ok": True}, cookie=cookie)

    def validate_user_fields(self, b, need_pw=True):
        un = norm_username(b.get("username"))
        name, email = s(b.get("name"), 80), s(b.get("email"), 120)
        if not USERNAME_RE.match(un):
            raise HTTPError(400, "Kullanıcı adı 3-32 karakter olmalı; harf, rakam, nokta, tire ve alt çizgi kullanılabilir.")
        if not name:
            raise HTTPError(400, "Ad soyad gerekli.")
        if email and not EMAIL_RE.match(email):
            raise HTTPError(400, "Geçersiz e-posta adresi.")
        if need_pw:
            pr = pw_problem(str(b.get("password") or ""))
            if pr:
                raise HTTPError(400, pr)
        return un, name, email

    def first_admin(self):
        b = self.body()
        with tx(self.con):
            if self.con.execute("SELECT COUNT(*) FROM users").fetchone()[0]:
                raise HTTPError(403, "Kurulum zaten tamamlanmış.")
            if SETUP_CODE:
                key, t = (self.client_ip(), "#kurulum"), time.time()
                FAILS[key] = [x for x in FAILS.get(key, []) if t - x < 300]
                if len(FAILS[key]) >= 5:
                    raise HTTPError(429, "Çok fazla hatalı deneme. 5 dakika sonra tekrar deneyin.")
                if not hmac.compare_digest(norm_code(b.get("kod")), SETUP_CODE):
                    FAILS[key].append(t)
                    raise HTTPError(403, "Kurulum kodu hatalı. Kod, sunucu açılırken loglara yazılır.")
            un, name, email = self.validate_user_fields(b)
            add_user(self.con, un, name, email, True, "duzenleyici", "uzman", hash_pw(b["password"]))
            link_assignees(self.con)
            cookie = self.start_session(un, False)
        return self.send_json({"ok": True}, cookie=cookie)

    def change_own_pw(self):
        u = self.require(allow_must_change=True)
        b = self.body()
        with tx(self.con):
            if not check_pw(str(b.get("old") or ""), self.con.execute("SELECT pw FROM users WHERE username=?", (u["username"],)).fetchone()["pw"]):
                raise HTTPError(400, "Mevcut şifre hatalı.")
            pr = pw_problem(str(b.get("new") or ""))
            if pr:
                raise HTTPError(400, pr)
            self.con.execute("UPDATE users SET pw=?, must_change=0 WHERE username=?", (hash_pw(b["new"]), u["username"]))
            self.con.execute("DELETE FROM sessions WHERE username=? AND token_hash<>?", (u["username"], u["_th"]))
            self.con.commit()
        return self.send_json({"ok": True})

    def users_api(self, m, p):
        me = self.require(admin=True)
        con = self.con

        def roles(b, cur=None):
            admin = bool(b["admin"]) if "admin" in b else (cur["role"] == "yonetici" if cur else False)
            crm = b.get("crm_role", cur["crm_role"] if cur else "") or ""
            srv = b.get("servis_role", cur["servis_role"] if cur else "") or ""
            if crm not in CRM_LV or srv not in SRV_LV:
                raise HTTPError(400, "Geçersiz yetki.")
            return admin, crm, srv

        if len(p) == 1 and m == "GET":
            with tx(con):
                rows = con.execute("SELECT * FROM users ORDER BY name").fetchall()
            return self.send_json({"users": [dict(username=r["username"], name=r["name"], email=r["email"] or "", admin=r["role"] == "yonetici",
                                                  crm_role=r["crm_role"] or "", servis_role=r["servis_role"] or "", active=bool(r["active"]),
                                                  mustChange=bool(r["must_change"]), lastLogin=r["last_login"]) for r in rows]})
        if len(p) == 1 and m == "POST":
            b = self.body()
            un, name, email = self.validate_user_fields(b)
            admin, crm, srv = roles(b)
            if not (admin or crm or srv):
                raise HTTPError(400, "Kullanıcıya en az bir bölümde yetki verin.")
            with tx(con):
                if con.execute("SELECT 1 FROM users WHERE username=?", (un,)).fetchone():
                    raise HTTPError(409, "Bu kullanıcı adı zaten var.")
                add_user(con, un, name, email, admin, crm, srv, hash_pw(b["password"]), must_change=1)
                link_assignees(con); bump(con); con.commit()
            return self.send_json({"ok": True}, 201)
        if len(p) == 2 and m == "PUT":
            target, b = p[1], self.body()
            with tx(con):
                cur = con.execute("SELECT * FROM users WHERE username=?", (target,)).fetchone()
                if not cur:
                    raise HTTPError(404, "Kullanıcı bulunamadı.")
                admin, crm, srv = roles(b, cur)
                active = (1 if b["active"] else 0) if "active" in b else cur["active"]
                if target == me["username"] and (not admin or not active):
                    raise HTTPError(400, "Kendi yönetici yetkinizi kaldıramaz veya hesabınızı kapatamazsınız.")
                if (not admin or not active) and not con.execute(
                        "SELECT COUNT(*) FROM users WHERE role='yonetici' AND active=1 AND username<>?", (target,)).fetchone()[0]:
                    raise HTTPError(400, "En az bir etkin yönetici kalmalı.")
                con.execute("UPDATE users SET role=?, crm_role=?, servis_role=?, active=? WHERE username=?",
                            ("yonetici" if admin else "kullanici", crm, srv, active, target))
                if "name" in b and s(b["name"], 80):
                    con.execute("UPDATE users SET name=? WHERE username=?", (s(b["name"], 80), target))
                if "email" in b:
                    em = s(b["email"], 120)
                    if em and not EMAIL_RE.match(em):
                        raise HTTPError(400, "Geçersiz e-posta adresi.")
                    con.execute("UPDATE users SET email=? WHERE username=?", (em, target))
                if b.get("password"):
                    pr = pw_problem(str(b["password"]))
                    if pr:
                        raise HTTPError(400, pr)
                    con.execute("UPDATE users SET pw=?, must_change=1 WHERE username=?", (hash_pw(str(b["password"])), target))
                    con.execute("DELETE FROM sessions WHERE username=?", (target,))
                if not active:
                    con.execute("DELETE FROM sessions WHERE username=?", (target,))
                bump(con); con.commit()
            return self.send_json({"ok": True})
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ satış: fırsatlar ve projeler
    def deals_api(self, m, p):
        con = self.con
        if p == ["deals"] and m == "GET":
            self.require("crm")
            with tx(con):
                rows = con.execute("SELECT id, data FROM deals").fetchall()
                r = rev(con)
            return self.send_json({"rev": r, "deals": [dict(json.loads(d), id=i) for i, d in rows]})
        if p == ["deals"] and m == "POST":
            u = self.require("crm", 2)
            data = clean_deal(self.body(), partial=False)
            data.setdefault("stage", "ilgi")
            did = "d" + secrets.token_hex(5)
            with tx(con):
                order = con.execute("SELECT COALESCE(MAX(json_extract(data,'$.order')),0) FROM deals").fetchone()[0] + 1
                data.update(order=order, notes=[], createdAt=now(), updatedAt=now(), updatedBy=u["name"], createdBy=u["name"])
                con.execute("INSERT INTO deals VALUES (?,?,?)", (did, json.dumps(data, ensure_ascii=False), now()))
                bump(con); con.commit()
            return self.send_json(dict(data, id=did), 201)
        if len(p) == 2 and m == "PUT":
            u = self.require("crm", 2)
            patch = clean_deal(self.body(), partial=True)
            with tx(con):
                d = self.get_deal(p[1])
                d.update(patch, updatedAt=now(), updatedBy=u["name"])
                self.put_deal(p[1], d)
            return self.send_json(dict(d, id=p[1]))
        if len(p) == 2 and m == "DELETE":
            self.require("crm", 2)
            with tx(con):
                con.execute("DELETE FROM deals WHERE id=?", (p[1],))
                bump(con); con.commit()
            return self.send_json({"ok": True})
        if len(p) == 3 and p[2] == "notes" and m == "POST":
            u = self.require("crm", 2)
            text = s(self.body().get("text"), 5000)
            if not text:
                raise HTTPError(400, "Not boş olamaz")
            with tx(con):
                d = self.get_deal(p[1])
                d.setdefault("notes", []).append({"text": text, "at": now(), "by": u["name"], "user": u["username"]})
                d["updatedAt"] = now()
                self.put_deal(p[1], d)
            return self.send_json(dict(d, id=p[1]))
        raise HTTPError(404, "Bulunamadı")

    def get_deal(self, did):
        row = self.con.execute("SELECT data FROM deals WHERE id=?", (did,)).fetchone()
        if not row:
            raise HTTPError(404, "Fırsat bulunamadı")
        return json.loads(row[0])

    def put_deal(self, did, d):
        self.con.execute("UPDATE deals SET data=?, updated_at=? WHERE id=?", (json.dumps(d, ensure_ascii=False), now(), did))
        bump(self.con); self.con.commit()

    # ------------------------------------------------------------ envanter
    def inv_dict(self, r):
        code, text = warranty(r["w_start"], r["w_end"])
        return {"id": r["id"], "firma": r["firma"], "serial": r["serial"], "brand": r["brand"], "part_no": r["part_no"],
                "description": r["description"], "service": r["service"], "w_start": r["w_start"], "w_end": r["w_end"],
                "w_info": r["w_info"], "warranty": code, "warrantyText": text}

    def inventory_api(self, m, p):
        con = self.con
        if len(p) == 2 and p[1] == "kontrol" and m == "GET":
            self.require("servis")
            sn = norm_serial(self.q.get("seri"))
            if len(sn) < 3:
                raise HTTPError(400, "Seri numarası en az 3 karakter olmalı.")
            with tx(self.con):
                exact = con.execute("SELECT * FROM inventory WHERE serial_norm=?", (sn,)).fetchall()
                similar = [] if exact else con.execute("SELECT * FROM inventory WHERE serial_norm LIKE ? LIMIT 8", (f"%{sn}%",)).fetchall()
                open_t = {}
                for r in exact:
                    open_t[r["id"]] = [dict(no=x["no"], id=x["id"], stage=STAGE_T.get(x["stage"])) for x in con.execute(
                        "SELECT id, no, stage FROM tickets WHERE inventory_id=? AND stage NOT IN ('cozuldu','iade','kapandi','iptal')", (r["id"],))]
            return self.send_json({"serial": sn, "matches": [{**self.inv_dict(r), "openTickets": open_t.get(r["id"], [])} for r in exact],
                                   "similar": [self.inv_dict(r) for r in similar]})
        if len(p) == 1 and m == "GET":
            self.require("servis")
            q, firma, durum = s(self.q.get("q"), 100), self.q.get("firma", ""), self.q.get("durum", "")
            limit, offset = min(int(self.q.get("limit", 100)), 500), max(int(self.q.get("offset", 0)), 0)
            where, args = [], []
            if q:
                where.append("(serial_norm LIKE ? OR part_no LIKE ? OR description LIKE ? OR brand LIKE ?)")
                args += [f"%{norm_serial(q)}%", f"%{q}%", f"%{q}%", f"%{q}%"]
            if firma:
                where.append("firma=?"); args.append("" if firma == "__yok__" else firma)
            t = today()
            cond = {"kapsamda": "(w_end<>'' AND w_end>=? AND (w_start='' OR w_start<=?))", "bitti": "(w_end<>'' AND w_end<?)",
                    "bilgi_yok": "(w_end='')", "yakinda": "(w_end>=? AND w_end<=?)"}
            if durum in cond:
                where.append(cond[durum])
                args += {"kapsamda": [t, t], "bitti": [t], "bilgi_yok": [], "yakinda": [t, (date.today() + timedelta(days=90)).isoformat()]}[durum]
            w = (" WHERE " + " AND ".join(where)) if where else ""
            with tx(self.con):
                total = con.execute("SELECT COUNT(*) FROM inventory" + w, args).fetchone()[0]
                rows = con.execute(f"SELECT * FROM inventory{w} ORDER BY firma, serial LIMIT ? OFFSET ?", args + [limit, offset]).fetchall()
                firms = [dict(name=r[0], n=r[1]) for r in con.execute("SELECT firma, COUNT(*) FROM inventory GROUP BY firma ORDER BY firma")]
                counts = {k: con.execute("SELECT COUNT(*) FROM inventory WHERE " + c, a).fetchone()[0] for k, c, a in [
                    ("kapsamda", cond["kapsamda"], [t, t]), ("bitti", cond["bitti"], [t]), ("bilgi_yok", cond["bilgi_yok"], []),
                    ("yakinda", cond["yakinda"], [t, (date.today() + timedelta(days=90)).isoformat()])]}
                counts["toplam"] = con.execute("SELECT COUNT(*) FROM inventory").fetchone()[0]
            return self.send_json({"total": total, "items": [self.inv_dict(r) for r in rows], "firms": firms, "counts": counts})
        if m in ("POST", "PUT"):
            u = self.require("servis", 2)
            b = self.body()
            rec = {k: s(b.get(k), 500) for k in ("firma", "serial", "brand", "part_no", "description", "service", "w_start", "w_end", "w_info")}
            if not norm_serial(rec["serial"]):
                raise HTTPError(400, "Seri numarası gerekli.")
            for k in ("w_start", "w_end"):
                if rec[k] and not DATE_RE.match(rec[k]):
                    raise HTTPError(400, "Tarih biçimi YYYY-AA-GG olmalı.")
            if rec["w_start"] and rec["w_end"] and rec["w_end"] < rec["w_start"]:
                raise HTTPError(400, "Garanti bitişi başlangıçtan önce olamaz.")
            with tx(self.con):
                if m == "POST" and len(p) == 1:
                    cur = con.execute("INSERT INTO inventory (firma, serial, serial_norm, brand, part_no, description, service, w_start, w_end, w_info, created_at, updated_at) "
                                      "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", (rec["firma"], rec["serial"], norm_serial(rec["serial"]), rec["brand"], rec["part_no"],
                                                                           rec["description"], rec["service"], rec["w_start"], rec["w_end"], rec["w_info"], now(), now()))
                    iid = cur.lastrowid
                elif m == "PUT" and len(p) == 2:
                    iid = int(p[1])
                    if not con.execute("SELECT 1 FROM inventory WHERE id=?", (iid,)).fetchone():
                        raise HTTPError(404, "Ürün bulunamadı.")
                    con.execute("UPDATE inventory SET firma=?, serial=?, serial_norm=?, brand=?, part_no=?, description=?, service=?, w_start=?, w_end=?, w_info=?, updated_at=? WHERE id=?",
                                (rec["firma"], rec["serial"], norm_serial(rec["serial"]), rec["brand"], rec["part_no"], rec["description"], rec["service"],
                                 rec["w_start"], rec["w_end"], rec["w_info"], now(), iid))
                else:
                    raise HTTPError(404, "Bulunamadı")
                bump(con); con.commit()
                row = con.execute("SELECT * FROM inventory WHERE id=?", (iid,)).fetchone()
            return self.send_json(self.inv_dict(row))
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ firmalar
    def firms_api(self, m, p):
        con = self.con
        if len(p) == 1 and m == "GET":
            self.require("servis")
            with tx(self.con):
                names = {r[0] for r in con.execute("SELECT DISTINCT firma FROM inventory WHERE firma<>''")}
                names |= {r[0] for r in con.execute("SELECT DISTINCT firma FROM tickets WHERE firma<>''")}
                names |= {r[0] for r in con.execute("SELECT name FROM firms")}
                info = {r["name"]: dict(r) for r in con.execute("SELECT * FROM firms")}
                inv = dict(con.execute("SELECT firma, COUNT(*) FROM inventory GROUP BY firma").fetchall())
                tic = dict(con.execute("SELECT firma, COUNT(*) FROM tickets WHERE stage NOT IN ('cozuldu','iade','kapandi','iptal') GROUP BY firma").fetchall())
            return self.send_json({"firms": [dict(name=n, emails=info.get(n, {}).get("emails") or "", phone=info.get(n, {}).get("phone") or "",
                                                  note=info.get(n, {}).get("note") or "", inventory=inv.get(n, 0), openTickets=tic.get(n, 0))
                                             for n in sorted(names, key=str.lower)]})
        if len(p) == 1 and m == "PUT":
            self.require("servis", 2)
            b = self.body()
            name = s(b.get("name"), 200)
            if not name:
                raise HTTPError(400, "Firma adı gerekli.")
            emails = ", ".join(emails_in(b.get("emails")))
            with tx(self.con):
                con.execute("INSERT INTO firms VALUES (?,?,?,?) ON CONFLICT(name) DO UPDATE SET emails=excluded.emails, phone=excluded.phone, note=excluded.note",
                            (name, emails, s(b.get("phone"), 60), s(b.get("note"), 1000)))
                bump(con); con.commit()
            return self.send_json({"ok": True})
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ çağrılar
    def t_dict(self, r, full=False):
        d = dict(r)
        d["product"] = json.loads(d["product"] or "{}")
        d["stageName"] = STAGE_T.get(d["stage"], d["stage"])
        d["closed"] = d["stage"] in CLOSED
        if not full:
            d.pop("description", None)
        return d

    def tickets_api(self, m, p):
        con = self.con
        if len(p) == 1 and m == "GET":
            self.require("servis")
            with tx(self.con):
                rows = con.execute("SELECT * FROM tickets ORDER BY no DESC").fetchall()
                r = rev(con)
            return self.send_json({"rev": r, "tickets": [self.t_dict(x) for x in rows]})
        if len(p) == 1 and m == "POST":
            return self.create_ticket()
        if len(p) >= 2:
            tid = int(p[1]) if p[1].isdigit() else -1
            if len(p) == 2 and m == "GET":
                self.require("servis")
                with tx(self.con):
                    t = con.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
                    if not t:
                        raise HTTPError(404, "Çağrı bulunamadı.")
                    ev = [dict(x) for x in con.execute("SELECT * FROM events WHERE ticket_id=? ORDER BY id", (tid,))]
                    em = [dict(x) for x in con.execute("SELECT id, kind, to_addr, subject, body, status, error, created_at, sent_at FROM emails WHERE ticket_id=? ORDER BY id", (tid,))]
                    inv = con.execute("SELECT * FROM inventory WHERE id=?", (t["inventory_id"],)).fetchone() if t["inventory_id"] else None
                return self.send_json({"ticket": self.t_dict(t, True), "events": ev, "emails": em, "inventory": self.inv_dict(inv) if inv else None})
            if len(p) == 2 and m == "PUT":
                return self.update_ticket(tid)
            if len(p) == 3 and p[2] == "yorum" and m == "POST":
                return self.comment(tid)
        raise HTTPError(404, "Bulunamadı")

    def assignee_name(self, un):
        if not un:
            return ""
        r = self.con.execute("SELECT name FROM users WHERE username=? AND active=1 AND (servis_role='uzman' OR role='yonetici')", (un,)).fetchone()
        if not r:
            raise HTTPError(400, "Seçilen uzman bulunamadı veya etkin değil.")
        return r["name"]

    def create_ticket(self):
        u = self.require("servis", 2)
        b = self.body()
        con = self.con
        subject = s(b.get("subject"), 200)
        if not subject:
            raise HTTPError(400, "Konu gerekli.")
        prio = b.get("priority") if b.get("priority") in PRIORITIES else "normal"
        contact_email = ", ".join(emails_in(b.get("contact_email")))
        sla = s(b.get("sla_due"), 10)
        if sla and not DATE_RE.match(sla):
            raise HTTPError(400, "Hedef tarih biçimi YYYY-AA-GG olmalı.")
        with tx(self.con):
            inv = con.execute("SELECT * FROM inventory WHERE id=?", (int(b.get("inventory_id") or 0),)).fetchone()
            if not inv:
                raise HTTPError(400, "Çağrı açmak için envanterde kayıtlı bir ürün seçilmeli. Seri numarasını kontrol edin.")
            code, text = warranty(inv["w_start"], inv["w_end"])
            out = 0
            reason = ""
            if code in ("bitti", "baslamadi"):
                reason = s(b.get("override_reason"), 500)
                if not u["admin"]:
                    raise HTTPError(403, f"Ürün garanti kapsamında değil ({text}). Garanti dışı çağrıyı yalnızca yönetici açabilir.")
                if not reason:
                    raise HTTPError(400, f"Ürün garanti kapsamında değil ({text}). Garanti dışı açmak için gerekçe yazın.")
                out = 1
            assignee = s(b.get("assignee"), 40)
            if not assignee:
                raise HTTPError(400, "Sorumlu uzmanı seçin.")
            aname = self.assignee_name(assignee)
            sla = sla or add_business_days(date.today(), int(settings(con).get("sla_days") or 5)).isoformat()
            last = con.execute("SELECT MAX(CAST(no AS INTEGER)) FROM tickets").fetchone()[0] or 0
            no = f"{last + 1:05d}"
            product = json.dumps({"brand": inv["brand"], "part_no": inv["part_no"], "description": inv["description"], "service": inv["service"],
                                  "w_start": inv["w_start"], "w_end": inv["w_end"], "w_info": inv["w_info"]}, ensure_ascii=False)
            stage = "acik"
            cur = con.execute(
                "INSERT INTO tickets (no, subject, description, priority, stage, firma, contact_name, contact_email, contact_phone, location, serial, inventory_id, "
                "product, warranty_code, warranty_text, out_of_warranty, override_reason, assignee, assignee_name, sla_due, source, created_at, created_by, updated_at) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (no, subject, s(b.get("description"), 5000), prio, stage, inv["firma"], s(b.get("contact_name"), 120), contact_email,
                 s(b.get("contact_phone"), 60), s(b.get("location"), 300), inv["serial"], inv["id"], product, code, text, out, reason,
                 assignee, aname, sla, "Uygulama", now(), u["name"], now()))
            tid = cur.lastrowid
            msg = f"Çağrı açıldı. {text}." + (f" Garanti dışı açıldı; gerekçe: {reason}" if out else "") + \
                (" Garanti doğrulanamadı; envanterde garanti tarihleri girilmemiş." if code == "bilgi_yok" else "")
            con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,1)", (tid, now(), u["name"], "olusturuldu", msg))
            if aname:
                con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,1)", (tid, now(), u["name"], "atama", f"Sorumlu uzman: {aname}"))
            t = con.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
            mail_open(con, t, u)
            bump(con); con.commit()
        return self.send_json({"id": tid, "no": no}, 201)

    def update_ticket(self, tid):
        u = self.require("servis", 2)
        b = self.body()
        con = self.con
        with tx(self.con):
            t = con.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
            if not t:
                raise HTTPError(404, "Çağrı bulunamadı.")
            new = {}
            for k in EDITABLE:
                if k not in b:
                    continue
                v = s(b[k], 5000 if k in ("description", "solution") else 300)
                if k == "stage" and v not in STAGE_T:
                    raise HTTPError(400, "Geçersiz aşama.")
                if k == "priority" and v not in PRIORITIES:
                    raise HTTPError(400, "Geçersiz öncelik.")
                if k == "sla_due" and v and not DATE_RE.match(v):
                    raise HTTPError(400, "Tarih biçimi YYYY-AA-GG olmalı.")
                if k == "contact_email":
                    v = ", ".join(emails_in(v))
                if k == "subject" and not v:
                    raise HTTPError(400, "Konu boş olamaz.")
                if (t[k] or "") != v:
                    new[k] = v
            if "stage" in new and new["stage"] in NEEDS_SOLUTION and not (new.get("solution", t["solution"]) or "").strip():
                raise HTTPError(400, f"Çağrıyı “{STAGE_T[new['stage']]}” olarak işaretlemeden önce çözüm açıklamasını yazın.")
            changes = []
            if "assignee" in new:
                new["assignee_name"] = self.assignee_name(new["assignee"])
            for k, v in new.items():
                if k in ("assignee_name",):
                    continue
                old_disp, new_disp = t[k] or "", v
                if k == "stage":
                    old_disp, new_disp = STAGE_T.get(old_disp, old_disp), STAGE_T[v]
                elif k == "priority":
                    old_disp, new_disp = PRIORITIES.get(old_disp, old_disp), PRIORITIES[v]
                elif k == "assignee":
                    old_disp, new_disp = t["assignee_name"] or "", new["assignee_name"]
                elif k == "sla_due":
                    old_disp, new_disp = fmt_d(old_disp), fmt_d(v)
                if k in TRACKED:
                    changes.append((TRACKED[k], old_disp, new_disp))
            order = list(TRACKED.values())
            changes.sort(key=lambda c: order.index(c[0]))
            comment = s(b.get("comment"), 5000)
            if not new and not comment:
                return self.send_json({"ok": True, "unchanged": True})
            if new:
                if "stage" in new:
                    new["closed_at"] = now() if new["stage"] in CLOSED else None
                new["updated_at"] = now()
                con.execute(f"UPDATE tickets SET {', '.join(k + '=?' for k in new)} WHERE id=?", list(new.values()) + [tid])
                for label, o, n in changes:
                    kind = "asama" if label == "Aşama" else ("atama" if label == "Sorumlu uzman" else "guncelleme")
                    con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,1)",
                                (tid, now(), u["name"], kind, f"{label}: {n or '(boş)'}" + (f" (önceki: {o})" if o else "")))
            if comment:
                con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,1)", (tid, now(), u["name"], "yorum", comment))
            t2 = con.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
            notify = b.get("notify", True) is not False
            if notify:
                mail_update(con, t2, changes, comment)
            if "assignee" in new and new["assignee"]:
                st = settings(con)
                ue = con.execute("SELECT email, name FROM users WHERE username=?", (new["assignee"],)).fetchone()
                if st.get("notify_expert") == "1" and ue and ue["email"]:
                    base = public_base(st)
                    body = "\n".join([f"Merhaba {ue['name']},", "", "Aşağıdaki teknik servis çağrısı size atandı.", ""] + ticket_lines(t2, st) +
                                     ["", "Arıza tanımı:", t2["description"] or "-"] + ([f"", f"Çağrıya git: {base}/#cagri-{tid}"] if base else []) + ["", st.get("signature") or ""])
                    queue_mail(con, tid, "uzman", [ue["email"]], f"[Çağrı {t2['no']}] Size çağrı atandı: {t2['subject']}", body)
            bump(con); con.commit()
        return self.send_json({"ok": True, "changes": len(changes)})

    def comment(self, tid):
        u = self.require("servis", 2)
        b = self.body()
        text = s(b.get("text"), 5000)
        if not text:
            raise HTTPError(400, "Not boş olamaz.")
        public = bool(b.get("public"))
        with tx(self.con):
            t = self.con.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
            if not t:
                raise HTTPError(404, "Çağrı bulunamadı.")
            self.con.execute("INSERT INTO events (ticket_id, at, user, kind, text, public) VALUES (?,?,?,?,?,?)",
                             (tid, now(), u["name"], "yorum" if public else "not", text, 1 if public else 0))
            self.con.execute("UPDATE tickets SET updated_at=? WHERE id=?", (now(), tid))
            if public:
                mail_update(self.con, t, [], text)
            bump(self.con); self.con.commit()
        return self.send_json({"ok": True})

    # ------------------------------------------------------------ e-posta
    def mail_api(self, m, p):
        self.require("servis", 2 if m != "GET" else 1)
        if len(p) == 3 and p[2] == "tekrar" and m == "POST":
            with tx(self.con):
                st = settings(self.con)
                if not (st.get("smtp_host") and st.get("from_addr")):
                    raise HTTPError(400, "Önce Ayarlar ekranından e-posta sunucusunu tanımlayın.")
                self.con.execute("UPDATE emails SET status='bekliyor', error=NULL WHERE id=? AND status IN ('hata','ayar_yok')", (int(p[1]),))
                bump(self.con); self.con.commit()
            MAILQ.put(int(p[1]))
            return self.send_json({"ok": True})
        if p == ["eposta", "bekleyenler"] and m == "POST":
            with tx(self.con):
                st = settings(self.con)
                if not (st.get("smtp_host") and st.get("from_addr")):
                    raise HTTPError(400, "Önce e-posta sunucusunu tanımlayın.")
                ids = [r[0] for r in self.con.execute("SELECT id FROM emails WHERE status IN ('ayar_yok','hata')")]
                self.con.execute("UPDATE emails SET status='bekliyor', error=NULL WHERE status IN ('ayar_yok','hata')")
                bump(self.con); self.con.commit()
            for i in ids:
                MAILQ.put(i)
            return self.send_json({"ok": True, "n": len(ids)})
        if p == ["eposta"] and m == "GET":
            with tx(self.con):
                rows = self.con.execute("SELECT e.id, e.ticket_id, t.no, e.kind, e.to_addr, e.subject, e.status, e.error, e.created_at, e.sent_at "
                                        "FROM emails e LEFT JOIN tickets t ON t.id=e.ticket_id ORDER BY e.id DESC LIMIT 200").fetchall()
            return self.send_json({"emails": [dict(r) for r in rows]})
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ ayarlar
    def settings_api(self, m, p):
        self.require(admin=True)
        con = self.con
        if p == ["ayarlar"] and m == "GET":
            with tx(self.con):
                st = settings(con)
                pend = dict(con.execute("SELECT status, COUNT(*) FROM emails GROUP BY status").fetchall())
            st["smtp_password_set"] = bool(st.pop("smtp_password", ""))
            return self.send_json({"settings": st, "mailCounts": pend})
        if p == ["ayarlar"] and m == "PUT":
            b = self.body()
            with tx(self.con):
                for k in DEFAULT_SETTINGS:
                    if k not in b:
                        continue
                    v = s(b[k], 2000)
                    if k == "smtp_password" and not v:
                        continue  # boş bırakılırsa mevcut şifre korunur
                    if k in ("from_addr", "reply_to") and v and not EMAIL_RE.match(v):
                        raise HTTPError(400, f"Geçersiz e-posta adresi: {v}")
                    if k in ("smtp_port", "sla_days") and v and not v.isdigit():
                        raise HTTPError(400, "Port ve gün sayısı rakam olmalı.")
                    if k == "smtp_security" and v not in ("starttls", "ssl", "none"):
                        raise HTTPError(400, "Geçersiz güvenlik seçeneği.")
                    con.execute("UPDATE settings SET v=? WHERE k=?", (v, k))
                con.commit()
            return self.send_json({"ok": True})
        if p == ["ayarlar", "test"] and m == "POST":
            to = s(self.body().get("to"), 120)
            if not EMAIL_RE.match(to):
                raise HTTPError(400, "Geçerli bir alıcı e-posta adresi yazın.")
            with tx(self.con):
                st = settings(con)
            if not (st.get("smtp_host") and st.get("from_addr")):
                raise HTTPError(400, "Önce sunucu ve gönderen adresini kaydedin.")
            try:
                smtp_send(st, to, f"{st.get('company') or ''} Çağrı Karşılama – test e-postası",
                          "Bu bir test e-postasıdır. Çağrı Karşılama uygulamasının e-posta ayarları çalışıyor.\n\n" + (st.get("signature") or ""))
            except Exception as e:
                raise HTTPError(400, f"Gönderilemedi: {type(e).__name__}: {e}")
            return self.send_json({"ok": True})
        raise HTTPError(404, "Bulunamadı")

    # ------------------------------------------------------------ dosyalar
    def static(self, path):
        if path in ("/giris", "/giris.html"):
            if self.current_user():
                return self.redirect("/")
            return self.file("giris.html")
        if path == "/favicon.ico":
            raise HTTPError(404, "Bulunamadı")
        u = self.require(allow_must_change=True)
        if path == "/":
            return self.redirect("/satis" if u["crmLevel"] or not u["servisLevel"] else "/servis")
        if path == "/satis":
            return self.file("satis.html")
        if path == "/servis":
            return self.file("servis.html")
        return self.file(path.lstrip("/"))

    def file(self, rel):
        full = os.path.normpath(os.path.join(STATIC, rel))
        if not full.startswith(STATIC + os.sep) or not os.path.isfile(full):
            raise HTTPError(404, "Bulunamadı")
        ctype = {".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8", ".css": "text/css; charset=utf-8"}.get(
            os.path.splitext(full)[1], "application/octet-stream")
        with open(full, "rb") as f:
            raw = f.read()
        self.send_response(200)
        for h, v in (("Content-Type", ctype), ("Content-Length", str(len(raw))), ("Cache-Control", "no-cache"),
                     ("X-Frame-Options", "DENY"), ("X-Content-Type-Options", "nosniff")):
            self.send_header(h, v)
        self.end_headers()
        self.wfile.write(raw)


# ---------------------------------------------------------------- satış: fırsat doğrulama
CRM_STAGES = {"ilgi", "teklif", "kazanildi", "tamamlandi", "iptal", "kayip"}
DEAL_FIELDS = {"name", "customer", "contact", "partner", "amount", "cost", "currency", "pm", "company", "stage", "nextStep", "nextDate", "order"}


def _date(v):
    v = str(v or "").strip()
    if v and not DATE_RE.match(v):
        raise ValueError("Tarih biçimi YYYY-AA-GG olmalı")
    return v


def _num(v):
    try:
        return round(float(v or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def clean_project(p):
    if not isinstance(p, dict):
        raise ValueError("Geçersiz proje bilgisi")
    out = {"contractStart": _date(p.get("contractStart")), "contractEnd": _date(p.get("contractEnd")),
           "contractAmount": _num(p.get("contractAmount")), "costAmount": _num(p.get("costAmount")),
           "currency": str(p.get("currency") or "")[:5], "note": str(p.get("note") or "")[:2000], "invoices": [], "supplierPayments": []}
    if out["contractStart"] and out["contractEnd"] and out["contractEnd"] < out["contractStart"]:
        raise ValueError("Sözleşme bitiş tarihi başlangıçtan önce olamaz")
    for r in (p.get("invoices") or [])[:200]:
        if isinstance(r, dict):
            row = {"id": str(r.get("id") or secrets.token_hex(4))[:16], "no": str(r.get("no") or "")[:60], "date": _date(r.get("date")),
                   "amount": _num(r.get("amount")), "collectedDate": _date(r.get("collectedDate"))}
            if row["date"] or row["amount"] or row["collectedDate"]:
                out["invoices"].append(row)
    for r in (p.get("supplierPayments") or [])[:200]:
        if isinstance(r, dict):
            row = {"id": str(r.get("id") or secrets.token_hex(4))[:16], "supplier": str(r.get("supplier") or "")[:120],
                   "date": _date(r.get("date")), "amount": _num(r.get("amount"))}
            if row["supplier"] or row["date"] or row["amount"]:
                out["supplierPayments"].append(row)
    return out


def clean_deal(body, partial):
    out = {}
    if "project" in body:
        out["project"] = clean_project(body["project"])
    for k, v in body.items():
        if k not in DEAL_FIELDS:
            continue
        if k in ("amount", "cost", "order"):
            v = _num(v)
        elif k == "stage":
            if v not in CRM_STAGES:
                raise ValueError("Geçersiz aşama")
        else:
            v = str(v or "")[:2000]
        out[k] = v
    if not partial and not out.get("name"):
        raise ValueError("Fırsat adı gerekli")
    return out


def import_crm_excel(con, path):
    """Odoo CRM'den aşamaya göre gruplanmış fırsat dışa aktarımı (Ekspa-Proje.xlsx biçimi)."""
    import openpyxl
    smap = {"İlgilenilen Fırsatlar": "ilgi", "Teklif Verilen Fırsatlar": "teklif", "Kazanılan Projeler": "kazanildi",
            "Tamamlanan Projeler": "tamamlandi", "İptal Edilen Fırsatlar": "iptal", "İptal Edilen Fırstlar": "iptal", "Kayıp": "kayip"}
    cmap = {"EKS-PA BILGISAYAR PAZARLAMA VE TICARET ANONIM SIRKETI": "EKS-PA",
            "EKSPRES-BIM BILGISAYAR HIZMETLERISANAYI VE TICARET ANONIM SIRKETI": "EKSPRES-BİM"}
    ws = openpyxl.load_workbook(path, data_only=True).active
    existing = {json.loads(d)["name"].strip().lower(): i for i, d in con.execute("SELECT id, data FROM deals")}
    order = con.execute("SELECT COALESCE(MAX(json_extract(data,'$.order')),0) FROM deals").fetchone()[0]
    added = updated = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or len(r) < 10 or r[9] not in smap:
            continue
        st = lambda v: str(v).strip() if v is not None else ""
        name = " ".join(st(r[0]).split())
        rec = dict(name=name, customer=st(r[1]), contact=st(r[2]), partner=st(r[3]), amount=float(r[4] or 0), cost=float(r[5] or 0),
                   currency={"EURO": "EUR"}.get(r[6], r[6]) or "", pm=st(r[7]), company=cmap.get(r[8], st(r[8])), stage=smap[r[9]])
        if name.lower() in existing:
            did = existing[name.lower()]
            d = json.loads(con.execute("SELECT data FROM deals WHERE id=?", (did,)).fetchone()[0])
            d.update(rec, updatedAt=now(), updatedBy="Odoo içe aktarma")
            con.execute("UPDATE deals SET data=?, updated_at=? WHERE id=?", (json.dumps(d, ensure_ascii=False), now(), did))
            updated += 1
        else:
            order += 1
            did = "d" + secrets.token_hex(5)
            rec.update(nextStep="", nextDate="", notes=[], order=order, createdAt=now(), updatedAt=now(), updatedBy="Odoo içe aktarma")
            con.execute("INSERT INTO deals VALUES (?,?,?)", (did, json.dumps(rec, ensure_ascii=False), now()))
            existing[name.lower()] = did
            added += 1
    bump(con); con.commit()
    return added, updated


# ---------------------------------------------------------------- Excel içe aktarma
def read_inventory_xlsx(path):
    """Odoo 'Teknik Servis Envanter' dışa aktarımı: firmaya göre gruplanmış satırlar."""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    out = []
    d = lambda v: v.date().isoformat() if hasattr(v, "date") else (str(v)[:10] if v else "")
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or (isinstance(r[0], str) and all(x is None for x in r[1:])):
            continue  # grup başlığı satırı, örn. "SDT-Göktürk (1179)"
        if not r[2]:
            continue
        out.append({"firma": str(r[1] or "").strip(), "serial": " ".join(str(r[2]).split()), "brand": str(r[3] or "").strip(),
                    "part_no": str(r[4] or "").strip(), "description": " ".join(str(r[5] or "").split()), "service": str(r[6] or "").strip(),
                    "w_start": d(r[7]), "w_end": d(r[8]), "w_info": str(r[9] or "").strip()})
    return out


ODOO_STAGE = {"Yeni": "acik", "Devam Eden": "inceleniyor", "Beklemede": "musteri_bekleniyor", "Çözüldü": "cozuldu", "İptal Edildi": "iptal", "Kapandı": "kapandi"}
ODOO_PRIO = {"Acil": "acil", "Yüksek öncelik": "yuksek", "Orta öncelik": "normal", "Düşük öncelik": "dusuk"}


def read_tickets_xlsx(path):
    """Odoo 'Yardım Masası Talebi' dışa aktarımı."""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        cust = str(r[4] or "").strip()
        firma, _, contact = cust.partition(", ")
        out.append({"no": str(r[0]).strip().zfill(5), "priority": ODOO_PRIO.get(r[1], "normal"), "subject": str(r[2] or "").strip() or "(konu yok)",
                    "assignee_name": str(r[3] or "").strip(), "firma": firma.strip(), "contact_name": contact.strip(),
                    "stage": ODOO_STAGE.get(str(r[8] or "").strip(), "kapandi"), "odoo_stage": str(r[8] or "")})
    return out


def link_assignees(con):
    for name, un in con.execute("SELECT name, username FROM users").fetchall():
        con.execute("UPDATE tickets SET assignee=? WHERE assignee_name=? AND (assignee IS NULL OR assignee='')", (un, name))


# ---------------------------------------------------------------- komut satırı
def ask_pw():
    while True:
        a = getpass.getpass("Şifre: ")
        pr = pw_problem(a)
        if pr:
            print(pr); continue
        if getpass.getpass("Şifre (tekrar): ") != a:
            print("Şifreler aynı değil."); continue
        return a


def main():
    ap = argparse.ArgumentParser(description="Ekspa İş Takibi sunucusu (Satış + Teknik Servis)")
    ap.add_argument("--host", default=os.environ.get("EKSPA_HOST", "0.0.0.0"))
    # Railway (ve benzeri platformlar) dinlenecek portu PORT değişkeniyle verir.
    ap.add_argument("--port", type=int, default=int(os.environ.get("EKSPA_PORT") or os.environ.get("PORT") or "8080"))
    ap.add_argument("--import", dest="xlsx", metavar="XLSX", help="Odoo CRM fırsat dışa aktarımını içe al")
    ap.add_argument("--import-envanter", metavar="XLSX")
    ap.add_argument("--import-talepler", metavar="XLSX")
    ap.add_argument("--kullanici-ekle", metavar="KULLANICI_ADI")
    ap.add_argument("--yonetici", action="store_true", help="--kullanici-ekle ile: yönetici yap")
    ap.add_argument("--satis", choices=["duzenleyici", "goruntuleyici", "yok"], default="yok")
    ap.add_argument("--servis", choices=["uzman", "goruntuleyici", "yok"], default="yok")
    ap.add_argument("--sifre-sifirla", metavar="KULLANICI_ADI")
    ap.add_argument("--kullanicilar", action="store_true")
    a = ap.parse_args()
    con = init_db()
    if a.xlsx:
        ad, up = import_crm_excel(con, a.xlsx)
        return print(f"Fırsatlar içe aktarıldı: {ad} yeni, {up} güncellendi.")
    if a.import_envanter:
        rows = read_inventory_xlsx(a.import_envanter)
        existing = {(r[0], r[1]) for r in con.execute("SELECT serial_norm, firma FROM inventory")}
        new = [x for x in rows if (norm_serial(x["serial"]), x["firma"]) not in existing]
        upd = 0
        for x in rows:
            if (norm_serial(x["serial"]), x["firma"]) in existing:
                con.execute("UPDATE inventory SET brand=?, part_no=?, description=?, service=?, w_start=?, w_end=?, w_info=?, updated_at=? WHERE serial_norm=? AND firma=?",
                            (x["brand"], x["part_no"], x["description"], x["service"], x["w_start"], x["w_end"], x["w_info"], now(), norm_serial(x["serial"]), x["firma"]))
                upd += 1
        insert_inventory(con, new)
        con.commit()
        return print(f"Envanter içe aktarıldı: {len(new)} yeni, {upd} güncellendi.")
    if a.import_talepler:
        n = insert_tickets(con, read_tickets_xlsx(a.import_talepler))
        link_assignees(con); con.commit()
        return print(f"Talepler içe aktarıldı: {n} yeni çağrı (aynı numaralı kayıtlar atlandı).")
    if a.kullanicilar:
        for r in con.execute("SELECT * FROM users ORDER BY name"):
            yetki = "Yönetici" if r["role"] == "yonetici" else f"Satış: {CRM_T[r['crm_role'] or '']} · Servis: {SRV_T[r['servis_role'] or '']}"
            print(f"{r['username']:<18} {r['name']:<26} {r['email'] or '-':<30} {yetki:<44} {'etkin' if r['active'] else 'kapalı'}")
        return
    if a.kullanici_ekle:
        un = norm_username(a.kullanici_ekle)
        if not USERNAME_RE.match(un) or con.execute("SELECT 1 FROM users WHERE username=?", (un,)).fetchone():
            sys.exit("Geçersiz ya da zaten var olan kullanıcı adı.")
        name = input("Ad soyad: ").strip() or un
        email = input("E-posta: ").strip()
        add_user(con, un, name, email, a.yonetici, "" if a.satis == "yok" else a.satis, "" if a.servis == "yok" else a.servis, hash_pw(ask_pw()))
        link_assignees(con); con.commit()
        return print(f"Kullanıcı eklendi: {un}")
    if a.sifre_sifirla:
        un = norm_username(a.sifre_sifirla)
        if not con.execute("SELECT 1 FROM users WHERE username=?", (un,)).fetchone():
            sys.exit("Böyle bir kullanıcı yok.")
        con.execute("UPDATE users SET pw=?, active=1, must_change=0 WHERE username=?", (hash_pw(ask_pw()), un))
        con.execute("DELETE FROM sessions WHERE username=?", (un,))
        con.commit()
        return print("Şifre güncellendi.")
    srv = ThreadingHTTPServer((a.host, a.port), Handler)
    srv.con = con
    threading.Thread(target=mail_worker, args=(con,), daemon=True).start()
    for (mid,) in con.execute("SELECT id FROM emails WHERE status='bekliyor'").fetchall():
        MAILQ.put(mid)
    shown = "localhost" if a.host in ("0.0.0.0", "") else a.host
    url = f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}" if os.environ.get("RAILWAY_PUBLIC_DOMAIN") else f"http://{shown}:{a.port}"
    print(f"Ekspa İş Takibi çalışıyor: {url}  (port {a.port}, veri klasörü {DATA_DIR})"
          + ("" if ON_RAILWAY else "  — durdurmak için Ctrl+C"))
    if ON_RAILWAY and not os.environ.get("RAILWAY_VOLUME_MOUNT_PATH"):
        print("!! UYARI: Bu servise Railway Volume bağlı değil. Veritabanı her yeniden deploy'da SİLİNİR.\n"
              "!! Railway'de servise bir Volume ekleyin (Mount path: /veri) ve yeniden deploy edin.")
    if not con.execute("SELECT COUNT(*) FROM users").fetchone()[0]:
        print(f">> Henüz kullanıcı yok. Tarayıcıdan {url} adresini açıp ilk yönetici hesabını hemen oluşturun.")
        if SETUP_CODE_ENV:
            print(">> Kurulum kodu: EKSPA_KURULUM_KODU değişkenine yazdığınız değer.")
        elif SETUP_CODE:
            print(f">> Kurulum kodu: {SETUP_CODE[:4]}-{SETUP_CODE[4:8]}-{SETUP_CODE[8:]}  (her yeniden başlatmada değişir)")

    def stop(*_):
        raise KeyboardInterrupt
    # Railway yeniden deploy ederken SIGTERM gönderir; konteynerde PID 1 olarak bu sinyali kendimiz ele almalıyız.
    signal.signal(signal.SIGTERM, stop)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nKapatılıyor.")
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
