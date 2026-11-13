"""
SQLite (local) + PostgreSQL (Render) persistence layer for BS Lead Manager.
"""
import os
import sqlite3
from datetime import datetime

DATABASE_URL = os.environ.get("DATABASE_URL", "")
USE_PG = bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))

if USE_PG:
    import psycopg2
    import psycopg2.extras

DB_PATH = os.path.join(os.path.dirname(__file__), "leads.db")

STATUS_VALUES = ('pending', 'contacted', 'interested', 'meeting', 'closed', 'lost')

def _conn():
    if USE_PG:
        return psycopg2.connect(DATABASE_URL, sslmode="require")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def _fetchall(cursor):
    if USE_PG:
        cols = [desc[0] for desc in cursor.description]
        return [dict(zip(cols, row)) for row in cursor.fetchall()]
    return [dict(row) for row in cursor.fetchall()]

def _fetchone(cursor):
    if USE_PG:
        row = cursor.fetchone()
        if row is None: return None
        cols = [desc[0] for desc in cursor.description]
        return dict(zip(cols, row))
    row = cursor.fetchone()
    return dict(row) if row else None

def p():
    return "%s" if USE_PG else "?"

def init_db():
    conn = _conn()
    cur  = conn.cursor()
    if USE_PG:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS weeks (
                id         SERIAL PRIMARY KEY,
                week_label TEXT NOT NULL,
                created_at TEXT NOT NULL
            )""")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS leads (
                id         SERIAL PRIMARY KEY,
                week_id    INTEGER NOT NULL,
                name       TEXT NOT NULL,
                category   TEXT NOT NULL DEFAULT 'restaurant',
                email      TEXT,
                phone      TEXT,
                address    TEXT,
                suburb     TEXT,
                state      TEXT,
                website    TEXT,
                priority   TEXT NOT NULL DEFAULT 'medium',
                reason     TEXT,
                status     TEXT NOT NULL DEFAULT 'pending',
                notes      TEXT DEFAULT '',
                sent       INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY (week_id) REFERENCES weeks(id)
            )""")
        for col_def in [
            "ALTER TABLE leads ADD COLUMN IF NOT EXISTS status TEXT NOT NULL DEFAULT 'pending'",
            "ALTER TABLE leads ADD COLUMN IF NOT EXISTS notes TEXT DEFAULT ''",
        ]:
            try: cur.execute(col_def)
            except: pass
        cur.execute("UPDATE leads SET status='contacted' WHERE sent=1 AND status='pending'")
    else:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS weeks (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                week_label TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS leads (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                week_id    INTEGER NOT NULL,
                name       TEXT NOT NULL,
                category   TEXT NOT NULL DEFAULT 'restaurant',
                email      TEXT,
                phone      TEXT,
                address    TEXT,
                suburb     TEXT,
                state      TEXT,
                website    TEXT,
                priority   TEXT NOT NULL DEFAULT 'medium',
                reason     TEXT,
                status     TEXT NOT NULL DEFAULT 'pending',
                notes      TEXT DEFAULT '',
                sent       INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY (week_id) REFERENCES weeks(id)
            );
        """)
        for col_def in [
            "ALTER TABLE leads ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'",
            "ALTER TABLE leads ADD COLUMN notes TEXT DEFAULT ''",
        ]:
            try: conn.execute(col_def)
            except: pass
        conn.execute("UPDATE leads SET status='contacted' WHERE sent=1 AND (status IS NULL OR status='pending')")
    conn.commit()
    conn.close()

def get_or_create_current_week() -> int:
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    if USE_PG:
        cur.execute("SELECT id FROM weeks WHERE DATE(created_at::timestamp) = CURRENT_DATE ORDER BY id DESC LIMIT 1")
    else:
        cur.execute("SELECT id FROM weeks WHERE DATE(created_at) = DATE('now') ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    if row:
        week_id = row[0] if not isinstance(row, dict) else row["id"]
        conn.close()
        return week_id
    today = datetime.now()
    label = f"Week of {today.strftime('%d %b %Y')}"
    if USE_PG:
        cur.execute(f"INSERT INTO weeks (week_label, created_at) VALUES ({ph},{ph}) RETURNING id", (label, today.isoformat()))
        week_id = cur.fetchone()[0]
    else:
        cur.execute(f"INSERT INTO weeks (week_label, created_at) VALUES ({ph},{ph})", (label, today.isoformat()))
        week_id = cur.lastrowid
    conn.commit()
    conn.close()
    return week_id

def get_week(week_id: int) -> dict | None:
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute(f"SELECT * FROM weeks WHERE id = {ph}", (week_id,))
    row = _fetchone(cur)
    conn.close()
    return row

def get_weeks() -> list:
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        SELECT w.id, w.week_label, w.created_at,
               COUNT(l.id) AS total_leads,
               SUM(CASE WHEN l.status IN ('contacted','interested','meeting','closed') THEN 1 ELSE 0 END) AS sent_leads,
               SUM(CASE WHEN l.email IS NOT NULL AND l.email != '' THEN 1 ELSE 0 END) AS email_leads
        FROM weeks w
        LEFT JOIN leads l ON l.week_id = w.id
        GROUP BY w.id, w.week_label, w.created_at
        ORDER BY w.created_at DESC
    """)
    rows = _fetchall(cur)
    conn.close()
    return rows

def get_all_business_names() -> list:
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("SELECT name, email FROM leads")
    rows = _fetchall(cur)
    conn.close()
    return [(r["name"], r["email"] or "") for r in rows]

def save_leads(leads: list, week_id: int) -> list:
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("SELECT name FROM leads")
    existing_names  = {r["name"].lower() for r in _fetchall(cur)}
    cur.execute("SELECT email FROM leads WHERE email IS NOT NULL AND email != ''")
    existing_emails = {r["email"].lower() for r in _fetchall(cur)}
    saved = []
    now   = datetime.now().isoformat()
    for lead in leads:
        name  = (lead.get("name")  or "").strip()
        email = (lead.get("email") or "").strip().lower()
        if not name or name.lower() in existing_names: continue
        if email and email in existing_emails: continue
        vals = (week_id, name, lead.get("category","restaurant"), email or None,
                lead.get("phone",""), lead.get("address",""), lead.get("suburb",""),
                lead.get("state",""), lead.get("website",""),
                lead.get("priority","medium"), lead.get("reason",""), now)
        sql = f"""INSERT INTO leads (week_id,name,category,email,phone,address,suburb,
                  state,website,priority,reason,status,notes,sent,created_at)
                  VALUES ({ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},{ph},'pending','',0,{ph})"""
        if USE_PG:
            cur.execute(sql + " RETURNING id", vals)
            new_id = cur.fetchone()[0]
        else:
            cur.execute(sql, vals)
            new_id = cur.lastrowid
        existing_names.add(name.lower())
        if email: existing_emails.add(email)
        saved.append({**lead, "id": new_id, "week_id": week_id, "status": "pending", "notes": "", "sent": 0})
    conn.commit()
    conn.close()
    return saved

def get_leads_by_week(week_id: int) -> list:
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute(f"""
        SELECT * FROM leads WHERE week_id = {ph}
        ORDER BY CASE priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 ELSE 3 END, name
    """, (week_id,))
    rows = _fetchall(cur)
    conn.close()
    return rows

def get_all_leads() -> list:
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("""
        SELECT l.*, w.week_label FROM leads l
        JOIN weeks w ON l.week_id = w.id
        ORDER BY l.created_at DESC
    """)
    rows = _fetchall(cur)
    conn.close()
    return rows

def get_stats() -> dict:
    conn = _conn()
    cur  = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM leads");                                           total      = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='pending'");                   pending    = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='contacted'");                 contacted  = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='interested'");                interested = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='meeting'");                   meeting    = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='closed'");                    closed     = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE status='lost'");                      lost       = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM leads WHERE email IS NOT NULL AND email != ''");  with_email = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM weeks");                                           weeks      = cur.fetchone()[0]
    conn.close()
    return {
        "total": total, "pending": pending, "contacted": contacted,
        "interested": interested, "meeting": meeting, "closed": closed,
        "lost": lost, "sent": total - pending, "weeks": weeks, "with_email": with_email,
    }

def update_status(lead_id: int, status: str):
    if status not in STATUS_VALUES:
        status = 'pending'
    ph      = p()
    conn    = _conn()
    cur     = conn.cursor()
    sent_val = 0 if status == 'pending' else 1
    cur.execute(f"UPDATE leads SET status={ph}, sent={ph} WHERE id={ph}", (status, sent_val, lead_id))
    conn.commit()
    conn.close()

def update_notes(lead_id: int, notes: str):
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute(f"UPDATE leads SET notes={ph} WHERE id={ph}", (notes or "", lead_id))
    conn.commit()
    conn.close()

def mark_sent(lead_id: int, sent: bool):
    update_status(lead_id, 'contacted' if sent else 'pending')

def delete_lead(lead_id: int):
    ph   = p()
    conn = _conn()
    cur  = conn.cursor()
    cur.execute(f"DELETE FROM leads WHERE id = {ph}", (lead_id,))
    conn.commit()
    conn.close()

def get_current_week_leads() -> list:
    conn = _conn()
    cur  = conn.cursor()
    if USE_PG:
        cur.execute("SELECT id FROM weeks WHERE DATE(created_at::timestamp) = CURRENT_DATE ORDER BY id DESC LIMIT 1")
    else:
        cur.execute("SELECT id FROM weeks WHERE DATE(created_at) = DATE('now') ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    if not row:
        conn.close()
        return []
    week_id = row[0] if not isinstance(row, dict) else row["id"]
    ph = p()
    cur.execute(f"""
        SELECT * FROM leads WHERE week_id = {ph}
        ORDER BY CASE priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 ELSE 3 END, name
    """, (week_id,))
    rows = _fetchall(cur)
    conn.close()
    return rows
