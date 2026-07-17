"""Small local SQLite store that remembers the adviser, dean and per-subject
faculty names typed in previous runs, so the wizard can suggest them next
time instead of asking from scratch. Nothing here ever leaves the machine.
"""
import os
import sqlite3
import sys
import time


def _data_dir():
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    path = os.path.join(base, "AdviseeDocFiller")
    os.makedirs(path, exist_ok=True)
    return path


DB_PATH = os.path.join(_data_dir(), "store.db")


def _connect():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS people (
        name TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('adviser', 'dean')),
        last_used REAL NOT NULL,
        PRIMARY KEY (name, role)
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS faculty_codes (
        subject_code TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        last_used REAL NOT NULL
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS faculty_names (
        name TEXT PRIMARY KEY,
        last_used REAL NOT NULL
    )""")
    return conn


def remember_person(role, name):
    """role: 'adviser' or 'dean'."""
    name = (name or "").strip()
    if not name:
        return
    with _connect() as conn:
        conn.execute(
            "INSERT INTO people (name, role, last_used) VALUES (?, ?, ?) "
            "ON CONFLICT(name, role) DO UPDATE SET last_used = excluded.last_used",
            (name, role, time.time()))


def known_people(role):
    """Most-recently-used first."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT name FROM people WHERE role = ? ORDER BY last_used DESC",
            (role,)).fetchall()
    return [r[0] for r in rows]


def remember_faculty(subject_code, name):
    subject_code = (subject_code or "").strip().replace(" ", "").upper()
    name = (name or "").strip()
    if not subject_code or not name:
        return
    now = time.time()
    with _connect() as conn:
        conn.execute(
            "INSERT INTO faculty_codes (subject_code, name, last_used) "
            "VALUES (?, ?, ?) ON CONFLICT(subject_code) DO UPDATE SET "
            "name = excluded.name, last_used = excluded.last_used",
            (subject_code, name, now))
        conn.execute(
            "INSERT INTO faculty_names (name, last_used) VALUES (?, ?) "
            "ON CONFLICT(name) DO UPDATE SET last_used = excluded.last_used",
            (name, now))


def faculty_for_codes(codes):
    """{subject_code: name} for previously remembered codes, blank for the
    rest so callers can prefill known ones and leave others empty."""
    codes = [c.strip().replace(" ", "").upper() for c in codes if c.strip()]
    if not codes:
        return {}
    with _connect() as conn:
        q = ("SELECT subject_code, name FROM faculty_codes "
             "WHERE subject_code IN ({})".format(",".join("?" * len(codes))))
        rows = conn.execute(q, codes).fetchall()
    return dict(rows)


def known_faculty_names():
    """Every distinct faculty name ever entered, most-recently-used first."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT name FROM faculty_names ORDER BY last_used DESC").fetchall()
    return [r[0] for r in rows]
