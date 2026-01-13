from flask import Flask, render_template, request, redirect, url_for, session
import sqlite3, os, uuid
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev_only_change_me")
DB_PATH = "orders.db"

# ----- Config -----
PREPARING_STATION  = "preparing"
CNC_STATIONS       = ["cnc1", "cnc2", "cnc3"]
TRAMMING1_STATION  = "tramming1"
TRAMMING2_STATION  = "tramming2"
EDGE_STATIONS      = ["edge1", "edge2", "edge3", "edge4"]
WRAPPING_STATION   = "wrapping"
LOADING_STATION    = "loading"
MANAGER_AREA       = "manager"

# ----- Training config -----
TRAINING_EMPLOYEES = [
    ("adam_smith",    "Adam Smith"),
    ("ben_johnson",   "Ben Johnson"),
    ("charlie_brown", "Charlie Brown"),
    ("daniel_harris", "Daniel Harris"),
    ("emily_clark",   "Emily Clark"),
    ("frank_lewis",   "Frank Lewis"),
    ("grace_wilson",  "Grace Wilson"),
    ("harry_walker",  "Harry Walker"),
    ("isla_young",    "Isla Young"),
    ("jack_king",     "Jack King"),
]

TRAINING_STATIONS = [
    ("preparing",          "Preparing"),
    ("cnc1",               "CNC 1"),
    ("cnc2",               "CNC 2"),
    ("cnc3",               "CNC 3"),
    ("tramming1",          "Tramming 1"),
    ("edge1",              "Edge 1"),
    ("edge2",              "Edge 2"),
    ("edge3",              "Edge 3"),
    ("edge4",              "Edge 4"),
    ("tramming2",          "Tramming 2"),
    ("wrapping",           "Wrapping"),
    ("loading",            "Loading"),
    ("forklift",           "Forklift"),
    ("extraction_system",  "Extraction system"),
    ("change_blades",      "Change blades"),
    ("tape_change",        "Tape change"),
    ("first_aid",          "First aid"),
    ("fire_marshal",       "Fire marshal"),
]

# Map station codes to labels
STATION_LABELS = {code: label for code, label in TRAINING_STATIONS}

# Stations that appear on Weekly staffing screen
STAFFING_STATIONS = [
    PREPARING_STATION,
    *CNC_STATIONS,
    TRAMMING1_STATION,
    *EDGE_STATIONS,
    TRAMMING2_STATION,
    WRAPPING_STATION,
    LOADING_STATION,
]

# Weekdays only, no weekend
DAYS_OF_WEEK = [
    ("mon", "Monday"),
    ("tue", "Tuesday"),
    ("wed", "Wednesday"),
    ("thu", "Thursday"),
    ("fri", "Friday"),
]

# Group multiple machines into single columns in Manager view
AREA_GROUPS = {
    "cnc": CNC_STATIONS,
    "edge": EDGE_STATIONS,
}

# Columns shown on Manager screen, left to right
MANAGER_STATIONS = [
    PREPARING_STATION,
    "cnc",
    TRAMMING1_STATION,
    "edge",
    TRAMMING2_STATION,
    WRAPPING_STATION,
    LOADING_STATION,
]

QUEUE_CAPACITY_PREP = 2
EDGE_CAPACITY       = {"edge1": 5, "edge2": 4, "edge3": 3, "edge4": 3}
WRAP_SLOTS          = [1, 2, 3]
LORRY_CAPACITY      = 12

# Settings keys
KEY_LORRY1 = "lorry1_num"     # current number displayed on LEFT slot
KEY_LORRY2 = "lorry2_num"     # current number displayed on RIGHT slot
KEY_NEXT   = "next_lorry_num" # next number to assign to whichever slot completes next

# ----- DB init -----
def init_db():
    first_time = not os.path.exists(DB_PATH)
    conn = sqlite3.connect(DB_PATH); c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_number TEXT NOT NULL,
        status TEXT NOT NULL,
        current_station TEXT,
        queued_at   TEXT DEFAULT CURRENT_TIMESTAMP,
        started_at  TEXT,
        finished_at TEXT,
        lorry       TEXT,
        wrap_slot   INTEGER,
        batch_id    TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL,
        area TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    # Order history log
    c.execute("""
    CREATE TABLE IF NOT EXISTS order_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        station TEXT NOT NULL,
        status TEXT NOT NULL,
        changed_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Training matrix table
    c.execute("""
    CREATE TABLE IF NOT EXISTS training (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee TEXT NOT NULL,
        station TEXT NOT NULL,
        status TEXT NOT NULL    -- 'not_trained', 'partial', 'full'
    )
    """)

    # Ensure (employee, station) pair is unique so we can upsert
    c.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_training_emp_station
        ON training(employee, station)
    """)

    # Weekly staffing plan: one employee per station per weekday
    c.execute("""
    CREATE TABLE IF NOT EXISTS staffing_plan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        station TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        employee TEXT NOT NULL
    )
    """)
    c.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_staffing_station_day
        ON staffing_plan(station, day_of_week)
    """)

    # Backfill columns, safe if already exist
    for alter in [
        "ALTER TABLE orders ADD COLUMN current_station TEXT",
        "ALTER TABLE orders ADD COLUMN queued_at TEXT DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE orders ADD COLUMN started_at TEXT",
        "ALTER TABLE orders ADD COLUMN finished_at TEXT",
        "ALTER TABLE orders ADD COLUMN lorry TEXT",
        "ALTER TABLE orders ADD COLUMN wrap_slot INTEGER",
        "ALTER TABLE orders ADD COLUMN batch_id TEXT",
    ]:
        try: c.execute(alter)
        except sqlite3.OperationalError: pass

    # Seed accounts on first run
    if first_time:
        def add_user(u, pw, role, area):
            c.execute("INSERT OR IGNORE INTO users (username,password_hash,role,area) VALUES (?,?,?,?)",
                      (u, generate_password_hash(pw), role, area))
        # station users
        add_user(PREPARING_STATION, "prep123", "station", PREPARING_STATION)
        for n in CNC_STATIONS: add_user(n, "cnc123", "station", n)
        add_user(TRAMMING1_STATION, "tram123", "station", TRAMMING1_STATION)
        add_user(TRAMMING2_STATION, "tram123", "station", TRAMMING2_STATION)
        for n in EDGE_STATIONS: add_user(n, "edge123", "station", n)
        add_user(WRAPPING_STATION, "wrap123", "station", WRAPPING_STATION)
        add_user(LOADING_STATION,  "load123", "station", LOADING_STATION)
        # manager user
        add_user("manager", "manager123", "manager", MANAGER_AREA)

    # Lorry numbers
    c.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)", (KEY_LORRY1, "1"))
    c.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)", (KEY_LORRY2, "2"))
    c.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)", (KEY_NEXT,   "3"))

    c.execute("SELECT value FROM settings WHERE key=?", (KEY_LORRY1,)); l1 = int((c.fetchone() or ["1"])[0])
    c.execute("SELECT value FROM settings WHERE key=?", (KEY_LORRY2,)); l2 = int((c.fetchone() or ["2"])[0])
    c.execute("SELECT value FROM settings WHERE key=?", (KEY_NEXT,));   nx = int((c.fetchone() or ["3"])[0])

    changed = False
    if l2 == l1:
        l2 = l1 + 1
        c.execute("UPDATE settings SET value=? WHERE key=?", (str(l2), KEY_LORRY2)); changed = True
    if nx <= max(l1, l2):
        nx = max(l1, l2) + 1
        c.execute("UPDATE settings SET value=? WHERE key=?", (str(nx), KEY_NEXT)); changed = True
    if changed: conn.commit()

    conn.commit(); conn.close()

def ensure_users():
    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
    def ensure(u, pw, role, area):
        c.execute("INSERT OR IGNORE INTO users (username,password_hash,role,area) VALUES (?,?,?,?)",
                  (u, generate_password_hash(pw), role, area))
    ensure(PREPARING_STATION, "prep123", "station", PREPARING_STATION)
    for n in CNC_STATIONS: ensure(n, "cnc123", "station", n)
    ensure(TRAMMING1_STATION, "tram123", "station", TRAMMING1_STATION)
    ensure(TRAMMING2_STATION, "tram123", "station", TRAMMING2_STATION)
    for n in EDGE_STATIONS: ensure(n, "edge123", "station", n)
    ensure(WRAPPING_STATION, "wrap123", "station", WRAPPING_STATION)
    ensure(LOADING_STATION,  "load123", "station", LOADING_STATION)
    ensure("manager", "manager123", "manager", MANAGER_AREA)
    conn.commit(); conn.close()

init_db(); ensure_users()

# ----- helpers -----
def get_db(): return sqlite3.connect(DB_PATH)

def get_setting(key, default="1"):
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = c.fetchone(); conn.close()
    return row[0] if row else default

def set_setting(key, value):
    conn = get_db(); c = conn.cursor()
    c.execute(
        "INSERT INTO settings(key,value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value)
    )
    conn.commit(); conn.close()

def get_lorry_state():
    l1 = int(get_setting(KEY_LORRY1, "1"))
    l2 = int(get_setting(KEY_LORRY2, "2"))
    nx = int(get_setting(KEY_NEXT,   str(max(l1, l2) + 1)))
    if l2 == l1:
        l2 = l1 + 1; set_setting(KEY_LORRY2, str(l2))
    if nx <= max(l1, l2):
        nx = max(l1, l2) + 1; set_setting(KEY_NEXT, str(nx))
    return f"Lorry {l1}", f"Lorry {l2}", l1, l2, nx

def advance_lorry(slot: int):
    l1 = int(get_setting(KEY_LORRY1, "1"))
    l2 = int(get_setting(KEY_LORRY2, "2"))
    nx = int(get_setting(KEY_NEXT,   str(max(l1, l2) + 1)))
    if slot == 1:
        set_setting(KEY_LORRY1, str(nx))
    else:
        set_setting(KEY_LORRY2, str(nx))
    set_setting(KEY_NEXT, str(nx + 1))

def login_required(f):
    from functools import wraps
    @wraps(f)
    def wrapped(*a, **kw):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return f(*a, **kw)
    return wrapped

# History helpers
def log_status(order_id: int, station: str, status: str):
    conn = get_db(); c = conn.cursor()
    c.execute("INSERT INTO order_history (order_id, station, status) VALUES (?,?,?)",
              (order_id, station, status))
    conn.commit(); conn.close()

def log_many(pairs):
    if not pairs:
        return
    conn = get_db(); c = conn.cursor()
    c.executemany("INSERT INTO order_history (order_id, station, status) VALUES (?,?,?)", pairs)
    conn.commit(); conn.close()

# Training helper, build default "not_trained" matrix then overlay DB values
def get_training_status_matrix():
    """
    Returns a dict:
        { emp_key: { station_code: status_string } }
    Default status is 'not_trained' for every cell.
    """
    matrix = {
        emp_key: {st_code: "not_trained" for st_code, _ in TRAINING_STATIONS}
        for emp_key, _ in TRAINING_EMPLOYEES
    }

    conn = get_db(); c = conn.cursor()
    c.execute("SELECT employee, station, status FROM training")
    for emp, st, status in c.fetchall():
        if emp in matrix and st in matrix[emp]:
            matrix[emp][st] = status
    conn.close()
    return matrix

# ----- auth -----
@app.route("/login", methods=["GET","POST"])
def login():
    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip().lower()
        password = request.form.get("password") or ""
        conn = get_db(); c = conn.cursor()
        c.execute("SELECT id, username, password_hash, role, area FROM users WHERE username=?", (username,))
        row = c.fetchone(); conn.close()
        if row and check_password_hash(row[2], password):
            session["user_id"] = row[0]
            session["username"] = row[1]
            session["role"] = row[3]
            session["area"] = row[4]
            return redirect(url_for("home"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
def home():
    if session.get("user_id"):
        area = (session.get("area") or "").lower()
        if area == MANAGER_AREA:        return redirect(url_for("manager_view"))
        if area == PREPARING_STATION:   return redirect(url_for("preparing_station"))
        if area in CNC_STATIONS:        return redirect(url_for("cnc_station"))
        if area == TRAMMING1_STATION:   return redirect(url_for("tramming1_station"))
        if area in EDGE_STATIONS:       return redirect(url_for("edge_station"))
        if area == TRAMMING2_STATION:   return redirect(url_for("tramming2_station"))
        if area == WRAPPING_STATION:    return redirect(url_for("wrapping_station"))
        if area == LOADING_STATION:     return redirect(url_for("loading_station"))
    return redirect(url_for("login"))

# ----- Preparing -----
@app.route("/preparing", methods=["GET","POST"])
@login_required
def preparing_station():
    if (session.get("area") or "").lower() != PREPARING_STATION:
        return ("Forbidden: not Preparing", 403)
    conn = get_db(); c = conn.cursor()

    if request.method == "POST":
        order_no = (request.form.get("order_number") or "").strip()
        target_cnc = (request.form.get("target_cnc") or "").strip().lower()
        c.execute("""
          SELECT id, order_number, status, current_station,
                 datetime(queued_at,'localtime'), datetime(started_at,'localtime'), datetime(finished_at,'localtime')
          FROM orders WHERE order_number=? LIMIT 1
        """, (order_no,))
        dup = c.fetchone()
        if dup:
            ts = dup[6] or dup[5] or dup[4] or "—"
            conn.close()
            return render_template("confirm.html",
                title="Order already exists",
                message=f"Order {order_no} already exists, status: {dup[2]}, at: {dup[3].upper() if dup[3] else '—'}, time: {ts}.",
                confirm_name=None,
                cancel_url=url_for("preparing_station"),
                post_url=None,
                hidden_fields={}
            )
        # capacity
        c.execute("SELECT COUNT(*) FROM orders WHERE current_station=? AND status='Pending'", (target_cnc,))
        (count_pending,) = c.fetchone()
        if count_pending >= QUEUE_CAPACITY_PREP:
            available = []
            for cnc in CNC_STATIONS:
                c.execute("SELECT COUNT(*) FROM orders WHERE current_station=? AND status='Pending'", (cnc,))
                (cnt,) = c.fetchone()
                if cnt < QUEUE_CAPACITY_PREP:
                    available.append(cnc.upper())
            if not available:
                conn.close()
                return render_template("confirm.html",
                    title="All CNC queues are full",
                    message="All CNCs are full at the moment. Please wait.",
                    confirm_name=None,
                    cancel_url=url_for("preparing_station"),
                    post_url=None,
                    hidden_fields={}
                )
            conn.close()
            return render_template("confirm.html",
                title="Queue full",
                message=f"{target_cnc.upper()} is full. Available: {', '.join(available)}",
                confirm_name=None,
                cancel_url=url_for("preparing_station"),
                post_url=None,
                hidden_fields={}
            )
        # confirm and insert
        if not request.form.get("confirm"):
            conn.close()
            return render_template("confirm.html",
                title="Confirm add",
                message=f"Add order {order_no} to {target_cnc.upper()} as Pending?",
                confirm_name="confirm",
                cancel_url=url_for("preparing_station"),
                post_url=url_for("preparing_station"),
                hidden_fields={"order_number": order_no, "target_cnc": target_cnc}
            )
        c.execute(
            "INSERT INTO orders (order_number, status, current_station) "
            "VALUES (?, 'Pending', ?)",
            (order_no, target_cnc)
        )
        oid = c.lastrowid
        conn.commit(); conn.close()
        # history
        log_many([(oid, PREPARING_STATION, "done"), (oid, target_cnc, "pending")])
        return redirect(url_for("preparing_station"))

    cnc_slots = {}
    for cnc in CNC_STATIONS:
        c.execute("""
            SELECT id, order_number, status, datetime(queued_at,'localtime')
            FROM orders
            WHERE current_station=? AND status='Pending'
            ORDER BY queued_at ASC
            LIMIT ?
        """, (cnc, QUEUE_CAPACITY_PREP))
        cnc_slots[cnc] = c.fetchall()
    conn.close()
    return render_template("station_preparing_one.html", cnc_slots=cnc_slots, capacity=QUEUE_CAPACITY_PREP)

# ----- CNC -----
@app.route("/cnc", methods=["GET","POST"])
@login_required
def cnc_station():
    area = (session.get("area") or "").lower()
    if area not in CNC_STATIONS:
        return ("Forbidden: not a CNC station", 403)
    conn = get_db(); c = conn.cursor()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='Pending'
        ORDER BY queued_at ASC
    """, (area,))
    pending = c.fetchall()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime'), datetime(started_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='In progress'
        ORDER BY started_at ASC NULLS LAST, queued_at ASC
    """, (area,))
    inprog = c.fetchall()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime'), datetime(finished_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='Done'
        ORDER BY finished_at ASC NULLS LAST, queued_at ASC
    """, (area,))
    done = c.fetchall()

    if request.method == "POST":
        action = request.form.get("action")
        oid = request.form.get("order_id")
        if action == "start":
            if inprog:
                conn.close()
                return render_template("confirm.html",
                    title="Cannot start",
                    message="You already have a job In progress. Finish it before starting another.",
                    confirm_name=None,
                    cancel_url=url_for("cnc_station"),
                    post_url=None,
                    hidden_fields={}
                )
            first_id = pending[0][0] if pending else None
            if str(oid) != str(first_id):
                conn.close()
                return render_template("confirm.html",
                    title="Cannot start",
                    message="You can only start the first Pending job.",
                    confirm_name=None,
                    cancel_url=url_for("cnc_station"),
                    post_url=None,
                    hidden_fields={}
                )
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm start",
                    message=f"Start this job on {area.upper()}?",
                    confirm_name="confirm",
                    cancel_url=url_for("cnc_station"),
                    post_url=url_for("cnc_station"),
                    hidden_fields={"action": "start", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='In progress', started_at=CURRENT_TIMESTAMP
                WHERE id=? AND current_station=? AND status='Pending'
            """, (oid, area))
            d.commit(); d.close(); conn.close()
            log_status(int(oid), area, "in_progress")
            return redirect(url_for("cnc_station"))

        if action == "finish":
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm finish",
                    message="Mark this job Finished on CNC?",
                    confirm_name="confirm",
                    cancel_url=url_for("cnc_station"),
                    post_url=url_for("cnc_station"),
                    hidden_fields={"action": "finish", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='Done', finished_at=CURRENT_TIMESTAMP
                WHERE id=? AND current_station=? AND status='In progress'
            """, (oid, area))
            d.commit(); d.close(); conn.close()
            log_many([(int(oid), area, "done"), (int(oid), TRAMMING1_STATION, "pending")])
            return redirect(url_for("cnc_station"))

    conn.close()
    return render_template(
        "station_cnc.html",
        station=area.upper(),
        pending=pending,
        inprog=inprog,
        done=done
    )

# ----- Tramming 1 -----
@app.route("/tramming1", methods=["GET","POST"])
@login_required
def tramming1_station():
    if (session.get("area") or "").lower() != TRAMMING1_STATION:
        return ("Forbidden: not Tramming 1", 403)
    conn = get_db(); c = conn.cursor()

    cnc_done = {}
    for cnc in CNC_STATIONS:
        c.execute("""
            SELECT id, order_number, datetime(finished_at,'localtime'), datetime(queued_at,'localtime')
            FROM orders
            WHERE current_station=? AND status='Done'
            ORDER BY finished_at ASC NULLS LAST, queued_at ASC
        """, (cnc,))
        cnc_done[cnc] = c.fetchall()

    edge_pending = {}
    for ed in EDGE_STATIONS:
        c.execute("""
            SELECT id, order_number, datetime(queued_at,'localtime')
            FROM orders
            WHERE current_station=? AND status='Pending'
            ORDER BY queued_at ASC
        """, (ed,))
        edge_pending[ed] = c.fetchall()

    if request.method == "POST" and request.form.get("action") == "assign":
        order_id = request.form.get("order_id")
        src_cnc = request.form.get("src_cnc")
        tgt_edge = request.form.get("tgt_edge")
        first_id = cnc_done.get(src_cnc, [None])
        first_id = first_id[0][0] if first_id else None
        if str(order_id) != str(first_id):
            conn.close()
            return render_template("confirm.html",
                title="Cannot assign",
                message="Only the first Finished job in each CNC lane can be assigned.",
                confirm_name=None,
                cancel_url=url_for("tramming1_station"),
                post_url=None,
                hidden_fields={}
            )
        if not tgt_edge:
            conn.close()
            return render_template("confirm.html",
                title="Choose Edge Bander",
                message="Please select an Edge Bander.",
                confirm_name=None,
                cancel_url=url_for("tramming1_station"),
                post_url=None,
                hidden_fields={}
            )
        cap = EDGE_CAPACITY.get(tgt_edge, 0)
        c.execute("SELECT COUNT(*) FROM orders WHERE current_station=? AND status='Pending'", (tgt_edge,))
        (count_pending,) = c.fetchone()
        if count_pending >= cap:
            available = []
            for ed in EDGE_STATIONS:
                c.execute("SELECT COUNT(*) FROM orders WHERE current_station=? AND status='Pending'", (ed,))
                (cnt,) = c.fetchone()
                if cnt < EDGE_CAPACITY[ed]:
                    available.append(ed.upper())
            msg = f"{tgt_edge.upper()} is full ({count_pending}/{cap})."
            if available:
                msg += " Available: " + ", ".join(available)
            conn.close()
            return render_template("confirm.html",
                title="Capacity full",
                message=msg,
                confirm_name=None,
                cancel_url=url_for("tramming1_station"),
                post_url=None,
                hidden_fields={}
            )
        if not request.form.get("confirm"):
            conn.close()
            return render_template("confirm.html",
                title="Confirm assignment",
                message=f"Assign order to {tgt_edge.upper()}?",
                confirm_name="confirm",
                cancel_url=url_for("tramming1_station"),
                post_url=url_for("tramming1_station"),
                hidden_fields={"action": "assign", "order_id": order_id, "src_cnc": src_cnc, "tgt_edge": tgt_edge}
            )
        d = get_db(); dc = d.cursor()
        dc.execute("""
            UPDATE orders
            SET current_station=?, status='Pending', queued_at=CURRENT_TIMESTAMP
            WHERE id=? AND current_station=? AND status='Done'
        """, (tgt_edge, order_id, src_cnc))
        d.commit(); d.close(); conn.close()
        log_many([(int(order_id), TRAMMING1_STATION, "done"), (int(order_id), tgt_edge, "pending")])
        return redirect(url_for("tramming1_station"))

    conn.close()
    return render_template(
        "station_tramming1.html",
        cnc_done=cnc_done,
        edge_pending=edge_pending,
        capacities=EDGE_CAPACITY
    )

# ----- Edge bander -----
@app.route("/edge", methods=["GET","POST"])
@login_required
def edge_station():
    area = (session.get("area") or "").lower()
    if area not in EDGE_STATIONS:
        return ("Forbidden: not an Edge station", 403)
    conn = get_db(); c = conn.cursor()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='Pending'
        ORDER BY queued_at ASC
    """, (area,))
    pending = c.fetchall()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime'), datetime(started_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='In progress'
        ORDER BY started_at ASC NULLS LAST, queued_at ASC
    """, (area,))
    inprog = c.fetchall()
    c.execute("""
        SELECT id, order_number, datetime(queued_at,'localtime'), datetime(finished_at,'localtime')
        FROM orders
        WHERE current_station=? AND status='Done'
        ORDER BY finished_at ASC NULLS LAST, queued_at ASC
    """, (area,))
    done = c.fetchall()

    if request.method == "POST":
        action = request.form.get("action")
        oid = request.form.get("order_id")
        if action == "start":
            if inprog:
                conn.close()
                return render_template("confirm.html",
                    title="Cannot start",
                    message="You already have a job In progress. Finish it before starting another.",
                    confirm_name=None,
                    cancel_url=url_for("edge_station"),
                    post_url=None,
                    hidden_fields={}
                )
            first_id = pending[0][0] if pending else None
            if str(oid) != str(first_id):
                conn.close()
                return render_template("confirm.html",
                    title="Cannot start",
                    message="You can only start the first Pending job.",
                    confirm_name=None,
                    cancel_url=url_for("edge_station"),
                    post_url=None,
                    hidden_fields={}
                )
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm start",
                    message=f"Start this job on {area.upper()}?",
                    confirm_name="confirm",
                    cancel_url=url_for("edge_station"),
                    post_url=url_for("edge_station"),
                    hidden_fields={"action": "start", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='In progress', started_at=CURRENT_TIMESTAMP
                WHERE id=? AND current_station=? AND status='Pending'
            """, (oid, area))
            d.commit(); d.close(); conn.close()
            log_status(int(oid), area, "in_progress")
            return redirect(url_for("edge_station"))

        if action == "finish":
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm finish",
                    message="Mark this job Finished on Edge?",
                    confirm_name="confirm",
                    cancel_url=url_for("edge_station"),
                    post_url=url_for("edge_station"),
                    hidden_fields={"action": "finish", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='Done', finished_at=CURRENT_TIMESTAMP
                WHERE id=? AND current_station=? AND status='In progress'
            """, (oid, area))
            d.commit(); d.close(); conn.close()
            log_many([(int(oid), area, "done"), (int(oid), TRAMMING2_STATION, "pending")])
            return redirect(url_for("edge_station"))

    conn.close()
    return render_template(
        "station_edge.html",
        station=area.upper(),
        cap=EDGE_CAPACITY.get(area, 0),
        pending=pending,
        inprog=inprog,
        done=done
    )

# ----- Tramming 2 -----
@app.route("/tramming2", methods=["GET","POST"])
@login_required
def tramming2_station():
    if (session.get("area") or "").lower() != TRAMMING2_STATION:
        return ("Forbidden: not Tramming 2", 403)
    conn = get_db(); c = conn.cursor()

    edge_done = {}
    for ed in EDGE_STATIONS:
        c.execute("""
            SELECT id, order_number, datetime(finished_at,'localtime'), datetime(queued_at,'localtime')
            FROM orders
            WHERE current_station=? AND status='Done'
            ORDER BY finished_at ASC NULLS LAST, queued_at ASC
        """, (ed,))
        edge_done[ed] = c.fetchall()

    c.execute("""
        SELECT wrap_slot, id, order_number, status,
               datetime(queued_at,'localtime'), datetime(started_at,'localtime')
        FROM orders
        WHERE current_station='wrapping'
          AND wrap_slot IS NOT NULL
          AND status IN ('Pending','In progress')
        ORDER BY wrap_slot ASC
    """)
    wrap_occ = {r[0]: r[1:] for r in c.fetchall()}
    all_full = len(wrap_occ) >= len(WRAP_SLOTS)

    if request.method == "POST" and request.form.get("action") == "assign_wrap":
        if all_full:
            conn.close()
            return render_template("confirm.html",
                title="Wrapping full",
                message="All Wrapping slots are occupied.",
                confirm_name=None,
                cancel_url=url_for("tramming2_station"),
                post_url=None,
                hidden_fields={}
            )
        order_id = request.form.get("order_id")
        src_edge = request.form.get("src_edge")
        try:
            slot = int(request.form.get("wrap_slot") or 0)
        except ValueError:
            slot = 0
        if slot not in WRAP_SLOTS:
            conn.close()
            return render_template("confirm.html",
                title="Choose slot",
                message="Please choose Wrapping slot 1, 2, or 3.",
                confirm_name=None,
                cancel_url=url_for("tramming2_station"),
                post_url=None,
                hidden_fields={}
            )
        first_id = edge_done.get(src_edge, [None])
        first_id = first_id[0][0] if first_id else None
        if str(order_id) != str(first_id):
            conn.close()
            return render_template("confirm.html",
                title="Cannot move",
                message="Only the first Finished job in each Edge lane can be moved.",
                confirm_name=None,
                cancel_url=url_for("tramming2_station"),
                post_url=None,
                hidden_fields={}
            )
        if slot in wrap_occ:
            free = [str(s) for s in WRAP_SLOTS if s not in wrap_occ]
            msg = f"Wrapping slot {slot} is occupied."
            if free:
                msg += " Available slots: " + ", ".join(free)
            conn.close()
            return render_template("confirm.html",
                title="Slot occupied",
                message=msg,
                confirm_name=None,
                cancel_url=url_for("tramming2_station"),
                post_url=None,
                hidden_fields={}
            )
        if not request.form.get("confirm"):
            conn.close()
            return render_template("confirm.html",
                title="Confirm move",
                message=f"Move this job to WRAPPING, slot {slot}?",
                confirm_name="confirm",
                cancel_url=url_for("tramming2_station"),
                post_url=url_for("tramming2_station"),
                hidden_fields={
                    "action": "assign_wrap",
                    "order_id": order_id,
                    "src_edge": src_edge,
                    "wrap_slot": slot,
                }
            )
        d = get_db(); dc = d.cursor()
        dc.execute("""
            UPDATE orders
            SET current_station='wrapping',
                status='Pending',
                queued_at=CURRENT_TIMESTAMP,
                wrap_slot=?
            WHERE id=? AND current_station=? AND status='Done'
        """, (slot, order_id, src_edge))
        d.commit(); d.close(); conn.close()
        log_many([(int(order_id), TRAMMING2_STATION, "done"), (int(order_id), WRAPPING_STATION, "pending")])
        return redirect(url_for("tramming2_station"))

    conn.close()
    return render_template(
        "station_tramming2.html",
        edge_done=edge_done,
        wrap_occ=wrap_occ,
        wrap_slots=WRAP_SLOTS
    )

# ----- Wrapping -----
@app.route("/wrapping", methods=["GET","POST"])
@login_required
def wrapping_station():
    if (session.get("area") or "").lower() != WRAPPING_STATION:
        return ("Forbidden: not Wrapping", 403)
    conn = get_db(); c = conn.cursor()
    c.execute("""
        SELECT wrap_slot, id, order_number, status,
               datetime(queued_at,'localtime'), datetime(started_at,'localtime')
        FROM orders
        WHERE current_station='wrapping'
          AND wrap_slot IS NOT NULL
          AND status IN ('Pending','In progress')
        ORDER BY wrap_slot ASC
    """)
    wrap_occ = {r[0]: r[1:] for r in c.fetchall()}
    c.execute("""
        SELECT id, order_number, datetime(finished_at,'localtime')
        FROM orders
        WHERE current_station='wrapping'
          AND status='Done'
        ORDER BY finished_at DESC
    """)
    done = c.fetchall()

    if request.method == "POST":
        action = request.form.get("action")
        oid = request.form.get("order_id")
        if action == "start":
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm start",
                    message="Start this job in WRAPPING?",
                    confirm_name="confirm",
                    cancel_url=url_for("wrapping_station"),
                    post_url=url_for("wrapping_station"),
                    hidden_fields={"action": "start", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='In progress', started_at=CURRENT_TIMESTAMP
                WHERE id=? AND current_station='wrapping' AND status='Pending'
            """, (oid,))
            d.commit(); d.close(); conn.close()
            log_status(int(oid), WRAPPING_STATION, "in_progress")
            return redirect(url_for("wrapping_station"))

        if action == "finish":
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm finish",
                    message="Mark this job Finished in WRAPPING?",
                    confirm_name="confirm",
                    cancel_url=url_for("wrapping_station"),
                    post_url=url_for("wrapping_station"),
                    hidden_fields={"action": "finish", "order_id": oid}
                )
            d = get_db(); dc = d.cursor()
            dc.execute("""
                UPDATE orders
                SET status='Done',
                    finished_at=CURRENT_TIMESTAMP,
                    wrap_slot=NULL
                WHERE id=? AND current_station='wrapping' AND status='In progress'
            """, (oid,))
            d.commit(); d.close(); conn.close()
            log_status(int(oid), WRAPPING_STATION, "done")
            return redirect(url_for("wrapping_station"))

    conn.close()
    return render_template(
        "station_wrapping.html",
        wrap_occ=wrap_occ,
        wrap_slots=WRAP_SLOTS,
        done=done
    )

# ----- Loading -----
@app.route("/loading", methods=["GET","POST"])
@login_required
def loading_station():
    if (session.get("area") or "").lower() != LOADING_STATION:
        return ("Forbidden: not Loading", 403)

    lorry1_label, lorry2_label, l1_num, l2_num, next_num = get_lorry_state()

    conn = get_db(); c = conn.cursor()

    c.execute("""
        SELECT id, order_number, datetime(finished_at,'localtime')
        FROM orders
        WHERE current_station='wrapping' AND status='Done'
        ORDER BY finished_at ASC
    """)
    ready = c.fetchall()

    c.execute("""
        SELECT id, order_number, datetime(started_at,'localtime'), lorry, batch_id
        FROM orders
        WHERE current_station='loading' AND status='In progress'
        ORDER BY started_at ASC
    """)
    inprog_all = c.fetchall()
    inprog_l1 = [r for r in inprog_all if (r[3] or "") == lorry1_label]
    inprog_l2 = [r for r in inprog_all if (r[3] or "") == lorry2_label]

    c.execute("""
        SELECT id, order_number, datetime(finished_at,'localtime'), lorry
        FROM orders
        WHERE current_station='loading' AND status='Done'
        ORDER BY finished_at DESC
    """)
    fin_all = c.fetchall()
    fin_l1 = [r for r in fin_all if (r[3] or "") == lorry1_label]
    fin_l2 = [r for r in fin_all if (r[3] or "") == lorry2_label]
    count_l1, count_l2 = len(fin_l1), len(fin_l2)

    if request.method == "POST":
        action = request.form.get("action")

        if action == "start_batch":
            ids = [i for i in (request.form.get("order_ids", "").split(",")) if i]
            lorry_sel = (request.form.get("lorry") or "").strip()
            if lorry_sel not in ("1", "2"):
                conn.close()
                return render_template("confirm.html",
                    title="Choose lorry",
                    message="Please select Lorry 1 or Lorry 2.",
                    confirm_name=None,
                    cancel_url=url_for("loading_station"),
                    post_url=None,
                    hidden_fields={}
                )
            if not ids or len(ids) > 2:
                conn.close()
                return render_template("confirm.html",
                    title="Selection error",
                    message="Select 1 or 2 jobs to start loading.",
                    confirm_name=None,
                    cancel_url=url_for("loading_station"),
                    post_url=None,
                    hidden_fields={}
                )
            label = lorry1_label if lorry_sel == "1" else lorry2_label
            if not request.form.get("confirm"):
                conn.close()
                return render_template("confirm.html",
                    title="Confirm start loading",
                    message=f"Start loading {len(ids)} job(s) to {label}?",
                    confirm_name="confirm",
                    cancel_url=url_for("loading_station"),
                    post_url=url_for("loading_station"),
                    hidden_fields={
                        "action": "start_batch",
                        "lorry": lorry_sel,
                        "order_ids": ",".join(ids),
                    }
                )
            d = get_db(); dc = d.cursor()
            batch = str(uuid.uuid4()) if len(ids) > 1 else None
            for oid in ids:
                dc.execute("""
                    UPDATE orders
                    SET current_station='loading',
                        status='In progress',
                        started_at=CURRENT_TIMESTAMP,
                        lorry=?,
                        batch_id=?
                    WHERE id=? AND current_station='wrapping' AND status='Done'
                """, (label, batch, oid))
            d.commit(); d.close(); conn.close()
            log_many([(int(oid), LOADING_STATION, "in_progress") for oid in ids])
            return redirect(url_for("loading_station"))

        if action == "finish":
            oid = request.form.get("order_id")
            d = get_db(); dc = d.cursor()
            dc.execute("""
                SELECT batch_id
                FROM orders
                WHERE id=? AND current_station='loading' AND status='In progress'
            """, (oid,))
            row = dc.fetchone()
            if not request.form.get("confirm"):
                d.close()
                return render_template("confirm.html",
                    title="Confirm finish",
                    message="Mark this job loaded? If it is part of a pair, both will be marked done.",
                    confirm_name="confirm",
                    cancel_url=url_for("loading_station"),
                    post_url=url_for("loading_station"),
                    hidden_fields={"action": "finish", "order_id": oid}
                )
            if row and row[0]:
                batch_id = row[0]
                dc.execute("""
                    UPDATE orders
                    SET status='Done', finished_at=CURRENT_TIMESTAMP
                    WHERE batch_id=? AND current_station='loading' AND status='In progress'
                """, (batch_id,))
                dc.execute("""
                    SELECT id FROM orders
                    WHERE batch_id=? AND current_station='loading' AND status='Done'
                """, (batch_id,))
                ids_done = [r[0] for r in dc.fetchall()]
                d.commit(); d.close()
                log_many([(int(i), LOADING_STATION, "done") for i in ids_done])
            else:
                dc.execute("""
                    UPDATE orders
                    SET status='Done', finished_at=CURRENT_TIMESTAMP
                    WHERE id=? AND current_station='loading' AND status='In progress'
                """, (oid,))
                d.commit(); d.close()
                log_status(int(oid), LOADING_STATION, "done")
            return redirect(url_for("loading_station"))

        if action in ("complete_lorry1", "complete_lorry2"):
            left_side = action.endswith("1")
            fin_count = count_l1 if left_side else count_l2
            inprog_count = len(inprog_l1) if left_side else len(inprog_l2)
            if fin_count < LORRY_CAPACITY or inprog_count > 0:
                conn.close()
                return render_template("confirm.html",
                    title="Cannot complete",
                    message="Lorry is not fully loaded yet.",
                    confirm_name=None,
                    cancel_url=url_for("loading_station"),
                    post_url=None,
                    hidden_fields={}
                )
            advance_lorry(1 if left_side else 2)
            return redirect(url_for("loading_station"))

    conn.close()
    return render_template(
        "station_loading.html",
        ready=ready,
        inprog_l1=inprog_l1,
        inprog_l2=inprog_l2,
        fin_l1=fin_l1,
        fin_l2=fin_l2,
        lorry1_label=lorry1_label,
        lorry2_label=lorry2_label,
        count_l1=count_l1,
        count_l2=count_l2,
        capacity=LORRY_CAPACITY
    )

# ====== Manager training matrix screen ======
@app.route("/manager/training", methods=["GET", "POST"])
@login_required
def manager_training():
    if (session.get("area") or "").lower() != MANAGER_AREA:
        return ("Forbidden: not Manager", 403)

    if request.method == "POST":
        conn = get_db(); c = conn.cursor()
        for emp_key, emp_label in TRAINING_EMPLOYEES:
            for st_code, st_label in TRAINING_STATIONS:
                field_name = f"status_{emp_key}_{st_code}"
                val = request.form.get(field_name, "not_trained")
                if val not in ("not_trained", "partial", "full"):
                    val = "not_trained"
                c.execute("""
                    INSERT INTO training (employee, station, status)
                    VALUES (?, ?, ?)
                    ON CONFLICT(employee, station)
                    DO UPDATE SET status = excluded.status
                """, (emp_key, st_code, val))
        conn.commit(); conn.close()
        return redirect(url_for("manager_training"))

    mode = request.args.get("mode", "view")
    if mode not in ("view", "edit"):
        mode = "view"

    matrix = get_training_status_matrix()
    return render_template(
        "training_matrix.html",
        employees=TRAINING_EMPLOYEES,
        stations=TRAINING_STATIONS,
        matrix=matrix,
        mode=mode
    )

# ====== Weekly staffing screen ======
@app.route("/manager/weekly_staffing", methods=["GET", "POST"])
@login_required
def weekly_staffing():
    if (session.get("area") or "").lower() != MANAGER_AREA:
        return ("Forbidden: not Manager", 403)

    conn = get_db(); c = conn.cursor()

    if request.method == "POST":
        # Save weekly assignments, one employee per station per weekday
        for station in STAFFING_STATIONS:
            for day_code, day_label in DAYS_OF_WEEK:
                field_name = f"assign_{station}_{day_code}"
                val = (request.form.get(field_name) or "").strip()
                if not val:
                    # Empty selection, remove any existing assignment
                    c.execute(
                        "DELETE FROM staffing_plan WHERE station=? AND day_of_week=?",
                        (station, day_code)
                    )
                else:
                    # Upsert
                    c.execute("""
                        INSERT INTO staffing_plan (station, day_of_week, employee)
                        VALUES (?, ?, ?)
                        ON CONFLICT(station, day_of_week)
                        DO UPDATE SET employee = excluded.employee
                    """, (station, day_code, val))
        conn.commit()
        conn.close()
        return redirect(url_for("weekly_staffing"))

    # GET, load existing plan into dict[(station, day_code)] = employee_key
    c.execute("SELECT station, day_of_week, employee FROM staffing_plan")
    rows = c.fetchall()
    conn.close()

    plan = {}
    for station, day_code, emp in rows:
        plan[(station, day_code)] = emp

    # Training matrix for colour coding
    matrix = get_training_status_matrix()

    return render_template(
        "weekly_staffing.html",
        title="Weekly staffing",
        stations=STAFFING_STATIONS,
        station_labels=STATION_LABELS,
        days=DAYS_OF_WEEK,
        employees=TRAINING_EMPLOYEES,
        training_matrix=matrix,
        plan=plan
    )

# ====== Manager view with search + grouped columns + lorry labels ======
@app.route("/manager")
@login_required
def manager_view():
    if (session.get("area") or "").lower() != MANAGER_AREA:
        return ("Forbidden: not Manager", 403)

    q = (request.args.get("q") or "").strip()

    conn = get_db(); c = conn.cursor()

    # Orders list
    c.execute("SELECT id, order_number FROM orders ORDER BY id DESC")
    orders_rows = c.fetchall()
    order_ids = [r[0] for r in orders_rows]
    id_to_order = {r[0]: r[1] for r in orders_rows}

    # Latest history per (order, station)
    c.execute("""
        WITH latest AS (
            SELECT order_id, station, MAX(changed_at) AS latest_ts
            FROM order_history
            GROUP BY order_id, station
        )
        SELECT h.order_id, h.station, h.status, h.changed_at
        FROM order_history h
        JOIN latest l
          ON l.order_id = h.order_id
         AND l.station = h.station
         AND l.latest_ts = h.changed_at
    """)
    hist_rows = c.fetchall()

    # Build raw map then consolidate CNC and EDGE
    per_order_raw = {}
    for oid, station, status, ts in hist_rows:
        per_order_raw.setdefault(oid, {})[station] = {"status": status, "ts": ts}

    # Also pull a raw lookup of the orders table, including lorry label
    cur = get_db().cursor()
    cur.execute("""
      SELECT id, current_station, status,
             datetime(finished_at,'localtime'),
             datetime(queued_at,'localtime'),
             datetime(started_at,'localtime'),
             lorry
      FROM orders
    """)
    raw = {
        r[0]: {
            "cur": r[1], "st": r[2],
            "fin": r[3], "qts": r[4], "sts": r[5],
            "lr":  r[6]
        } for r in cur.fetchall()
    }
    cur.connection.close()

    def pick_latest(items):
        return max(items, key=lambda x: (x[2] or "")) if items else None

    per_order = {}
    for oid, station_map in per_order_raw.items():
        per_order[oid] = {}
        # simple stations
        for st in [
            PREPARING_STATION,
            TRAMMING1_STATION,
            TRAMMING2_STATION,
            WRAPPING_STATION,
            LOADING_STATION,
        ]:
            if st in station_map:
                per_order[oid][st] = {
                    "status": station_map[st]["status"],
                    "ts": station_map[st]["ts"],
                }
        # CNC consolidated
        cnc_cells = [
            (m, station_map[m]["status"], station_map[m]["ts"])
            for m in AREA_GROUPS["cnc"]
            if m in station_map
        ]
        latest_cnc = pick_latest(cnc_cells)
        if latest_cnc:
            mname, st_status, st_ts = latest_cnc
            per_order[oid]["cnc"] = {
                "status": st_status,
                "ts": st_ts,
                "machine": mname.upper(),
            }
        # EDGE consolidated
        edge_cells = [
            (m, station_map[m]["status"], station_map[m]["ts"])
            for m in AREA_GROUPS["edge"]
            if m in station_map
        ]
        latest_edge = pick_latest(edge_cells)
        if latest_edge:
            mname, st_status, st_ts = latest_edge
            per_order[oid]["edge"] = {
                "status": st_status,
                "ts": st_ts,
                "machine": mname.upper(),
            }
        # Attach lorry name to LOADING cell if present in orders row
        lr = raw.get(oid, {}).get("lr")
        if lr and LOADING_STATION in per_order[oid]:
            per_order[oid][LOADING_STATION]["machine"] = lr.upper()

    # Fallback to orders table for completion detection
    active_ids, completed_ids = [], []
    for oid in order_ids:
        hist_loading = (per_order.get(oid, {}) or {}).get(LOADING_STATION)
        order_row = raw.get(oid, {})
        if hist_loading and hist_loading["status"] == "done":
            completed_ids.append(oid)
        elif order_row.get("cur") == LOADING_STATION and order_row.get("st") == "Done":
            per_order.setdefault(oid, {})[LOADING_STATION] = {
                "status": "done",
                "ts": order_row.get("fin") or "—",
                "machine": (order_row.get("lr") or "").upper() or None,
            }
            completed_ids.append(oid)
        else:
            active_ids.append(oid)

    # KPIs
    l1_label, l2_label, l1n, l2n, nextn = get_lorry_state()
    lorries_completed_total = max(0, nextn - 3)
    c.execute("""
        SELECT COUNT(*)
        FROM orders
        WHERE current_station='loading' AND status='Done'
    """)
    (orders_fully_done,) = c.fetchone()

    per_area_done = {}
    for st in [
        PREPARING_STATION,
        *CNC_STATIONS,
        TRAMMING1_STATION,
        *EDGE_STATIONS,
        TRAMMING2_STATION,
        WRAPPING_STATION,
        LOADING_STATION,
    ]:
        c.execute("""
            SELECT COUNT(*)
            FROM order_history
            WHERE station=? AND status='done'
        """, (st,))
        per_area_done[st] = (c.fetchone() or [0])[0]

    def lorry_progress(label):
        c2 = get_db().cursor()
        c2.execute("""
            SELECT COUNT(*)
            FROM orders
            WHERE current_station='loading'
              AND status='Done'
              AND lorry=?
        """, (label,))
        (cnt_done,) = c2.fetchone()
        c2.connection.close()
        pct = int(min(100, round((cnt_done / LORRY_CAPACITY) * 100))) if LORRY_CAPACITY else 0
        return cnt_done, pct

    l1_done, l1_pct = lorry_progress(l1_label)
    l2_done, l2_pct = lorry_progress(l2_label)

    # Search handling
    search_mode = bool(q)
    search_found = False
    search_oid = None
    search_onum = None
    search_current_line = None

    def human_group(station_code):
        s = (station_code or "").lower()
        if s in AREA_GROUPS.get("cnc", []):
            return "CNC", f" ({s.upper()})"
        if s in AREA_GROUPS.get("edge", []):
            return "EDGE", f" ({s.upper()})"
        return (s.upper() if s else "UNKNOWN"), ""

    if search_mode:
        c.execute("""
            SELECT id, order_number, current_station, status,
                   datetime(queued_at,'localtime'),
                   datetime(started_at,'localtime'),
                   datetime(finished_at,'localtime'),
                   lorry
            FROM orders
            WHERE lower(order_number) = lower(?)
            LIMIT 1
        """, (q,))
        row = c.fetchone()
        if row:
            search_found = True
            (
                search_oid,
                search_onum,
                cur_st,
                cur_status,
                qts,
                sts,
                fts,
                lorry_lbl,
            ) = row

            label, machine_hint = human_group(cur_st)
            lorry_hint = ""
            if cur_st == LOADING_STATION and lorry_lbl:
                lorry_hint = f" ({lorry_lbl})"

            if cur_st == LOADING_STATION and cur_status == "Done":
                when = fts or "—"
                search_current_line = f"Current location: COMPLETED{lorry_hint} at {when}"
            elif cur_st == LOADING_STATION and cur_status in ("In progress", "Pending", "Done"):
                when = (
                    sts
                    if cur_status == "In progress"
                    else (qts if cur_status == "Pending" else fts)
                ) or "—"
                search_current_line = (
                    f"Current location: LOADING{lorry_hint} "
                    f"({cur_status.lower()}) since {when}"
                )
            elif cur_status == "In progress":
                when = sts or "—"
                search_current_line = (
                    f"Current location: {label}{machine_hint} "
                    f"(in progress) since {when}"
                )
            elif cur_status == "Pending":
                when = qts or "—"
                search_current_line = (
                    f"Current location: {label}{machine_hint} "
                    f"(pending) since {when}"
                )
            elif cur_status == "Done":
                when = fts or "—"
                search_current_line = (
                    f"Current location: {label}{machine_hint} "
                    f"(done) at {when}"
                )
            else:
                cells = per_order.get(search_oid, {})
                latest = None
                for st_code, cell in cells.items():
                    if (not latest) or (cell["ts"] > latest["ts"]):
                        latest = {
                            "st": st_code,
                            "status": cell["status"],
                            "ts": cell["ts"],
                        }
                if latest:
                    lab, mh = human_group(latest["st"])
                    search_current_line = (
                        f"Current location: {lab}{mh} "
                        f"({latest['status'].replace('_',' ')}) since {latest['ts']}"
                    )
                else:
                    search_current_line = "Current location: —"

            # Ensure minimal cell present for current station if history is missing
            if search_oid not in per_order:
                per_order[search_oid] = {}
            if cur_st:
                key = (
                    "cnc"
                    if cur_st in AREA_GROUPS["cnc"]
                    else ("edge" if cur_st in AREA_GROUPS["edge"] else cur_st)
                )
                ts_guess = (
                    qts
                    if cur_status == "Pending"
                    else (sts if cur_status == "In progress" else (fts or "—"))
                )
                per_order[search_oid][key] = {
                    "status": cur_status.replace(" ", "_").lower(),
                    "ts": ts_guess,
                    "machine": (
                        cur_st.upper()
                        if key in ("cnc", "edge")
                        else (lorry_lbl or "").upper()
                        if key == LOADING_STATION
                        else None
                    ),
                }

    conn.close()

    active = [(oid, id_to_order.get(oid, f"#{oid}")) for oid in active_ids]
    completed = [(oid, id_to_order.get(oid, f"#{oid}")) for oid in completed_ids]

    return render_template(
        "manager.html",
        stations=MANAGER_STATIONS,
        per_order=per_order,
        active_orders=active,
        completed_orders=completed,
        lorry1_label=l1_label,
        lorry2_label=l2_label,
        l1_done=l1_done,
        l2_done=l2_done,
        l1_pct=l1_pct,
        l2_pct=l2_pct,
        lorry_capacity=LORRY_CAPACITY,
        lorries_completed_total=lorries_completed_total,
        orders_fully_done=orders_fully_done,
        per_area_done=per_area_done,
        q=q,
        search_mode=search_mode,
        search_found=search_found,
        search_oid=search_oid,
        search_onum=search_onum,
        search_current_line=search_current_line,
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)