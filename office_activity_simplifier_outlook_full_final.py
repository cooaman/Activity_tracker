# office_activity_simplifier_outlook_full_final_windows.py
import sys
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, json, os, csv
from datetime import datetime, date, timedelta

# Optional Windows-specific/outlook imports
try:
    import win32com.client
    HAS_OUTLOOK = True
except Exception:
    HAS_OUTLOOK = False

# Safe import for win10toast: don't allow import-time pkg_errors to kill the app.
try:
    import importlib
    _win10toast = importlib.import_module("win10toast")
    def _create_toaster():
        try:
            return _win10toast.ToastNotifier()
        except Exception:
            return None
    toaster = _create_toaster()
    HAS_NOTIFY = toaster is not None
except Exception:
    toaster = None
    HAS_NOTIFY = False

try:
    from tkhtmlview import HTMLLabel
    HAS_HTML = True
except Exception:
    HAS_HTML = False

# optional calendar widget
try:
    from tkcalendar import DateEntry
    HAS_DATEENTRY = True
except Exception:
    DateEntry = None
    HAS_DATEENTRY = False

# optional modern theme: ttkbootstrap
try:
    import ttkbootstrap as tb
    HAS_BOOTSTRAP = True
except Exception:
    tb = None
    HAS_BOOTSTRAP = False

DB_FILE = "office_tasks.db"
SETTINGS_FILE = "settings.json"

PRIORITIES = ["Low", "Medium", "High"]
STATUSES = ["Pending", "In-Progress", "Done"]


def _now_iso():
    return datetime.now().isoformat(timespec="seconds")


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"outlook_refresh_minutes": 30, "show_description": False}


def save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f)


def _safe_show_toast(title, msg, duration=5):
    """
    Show a Windows toast if available. Catch and swallow all exceptions
    so toast library errors don't crash the app's callbacks.
    """
    global toaster, HAS_NOTIFY
    if not HAS_NOTIFY or toaster is None:
        return
    try:
        toaster.show_toast(title, (msg or "")[:200], duration=duration, threaded=True)
    except Exception as e:
        print("Toast error (ignored):", e)


# ---------- Simple Tooltip helper ----------
class _ToolTip:
    """Simple tooltip for widgets (small, cross-platform)."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwin = None
        widget.bind("<Enter>", self._enter)
        widget.bind("<Leave>", self._leave)

    def _enter(self, _):
        if self.tipwin:
            return
        x, y, cx, cy = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 20
        self.tipwin = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(tw, text=self.text, justify=tk.LEFT, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 9))
        lbl.pack(ipadx=4, ipady=2)

    def _leave(self, _):
        if self.tipwin:
            try:
                self.tipwin.destroy()
            except Exception:
                pass
            self.tipwin = None


# -------------------- Database --------------------
class TaskDB:
    def __init__(self, path=DB_FILE):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def _init_db(self):
        cur = self.conn.cursor()
        cur.execute(
            """CREATE TABLE IF NOT EXISTS tasks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                description TEXT,
                due_date TEXT,
                priority TEXT DEFAULT 'Medium',
                status TEXT DEFAULT 'Pending',
                created_at TEXT,
                updated_at TEXT,
                done_at TEXT,
                outlook_id TEXT,
                progress_log TEXT,
                attachments TEXT,
                reminder_minutes INTEGER,
                reminder_set_at TEXT,
                reminder_sent_at TEXT
            );"""
        )

        # Ensure schema migrations (if db already exists but lacks these columns)
        for col in ["outlook_id", "progress_log", "attachments", "reminder_minutes", "reminder_set_at", "reminder_sent_at"]:
            try:
                cur.execute(f"ALTER TABLE tasks ADD COLUMN {col} TEXT;")
            except sqlite3.OperationalError:
                pass

        self.conn.commit()

    def add(self, title, description, due_date, priority, status="Pending", outlook_id=None, reminder_minutes=None, reminder_set_at=None):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                """INSERT INTO tasks(title, description, due_date, priority, status,
                   created_at, updated_at, done_at, outlook_id, progress_log, reminder_minutes, reminder_set_at, reminder_sent_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (title, description, due_date, priority, status, now, now, done_at, outlook_id, "", reminder_minutes, reminder_set_at, None),
            )

    def update(self, task_id, title, description, due_date, priority, status, reminder_minutes=None, reminder_set_at=None):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            if reminder_minutes is None and reminder_set_at is None:
                self.conn.execute(
                    """UPDATE tasks SET title=?, description=?, due_date=?, priority=?, 
                       status=?, updated_at=?, done_at=? WHERE id=?""",
                    (title, description, due_date, priority, status, now, done_at, task_id),
                )
            else:
                self.conn.execute(
                    """UPDATE tasks SET title=?, description=?, due_date=?, priority=?, 
                       status=?, updated_at=?, done_at=?, reminder_minutes=?, reminder_set_at=? WHERE id=?""",
                    (title, description, due_date, priority, status, now, done_at, reminder_minutes, reminder_set_at, task_id),
                )

    def update_progress(self, task_id, progress_log):
        now = _now_iso()
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET progress_log=?, updated_at=? WHERE id=?",
                (progress_log, now, task_id),
            )

    def delete(self, task_id):
        with self.conn:
            self.conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))

    def mark_done(self, task_id):
        now = _now_iso()
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET status='Done', updated_at=?, done_at=? WHERE id=?",
                (now, now, task_id),
            )

    def fetch(self):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM tasks ORDER BY due_date IS NULL, due_date ASC, priority DESC")
        return cur.fetchall()

    def fetch_by_status(self, status):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE status=? ORDER BY priority DESC, due_date ASC", (status,))
        return cur.fetchall()

    def fetch_due_today(self):
        today = date.today().isoformat()
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE status!='Done' AND due_date=? ORDER BY priority DESC", (today,))
        return cur.fetchall()

    def fetch_overdue(self):
        today = date.today().isoformat()
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE status!='Done' AND due_date IS NOT NULL AND due_date < ?", (today,))
        return cur.fetchall()

    def bulk_add(self, rows):
        now = _now_iso()
        with self.conn:
            for r in rows:
                done_at = now if r.get("status") == "Done" else None
                self.conn.execute(
                    """INSERT INTO tasks(title, description, due_date, priority, status,
                                        created_at, updated_at, done_at, outlook_id, progress_log, reminder_minutes, reminder_set_at, reminder_sent_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        r.get("title"),
                        r.get("description"),
                        r.get("due_date"),
                        r.get("priority", "Medium"),
                        r.get("status", "Pending"),
                        now,
                        now,
                        done_at,
                        r.get("outlook_id"),
                        r.get("progress_log", ""),
                        None,
                        None,
                        None,
                    ),
                )


# -------------------- App --------------------
# Choose base window depending on presence of ttkbootstrap
if HAS_BOOTSTRAP:
    BaseWindow = tb.Window
else:
    BaseWindow = tk.Tk


class TaskApp(BaseWindow):
    def _init_styles(self):
        """Initialize theme, fonts and button styles (Option 1: custom default button)."""
        style = ttk.Style()

        # Prefer a theme that allows predictable styling across platforms
        try:
            style.theme_use("clam")
        except Exception:
            # If clam not available, fall back to default
            try:
                style.theme_use(style.theme_names()[0])
            except Exception:
                pass

        # Fonts
        default_font = ("Segoe UI", 10) if os.name == "nt" else ("Helvetica", 10)
        heading_font = ("Segoe UI", 11, "bold") if os.name == "nt" else ("Helvetica", 11, "bold")

        # Base UI font and heading font
        try:
            style.configure(".", font=default_font)
            style.configure("Treeview.Heading", font=heading_font)
        except Exception:
            pass

        # Define colors for themes (you already had these; keep them)
        self._themes = {
            "Light": {
                "bg": "#f7f7f7",
                "panel": "#ffffff",
                "kanban_bg": "#f0f0f0",
                "text": "#222222",
                "muted": "#666666"
            },
            "Dark": {
                "bg": "#2b2b2b",
                "panel": "#333333",
                "kanban_bg": "#3a3a3a",
                "text": "#ffffff",
                "muted": "#cccccc"
            }
        }
        self._current_theme = getattr(self, "_current_theme", "Light")

        # -- Button styling: override TButton so all ttk.Buttons pick this up --
        # This produces a neutral, flat, light-gray button with subtle hover feedback.
        try:
            # Primary look for normal buttons
            style.configure(
                "TButton",
                padding=6,
                relief="flat",
                background="#f0f0f0",   # normal background (light)
                foreground="#000000",
                borderwidth=0,
            )

            # Hover / active / pressed mappings
            style.map(
                "TButton",
                background=[
                    ("active", "#e6e6e6"),
                    ("pressed", "#dcdcdc"),
                    ("disabled", "#f0f0f0")
                ],
                foreground=[
                    ("disabled", "#9a9a9a")
                ]
            )

            # Also create a Custom.TButton alias for explicit use if needed
            style.configure(
                "Custom.TButton",
                padding=6,
                relief="flat",
                background="#f0f0f0",
                foreground="#000000",
                borderwidth=0,
            )
            style.map(
                "Custom.TButton",
                background=[("active", "#e6e6e6"), ("pressed", "#dcdcdc")],
                foreground=[("disabled", "#9a9a9a")]
            )
        except Exception as e:
            # If button styling fails for any reason, don't crash — just continue.
            print("Button style init warning:", e)

        # Tweak label and tree fonts/padding a little
        try:
            style.configure("TLabel", padding=2)
            style.configure("Treeview", rowheight=22)
        except Exception:
            pass

        # Tag colors for Treeview rows (used by _set_theme/_populate)
        try:
            if getattr(self, "_current_theme", "Light") == "Dark":
                style.configure("Treeview", background="#2f2f2f", fieldbackground="#2f2f2f", foreground="#ffffff")
            else:
                style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#000000")
        except Exception:
            pass

        # Configure tree tags (used in _populate)
        try:
            # light theme tags
            style.configure("Treeview.priority_high", background="#FFD6D6")
            style.configure("Treeview.priority_medium", background="#FFF5CC")
            style.configure("Treeview.priority_low", background="#E6FFEA")
        except Exception:
            # fallback — some ttk backends ignore widget-specific tag styles; you'll still get tag_configure() calls in _set_theme/_populate
            pass

    def _set_theme(self, theme_name):
        if theme_name not in self._themes:
            return
        self._current_theme = theme_name
        palette = self._themes[theme_name]
        try:
            self.configure(bg=palette["bg"])
        except Exception:
            pass

        # adjust tree tags for readability
        if getattr(self, "tree", None):
            try:
                if theme_name == "Dark":
                    self.tree.tag_configure("priority_high", background="#4a1a1a")
                    self.tree.tag_configure("priority_medium", background="#4a3b1a")
                    self.tree.tag_configure("priority_low", background="#1a4a2a")
                    self.tree.tag_configure("oddrow", background="#2f2f2f")
                    self.tree.tag_configure("evenrow", background="#353535")
                else:
                    self.tree.tag_configure("priority_high", background="#FFD6D6")
                    self.tree.tag_configure("priority_medium", background="#FFF5CC")
                    self.tree.tag_configure("priority_low", background="#E6FFEA")
                    self.tree.tag_configure("oddrow", background="#FFFFFF")
                    self.tree.tag_configure("evenrow", background="#F6F6F6")
            except Exception:
                pass
        self._populate()

    def _format_timedelta(self, td):
        total_seconds = int(td.total_seconds())
        if total_seconds <= 0:
            return "Now"
        hours, rem = divmod(total_seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        if hours > 0:
            return f"{hours}h {minutes}m"
        if minutes >= 1:
            return f"{minutes}m {seconds}s"
        return f"{seconds}s"

    def _refresh_reminder_display(self):
        """Update the 'reminder' cell in the Task List Treeview with live countdowns."""
        try:
            cur = self.db.conn.cursor()
            for iid in self.tree.get_children():
                vals = self.tree.item(iid, "values")
                if not vals:
                    continue
                try:
                    task_id = int(vals[0])
                except Exception:
                    continue

                cur.execute("SELECT reminder_minutes, reminder_set_at, reminder_sent_at FROM tasks WHERE id=?", (task_id,))
                row = cur.fetchone()
                display = "—"
                if row:
                    rm = row["reminder_minutes"]
                    set_at = row["reminder_set_at"]
                    sent_at = row["reminder_sent_at"]
                    if rm is None or rm == "" or set_at is None:
                        display = "—"
                    else:
                        try:
                            rm_int = int(rm)
                        except Exception:
                            display = "—"
                            self.tree.set(iid, "reminder", display)
                            continue
                        try:
                            set_dt = datetime.fromisoformat(set_at)
                        except Exception:
                            display = "—"
                            self.tree.set(iid, "reminder", display)
                            continue

                        target = set_dt + timedelta(minutes=rm_int)
                        if sent_at:
                            try:
                                sent_dt = datetime.fromisoformat(sent_at)
                                if sent_dt >= target:
                                    display = "Sent"
                                else:
                                    remaining = target - datetime.now()
                                    display = self._format_timedelta(remaining)
                            except Exception:
                                remaining = target - datetime.now()
                                display = self._format_timedelta(remaining)
                        else:
                            remaining = target - datetime.now()
                            display = self._format_timedelta(remaining)
                try:
                    self.tree.set(iid, "reminder", display)
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self.after(1000, self._refresh_reminder_display)

    # --------------- Reminder helpers ----------------
    def _schedule_task_reminder_checker(self):
        """
        Schedule the periodic reminder check (every few seconds).
        Keeps rescheduling even if an error occurs so a single failure won't stop reminders.
        """
        try:
            self._check_task_reminders()
        except Exception as e:
            # log but keep loop alive
            print("Unhandled error during reminder check (caught):", e)
        finally:
            try:
                # 5-second granularity for stopwatch-style reminders
                self.after(5 * 1000, self._schedule_task_reminder_checker)
            except Exception as e:
                print("Failed to schedule next reminder check:", e)

    def _check_task_reminders(self):
        """
        Query DB for tasks whose stopwatch-style reminder target has passed.
        IMPORTANT: Do NOT mark 'reminder_sent_at' here. Marking must happen on user action (Dismiss),
        otherwise snooze won't work correctly (user snoozes but record already marked sent).
        """
        try:
            cur = self.db.conn.cursor()
            cur.execute("""
                SELECT id, title, description, reminder_minutes, reminder_set_at, reminder_sent_at
                FROM tasks
                WHERE reminder_minutes IS NOT NULL AND reminder_minutes != '' AND reminder_set_at IS NOT NULL
                AND status != 'Done'
            """)
            rows = cur.fetchall()
            now_dt = datetime.now()

            for r in rows:
                # parse minutes and set time
                try:
                    rm_min = int(r["reminder_minutes"])
                except Exception:
                    continue
                try:
                    set_at = datetime.fromisoformat(r["reminder_set_at"])
                except Exception:
                    continue

                target = set_at + timedelta(minutes=rm_min)

                # if we have reminder_sent_at, parse it
                sent = None
                if r["reminder_sent_at"]:
                    try:
                        sent = datetime.fromisoformat(r["reminder_sent_at"])
                    except Exception:
                        sent = None

                # Only show if current time >= target and we have NOT already recorded a send at/after target
                if now_dt >= target and (sent is None or sent < target):
                    # Use after(0, ...) to ensure popup is created on the Tk event loop (UI thread)
                    try:
                        self.after(0, lambda _id=r["id"], _t=r["title"], _d=r["description"]: self._show_reminder_popup(_id, _t, _d))
                    except Exception as e:
                        print("Failed to schedule reminder popup on UI thread:", e)

                    # NOTE: do not update reminder_sent_at here — only update when the user dismisses the popup.
        except Exception as e:
            print("Reminder check error:", e)

    def _show_reminder_popup(self, task_id, title, description):
        """
        Improved cross-platform reminder popup:
        - Avoid modal grab_set() which can freeze UI on some platforms.
        - Bring window to front robustly (lift + focus_force + toggle -topmost).
        - Mark reminder_sent_at only when user dismisses, not when popup is shown.
        """

        # Windows toast (safe wrapper) — non-blocking
        if HAS_NOTIFY:
            _safe_show_toast(f"Reminder: {title}", description or "Task due soon")

        win = tk.Toplevel(self)
        win.title("🔔 Task Reminder")
        # slightly larger so buttons don't wrap
        win.geometry("700x300")
        win.transient(self)

        # Do not call grab_set() — keep popup non-modal so UI remains responsive.
        try:
            # try to raise and focus the popup in a way that works across platforms
            win.lift()
            win.attributes("-topmost", True)
            win.update()
            # small toggle may help macOS bring to front
            try:
                win.attributes("-topmost", False)
            except Exception:
                pass
            win.focus_force()
        except Exception:
            pass

        # ensure closing via window manager triggers the same cleanup as Dismiss
        def _on_close():
            # treat closing as dismiss: mark reminder_sent_at
            now_iso = datetime.now().isoformat(timespec="seconds")
            try:
                self.db.conn.execute("UPDATE tasks SET reminder_sent_at=? WHERE id=?", (now_iso, task_id))
                self.db.conn.commit()
            except Exception as e:
                print("Error marking reminder_sent_at on close:", e)
            try:
                win.destroy()
            except Exception:
                pass

        win.protocol("WM_DELETE_WINDOW", _on_close)

        header = ttk.Label(win, text=title, font=("", 14, "bold"))
        header.pack(padx=12, pady=(12, 6), anchor="w")

        txt = tk.Text(win, height=6, wrap="word", padx=8, pady=4)
        txt.insert("1.0", description or "(no description)")
        txt.config(state="disabled")
        txt.pack(fill=tk.BOTH, expand=False, padx=12, pady=(0,8))

        btnf = ttk.Frame(win)
        btnf.pack(fill=tk.X, padx=12, pady=8)

        def open_task():
            try:
                win.destroy()
            except Exception:
                pass
            try:
                self._open_edit_window(task_id)
            except Exception as e:
                print("Error opening task editor from reminder popup:", e)

        def _snooze(minutes):
            # When snoozing, restart countdown from now and clear reminder_sent_at so it can fire again.
            new_set = datetime.now().isoformat(timespec="seconds")
            try:
                # store integer minutes; reminder_sent_at cleared (NULL)
                self.db.conn.execute(
                    "UPDATE tasks SET reminder_minutes=?, reminder_set_at=?, reminder_sent_at=NULL WHERE id=?",
                    (int(minutes), new_set, task_id)
                )
                self.db.conn.commit()
            except Exception as e:
                print("Snooze update error:", e)
            try:
                win.destroy()
            except Exception:
                pass

        def dismiss():
            # Mark as reminded now (so it won't re-fire)
            now_iso = datetime.now().isoformat(timespec="seconds")
            try:
                self.db.conn.execute("UPDATE tasks SET reminder_sent_at=? WHERE id=?", (now_iso, task_id))
                self.db.conn.commit()
            except Exception as e:
                print("Dismiss update error:", e)
            try:
                win.destroy()
            except Exception:
                pass

        # Buttons laid out to avoid wrapping
        ttk.Button(btnf, text="Open Task", command=open_task).pack(side=tk.LEFT, padx=6)
        ttk.Button(btnf, text="Snooze 5m", command=lambda: _snooze(5)).pack(side=tk.LEFT, padx=6)
        ttk.Button(btnf, text="Snooze 10m", command=lambda: _snooze(10)).pack(side=tk.LEFT, padx=6)
        ttk.Button(btnf, text="Snooze 30m", command=lambda: _snooze(30)).pack(side=tk.LEFT, padx=6)
        ttk.Button(btnf, text="Dismiss", command=dismiss).pack(side=tk.RIGHT, padx=6)

        # final lift/focus to ensure visible
        try:
            win.lift()
            win.focus_force()
        except Exception:
            pass

    # -------------------- UI / CRUD / Kanban --------------------
    def _save_inline_from_form(self):
        sel = self.tree.selection()
        if sel:
            try:
                task_id = int(self.tree.item(sel[0], "values")[0])
            except Exception:
                messagebox.showwarning("Save", "Could not determine selected task id.")
                return
            self._open_edit_window(task_id)
        else:
            self._open_edit_window(None)

    def _on_task_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            task_id = int(self.tree.item(sel[0], "values")[0])
        except Exception:
            return
        self._open_edit_window(task_id)

    def _on_kanban_double_click(self, event):
        lb = event.widget
        idx = lb.nearest(event.y)
        if idx < 0:
            return
        status = lb.status_name
        try:
            task_id = self.kanban_item_map[status][idx]
        except Exception:
            return
        self._open_edit_window(task_id)

    def _open_edit_window(self, task_id=None):
        # local DateEntry import (if available)
        try:
            from tkcalendar import DateEntry
            has_dateentry = True
        except Exception:
            DateEntry = None
            has_dateentry = False

        win = tk.Toplevel(self)
        win.transient(self)
        try:
            win.grab_set()
        except Exception:
            pass
        win.title("Edit Task" if task_id else "Add Task")
        win.geometry("680x560")

        title_var = tk.StringVar()
        due_var = tk.StringVar()
        priority_var = tk.StringVar(value="Medium")
        status_var = tk.StringVar(value="Pending")
        reminder_var = tk.StringVar(value="")

        staged_attachments = []
        existing_attachments = []

        frm = ttk.Frame(win, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Title *").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=title_var, width=50).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Due Date (YYYY-MM-DD)").grid(row=1, column=0, sticky="w")
        if has_dateentry:
            due_widget = DateEntry(frm, date_pattern="yyyy-mm-dd", textvariable=due_var, width=20)
            due_widget.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        else:
            due_widget = ttk.Entry(frm, textvariable=due_var, width=20)
            due_widget.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Priority").grid(row=2, column=0, sticky="w")
        ttk.Combobox(frm, textvariable=priority_var, values=PRIORITIES, state="readonly", width=12).grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Status").grid(row=3, column=0, sticky="w")
        ttk.Combobox(frm, textvariable=status_var, values=STATUSES, state="readonly", width=12).grid(row=3, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Reminder (minutes before due)").grid(row=4, column=0, sticky="w")
        reminder_choices = ["", "5", "10", "30", "60", "120", "1440"]
        reminder_cb = ttk.Combobox(frm, textvariable=reminder_var, values=reminder_choices, width=18)
        reminder_cb.grid(row=4, column=1, sticky="w", padx=6, pady=4)
        reminder_cb.set("")

        ttk.Label(frm, text="Description").grid(row=5, column=0, sticky="nw", pady=(6,0))
        desc_text = tk.Text(frm, height=10, width=60, wrap="word")
        desc_text.grid(row=5, column=1, sticky="we", padx=6, pady=(6,0))

        if task_id:
            cur = self.db.conn.cursor()
            cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
            r = cur.fetchone()
            if r:
                title_var.set(r["title"])
                due_var.set(r["due_date"] or "")
                priority_var.set(r["priority"] or "Medium")
                status_var.set(r["status"] or "Pending")
                desc_text.insert(tk.END, r["description"] or "")
                if "attachments" in r.keys() and r["attachments"]:
                    try:
                        existing_attachments = json.loads(r["attachments"])
                    except Exception:
                        existing_attachments = []
                try:
                    rm = r["reminder_minutes"]
                    if rm not in (None, "", "None"):
                        reminder_var.set(str(rm))
                    else:
                        reminder_var.set("")
                except Exception:
                    reminder_var.set("")

        ttk.Label(frm, text="Attachments").grid(row=6, column=0, sticky="nw", pady=(10,0))
        attachments_frame = ttk.Frame(frm)
        attachments_frame.grid(row=6, column=1, sticky="we", padx=6, pady=(10,0))
        attachments_list_var = tk.StringVar(value=", ".join(os.path.basename(p) for p in existing_attachments))
        attachments_label = ttk.Label(attachments_frame, textvariable=attachments_list_var, wraplength=500)
        attachments_label.pack(anchor="w", fill=tk.X)

        def add_file_to_attachments():
            path = filedialog.askopenfilename(parent=win)
            if not path:
                return
            os.makedirs("attachments", exist_ok=True)
            fname = os.path.basename(path)
            dest = os.path.join("attachments", fname)
            if os.path.exists(dest):
                base, ext = os.path.splitext(fname)
                dest = os.path.join("attachments", f"{base}_{int(datetime.now().timestamp())}{ext}")
            with open(path, "rb") as fsrc, open(dest, "wb") as fdst:
                fdst.write(fsrc.read())

            if task_id:
                cur = self.db.conn.cursor()
                cur.execute("SELECT attachments FROM tasks WHERE id=?", (task_id,))
                row = cur.fetchone()
                files = []
                if row and row["attachments"]:
                    try:
                        files = json.loads(row["attachments"])
                    except Exception:
                        files = []
                files.append(dest)
                self.db.conn.execute("UPDATE tasks SET attachments=? WHERE id=?", (json.dumps(files), task_id))
                self.db.conn.commit()
                attachments_list_var.set(", ".join(os.path.basename(p) for p in files))
            else:
                staged_attachments.append(dest)
                attachments_list_var.set(", ".join(os.path.basename(p) for p in staged_attachments))

        def open_attachments():
            files = list(existing_attachments) + list(staged_attachments)
            if not files:
                messagebox.showinfo("Attachments", "No attachments to open.", parent=win)
                return
            for f in files:
                try:
                    if os.name == "nt":
                        os.startfile(f)
                    elif sys.platform == "darwin":
                        os.system(f"open '{f}'")
                    else:
                        os.system(f"xdg-open '{f}'")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open {f}: {e}", parent=win)

        btns_attach = ttk.Frame(attachments_frame)
        btns_attach.pack(anchor="w", pady=(6,0))
        ttk.Button(btns_attach, text="Add File", command=add_file_to_attachments).pack(side=tk.LEFT)
        ttk.Button(btns_attach, text="Open", command=open_attachments).pack(side=tk.LEFT, padx=6)

        def save():
            title = title_var.get().strip()
            if not title:
                messagebox.showwarning("Validation", "Title is required", parent=win)
                return
            due = due_var.get().strip()
            if due:
                try:
                    datetime.strptime(due, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Validation", "Date must be YYYY-MM-DD", parent=win)
                    return

            desc = desc_text.get("1.0", tk.END).strip()
            reminder_value = reminder_var.get().strip() or None

            if reminder_value not in (None, "", "None"):
                try:
                    reminder_minutes_int = int(reminder_value)
                except Exception:
                    messagebox.showwarning("Validation", "Reminder must be an integer number of minutes or blank.", parent=win)
                    return
                reminder_set_at_iso = datetime.now().isoformat(timespec="seconds")
            else:
                reminder_minutes_int = None
                reminder_set_at_iso = None

            if task_id:
                self.db.update(task_id, title, desc, due or None, priority_var.get(), status_var.get(),
                               reminder_minutes=reminder_minutes_int, reminder_set_at=reminder_set_at_iso)
            else:
                self.db.add(title, desc, due or None, priority_var.get(), status_var.get(),
                            reminder_minutes=reminder_minutes_int, reminder_set_at=reminder_set_at_iso)
                cur = self.db.conn.cursor()
                cur.execute("SELECT last_insert_rowid() as id")
                new_id = cur.fetchone()["id"]
                if staged_attachments:
                    try:
                        self.db.conn.execute("UPDATE tasks SET attachments=? WHERE id=?", (json.dumps(staged_attachments), new_id))
                    except Exception:
                        pass

            self._populate()
            self._populate_kanban()
            win.destroy()

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=12)
        ttk.Button(btn_frame, text="Save", command=save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="Cancel", command=win.destroy).pack(side=tk.LEFT, padx=6)

    def _open_selected_kanban_attachments(self):
        if not self.kanban_selected_id:
            messagebox.showwarning("No Task", "Please select a task first.")
            return

        cur = self.db.conn.cursor()
        cur.execute("SELECT attachments FROM tasks WHERE id=?", (self.kanban_selected_id,))
        row = cur.fetchone()
        if not row or not row["attachments"]:
            messagebox.showinfo("No Attachments", "No attachments found for this task.")
            return

        files = json.loads(row["attachments"])
        for f in files:
            try:
                if os.name == "nt":
                    os.startfile(f)
                elif sys.platform == "darwin":
                    os.system(f"open '{f}'")
                else:
                    os.system(f"xdg-open '{f}'")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open {f}: {e}")

    def _add_attachment(self):
        path = filedialog.askopenfilename()
        if not path:
            return
        os.makedirs("attachments", exist_ok=True)
        fname = os.path.basename(path)
        dest = os.path.join("attachments", fname)

        if os.path.exists(dest):
            base, ext = os.path.splitext(fname)
            dest = os.path.join("attachments", f"{base}_{int(datetime.now().timestamp())}{ext}")

        with open(path, "rb") as fsrc, open(dest, "wb") as fdst:
            fdst.write(fsrc.read())

        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Task", "Select a task first.")
            return
        task_id = int(self.tree.item(sel[0], "values")[0])

        cur = self.db.conn.cursor()
        cur.execute("SELECT attachments FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        files = []
        if row and row["attachments"]:
            try:
                files = json.loads(row["attachments"])
            except Exception:
                files = []
        files.append(dest)
        self.db.conn.execute("UPDATE tasks SET attachments=? WHERE id=?", (json.dumps(files), task_id))
        self.db.conn.commit()

        self.attachments_var.set(", ".join(os.path.basename(f) for f in files))
        messagebox.showinfo("Attachment", f"File {os.path.basename(dest)} added.")

    def _open_attachment(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Task", "Select a task first.")
            return
        task_id = int(self.tree.item(sel[0], "values")[0])
        cur = self.db.conn.cursor()
        cur.execute("SELECT attachments FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        if not row or not row["attachments"]:
            messagebox.showinfo("No Attachments", "No attachments found for this task.")
            return
        files = json.loads(row["attachments"])
        for f in files:
            if os.name == "nt":
                os.startfile(f)
            elif sys.platform == "darwin":
                os.system(f"open '{f}'")
            else:
                os.system(f"xdg-open '{f}'")

    def _on_kanban_drag_start(self, event):
        lb = event.widget
        idx = lb.nearest(event.y)
        if idx >= 0:
            self.drag_data = {"listbox": lb, "index": idx, "task_line": lb.get(idx)}

    def _on_kanban_drag_motion(self, event):
        lb = event.widget
        lb.selection_clear(0, tk.END)
        lb.selection_set(lb.nearest(event.y))

    def _on_kanban_drag_drop(self, event):
        if not hasattr(self, "drag_data"):
            return
        src_lb = self.drag_data["listbox"]
        src_status = src_lb.status_name
        src_index = self.drag_data["index"]

        try:
            task_id = self.kanban_item_map[src_status][src_index]
        except Exception:
            self.drag_data = None
            return

        src_lb.delete(src_index)
        try:
            del self.kanban_item_map[src_status][src_index]
        except Exception:
            pass

        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        target_lb = widget if isinstance(widget, tk.Listbox) else None

        if target_lb and hasattr(target_lb, "status_name"):
            target_status = target_lb.status_name
            idx = target_lb.size()
            cur = self.db.conn.cursor()
            cur.execute("SELECT title FROM tasks WHERE id=?", (task_id,))
            row = cur.fetchone()
            display = row["title"] if row else "Untitled"
            target_lb.insert(tk.END, display)
            self.kanban_item_map[target_status].append(task_id)
            self._move_task(task_id, target_status)
        else:
            cur = self.db.conn.cursor()
            row = cur.execute("SELECT title FROM tasks WHERE id=?", (task_id,)).fetchone()
            title = row["title"] if row else "Untitled"
            src_lb.insert(tk.END, title)
            self.kanban_item_map[src_status].append(task_id)

        self.drag_data = None

    def __init__(self):
        # if using ttkbootstrap, the constructor signature differs a bit:
        if HAS_BOOTSTRAP:
            super().__init__(themename="flatly")
        else:
            super().__init__()

        self.title("Office Activity Simplifier")
        self.geometry("1700x950")
        self.db = TaskDB()
        self.settings = load_settings()

        self._init_styles()

        self.kanban_selected_id = None
        self.kanban_selected_status = None
        self.kanban_item_map = {status: [] for status in STATUSES}

        self.attachments_var = tk.StringVar(value="")

        self._build_ui()

        self.after(100, self._populate)
        self.after(100, self._populate_kanban)

        # Reminder systems
        self._schedule_task_reminder_checker()
        self._refresh_reminder_display()

        if HAS_OUTLOOK:
            self._schedule_outlook_refresh(self.settings.get("outlook_refresh_minutes", 30))

        self._check_reminders()

    def _treeview_sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        try:
            l.sort(key=lambda t: datetime.strptime(t[0], "%Y-%m-%d"), reverse=reverse)
        except Exception:
            l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, "", index)
        self.tree.heading(col, command=lambda: self._treeview_sort_column(col, not reverse))

    # -------------------- UI --------------------
    def _build_ui(self):
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill=tk.X)

        # toolbar buttons with small icons (unicode) for readability on Windows
        ttk.Button(toolbar, text="📥 Import Outlook Tasks", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="🔄 Refresh Outlook", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="📁 Import CSV", command=self._import_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="📤 Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="📅 Show Today", command=self._show_today_popup).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="⚠️ Show Overdue", command=self._show_overdue_popup).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="⚙️ Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=5)

        ttk.Label(toolbar, text="Theme:").pack(side=tk.RIGHT, padx=(6,4))
        self.theme_var = tk.StringVar(value=self._current_theme)
        theme_cb = ttk.Combobox(toolbar, textvariable=self.theme_var, values=list(self._themes.keys()), width=10, state="readonly")
        theme_cb.pack(side=tk.RIGHT, padx=(0,8))
        theme_cb.bind("<<ComboboxSelected>>", lambda e: self._set_theme(self.theme_var.get()))

        # Main notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # Task List tab
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Task List")

        if self.settings.get("show_description", False):
            cols = ["id", "title", "desc", "due", "priority", "status", "reminder"]
        else:
            cols = ["id", "title", "due", "priority", "status", "reminder"]

        # Filter bar
        filter_frame = ttk.Frame(list_tab, padding=(6,4))
        filter_frame.pack(fill=tk.X, padx=6, pady=(6,4))

        ttk.Label(filter_frame, text="Search:").pack(side=tk.LEFT, padx=(0,4))
        self.filter_text_var = tk.StringVar(value="")
        search_entry = ttk.Entry(filter_frame, textvariable=self.filter_text_var, width=30)
        search_entry.pack(side=tk.LEFT)
        search_entry.bind("<KeyRelease>", lambda e: self._apply_filters())

        ttk.Label(filter_frame, text="Priority:").pack(side=tk.LEFT, padx=(12,4))
        self.filter_priority_var = tk.StringVar(value="All")
        pri_vals = ["All"] + PRIORITIES
        pri_cb = ttk.Combobox(filter_frame, textvariable=self.filter_priority_var, values=pri_vals, width=10, state="readonly")
        pri_cb.pack(side=tk.LEFT)
        pri_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        ttk.Label(filter_frame, text="Status:").pack(side=tk.LEFT, padx=(12,4))
        self.filter_status_var = tk.StringVar(value="All")
        stat_vals = ["All"] + STATUSES
        stat_cb = ttk.Combobox(filter_frame, textvariable=self.filter_status_var, values=stat_vals, width=12, state="readonly")
        stat_cb.pack(side=tk.LEFT)
        stat_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        ttk.Label(filter_frame, text="Due on (YYYY-MM-DD):").pack(side=tk.LEFT, padx=(12,4))
        self.filter_due_var = tk.StringVar(value="")
        due_entry = ttk.Entry(filter_frame, textvariable=self.filter_due_var, width=12)
        due_entry.pack(side=tk.LEFT)

        ttk.Button(filter_frame, text="Apply", command=self._apply_filters).pack(side=tk.LEFT, padx=(12,4))
        ttk.Button(filter_frame, text="Clear", command=self._clear_filters).pack(side=tk.LEFT)

        # Treeview
        self.tree = ttk.Treeview(list_tab, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        for col in cols:
            self.tree.heading(col, text=col.title(), command=lambda _col=col: self._treeview_sort_column(_col, False))

        if self.settings.get("show_description", False):
            self.tree.column("id", width=50, anchor="center")
            self.tree.column("title", width=300, anchor="w")
            self.tree.column("desc", width=300, anchor="w")
            self.tree.column("due", width=100, anchor="center")
            self.tree.column("priority", width=90, anchor="center")
            self.tree.column("status", width=90, anchor="center")
            self.tree.column("reminder", width=120, anchor="center")
        else:
            self.tree.column("id", width=60, anchor="center")
            self.tree.column("title", width=480, anchor="w")
            self.tree.column("due", width=120, anchor="center")
            self.tree.column("priority", width=100, anchor="center")
            self.tree.column("status", width=100, anchor="center")
            self.tree.column("reminder", width=120, anchor="center")

        try:
            self.tree.tag_configure("priority_high", background="#FFD6D6")
            self.tree.tag_configure("priority_medium", background="#FFF5CC")
            self.tree.tag_configure("priority_low", background="#E6FFEA")
            self.tree.tag_configure("oddrow", background="#FFFFFF")
            self.tree.tag_configure("evenrow", background="#F6F6F6")
        except Exception:
            pass

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_task_double_click)

        btns = ttk.Frame(list_tab)
        btns.pack(fill=tk.X, padx=8, pady=5)
        btn_add = ttk.Button(btns, text="➕ Add", command=lambda: self._open_edit_window(None)); btn_add.pack(side=tk.LEFT)
        btn_edit = ttk.Button(btns, text="✏️ Edit", command=lambda: self._open_edit_window(self._selected_tree_task_id() or None)); btn_edit.pack(side=tk.LEFT, padx=5)
        btn_done = ttk.Button(btns, text="✅ Mark Done", command=self._mark_done); btn_done.pack(side=tk.LEFT, padx=5)
        btn_del = ttk.Button(btns, text="🗑️ Delete", command=self._delete_task); btn_del.pack(side=tk.LEFT, padx=5)

        # tooltips
        _ToolTip(btn_add, "Add a new task")
        _ToolTip(btn_edit, "Edit selected task")
        _ToolTip(btn_done, "Mark selected task as done")
        _ToolTip(btn_del, "Delete selected task(s)")

        # Kanban tab
        self.kanban_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.kanban_tab, text="Kanban Board")

        frame = ttk.Frame(self.kanban_tab)
        frame.pack(fill=tk.BOTH, expand=True)

        self.kanban_lists = {}
        for idx, status in enumerate(STATUSES):
            col = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
            col.grid(row=0, column=idx, sticky="nsew", padx=6)
            frame.columnconfigure(idx, weight=2)

            header = ttk.Label(col, text=status, font=("", 12, "bold"))
            header.pack()

            lb = tk.Listbox(col, height=45, width=55, selectmode=tk.EXTENDED)
            lb.pack(fill=tk.BOTH, expand=True)
            lb.status_name = status

            lb.bind("<<ListboxSelect>>", self._kanban_select)
            lb.bind("<ButtonPress-1>", self._on_kanban_drag_start)
            lb.bind("<B1-Motion>", self._on_kanban_drag_motion)
            lb.bind("<ButtonRelease-1>", self._on_kanban_drag_drop)
            lb.bind("<Double-1>", self._on_kanban_double_click)

            self.kanban_lists[status] = lb

        desc_frame = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
        desc_frame.grid(row=0, column=len(STATUSES), sticky="nsew", padx=6)
        frame.columnconfigure(len(STATUSES), weight=1)

        if HAS_HTML:
            self.kanban_html = HTMLLabel(desc_frame, html="", width=50, height=15)
            self.kanban_text = tk.Text(desc_frame, wrap="word", height=15, width=50)
        else:
            self.kanban_html = tk.Text(desc_frame, wrap="word", height=15, width=50)
            self.kanban_text = self.kanban_html

        ttk.Label(desc_frame, text="Task Description / Email").pack(anchor="w")
        self.kanban_text.pack(fill=tk.BOTH, expand=True)

        self.btn_save_desc = ttk.Button(desc_frame, text="Save Description", command=self._save_kanban_desc)
        self.btn_save_desc.pack(pady=5)

        ttk.Label(desc_frame, text="Progress Log").pack(anchor="w")
        self.kanban_progress = tk.Text(desc_frame, height=8, wrap="word", width=50)
        self.kanban_progress.pack(fill=tk.BOTH, expand=True)
        ttk.Button(desc_frame, text="Update Progress", command=self._update_progress).pack(pady=5)

        ttk.Label(desc_frame, text="Attachments").pack(anchor="w", pady=(10, 0))
        self.kanban_attachments_var = tk.StringVar()
        self.kanban_attachments_label = ttk.Label(desc_frame, textvariable=self.kanban_attachments_var, wraplength=350)
        self.kanban_attachments_label.pack(anchor="w", fill=tk.X, pady=2)
        ttk.Button(desc_frame, text="Open Attachments", command=self._open_selected_kanban_attachments).pack(anchor="w", pady=2)

        action_frame = ttk.Frame(self.kanban_tab, padding=5)
        action_frame.pack(fill=tk.X)
        self.btn_edit = ttk.Button(action_frame, text="Edit", command=self._edit_selected_kanban, state="disabled"); self.btn_edit.pack(side=tk.LEFT, padx=5)
        self.btn_delete = ttk.Button(action_frame, text="Delete", command=self._delete_selected_kanban, state="disabled"); self.btn_delete.pack(side=tk.LEFT, padx=5)
        self.btn_done = ttk.Button(action_frame, text="Mark Done", command=self._mark_done_selected_kanban, state="disabled"); self.btn_done.pack(side=tk.LEFT, padx=5)
        self.btn_prev = ttk.Button(action_frame, text="← Move Previous", command=self._move_prev_selected, state="disabled"); self.btn_prev.pack(side=tk.LEFT, padx=5)
        self.btn_next = ttk.Button(action_frame, text="Move Next →", command=self._move_next_selected, state="disabled"); self.btn_next.pack(side=tk.LEFT, padx=5)

    # -------------------- CRUD helper used by Buttons --------------------
    def _selected_tree_task_id(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(self.tree.item(sel[0], "values")[0])
        except Exception:
            return None

    def _add_task(self):
        self._open_edit_window(None)

    def _update_task(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Task", "Select a task first.")
            return
        try:
            task_id = int(self.tree.item(sel[0], "values")[0])
        except Exception:
            messagebox.showwarning("No Task", "Could not determine task id.")
            return
        self._open_edit_window(task_id)

    def _delete_task(self):
        sel = self.tree.selection()
        if not sel:
            return
        confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(sel)} selected task(s)?")
        if not confirm:
            return
        for s in sel:
            task_id = int(self.tree.item(s, "values")[0])
            self.db.delete(task_id)
            self._sync_outlook_task(task_id, {}, action="delete")
        self._populate()
        self._populate_kanban()

    def _mark_done(self):
        sel = self.tree.selection()
        if not sel: return
        task_id = int(self.tree.item(sel[0], "values")[0])
        self.db.mark_done(task_id)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {}, action="done")

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            self.attachments_var.set("")
            return
        vals = self.tree.item(sel[0], "values")
        task_id = int(vals[0])
        cur = self.db.conn.cursor()
        cur.execute("SELECT attachments FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        if row and row["attachments"]:
            try:
                files = json.loads(row["attachments"])
            except Exception:
                files = []
        else:
            files = []
        if hasattr(self, "attachments_var"):
            try:
                self.attachments_var.set(", ".join(os.path.basename(f) for f in files))
            except Exception:
                self.attachments_var.set("")

    def _apply_filters(self):
        fd = self.filter_due_var.get().strip() if hasattr(self, "filter_due_var") else ""
        if fd:
            try:
                datetime.strptime(fd, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Filter", "Due Date filter must be YYYY-MM-DD")
                return
        self._populate()

    def _clear_filters(self):
        if hasattr(self, "filter_text_var"):
            self.filter_text_var.set("")
        if hasattr(self, "filter_priority_var"):
            self.filter_priority_var.set("All")
        if hasattr(self, "filter_status_var"):
            self.filter_status_var.set("All")
        if hasattr(self, "filter_due_var"):
            self.filter_due_var.set("")
        self._populate()

    # -------------------- Populate --------------------
    def _populate(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        rows = self.db.fetch()

        ft = (self.filter_text_var.get().strip().lower() if hasattr(self, "filter_text_var") else "").strip()
        fpri = (self.filter_priority_var.get() if hasattr(self, "filter_priority_var") else "All")
        fstat = (self.filter_status_var.get() if hasattr(self, "filter_status_var") else "All")
        fdue = (self.filter_due_var.get().strip() if hasattr(self, "filter_due_var") else "").strip()

        def row_matches(r):
            if ft:
                hay = ((r["title"] or "") + " " + (r["description"] or "")).lower()
                if ft not in hay:
                    return False
            if fpri and fpri != "All":
                if (r["priority"] or "") != fpri:
                    return False
            if fstat and fstat != "All":
                if (r["status"] or "") != fstat:
                    return False
            if fdue:
                try:
                    if (r["due_date"] or "") != fdue:
                        return False
                except Exception:
                    return False
            return True

        insert_index = 0
        for r in rows:
            if not row_matches(r):
                continue

            desc = r["description"] or ""
            desc_preview = desc.replace("<body>", "").replace("</body>", "").replace("<html>", "").replace("</html>", "")
            desc_preview = desc_preview.replace("\n", " ")
            if len(desc_preview) > 80:
                desc_preview = desc_preview[:80] + "..."

            reminder_val = r["reminder_minutes"] if "reminder_minutes" in r.keys() else None
            reminder_display = str(reminder_val) if reminder_val not in (None, "", "None") else "—"

            if self.settings.get("show_description", False):
                values = [r["id"], r["title"], desc_preview, r["due_date"] or "—", r["priority"], r["status"], reminder_display]
            else:
                values = [r["id"], r["title"], r["due_date"] or "—", r["priority"], r["status"], reminder_display]

            tags = []
            pr = (r["priority"] or "").lower()
            if pr == "high":
                tags.append("priority_high")
            elif pr == "medium":
                tags.append("priority_medium")
            else:
                tags.append("priority_low")
            tags.append("evenrow" if insert_index % 2 == 0 else "oddrow")

            try:
                self.tree.insert("", tk.END, values=values, tags=tags)
            except Exception:
                self.tree.insert("", tk.END, values=values)
            insert_index += 1

    def _populate_kanban(self):
        for status, lb in self.kanban_lists.items():
            lb.delete(0, tk.END)
            self.kanban_item_map[status] = []

        today = date.today()

        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                task_id = r['id']
                title = r['title']
                due_date = r['due_date']
                display = f"{title}"
                idx = lb.size()
                lb.insert(tk.END, display)
                self.kanban_item_map[status].append(task_id)
                if due_date:
                    try:
                        due = datetime.strptime(due_date, "%Y-%m-%d").date()
                        if due < today:
                            lb.itemconfig(idx, bg="#FFCCCC", fg="black")
                        elif due == today:
                            lb.itemconfig(idx, bg="#FFFACD", fg="black")
                    except Exception:
                        pass

    # -------------------- Kanban Actions --------------------
    def _kanban_select(self, event):
        lb = event.widget
        idxs = lb.curselection()
        if not idxs:
            return
        idx = idxs[0]
        status = lb.status_name
        try:
            task_id = self.kanban_item_map[status][idx]
        except Exception:
            messagebox.showwarning("Error", "Could not resolve task id for selection.")
            return
        self.kanban_selected_id = task_id
        self.kanban_selected_status = status
        cur = self.db.conn.cursor()
        cur.execute("SELECT description, progress_log, outlook_id FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        if not row:
            messagebox.showwarning("Error", "Task not found in database.")
            return
        desc = row["description"] or ""
        prog = row["progress_log"] or ""
        outlook_id = row["outlook_id"]

        if outlook_id and HAS_HTML:
            clean = desc or ""
            clean = re.sub(r'<style.*?>.*?</style>', '', clean, flags=re.DOTALL | re.IGNORECASE)
            clean = re.sub(r'<font[^>]*>', '', clean, flags=re.IGNORECASE).replace("</font>", "")
            clean = re.sub(r'style="[^"]*font-size:[^";]*;?"', '', clean, flags=re.IGNORECASE)
            clean = re.sub(r'style="[^"]*font-family:[^";]*;?"', '', clean, flags=re.IGNORECASE)
            clean = re.sub(r'<span[^>]*>', '<span>', clean, flags=re.IGNORECASE)
            if os.name == "nt":
                wrapper_style = "font-family:Segoe UI, Arial; font-size:9pt; line-height:1.3; color:#333;"
            else:
                wrapper_style = "font-family:Arial; font-size:11px; line-height:1.3; color:#333;"
            clean = f"<div style='{wrapper_style}'>{clean}</div>"
            self.kanban_text.pack_forget()
            self.kanban_html.set_html(clean)
            self.kanban_html.pack(fill=tk.BOTH, expand=True)
        else:
            if HAS_HTML:
                try:
                    self.kanban_html.pack_forget()
                except Exception:
                    pass
            self.kanban_text.delete("1.0", tk.END)
            self.kanban_text.insert(tk.END, desc)
            self.kanban_text.pack(fill=tk.BOTH, expand=True)

        self.kanban_progress.delete("1.0", tk.END)
        self.kanban_progress.insert(tk.END, prog)

        cur.execute("SELECT attachments FROM tasks WHERE id=?", (task_id,))
        row2 = cur.fetchone()
        if row2 and row2["attachments"]:
            files = json.loads(row2["attachments"])
            self.kanban_attachments_var.set(", ".join(os.path.basename(f) for f in files))
        else:
            self.kanban_attachments_var.set("No attachments")

        self.btn_edit.config(state="normal")
        self.btn_delete.config(state="normal")
        self.btn_done.config(state="normal")
        self.btn_prev.config(state="normal" if self.kanban_selected_status != "Pending" else "disabled")
        self.btn_next.config(state="normal" if self.kanban_selected_status != "Done" else "disabled")

    def _edit_selected_kanban(self):
        if not self.kanban_selected_id:
            return
        self._open_edit_window(self.kanban_selected_id)

    def _delete_selected_kanban(self):
        for status, lb in self.kanban_lists.items():
            sel = list(lb.curselection())
            if not sel:
                continue
            confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(sel)} selected task(s)?")
            if not confirm:
                return
            for idx in sorted(sel, reverse=True):
                try:
                    task_id = self.kanban_item_map[status][idx]
                except Exception:
                    continue
                self.db.delete(task_id)
                self._sync_outlook_task(task_id, {}, action="delete")
                lb.delete(idx)
                del self.kanban_item_map[status][idx]
        self._populate()
        self._populate_kanban()

    def _mark_done_selected_kanban(self):
        if not self.kanban_selected_id: return
        self.db.mark_done(self.kanban_selected_id)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {}, action="done")

    def _move_prev_selected(self):
        if not self.kanban_selected_id: return
        idx = STATUSES.index(self.kanban_selected_status)
        if idx > 0: self._move_task(self.kanban_selected_id, STATUSES[idx-1])

    def _move_next_selected(self):
        if not self.kanban_selected_id: return
        idx = STATUSES.index(self.kanban_selected_status)
        if idx < len(STATUSES)-1: self._move_task(self.kanban_selected_id, STATUSES[idx+1])

    def _move_task(self, task_id, new_status):
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(task_id, r["title"], r["description"], r["due_date"], r["priority"], new_status)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {"status": new_status}, action="update")

    def _update_progress(self):
        if not self.kanban_selected_id:
            return
        new_line = self.kanban_progress.get("1.0", tk.END).strip()
        if not new_line:
            return
        now = date.today().isoformat()
        entry = f"[{now}] {new_line}\n"
        cur = self.db.conn.cursor()
        cur.execute("SELECT progress_log FROM tasks WHERE id=?", (self.kanban_selected_id,))
        old = cur.fetchone()[0] or ""
        new_log = entry + old
        self.db.update_progress(self.kanban_selected_id, new_log)
        self.kanban_progress.delete("1.0", tk.END)
        self.kanban_progress.insert(tk.END, new_log)
        self._populate_kanban()

    # -------------------- Outlook / CSV / Sync (unchanged) --------------------
    def _get_flagged_from_folder(self, folder, flagged):
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            flagged_items = items.Restrict("[FlagStatus] = 2")
            for item in flagged_items:
                if getattr(item, "Class", 0) == 43:
                    attachments = []
                    try:
                        if item.Attachments.Count > 0:
                            os.makedirs("attachments", exist_ok=True)
                            for att in item.Attachments:
                                fname = os.path.join("attachments", att.FileName)
                                att.SaveAsFile(fname)
                                attachments.append(fname)
                    except Exception as e:
                        print("Attachment import error:", e)
                    due = item.TaskDueDate.strftime("%Y-%m-%d") if getattr(item, "TaskDueDate", None) else None
                    desc = getattr(item, "HTMLBody", "") or getattr(item, "Body", "")
                    flagged.append({
                        "title": f"[Mail] {item.Subject}",
                        "description": desc,
                        "due_date": due,
                        "priority": "Medium",
                        "status": "Pending",
                        "outlook_id": item.EntryID,
                        "attachments": json.dumps(attachments)
                    })
            for sub in folder.Folders:
                self._get_flagged_from_folder(sub, flagged)
        except Exception:
            pass

    def _get_flagged_emails(self):
        if not HAS_OUTLOOK:
            return []
        flagged = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            try:
                todo_folder = outlook.GetDefaultFolder(28)
                for item in todo_folder.Items:
                    if getattr(item, "Class", 0) == 43 and getattr(item,"FlagStatus",0) == 2:
                        due = item.DueDate.strftime("%Y-%m-%d") if getattr(item, "DueDate", None) else None
                        flagged.append({
                            "title": f"[OM] {item.Subject}",
                            "description": item.Body or "",
                            "due_date": due,
                            "priority": "Medium",
                            "status": "Pending",
                            "outlook_id": item.EntryID
                        })
            except Exception as e:
                print("To-Do List fetch error:", e)
            try:
                inbox = outlook.GetDefaultFolder(6)
                self._get_flagged_from_folder(inbox, flagged)
            except Exception as e:
                print("Inbox flagged mail fetch error:", e)
            try:
                search_root = outlook.GetDefaultFolder(23)
                for folder in search_root.Folders:
                    if folder.Name.lower() == "for follow up":
                        self._get_flagged_from_folder(folder, flagged)
            except Exception as e:
                print("Search folder fetch error:", e)
        except Exception as e:
            print("Outlook fetch error:", e)
        return flagged

    def _import_outlook_flags(self):
        flagged = self._get_flagged_emails()
        if not flagged:
            messagebox.showinfo("Outlook", "No active tasks or flagged emails found.")
            return
        cur = self.db.conn.cursor()
        new_items = [f for f in flagged if not cur.execute("SELECT 1 FROM tasks WHERE outlook_id=?", (f["outlook_id"],)).fetchone()]
        self.db.bulk_add(new_items)
        self._populate(); self._populate_kanban()
        messagebox.showinfo("Outlook", f"Imported {len(new_items)} new tasks.")

    def _refresh_outlook_flags(self):
        self._import_outlook_flags()

    def _schedule_outlook_refresh(self, minutes):
        self.after(minutes*60*1000, self._refresh_outlook_flags)

    def _sync_outlook_task(self, task_id, data, action="update"):
        if not HAS_OUTLOOK:
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            cur = self.db.conn.cursor()
            cur.execute("SELECT outlook_id FROM tasks WHERE id=?", (task_id,))
            row = cur.fetchone()
            if not row or not row["outlook_id"]:
                return
            entryid = row["outlook_id"]
            item = None
            try:
                todo_folder = outlook.GetDefaultFolder(28)
                for i in todo_folder.Items:
                    if i.EntryID == entryid:
                        item = i
                        break
            except Exception:
                pass
            if not item:
                try:
                    inbox = outlook.GetDefaultFolder(6)
                    item = inbox.Items.Find(f"[EntryID] = '{entryid}'")
                except Exception:
                    pass
            if not item:
                return
            if action == "done" or (action == "update" and data.get("status") == "Done"):
                if getattr(item, "Class", 0) == 48:
                    item.MarkComplete()
                elif getattr(item, "Class", 0) == 43:
                    item.FlagStatus = 1
                    item.Categories = "Completed"
                item.Save()
            elif action == "delete":
                item.Delete()
        except Exception as e:
            print("Outlook sync error:", e)

    # CSV import/export
    def _import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv")])
        if not path: return
        rows = []
        with open(path,newline="",encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for r in reader:
                if not r.get("title"): continue
                rows.append({"title": r["title"], "description": r.get("description",""), "due_date": r.get("due_date"),
                             "priority": r.get("priority","Medium"), "status": r.get("status","Pending")})
        if rows:
            self.db.bulk_add(rows); self._populate(); self._populate_kanban()
            messagebox.showinfo("CSV Import", f"Imported {len(rows)} tasks.")

    def _export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv")])
        if not path: return
        rows = self.db.fetch()
        with open(path,"w",newline="",encoding="utf-8") as f:
            writer = csv.writer(f); writer.writerow(["title","description","due_date","priority","status"])
            for r in rows: writer.writerow([r["title"], r["description"], r["due_date"], r["priority"], r["status"]])
        messagebox.showinfo("CSV Export", f"Exported {len(rows)} tasks.")

    def _save_kanban_desc(self):
        if not self.kanban_selected_id:
            messagebox.showwarning("No Task", "Please select a task in Kanban first.")
            return
        cur = self.db.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if not r:
            return
        if r["outlook_id"]:
            messagebox.showinfo("Info", "Outlook tasks cannot be edited here. Update directly in Outlook.")
            return
        new_desc = self.kanban_text.get("1.0", tk.END).strip()
        self.db.update(self.kanban_selected_id, r["title"], new_desc, r["due_date"], r["priority"], r["status"])
        self._populate()
        self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {"desc": new_desc}, action="update")
        messagebox.showinfo("Saved", "Description updated successfully.")

    def _show_overdue_popup(self):
        rows = self.db.fetch_overdue()
        win = tk.Toplevel(self); win.title("Overdue Tasks")
        win.geometry("800x400")
        if not rows:
            tk.Label(win, text="✅ No overdue tasks!", font=("", 12, "bold")).pack(padx=20, pady=20)
            return
        cols = ["Title", "Due Date", "Priority", "Status"]
        tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
        tree.heading("Title", text="Title")
        tree.heading("Due Date", text="Due Date")
        tree.heading("Priority", text="Priority")
        tree.heading("Status", text="Status")
        tree.column("Title", width=int(800*0.6), anchor="w")
        tree.column("Due Date", width=int(800*0.14), anchor="center")
        tree.column("Priority", width=int(800*0.13), anchor="center")
        tree.column("Status", width=int(800*0.13), anchor="center")
        for r in rows:
            tree.insert("", tk.END, values=(r["title"], r["due_date"], r["priority"], r["status"]))
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def _show_today_popup(self):
        rows = self.db.fetch_due_today()
        win = tk.Toplevel(self); win.title("Today's Tasks")
        win.geometry("800x400")
        if not rows:
            tk.Label(win, text="🎉 No tasks due today!", font=("", 12, "bold")).pack(padx=20, pady=20)
            return
        cols = ["Title", "Due Date", "Priority", "Status"]
        tree = ttk.Treeview(win, columns=cols, show="headings", height=15)
        tree.heading("Title", text="Title")
        tree.heading("Due Date", text="Due Date")
        tree.heading("Priority", text="Priority")
        tree.heading("Status", text="Status")
        tree.column("Title", width=int(800*0.6), anchor="w")
        tree.column("Due Date", width=int(800*0.14), anchor="center")
        tree.column("Priority", width=int(800*0.13), anchor="center")
        tree.column("Status", width=int(800*0.13), anchor="center")
        for r in rows:
            tree.insert("", tk.END, values=(r["title"], r["due_date"], r["priority"], r["status"]))
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Settings
    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("350x240")
        tk.Label(win, text="Outlook Refresh Minutes").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        refresh_var = tk.IntVar(value=self.settings.get("outlook_refresh_minutes", 30))
        tk.Entry(win, textvariable=refresh_var, width=10).grid(row=0, column=1, padx=10, pady=5)
        show_desc_var = tk.BooleanVar(value=self.settings.get("show_description", False))
        tk.Checkbutton(win, text="Show Description in Task List", variable=show_desc_var).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        ttk.Label(win, text="Default Theme:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        default_theme_var = tk.StringVar(value=self._current_theme)
        ttk.Combobox(win, textvariable=default_theme_var, values=list(self._themes.keys()), state="readonly").grid(row=2, column=1, padx=10, pady=5)
        def save_and_close():
            self.settings["outlook_refresh_minutes"] = refresh_var.get()
            self.settings["show_description"] = show_desc_var.get()
            self._set_theme(default_theme_var.get())
            save_settings(self.settings)
            messagebox.showinfo("Settings", "Settings saved.\nRestart app to apply Task List layout changes.")
            win.destroy()
        ttk.Button(win, text="Save", command=save_and_close).grid(row=3, column=0, columnspan=2, pady=15)

    def _check_reminders(self):
        due_today = self.db.fetch_due_today()
        if due_today and HAS_NOTIFY:
            _safe_show_toast("Tasks Due Today", f"{len(due_today)} tasks due today")
        self.after(3600*1000, self._check_reminders)


# -------------------- Main --------------------
def main():
    app = TaskApp()
    app.mainloop()

if __name__ == "__main__":
    main()