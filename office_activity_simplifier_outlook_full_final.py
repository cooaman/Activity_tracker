import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, json, os, csv
from datetime import datetime, date

try:
    import win32com.client
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False

try:
    from win10toast import ToastNotifier
    toaster = ToastNotifier()
    HAS_NOTIFY = True
except ImportError:
    HAS_NOTIFY = False

try:
    from tkhtmlview import HTMLLabel
    HAS_HTML = True
except ImportError:
    HAS_HTML = False

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
                progress_log TEXT
            );"""
        )
        for col in ["outlook_id", "progress_log"]:
            try:
                cur.execute(f"ALTER TABLE tasks ADD COLUMN {col} TEXT;")
            except sqlite3.OperationalError:
                pass
        self.conn.commit()

    def add(self, title, description, due_date, priority, status="Pending", outlook_id=None):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                """INSERT INTO tasks(title, description, due_date, priority, status,
                   created_at, updated_at, done_at, outlook_id, progress_log)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (title, description, due_date, priority, status, now, now, done_at, outlook_id, ""),
            )

    def update(self, task_id, title, description, due_date, priority, status):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                """UPDATE tasks SET title=?, description=?, due_date=?, priority=?, 
                   status=?, updated_at=?, done_at=? WHERE id=?""",
                (title, description, due_date, priority, status, now, done_at, task_id),
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
        """Insert multiple tasks (used for Outlook/CSV imports)."""
        now = _now_iso()
        with self.conn:
            for r in rows:
                done_at = now if r.get("status") == "Done" else None
                self.conn.execute(
                    """INSERT INTO tasks(title, description, due_date, priority, status,
                                        created_at, updated_at, done_at, outlook_id, progress_log)
                    VALUES(?,?,?,?,?,?,?,?,?,?)""",
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
                    ),
                )

    # -------------------- App --------------------
class TaskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Activity Simplifier")
        self.geometry("1700x950")
        self.db = TaskDB()
        self.settings = load_settings()

        self.kanban_selected_id = None
        self.kanban_selected_status = None

        self._build_ui()
        self.after(100, self._populate)
        self.after(100, self._populate_kanban)

        if HAS_OUTLOOK:
            self._schedule_outlook_refresh(self.settings.get("outlook_refresh_minutes", 30))

        self._check_reminders()

    # -------------------- UI --------------------
    def _build_ui(self):
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill=tk.X)
        ttk.Button(toolbar, text="Import Outlook Tasks", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Refresh Outlook", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Import CSV", command=self._import_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="Show Overdue", command=self._show_overdue_popup).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Show Today", command=self._show_today_popup).pack(side=tk.LEFT, padx=5)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # -------------------- Task List --------------------
        # Task List
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Task List")

        # Columns depend on settings
        if self.settings.get("show_description", False):
            cols = ["id", "title", "desc", "due", "priority", "status"]
        else:
            cols = ["id", "title", "due", "priority", "status"]

        self.tree = ttk.Treeview(list_tab, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        for col in cols:
            self.tree.heading(col, text=col.title())

        # Apply widths
        if self.settings.get("show_description", False):
            self.tree.column("id", width=50, anchor="center")  # ~5%
            self.tree.column("title", width=350, anchor="w")   # ~35%
            self.tree.column("desc", width=350, anchor="w")    # ~35%
            self.tree.column("due", width=100, anchor="center")
            self.tree.column("priority", width=100, anchor="center")
            self.tree.column("status", width=100, anchor="center")
        else:
            self.tree.column("id", width=80, anchor="center")   # ~10%
            self.tree.column("title", width=480, anchor="w")    # ~60%
            self.tree.column("due", width=100, anchor="center") # ~10%
            self.tree.column("priority", width=100, anchor="center")
            self.tree.column("status", width=100, anchor="center")

        form = ttk.LabelFrame(list_tab, text="Task Details", padding=10)
        form.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.title_var = tk.StringVar()
        self.due_var = tk.StringVar()
        self.priority_var = tk.StringVar(value="Medium")
        self.status_var = tk.StringVar(value="Pending")

        ttk.Label(form, text="Title *").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.title_var, width=50).grid(row=0, column=1, sticky="w")

        ttk.Label(form, text="Due Date (YYYY-MM-DD)").grid(row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=self.due_var, width=15).grid(row=0, column=3, sticky="w")

        ttk.Label(form, text="Priority").grid(row=1, column=0, sticky="w")
        ttk.Combobox(form, textvariable=self.priority_var, values=PRIORITIES, state="readonly", width=12).grid(row=1, column=1, sticky="w")

        ttk.Label(form, text="Status").grid(row=1, column=2, sticky="w")
        ttk.Combobox(form, textvariable=self.status_var, values=STATUSES, state="readonly", width=12).grid(row=1, column=3, sticky="w")

        ttk.Label(form, text="Description").grid(row=2, column=0, sticky="nw")
        self.desc_text = tk.Text(form, height=4, width=80)
        self.desc_text.grid(row=2, column=1, columnspan=3, sticky="we")

        btns = ttk.Frame(list_tab)
        btns.pack(fill=tk.X, padx=8, pady=5)
        ttk.Button(btns, text="Add", command=self._add_task).pack(side=tk.LEFT)
        ttk.Button(btns, text="Update", command=self._update_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Mark Done", command=self._mark_done).pack(side=tk.LEFT, padx=5)
        ttk.Button(btns, text="Delete", command=self._delete_task).pack(side=tk.LEFT, padx=5)

        # -------------------- Kanban Board --------------------
        self.kanban_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.kanban_tab, text="Kanban Board")

        frame = ttk.Frame(self.kanban_tab)
        frame.pack(fill=tk.BOTH, expand=True)

        self.kanban_lists = {}
        for idx, status in enumerate(STATUSES):
            col = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
            col.grid(row=0, column=idx, sticky="nsew", padx=6)
            frame.columnconfigure(idx, weight=2)
            ttk.Label(col, text=status, font=("", 12, "bold")).pack()
            lb = tk.Listbox(col, height=30, width=55)
            lb.pack(fill=tk.BOTH, expand=True)
            lb.status_name = status
            lb.bind("<<ListboxSelect>>", self._kanban_select)
            self.kanban_lists[status] = lb

        # Right side details panel
        desc_frame = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
        desc_frame.grid(row=0, column=len(STATUSES), sticky="nsew", padx=6)
        frame.columnconfigure(len(STATUSES), weight=1)

        ttk.Label(desc_frame, text="Task Description / Email").pack(anchor="w")
        if HAS_HTML:
            self.kanban_html = HTMLLabel(desc_frame, html="", width=50, height=15)
            self.kanban_html.pack(fill=tk.BOTH, expand=True)
        else:
            self.kanban_html = tk.Text(desc_frame, wrap="word", height=15, width=50)
            self.kanban_html.pack(fill=tk.BOTH, expand=True)
        ttk.Button(desc_frame, text="Save Description", command=self._save_kanban_desc).pack(pady=5)

        ttk.Label(desc_frame, text="Progress Log").pack(anchor="w")
        self.kanban_progress = tk.Text(desc_frame, height=8, wrap="word", width=50)
        self.kanban_progress.pack(fill=tk.BOTH, expand=True)
        ttk.Button(desc_frame, text="Update Progress", command=self._update_progress).pack(pady=5)

        action_frame = ttk.Frame(self.kanban_tab, padding=5)
        action_frame.pack(fill=tk.X)
        self.btn_edit = ttk.Button(action_frame, text="Edit", command=self._edit_selected_kanban, state="disabled"); self.btn_edit.pack(side=tk.LEFT, padx=5)
        self.btn_delete = ttk.Button(action_frame, text="Delete", command=self._delete_selected_kanban, state="disabled"); self.btn_delete.pack(side=tk.LEFT, padx=5)
        self.btn_done = ttk.Button(action_frame, text="Mark Done", command=self._mark_done_selected_kanban, state="disabled"); self.btn_done.pack(side=tk.LEFT, padx=5)
        self.btn_prev = ttk.Button(action_frame, text="← Move Previous", command=self._move_prev_selected, state="disabled"); self.btn_prev.pack(side=tk.LEFT, padx=5)
        self.btn_next = ttk.Button(action_frame, text="Move Next →", command=self._move_next_selected, state="disabled"); self.btn_next.pack(side=tk.LEFT, padx=5)

    # -------------------- CRUD --------------------
    def _validate_form(self):
        title = self.title_var.get().strip()
        if not title:
            messagebox.showwarning("Validation", "Title is required")
            return None
        due = self.due_var.get().strip()
        if due:
            try:
                datetime.strptime(due, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Validation", "Date must be YYYY-MM-DD")
                return None
        return {"title": title, "desc": self.desc_text.get("1.0", tk.END).strip(), "due": due or None,
                "priority": self.priority_var.get(), "status": self.status_var.get()}

    def _add_task(self):
        d = self._validate_form()
        if not d: return
        self.db.add(d["title"], d["desc"], d["due"], d["priority"], d["status"])
        self._populate(); self._populate_kanban()

    def _update_task(self):
        sel = self.tree.selection()
        if not sel: return
        task_id = int(self.tree.item(sel[0], "values")[0])
        d = self._validate_form()
        if not d: return
        self.db.update(task_id, d["title"], d["desc"], d["due"], d["priority"], d["status"])
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, d, action="update")

    def _delete_task(self):
        sel = self.tree.selection()
        if not sel: return
        task_id = int(self.tree.item(sel[0], "values")[0])
        self.db.delete(task_id)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {}, action="delete")

    def _mark_done(self):
        sel = self.tree.selection()
        if not sel: return
        task_id = int(self.tree.item(sel[0], "values")[0])
        self.db.mark_done(task_id)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {}, action="done")

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        task_id = int(vals[0])
        self.title_var.set(vals[1]); self.due_var.set(vals[2]); self.priority_var.set(vals[3]); self.status_var.set(vals[4])
        cur = self.db.conn.cursor(); cur.execute("SELECT description FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone(); self.desc_text.delete("1.0", tk.END); self.desc_text.insert(tk.END, row[0] if row else "")

    
    
            # -------------------- Populate --------------------
    def _populate(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        for r in self.db.fetch():
            desc = r["description"] or ""
            # Clean HTML for preview
            desc = desc.replace("<body>", "").replace("</body>", "").replace("<html>", "").replace("</html>", "")
            desc = desc.replace("\n", " ")[:80] + "..." if desc else ""

            if self.settings.get("show_description", False):
                values = [r["id"], r["title"], desc, r["due_date"] or "—", r["priority"], r["status"]]
            else:
                values = [r["id"], r["title"], r["due_date"] or "—", r["priority"], r["status"]]
            self.tree.insert("", tk.END, values=values)

    def _populate_kanban(self):
        for lb in self.kanban_lists.values():
            lb.delete(0, tk.END)
        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                display = f"[{r['id']}] {r['title']}"
                idx = lb.size()
                lb.insert(tk.END, display)

                if r["priority"] == "High":
                    lb.itemconfig(idx, fg="red")
                elif r["priority"] == "Medium":
                    lb.itemconfig(idx, fg="orange")
                else:
                    lb.itemconfig(idx, fg="green")

                if r["due_date"] and r["status"] != "Done":
                    try:
                        if datetime.strptime(r["due_date"], "%Y-%m-%d").date() < date.today():
                            lb.itemconfig(idx, fg="red", font=("TkDefaultFont", 10, "bold"))
                    except:
                        pass

    # -------------------- Kanban Actions --------------------
    def _kanban_select(self, event):
        lb = event.widget
        idx = lb.curselection()
        if not idx: return
        line = lb.get(idx[0])
        task_id = int(line.split("]")[0][1:])
        self.kanban_selected_id = task_id
        self.kanban_selected_status = lb.status_name

        cur = self.db.conn.cursor()
        cur.execute("SELECT description, progress_log FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        desc = row["description"] or ""
        prog = row["progress_log"] or ""

        if HAS_HTML and ("<html" in desc.lower() or "<body" in desc.lower()):
            # Clean wrapper tags
            clean = desc.replace("<body>", "").replace("</body>", "").replace("<html>", "").replace("</html>", "")
            self.kanban_html.set_html(clean)
        else:
            if HAS_HTML:
                self.kanban_html.set_html(desc.replace("\n", "<br>"))
            else:
                self.kanban_html.delete("1.0", tk.END)
                self.kanban_html.insert(tk.END, desc)

        self.kanban_progress.delete("1.0", tk.END)
        self.kanban_progress.insert(tk.END, prog)

        self.btn_edit.config(state="normal")
        self.btn_delete.config(state="normal")
        self.btn_done.config(state="normal")
        self.btn_prev.config(state="normal" if self.kanban_selected_status != "Pending" else "disabled")
        self.btn_next.config(state="normal" if self.kanban_selected_status != "Done" else "disabled")

    def _edit_selected_kanban(self):
        if not self.kanban_selected_id: return
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if r:
            self.title_var.set(r["title"]); self.due_var.set(r["due_date"] or "")
            self.priority_var.set(r["priority"]); self.status_var.set(r["status"])
            self.desc_text.delete("1.0", tk.END); self.desc_text.insert(tk.END, r["description"] or "")
            self.notebook.select(0)

    def _delete_selected_kanban(self):
        if not self.kanban_selected_id: return
        self.db.delete(self.kanban_selected_id)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {}, action="delete")

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
        if not self.kanban_selected_id: return
        new_line = self.kanban_progress.get("1.0", tk.END).strip()
        if not new_line: return
        now = date.today().isoformat()
        entry = f"[{now}] {new_line}\n"
        cur = self.db.conn.cursor(); cur.execute("SELECT progress_log FROM tasks WHERE id=?", (self.kanban_selected_id,))
        old = cur.fetchone()[0] or ""
        new_log = entry + old
        self.db.update_progress(self.kanban_selected_id, new_log)
        self._populate_kanban()


    def _get_flagged_from_folder(self, folder, flagged):
        """Recursively fetch flagged mails from a folder + subfolders"""
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # must sort before Restrict()
            flagged_items = items.Restrict("[FlagStatus] = 2")

            for item in flagged_items:
                if getattr(item, "Class", 0) == 43:  # MailItem
                    due = item.TaskDueDate.strftime("%Y-%m-%d") if getattr(item, "TaskDueDate", None) else None
                    desc = getattr(item, "HTMLBody", "") or getattr(item, "Body", "")
                    flagged.append({
                        "title": f"[Mail] {item.Subject}",
                        "description": desc,
                        "due_date": due,
                        "priority": "Medium",
                        "status": "Pending",
                        "outlook_id": item.EntryID
                    })

            # recurse into subfolders
            for sub in folder.Folders:
                self._get_flagged_from_folder(sub, flagged)

        except Exception:
            pass
    # -------------------- Outlook --------------------
    def _get_flagged_emails(self):
        if not HAS_OUTLOOK:
            return []
        flagged = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

            # --- Tasks from To-Do List ---
            try:
                todo_folder = outlook.GetDefaultFolder(28)  # To-Do List
                items = todo_folder.Items
                for item in items:
                    try:
                        if getattr(item, "Class", 0) == 48 and not item.Complete:  # TaskItem
                            due = item.DueDate.strftime("%Y-%m-%d") if getattr(item, "DueDate", None) else None
                            flagged.append({
                                "title": f"[Task] {item.Subject}",
                                "description": item.Body or "",
                                "due_date": due,
                                "priority": "Medium",
                                "status": "Pending",
                                "outlook_id": item.EntryID
                            })
                    except Exception:
                        continue
            except Exception as e:
                print("To-Do List fetch error:", e)

            # --- Flagged Emails from Inbox only (not all subfolders) ---
            try:
                inbox = outlook.GetDefaultFolder(6)  # Inbox
                items = inbox.Items
                items.Sort("[ReceivedTime]", True)
                flagged_items = items.Restrict("[FlagStatus] = 2")
                for item in flagged_items:
                    if getattr(item, "Class", 0) == 43:  # MailItem
                        due = item.TaskDueDate.strftime("%Y-%m-%d") if getattr(item, "TaskDueDate", None) else None
                        desc = getattr(item, "HTMLBody", "") or getattr(item, "Body", "")
                        flagged.append({
                            "title": f"[Mail] {item.Subject}",
                            "description": desc,
                            "due_date": due,
                            "priority": "Medium",
                            "status": "Pending",
                            "outlook_id": item.EntryID
                        })
            except Exception as e:
                print("Inbox flagged mail fetch error:", e)

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

            # Search in Tasks + Inbox
            item = None
            try:
                todo_folder = outlook.GetDefaultFolder(28)  # To-Do List
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
                if getattr(item, "Class", 0) == 48:  # TaskItem
                    item.MarkComplete()
                elif getattr(item, "Class", 0) == 43:  # MailItem
                    item.FlagStatus = 1  # clear flag
                    item.Categories = "Completed"  # optional visual marker
                item.Save()
            elif action == "delete":
                item.Delete()
        except Exception as e:
            print("Outlook sync error:", e)

    # -------------------- CSV --------------------
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

        if HAS_HTML:
            # HTMLLabel is display-only → fallback: keep description unchanged
            messagebox.showinfo("Info", "HTML description cannot be edited directly. Use Task List tab to update.")
            return
        else:
            new_desc = self.kanban_html.get("1.0", tk.END).strip()

        cur = self.db.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(self.kanban_selected_id, r["title"], new_desc, r["due_date"], r["priority"], r["status"])
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {"desc": new_desc}, action="update")


    def _show_overdue_popup(self):
        """Popup window showing overdue tasks in a table"""
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

        # Set widths (approx. % of 800px window)
        tree.column("Title", width=int(800*0.6), anchor="w")
        tree.column("Due Date", width=int(800*0.14), anchor="center")
        tree.column("Priority", width=int(800*0.13), anchor="center")
        tree.column("Status", width=int(800*0.13), anchor="center")

        for r in rows:
            tree.insert("", tk.END, values=(r["title"], r["due_date"], r["priority"], r["status"]))

        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)


    def _show_today_popup(self):
        """Popup window showing today's tasks in a table"""
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


    # -------------------- Settings --------------------
    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("350x200")

        # Outlook refresh setting
        tk.Label(win, text="Outlook Refresh Minutes").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        refresh_var = tk.IntVar(value=self.settings.get("outlook_refresh_minutes", 30))
        tk.Entry(win, textvariable=refresh_var, width=10).grid(row=0, column=1, padx=10, pady=5)

        # Show description in task list setting
        show_desc_var = tk.BooleanVar(value=self.settings.get("show_description", False))
        tk.Checkbutton(win, text="Show Description in Task List", variable=show_desc_var).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5
        )

        def save_and_close():
            self.settings["outlook_refresh_minutes"] = refresh_var.get()
            self.settings["show_description"] = show_desc_var.get()
            save_settings(self.settings)
            messagebox.showinfo("Settings", "Settings saved.\nRestart app to apply Task List layout changes.")
            win.destroy()

        ttk.Button(win, text="Save", command=save_and_close).grid(row=2, column=0, columnspan=2, pady=15)

    # -------------------- Reminders --------------------
    def _check_reminders(self):
        due_today = self.db.fetch_due_today()
        if due_today and HAS_NOTIFY:
            toaster.show_toast("Tasks Due Today", f"{len(due_today)} tasks due today", duration=5)
        self.after(3600*1000, self._check_reminders)


# -------------------- Main --------------------
def main():
    app = TaskApp()
    app.mainloop()

if __name__=="__main__":
    main()