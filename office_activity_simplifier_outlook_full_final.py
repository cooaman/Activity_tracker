import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, json, os, csv
from datetime import datetime, date

try:
    import win32com.client
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False

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
    return {"outlook_refresh_minutes": 30}

def save_settings(settings):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f)

# -------------------- Database Layer --------------------
class TaskDB:
    def __init__(self, path=DB_FILE):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self._init_db()

    def _init_db(self):
        cur = self.conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS tasks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                description TEXT,
                due_date TEXT,
                priority TEXT DEFAULT 'Medium',
                status TEXT DEFAULT 'Pending',
                created_at TEXT,
                updated_at TEXT,
                done_at TEXT
            );
            """
        )
        self.conn.commit()

    def add(self, title, description, due_date, priority, status="Pending"):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                "INSERT INTO tasks(title, description, due_date, priority, status, created_at, updated_at, done_at) VALUES(?,?,?,?,?,?,?,?)",
                (title, description, due_date, priority, status, now, now, done_at),
            )

    def update(self, task_id, title, description, due_date, priority, status):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET title=?, description=?, due_date=?, priority=?, status=?, updated_at=?, done_at=? WHERE id=?",
                (title, description, due_date, priority, status, now, done_at, task_id),
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

    def fetch(self, status_filter=None):
        cur = self.conn.cursor()
        if status_filter and status_filter != "All":
            cur.execute(
                "SELECT * FROM tasks WHERE status=? ORDER BY due_date IS NULL, due_date ASC, priority DESC",
                (status_filter,),
            )
        else:
            cur.execute(
                "SELECT * FROM tasks ORDER BY due_date IS NULL, due_date ASC, priority DESC"
            )
        return cur.fetchall()

    def fetch_by_status(self, status):
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM tasks WHERE status=? ORDER BY priority DESC, due_date ASC",
            (status,),
        )
        return cur.fetchall()

    def fetch_due_today(self):
        today = date.today().isoformat()
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM tasks WHERE status!='Done' AND due_date=? ORDER BY priority DESC",
            (today,),
        )
        return cur.fetchall()

    def fetch_overdue(self):
        today = date.today().isoformat()
        cur = self.conn.cursor()
        cur.execute(
            "SELECT * FROM tasks WHERE status!='Done' AND due_date IS NOT NULL AND due_date < ? ORDER BY due_date ASC",
            (today,),
        )
        return cur.fetchall()

    def bulk_add(self, rows):
        now = _now_iso()
        with self.conn:
            for r in rows:
                self.conn.execute(
                    "INSERT INTO tasks(title, description, due_date, priority, status, created_at, updated_at) VALUES(?,?,?,?,?,?,?)",
                    (
                        r["title"],
                        r.get("description", ""),
                        r.get("due_date"),
                        r.get("priority", "Medium"),
                        r.get("status", "Pending"),
                        now,
                        now,
                    ),
                )

# -------------------- Application Layer --------------------
class TaskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Activity Simplifier")
        self.geometry("1250x820")
        self.db = TaskDB()
        self.settings = load_settings()

        self._build_ui()
        self._populate()
        self._populate_kanban()

        if HAS_OUTLOOK:
            self._schedule_outlook_refresh(
                interval_minutes=self.settings.get("outlook_refresh_minutes", 30)
            )

    # -------------------- UI --------------------
    def _build_ui(self):
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill=tk.X)
        ttk.Button(toolbar, text="Show Overdue", command=self._show_overdue_popup).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Show Today", command=self._show_today_popup).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(toolbar, text="Import CSV (Bulk)", command=self._import_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=6)

        # Outlook buttons always visible, but handle gracefully
        ttk.Button(toolbar, text="Import Outlook Flags", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Refresh Outlook Flags", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=6)

        ttk.Button(toolbar, text="Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # Task List tab
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Task List")

        self.tree = ttk.Treeview(
            list_tab,
            columns=("id", "title", "due", "priority", "status"),
            show="headings",
        )
        self.tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        for col in ("id", "title", "due", "priority", "status"):
            self.tree.heading(col, text=col.title())

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ttk.LabelFrame(list_tab, text="Task Details", padding=10)
        form.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.title_var = tk.StringVar()
        self.due_var = tk.StringVar()
        self.priority_var = tk.StringVar(value="Medium")
        self.status_var = tk.StringVar(value="Pending")

        ttk.Label(form, text="Title *").grid(row=0, column=0, sticky="w")
        self.title_entry = ttk.Entry(form, textvariable=self.title_var, width=50)
        self.title_entry.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Due Date (YYYY-MM-DD)").grid(row=0, column=2, sticky="w")
        self.due_entry = ttk.Entry(form, textvariable=self.due_var, width=16)
        self.due_entry.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Priority").grid(row=1, column=0, sticky="w")
        self.priority_combo = ttk.Combobox(
            form, textvariable=self.priority_var, values=PRIORITIES, state="readonly", width=12
        )
        self.priority_combo.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Status").grid(row=1, column=2, sticky="w")
        self.status_combo = ttk.Combobox(
            form, textvariable=self.status_var, values=STATUSES, state="readonly", width=12
        )
        self.status_combo.grid(row=1, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Description").grid(row=2, column=0, sticky="nw")
        self.desc_text = tk.Text(form, height=4, width=80)
        self.desc_text.grid(row=2, column=1, columnspan=3, sticky="we", padx=6, pady=4)

        btns = ttk.Frame(list_tab)
        btns.pack(fill=tk.X, padx=8, pady=(0, 10))
        ttk.Button(btns, text="Add New", command=self._add_task).pack(side=tk.LEFT)
        ttk.Button(btns, text="Update Selected", command=self._update_task).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Mark Done", command=self._mark_done).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Delete Selected", command=self._delete_task).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Clear Form", command=self._clear_form).pack(side=tk.LEFT, padx=6)

        # Kanban tab
        self.kanban_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.kanban_tab, text="Kanban Board")
        frame = ttk.Frame(self.kanban_tab)
        frame.pack(fill=tk.BOTH, expand=True)
        self.kanban_lists = {}
        for idx, status in enumerate(STATUSES):
            col = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
            col.grid(row=0, column=idx, sticky="nsew", padx=6)
            frame.columnconfigure(idx, weight=1)
            ttk.Label(col, text=status, font=("", 12, "bold")).pack()
            lb = tk.Listbox(col, height=25)
            lb.pack(fill=tk.BOTH, expand=True)
            self.kanban_lists[status] = lb

    # -------------------- CRUD --------------------
    def _validate_form(self):
        title = self.title_var.get().strip()
        if not title:
            messagebox.showwarning("Validation", "Title is required.")
            return None
        due = self.due_var.get().strip()
        if due:
            try:
                datetime.strptime(due, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Validation", "Due date must be YYYY-MM-DD.")
                return None
        return {
            "title": title,
            "desc": self.desc_text.get("1.0", tk.END).strip(),
            "due": due if due else None,
            "priority": self.priority_var.get(),
            "status": self.status_var.get(),
        }

    def _add_task(self):
        data = self._validate_form()
        if not data:
            return
        self.db.add(data["title"], data["desc"], data["due"], data["priority"], data["status"])
        self._populate(); self._populate_kanban(); self._clear_form()

    def _update_task(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select a task to update.")
            return
        task_id = int(self.tree.item(sel[0], "values")[0])
        data = self._validate_form()
        if not data:
            return
        self.db.update(task_id, data["title"], data["desc"], data["due"], data["priority"], data["status"])
        self._populate(); self._populate_kanban()

    def _delete_task(self):
        sel = self.tree.selection()
        if not sel:
            return
        task_id = int(self.tree.item(sel[0], "values")[0])
        if messagebox.askyesno("Confirm", "Delete selected task?"):
            self.db.delete(task_id)
            self._populate(); self._populate_kanban(); self._clear_form()

    def _mark_done(self):
        sel = self.tree.selection()
        if not sel:
            return
        task_id = int(self.tree.item(sel[0], "values")[0])
        self.db.mark_done(task_id)
        self._populate(); self._populate_kanban()

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        task_id = int(vals[0])
        self.title_var.set(vals[1])
        self.due_var.set(vals[2])
        self.priority_var.set(vals[3])
        self.status_var.set(vals[4])
        cur = self.db.conn.cursor()
        cur.execute("SELECT description FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        self.desc_text.delete("1.0", tk.END)
        self.desc_text.insert(tk.END, row[0] if row else "")

    def _clear_form(self):
        self.title_var.set(""); self.due_var.set("")
        self.priority_var.set("Medium"); self.status_var.set("Pending")
        self.desc_text.delete("1.0", tk.END)

    def _populate(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for r in self.db.fetch():
            self.tree.insert("", tk.END, values=(r["id"], r["title"], r["due_date"] or "—", r["priority"], r["status"]))

    def _kanban_label(self, r):
        label = f"[#{r['id']}] {r['title']}"
        if r["status"] == "Done" and r["done_at"]:
            label += f" (Done: {r['done_at'][:10]})"
        return label

    def _populate_kanban(self):
        for lb in self.kanban_lists.values(): lb.delete(0, tk.END)
        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                idx = lb.size(); lb.insert(tk.END, self._kanban_label(r))
                if status == "Pending": lb.itemconfig(idx, bg="#fff9d6")
                elif status == "In-Progress": lb.itemconfig(idx, bg="#d6ecff")
                elif status == "Done": lb.itemconfig(idx, bg="#d6ffd6", fg="#555555")
                if r["priority"] == "High": lb.itemconfig(idx, fg="red")

    # -------------------- Outlook Integration --------------------
    def _get_flagged_emails(self):
        if not HAS_OUTLOOK:
            return []
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        folders = [inbox] + list(inbox.Folders)  # include subfolders
        flagged = []
        for folder in folders:
            try:
                messages = folder.Items
                for msg in messages:
                    try:
                        if hasattr(msg, "FlagStatus") and msg.FlagStatus == 2:  # only flagged
                            due_date = msg.FlagDueBy.strftime("%Y-%m-%d") if msg.FlagDueBy else None
                            flagged.append({
                                "title": f"[Outlook] {msg.Subject}",
                                "description": (msg.Body or "").strip()[:500],
                                "due_date": due_date,
                                "priority": "Medium",
                                "status": "Pending",
                            })
                    except Exception:
                        continue
            except Exception:
                continue
        return flagged

    def _import_outlook_flags(self):
        rows = self._get_flagged_emails()
        if not rows:
            messagebox.showinfo("Outlook", "No flagged emails found or Outlook not available.")
            return
        self.db.bulk_add(rows)
        self._populate(); self._populate_kanban()
        messagebox.showinfo("Outlook", f"Imported {len(rows)} flagged emails as tasks.")

    def _refresh_outlook_flags(self):
        self._import_outlook_flags()

    def _schedule_outlook_refresh(self, interval_minutes=30):
        ms = interval_minutes * 60 * 1000
        def callback():
            try: self._refresh_outlook_flags()
            finally: self.after(ms, callback)
        self.after(ms, callback)

    # -------------------- CSV Import/Export --------------------
    def _import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path: return
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = [{"title": row["Title"], "description": row.get("Description",""),
                     "due_date": row.get("Due Date"), "priority": row.get("Priority","Medium"),
                     "status": row.get("Status","Pending")} for row in reader]
        self.db.bulk_add(rows); self._populate(); self._populate_kanban()

    def _export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not path: return
        rows = self.db.fetch()
        with open(path,"w",newline="",encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Title","Description","Due Date","Priority","Status","Created","Updated","Done At","ID"])
            for r in rows:
                writer.writerow([r["title"],r["description"],r["due_date"],r["priority"],r["status"],r["created_at"],r["updated_at"],r["done_at"],r["id"]])

    # -------------------- Settings --------------------
    def _open_settings(self):
        dlg = tk.Toplevel(self); dlg.title("Settings"); dlg.geometry("320x160")
        interval_var = tk.IntVar(value=self.settings.get("outlook_refresh_minutes",30))
        ttk.Label(dlg, text="Outlook Auto-Refresh Interval (minutes):").pack(pady=10)
        entry = ttk.Entry(dlg, textvariable=interval_var); entry.pack()
        def save_and_close():
            self.settings["outlook_refresh_minutes"] = interval_var.get()
            save_settings(self.settings)
            messagebox.showinfo("Settings", "Saved. Restart app to apply.", parent=dlg)
            dlg.destroy()
        ttk.Button(dlg, text="Save", command=save_and_close).pack(pady=10)
        dlg.transient(self); dlg.grab_set()

    # -------------------- Popups --------------------
    def _show_today_popup(self):
        due_today = self.db.fetch_due_today()
        if due_today:
            messagebox.showinfo("Today's Tasks", "\n".join([f"{r['title']} (Due: {r['due_date']})" if r["due_date"] else r["title"] for r in due_today]))
        else:
            messagebox.showinfo("Today's Tasks", "No tasks due today.")

    def _show_overdue_popup(self):
        overdue = self.db.fetch_overdue()
        if overdue:
            messagebox.showwarning("Overdue Tasks", "\n".join([f"{r['title']} (Due: {r['due_date']})" for r in overdue]))
        else:
            messagebox.showinfo("Overdue Tasks", "No overdue tasks.")

# -------------------- Entry Point --------------------
def main():
    app = TaskApp()
    app.mainloop()

if __name__ == "__main__":
    main()