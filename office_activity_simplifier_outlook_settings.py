
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, json, os
from datetime import datetime, date
import csv

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
            with open(SETTINGS_FILE,"r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"outlook_refresh_minutes": 30}

def save_settings(settings):
    with open(SETTINGS_FILE,"w") as f:
        json.dump(settings,f)

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

    def bulk_add(self, rows):
        now = _now_iso()
        with self.conn:
            self.conn.executemany(
                "INSERT INTO tasks(title, description, due_date, priority, status, created_at, updated_at, done_at) VALUES(?,?,?,?,?,?,?,?)",
                [(r['title'], r.get('description') or '', r.get('due_date'), r.get('priority') or 'Medium',
                  r.get('status') or 'Pending', now, now, (now if (r.get('status') or 'Pending')=='Done' else None))
                 for r in rows]
            )

    def update_status(self, task_id, status):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET status=?, updated_at=?, done_at=? WHERE id=?",
                (status, now, done_at, task_id),
            )

    def fetch(self, status_filter=None):
        cur = self.conn.cursor()
        if status_filter and status_filter != "All":
            cur.execute("SELECT * FROM tasks WHERE status=? ORDER BY due_date IS NULL, due_date ASC, priority DESC", (status_filter,))
        else:
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
        cur.execute("SELECT * FROM tasks WHERE status!='Done' AND due_date IS NOT NULL AND due_date < ? ORDER BY due_date ASC", (today,))
        return cur.fetchall()

class TaskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Activity Simplifier")
        self.geometry("1200x780")
        self.db = TaskDB()
        self.settings = load_settings()

        self._build_ui()
        self._populate()
        self._populate_kanban()

        if HAS_OUTLOOK:
            self._schedule_outlook_refresh(interval_minutes=self.settings.get("outlook_refresh_minutes",30))

    def _build_ui(self):
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill=tk.X)
        ttk.Button(toolbar, text="Show Overdue", command=self._show_overdue_popup).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Show Today", command=self._show_today_popup).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(toolbar, text="Import CSV (Bulk)", command=self._import_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=6)
        if HAS_OUTLOOK:
            ttk.Button(toolbar, text="Import Outlook Flags", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=6)
            ttk.Button(toolbar, text="Refresh Outlook Flags", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))

        # List tab
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="List")
        self.tree = ttk.Treeview(list_tab, columns=("id","title","status"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        for col in ("id","title","status"):
            self.tree.heading(col, text=col.title())

        # Kanban tab
        self.kanban_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.kanban_tab, text="Kanban")
        frame = ttk.Frame(self.kanban_tab)
        frame.pack(fill=tk.BOTH, expand=True)
        self.kanban_lists = {}
        for idx, status in enumerate(STATUSES):
            col = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
            col.grid(row=0, column=idx, sticky="nsew", padx=6)
            frame.columnconfigure(idx, weight=1)
            ttk.Label(col, text=status, font=("",12,"bold")).pack()
            lb = tk.Listbox(col, height=25)
            lb.pack(fill=tk.BOTH, expand=True)
            self.kanban_lists[status] = lb

    def _populate(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for r in self.db.fetch():
            self.tree.insert("", tk.END, values=(r["id"], r["title"], r["status"]))

    def _kanban_label(self, r):
        label = f"[#{r['id']}] {r['title']}"
        if r["status"] == "Done" and r["done_at"]:
            label += f" (Done: {r['done_at'][:10]})"
        return label

    def _populate_kanban(self):
        for lb in self.kanban_lists.values():
            lb.delete(0, tk.END)
        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                idx = lb.size()
                lb.insert(tk.END, self._kanban_label(r))
                if status == "Pending":
                    lb.itemconfig(idx, bg="#fff9d6")
                elif status == "In-Progress":
                    lb.itemconfig(idx, bg="#d6ecff")
                elif status == "Done":
                    lb.itemconfig(idx, bg="#d6ffd6", fg="#555555")
                if r["priority"] == "High":
                    lb.itemconfig(idx, fg="red", font=("TkDefaultFont", 10, "bold"))

    # Outlook
    def _get_flagged_emails(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        flagged = []
        for msg in messages:
            try:
                if msg.IsMarkedAsTask:
                    due_date = msg.TaskDueDate.strftime("%Y-%m-%d") if msg.TaskDueDate else None
                    flagged.append({
                        "title": f"[Outlook] {msg.Subject}",
                        "description": (msg.Body or "").strip()[:500],
                        "due_date": due_date,
                        "priority": "Medium",
                        "status": "Pending"
                    })
            except Exception:
                continue
        return flagged

    def _import_outlook_flags(self):
        rows = self._get_flagged_emails()
        if rows:
            self.db.bulk_add(rows)
            self._populate(); self._populate_kanban()
            messagebox.showinfo("Outlook", f"Imported {len(rows)} flagged emails as tasks.")
        else:
            messagebox.showinfo("Outlook", "No flagged emails found.")

    def _refresh_outlook_flags(self):
        cur = self.db.conn.cursor()
        cur.execute("DELETE FROM tasks WHERE title LIKE '[Outlook]%'")
        self.db.conn.commit()
        self._import_outlook_flags()

    def _schedule_outlook_refresh(self, interval_minutes=30):
        ms = interval_minutes * 60 * 1000
        def callback():
            try:
                self._refresh_outlook_flags()
            finally:
                self.after(ms, callback)
        self.after(ms, callback)

    # Settings UI
    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("320x160")
        interval_var = tk.IntVar(value=self.settings.get("outlook_refresh_minutes",30))

        ttk.Label(dlg,text="Outlook Auto-Refresh Interval (minutes):").pack(pady=10)
        entry = ttk.Entry(dlg,textvariable=interval_var)
        entry.pack()

        def save_and_close():
            self.settings["outlook_refresh_minutes"]=interval_var.get()
            save_settings(self.settings)
            messagebox.showinfo("Settings","Saved. Restart app to apply.",parent=dlg)
            dlg.destroy()

        ttk.Button(dlg,text="Save",command=save_and_close).pack(pady=10)
        dlg.transient(self)
        dlg.grab_set()

    def _import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV","*.csv")])
        if not path: return
        with open(path,"r",encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows=[{"title":row["Title"],"description":row.get("Description",""),
                   "due_date":row.get("Due Date"),"priority":row.get("Priority","Medium"),
                   "status":row.get("Status","Pending")} for row in reader]
        self.db.bulk_add(rows)
        self._populate(); self._populate_kanban()

    def _export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not path: return
        rows = self.db.fetch()
        with open(path,"w",newline="",encoding="utf-8") as f:
            writer=csv.writer(f)
            writer.writerow(["Title","Description","Due Date","Priority","Status","Created","Updated","Done At","ID"])
            for r in rows:
                writer.writerow([r["title"],r["description"],r["due_date"],r["priority"],r["status"],r["created_at"],r["updated_at"],r["done_at"],r["id"]])

    def _show_today_popup(self):
        due_today=self.db.fetch_due_today()
        if due_today:
            messagebox.showinfo("Today's Tasks","\n".join([r["title"] for r in due_today]))
        else:
            messagebox.showinfo("Today's Tasks","No tasks due today.")

    def _show_overdue_popup(self):
        overdue=self.db.fetch_overdue()
        if overdue:
            messagebox.showwarning("Overdue Tasks","\n".join([r["title"] for r in overdue]))
        else:
            messagebox.showinfo("Overdue Tasks","No overdue tasks.")

def main():
    app=TaskApp()
    app.mainloop()

if __name__=="__main__":
    main()
