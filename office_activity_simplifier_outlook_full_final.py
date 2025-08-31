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
    return {"outlook_refresh_minutes": 30, "show_description": False}

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
                done_at TEXT,
                outlook_id TEXT
            );
            """
        )
        # Add column if missing
        try:
            cur.execute("ALTER TABLE tasks ADD COLUMN outlook_id TEXT;")
        except sqlite3.OperationalError:
            pass
        self.conn.commit()

    def add(self, title, description, due_date, priority, status="Pending", outlook_id=None):
        now = _now_iso()
        done_at = now if status == "Done" else None
        with self.conn:
            self.conn.execute(
                "INSERT INTO tasks(title, description, due_date, priority, status, created_at, updated_at, done_at, outlook_id) VALUES(?,?,?,?,?,?,?,?,?)",
                (title, description, due_date, priority, status, now, now, done_at, outlook_id),
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
                    "INSERT INTO tasks(title, description, due_date, priority, status, created_at, updated_at, outlook_id) VALUES(?,?,?,?,?,?,?,?)",
                    (
                        r["title"],
                        r.get("description", ""),
                        r.get("due_date"),
                        r.get("priority", "Medium"),
                        r.get("status", "Pending"),
                        now,
                        now,
                        r.get("outlook_id"),
                    ),
                )

# -------------------- Application Layer --------------------
class TaskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Activity Simplifier")
        self.geometry("1500x850")
        self.db = TaskDB()
        self.settings = load_settings()
        self.kanban_selected_id = None
        self.kanban_selected_status = None

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
        ttk.Button(toolbar, text="Import Outlook Tasks", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Refresh Outlook Tasks", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # -------------------- Task List tab --------------------
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Task List")

        cols = ["id", "title"]
        if self.settings.get("show_description", False):
            cols.append("desc")
        cols += ["due", "priority", "status"]

        self.tree = ttk.Treeview(list_tab, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        for col in cols:
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
        self.priority_combo = ttk.Combobox(form, textvariable=self.priority_var, values=PRIORITIES, state="readonly", width=12)
        self.priority_combo.grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Status").grid(row=1, column=2, sticky="w")
        self.status_combo = ttk.Combobox(form, textvariable=self.status_var, values=STATUSES, state="readonly", width=12)
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

        # -------------------- Kanban tab --------------------
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
            lb = tk.Listbox(col, height=20, width=40)
            lb.pack(fill=tk.BOTH, expand=True)
            lb.bind("<<ListboxSelect>>", self._kanban_select)
            lb.status_name = status
            self.kanban_lists[status] = lb

        desc_frame = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
        desc_frame.grid(row=0, column=len(STATUSES), sticky="nsew", padx=6)
        frame.columnconfigure(len(STATUSES), weight=1)

        ttk.Label(desc_frame, text="Task Description (editable)").pack(anchor="w")
        self.kanban_desc = tk.Text(desc_frame, height=20, wrap="word", width=50)
        self.kanban_desc.pack(fill=tk.BOTH, expand=True)
        ttk.Button(desc_frame, text="Save Description", command=self._save_kanban_desc).pack(pady=6)

        action_frame = ttk.Frame(self.kanban_tab, padding=10)
        action_frame.pack(fill=tk.X, pady=5)

        self.btn_edit = ttk.Button(action_frame, text="Edit", command=self._edit_selected_kanban, state="disabled")
        self.btn_edit.pack(side=tk.LEFT, padx=5)

        self.btn_delete = ttk.Button(action_frame, text="Delete", command=self._delete_selected_kanban, state="disabled")
        self.btn_delete.pack(side=tk.LEFT, padx=5)

        self.btn_done = ttk.Button(action_frame, text="Mark Done", command=self._mark_done_selected_kanban, state="disabled")
        self.btn_done.pack(side=tk.LEFT, padx=5)

        self.btn_prev = ttk.Button(action_frame, text="← Move Previous", command=self._move_prev_selected, state="disabled")
        self.btn_prev.pack(side=tk.LEFT, padx=5)

        self.btn_next = ttk.Button(action_frame, text="Move Next →", command=self._move_next_selected, state="disabled")
        self.btn_next.pack(side=tk.LEFT, padx=5)

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
        d = self._validate_form()
        if not d: return
        self.db.add(d["title"], d["desc"], d["due"], d["priority"], d["status"])
        self._populate(); self._populate_kanban(); self._clear_form()

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
        if messagebox.askyesno("Confirm","Delete selected task?"):
            self.db.delete(task_id); self._populate(); self._populate_kanban(); self._clear_form()
            self._sync_outlook_task(task_id, {}, action="delete")

    def _mark_done(self):
        sel = self.tree.selection()
        if not sel: return
        task_id = int(self.tree.item(sel[0], "values")[0])
        self.db.mark_done(task_id); self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {}, action="done")

    def _on_select(self,event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        task_id = int(vals[0])
        self.title_var.set(vals[1])
        if self.settings.get("show_description", False):
            self.due_var.set(vals[3]); self.priority_var.set(vals[4]); self.status_var.set(vals[5])
        else:
            self.due_var.set(vals[2]); self.priority_var.set(vals[3]); self.status_var.set(vals[4])
        cur = self.db.conn.cursor(); cur.execute("SELECT description FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        self.desc_text.delete("1.0", tk.END); self.desc_text.insert(tk.END, row[0] if row else "")

    def _clear_form(self):
        self.title_var.set(""); self.due_var.set("")
        self.priority_var.set("Medium"); self.status_var.set("Pending")
        self.desc_text.delete("1.0", tk.END)

    def _populate(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        for r in self.db.fetch():
            desc = (r["description"] or "").replace("\n"," ")[:80]+"..." if r["description"] else ""
            values=[r["id"], r["title"]]
            if self.settings.get("show_description", False): values.append(desc)
            values += [r["due_date"] or "—", r["priority"], r["status"]]
            self.tree.insert("", tk.END, values=values)

    # -------------------- Kanban --------------------
    def _populate_kanban(self):
        for lb in self.kanban_lists.values(): lb.delete(0, tk.END)
        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                lb.insert(tk.END, f"[#{r['id']}] {r['title']}")

    def _kanban_select(self, event):
        lb = event.widget; idx = lb.curselection()
        if not idx: return
        line = lb.get(idx[0])
        if not line.startswith("[#"): return
        task_id = int(line.split("]")[0][2:])
        self.kanban_selected_id = task_id
        self.kanban_selected_status = lb.status_name

        cur = self.db.conn.cursor(); cur.execute("SELECT description FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        self.kanban_desc.delete("1.0", tk.END)
        if row and row[0]:
            self.kanban_desc.insert(tk.END, row[0])

        self._enable_kanban_buttons(lb.status_name)

    def _save_kanban_desc(self):
        if not self.kanban_selected_id:
            messagebox.showinfo("No Task","Select a task in Kanban first."); return
        new_desc = self.kanban_desc.get("1.0", tk.END).strip()
        cur = self.db.conn.cursor()
        cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(self.kanban_selected_id, r["title"], new_desc, r["due_date"], r["priority"], r["status"])
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {"title": r["title"], "desc": new_desc, "due": r["due_date"]}, action="update")
        messagebox.showinfo("Saved","Description updated successfully.")

    def _enable_kanban_buttons(self, status):
        self.btn_edit.config(state="normal")
        self.btn_delete.config(state="normal")
        self.btn_done.config(state="normal")
        idx_status = STATUSES.index(status)
        self.btn_prev.config(state="normal" if idx_status > 0 else "disabled")
        self.btn_next.config(state="normal" if idx_status < len(STATUSES)-1 else "disabled")

    def _edit_selected_kanban(self):
        if not self.kanban_selected_id: return
        self._edit_task(self.kanban_selected_id)

    def _delete_selected_kanban(self):
        if not self.kanban_selected_id: return
        self._delete_task_kanban(self.kanban_selected_id)
        self._sync_outlook_task(self.kanban_selected_id, {}, action="delete")

    def _mark_done_selected_kanban(self):
        if not self.kanban_selected_id: return
        self._mark_done_kanban(self.kanban_selected_id)
        self._sync_outlook_task(self.kanban_selected_id, {}, action="done")

    def _move_prev_selected(self):
        if not self.kanban_selected_id: return
        idx_status = STATUSES.index(self.kanban_selected_status)
        if idx_status > 0:
            self._move_task(self.kanban_selected_id, STATUSES[idx_status-1])

    def _move_next_selected(self):
        if not self.kanban_selected_id: return
        idx_status = STATUSES.index(self.kanban_selected_status)
        if idx_status < len(STATUSES)-1:
            self._move_task(self.kanban_selected_id, STATUSES[idx_status+1])

    def _edit_task(self, task_id):
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        r = cur.fetchone()
        if not r: return
        self.notebook.select(0)
        self.title_var.set(r["title"]); self.due_var.set(r["due_date"] or "")
        self.priority_var.set(r["priority"]); self.status_var.set(r["status"])
        self.desc_text.delete("1.0", tk.END); self.desc_text.insert(tk.END, r["description"] or "")

    def _delete_task_kanban(self, task_id):
        self.db.delete(task_id); self._populate(); self._populate_kanban()

    def _mark_done_kanban(self, task_id):
        self.db.mark_done(task_id); self._populate(); self._populate_kanban()

    def _move_task(self, task_id, new_status):
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(task_id, r["title"], r["description"], r["due_date"], r["priority"], new_status)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {"title": r["title"], "desc": r["description"], "due": r["due_date"]}, action="update")

    # -------------------- Outlook --------------------
    def _get_flagged_emails(self):
        if not HAS_OUTLOOK: return []
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        flagged = []
        try:
            todo_folder = outlook.GetDefaultFolder(28)  # olFolderToDo
            for task in todo_folder.Items:
                try:
                    status = getattr(task, "Status", None)
                    if status in (0, 1):  # Active tasks only
                        due = task.DueDate.strftime("%Y-%m-%d") if getattr(task, "DueDate", None) else None
                        flagged.append({
                            "title": f"[Outlook] {task.Subject}",
                            "description": (getattr(task, "Body", "") or "").strip()[:500],
                            "due_date": due,
                            "priority": "Medium",
                            "status": "Pending",
                            "outlook_id": task.EntryID
                        })
                except Exception:
                    continue
        except Exception as e:
            print("Error accessing To-Do List:", e)
        return flagged

    def _import_outlook_flags(self):
        rows=self._get_flagged_emails()
        if not rows:
            messagebox.showinfo("Outlook","No active tasks found in To-Do List."); return
        self.db.bulk_add(rows); self._populate(); self._populate_kanban()
        messagebox.showinfo("Outlook", f"Imported {len(rows)} active tasks.")

    def _refresh_outlook_flags(self): 
        self._import_outlook_flags()

    def _schedule_outlook_refresh(self, interval_minutes=30):
        ms=interval_minutes*60*1000
        def cb():
            try: self._refresh_outlook_flags()
            finally: self.after(ms,cb)
        self.after(ms,cb)

    def _sync_outlook_task(self, task_id, new_data, action="update"):
        if not HAS_OUTLOOK: return
        cur = self.db.conn.cursor()
        cur.execute("SELECT outlook_id FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        if not row or not row["outlook_id"]: return
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        try:
            task = outlook.GetItemFromID(row["outlook_id"])
            if action == "delete":
                task.Delete()
                return
            if action == "done":
                task.MarkComplete()
            else:  # update
                task.Subject = new_data["title"]
                task.Body = new_data["desc"]
                if new_data["due"]:
                    task.DueDate = datetime.strptime(new_data["due"], "%Y-%m-%d")
            task.Save()
        except Exception as e:
            print("Outlook sync error:", e)

    # -------------------- CSV --------------------
    def _import_csv(self):
        path=filedialog.askopenfilename(filetypes=[("CSV","*.csv")])
        if not path: return
        with open(path,"r",encoding="utf-8") as f:
            reader=csv.DictReader(f)
            rows=[{"title":row["Title"],"description":row.get("Description",""),
                   "due_date":row.get("Due Date"),
                   "priority":row.get("Priority","Medium"),
                   "status":row.get("Status","Pending")} for row in reader]
        self.db.bulk_add(rows); self._populate(); self._populate_kanban()

    def _export_csv(self):
        path=filedialog.asksaveasfilename(defaultextension=".csv")
        if not path: return
        rows=self.db.fetch()
        with open(path,"w",newline="",encoding="utf-8") as f:
            writer=csv.writer(f)
            writer.writerow(["Title","Description","Due Date","Priority","Status","Created","Updated","Done At","ID"])
            for r in rows:
                writer.writerow([r["title"],r["description"],r["due_date"],r["priority"],
                                 r["status"],r["created_at"],r["updated_at"],r["done_at"],r["id"]])

    # -------------------- Settings --------------------
    def _open_settings(self):
        dlg=tk.Toplevel(self); dlg.title("Settings"); dlg.geometry("400x220")
        interval_var=tk.IntVar(value=self.settings.get("outlook_refresh_minutes",30))
        show_desc_var=tk.BooleanVar(value=self.settings.get("show_description",False))
        ttk.Label(dlg,text="Outlook Auto-Refresh Interval (minutes):").pack(pady=10)
        entry=ttk.Entry(dlg,textvariable=interval_var); entry.pack()
        ttk.Checkbutton(dlg,text="Show Description in Task List",variable=show_desc_var).pack(pady=10)
        def save_and_close():
            self.settings["outlook_refresh_minutes"]=interval_var.get()
            self.settings["show_description"]=show_desc_var.get()
            save_settings(self.settings)
            messagebox.showinfo("Settings","Saved. Restart app to apply.",parent=dlg)
            dlg.destroy()
        ttk.Button(dlg,text="Save",command=save_and_close).pack(pady=10)
        dlg.transient(self); dlg.grab_set()

    # -------------------- Popups --------------------
    def _show_today_popup(self):
        due=self.db.fetch_due_today()
        if due:
            messagebox.showinfo("Today's Tasks","\n".join(
                [f"{r['title']} (Due: {r['due_date']})" if r["due_date"] else r["title"] for r in due]))
        else:
            messagebox.showinfo("Today's Tasks","No tasks due today.")

    def _show_overdue_popup(self):
        overdue=self.db.fetch_overdue()
        if overdue:
            messagebox.showwarning("Overdue Tasks","\n".join(
                [f"{r['title']} (Due: {r['due_date']})" for r in overdue]))
        else:
            messagebox.showinfo("Overdue Tasks","No overdue tasks.")

# -------------------- Entry Point --------------------
def main():
    app=TaskApp()
    app.mainloop()

if __name__=="__main__":
    main()