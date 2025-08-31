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
    from plyer import notification
    HAS_NOTIFY = True
except ImportError:
    HAS_NOTIFY = False

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
            'CREATE TABLE IF NOT EXISTS tasks('
            'id INTEGER PRIMARY KEY AUTOINCREMENT,'
            'title TEXT NOT NULL,'
            'description TEXT,'
            'due_date TEXT,'
            'priority TEXT DEFAULT "Medium",'
            'status TEXT DEFAULT "Pending",'
            'created_at TEXT,'
            'updated_at TEXT,'
            'done_at TEXT,'
            'outlook_id TEXT'
            ');'
        )
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

# -------------------- App --------------------
class TaskApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Activity Simplifier")
        self.geometry("1700x900")
        self.db = TaskDB()
        self.settings = load_settings()
        self.kanban_selected_id = None
        self.kanban_selected_status = None
        self.drag_data = {"task": None, "source": None}

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
        ttk.Button(toolbar, text="Show Overdue", command=self._show_overdue_popup).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Show Today", command=self._show_today_popup).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(toolbar, text="Import CSV", command=self._import_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Import Outlook Tasks", command=self._import_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Refresh Outlook", command=self._refresh_outlook_flags).pack(side=tk.LEFT, padx=6)
        ttk.Button(toolbar, text="Settings", command=self._open_settings).pack(side=tk.RIGHT, padx=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # Task List tab
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Task List")

        cols = ["id", "title"]
        if self.settings.get("show_description", False): cols.append("desc")
        cols += ["due", "priority", "status"]

        self.tree = ttk.Treeview(list_tab, columns=cols, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        for col in cols: self.tree.heading(col, text=col.title())
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

        # Kanban Board
        self.kanban_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.kanban_tab, text="Kanban Board")

        frame = ttk.Frame(self.kanban_tab); frame.pack(fill=tk.BOTH, expand=True)
        self.kanban_lists = {}
        for idx, status in enumerate(STATUSES):
            col = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
            col.grid(row=0, column=idx, sticky="nsew", padx=6)
            frame.columnconfigure(idx, weight=3)  # wider columns
            ttk.Label(col, text=status, font=("", 12, "bold")).pack()
            lb = tk.Listbox(col, height=25, width=60)
            lb.pack(fill=tk.BOTH, expand=True)
            lb.bind("<<ListboxSelect>>", self._kanban_select)
            lb.bind("<ButtonPress-1>", self._on_drag_start)
            lb.bind("<ButtonRelease-1>", self._on_drag_stop)
            lb.status_name = status
            self.kanban_lists[status] = lb

        # Right-side description panel
        desc_frame = ttk.Frame(frame, padding=6, borderwidth=1, relief="groove")
        desc_frame.grid(row=0, column=len(STATUSES), sticky="nsew", padx=6)
        frame.columnconfigure(len(STATUSES), weight=1)

        ttk.Label(desc_frame, text="Task Details (editable)").pack(anchor="w")
        self.kanban_desc = tk.Text(desc_frame, height=25, wrap="word", width=50)
        self.kanban_desc.pack(fill=tk.BOTH, expand=True)
        ttk.Button(desc_frame, text="Save Description", command=self._save_kanban_desc).pack(pady=6)

        # Action buttons
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

        self.statusbar = tk.Label(self, text="", anchor="w", relief="sunken")
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

    # -------------------- CRUD + Kanban logic continues...


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

        overdue = len(self.db.fetch_overdue())
        today = len(self.db.fetch_due_today())
        pending = len(self.db.fetch_by_status("Pending"))
        self.statusbar.config(text=f"Overdue: {overdue} | Today: {today} | Pending: {pending}")

    # -------------------- Kanban --------------------
    def _populate_kanban(self):
        for lb in self.kanban_lists.values(): lb.delete(0, tk.END)
        for status, lb in self.kanban_lists.items():
            for r in self.db.fetch_by_status(status):
                due = r["due_date"] or "—"
                desc = (r["description"] or "").replace("\n"," ")[:50]
                display = f"[#{r['id']}] {r['title']}\n(Priority: {r['priority']}, Due: {due})\nDesc: {desc}"

                idx = lb.size()
                lb.insert(tk.END, display)

                # Priority color
                if r["priority"] == "High":
                    lb.itemconfig(idx, fg="red")
                elif r["priority"] == "Medium":
                    lb.itemconfig(idx, fg="orange")
                else:
                    lb.itemconfig(idx, fg="green")

                # Overdue bold
                if r["due_date"] and r["status"] != "Done":
                    try:
                        if datetime.strptime(r["due_date"], "%Y-%m-%d").date() < date.today():
                            lb.itemconfig(idx, fg="red", font=("TkDefaultFont", 10, "bold"))
                    except:
                        pass

    def _save_kanban_desc(self):
        if not self.kanban_selected_id:
            messagebox.showinfo("No Task","Select a task in Kanban first."); return
        new_desc = self.kanban_desc.get("1.0", tk.END).strip()
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(self.kanban_selected_id, r["title"], new_desc, r["due_date"], r["priority"], r["status"])
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(self.kanban_selected_id, {"title": r["title"], "desc": new_desc, "due": r["due_date"], "status": r["status"]}, action="update")

    def _enable_kanban_buttons(self, status):
        self.btn_edit.config(state="normal")
        self.btn_delete.config(state="normal")
        self.btn_done.config(state="normal")
        idx_status = STATUSES.index(status)
        self.btn_prev.config(state="normal" if idx_status > 0 else "disabled")
        self.btn_next.config(state="normal" if idx_status < len(STATUSES)-1 else "disabled")

    def _kanban_select(self, event):
        lb = event.widget; idx = lb.curselection()
        if not idx: return
        line = lb.get(idx[0])
        if not line.startswith("[#"): return
        task_id = int(line.split("]")[0][2:])
        self.kanban_selected_id = task_id
        self.kanban_selected_status = lb.status_name

        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        row = cur.fetchone()
        self.kanban_desc.delete("1.0", tk.END)
        if row and row["description"]:
            self.kanban_desc.insert(tk.END, row["description"])

        # due date info
        if row and row["due_date"]:
            try:
                due_dt = datetime.strptime(row["due_date"], "%Y-%m-%d").date()
                delta = (due_dt - date.today()).days
                if delta < 0:
                    self.kanban_desc.insert(tk.END, f"\n⚠️ Due date passed by {abs(delta)} days")
                    self.kanban_desc.config(fg="red")
                elif delta == 0:
                    self.kanban_desc.insert(tk.END, "\n⚠️ Due today!")
                    self.kanban_desc.config(fg="orange")
                else:
                    self.kanban_desc.insert(tk.END, f"\n✅ Due in {delta} days")
                    self.kanban_desc.config(fg="green")
            except:
                pass
        else:
            self.kanban_desc.insert(tk.END, "\n(No due date)")
            self.kanban_desc.config(fg="black")

        self._enable_kanban_buttons(lb.status_name)

    def _edit_selected_kanban(self):
        if not self.kanban_selected_id: return
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (self.kanban_selected_id,))
        r = cur.fetchone()
        if r:
            self.title_var.set(r["title"])
            self.due_var.set(r["due_date"] or "")
            self.priority_var.set(r["priority"])
            self.status_var.set(r["status"])
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
        idx_status = STATUSES.index(self.kanban_selected_status)
        if idx_status > 0:
            self._move_task(self.kanban_selected_id, STATUSES[idx_status-1])

    def _move_next_selected(self):
        if not self.kanban_selected_id: return
        idx_status = STATUSES.index(self.kanban_selected_status)
        if idx_status < len(STATUSES)-1:
            self._move_task(self.kanban_selected_id, STATUSES[idx_status+1])

    def _on_drag_start(self, event):
        lb = event.widget
        idx = lb.nearest(event.y)
        if idx >= 0:
            self.drag_data["task"] = lb.get(idx)
            self.drag_data["source"] = lb

    def _on_drag_stop(self, event):
        if not self.drag_data["task"]: return
        target_lb = event.widget
        if isinstance(target_lb, tk.Listbox) and target_lb != self.drag_data["source"]:
            line = self.drag_data["task"]
            if line.startswith("[#"):
                task_id = int(line.split("]")[0][2:])
                self._move_task(task_id, target_lb.status_name)
        self.drag_data = {"task": None, "source": None}

    def _move_task(self, task_id, new_status):
        cur = self.db.conn.cursor(); cur.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        r = cur.fetchone()
        if not r: return
        self.db.update(task_id, r["title"], r["description"], r["due_date"], r["priority"], new_status)
        self._populate(); self._populate_kanban()
        self._sync_outlook_task(task_id, {"title": r["title"], "desc": r["description"], "due": r["due_date"], "status": new_status}, action="update")




        # Part 3

            # -------------------- Outlook Sync --------------------
    def _get_flagged_emails(self):
        """
        Import Active Tasks and Flagged Emails from Outlook 'To-Do List'.
        Only imports tasks that are not complete and emails that are flagged but not completed.
        """
        if not HAS_OUTLOOK:
            return []

        flagged = []
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            todo_folder = outlook.GetDefaultFolder(28)  # olFolderToDo
            items = todo_folder.Items

            for item in items:
                try:
                    cls = getattr(item, "Class", 0)

                    # TaskItem (Class 48)
                    if cls == 48:
                        if not item.Complete:
                            due = item.DueDate.strftime("%Y-%m-%d") if getattr(item, "DueDate", None) else None
                            flagged.append({
                                "title": f"[Task] {item.Subject}",
                                "description": (getattr(item, "Body", "") or "")[:500],
                                "due_date": due,
                                "priority": "Medium",
                                "status": "Pending",
                                "outlook_id": item.EntryID
                            })

                    # MailItem (Class 43)
                    elif cls == 43:
                        flag_status = getattr(item, "FlagStatus", None)
                        is_done = getattr(item, "IsMarkedAsTaskComplete", None)
                        due_date = getattr(item, "TaskDueDate", None)

                        if flag_status == 2 and not is_done:  # olFlagMarked and not completed
                            due = due_date.strftime("%Y-%m-%d") if due_date else None
                            flagged.append({
                                "title": f"[Mail] {item.Subject}",
                                "description": (getattr(item, "Body", "") or "")[:500],
                                "due_date": due,
                                "priority": "Medium",
                                "status": "Pending",
                                "outlook_id": item.EntryID
                            })

                except Exception as inner_e:
                    print("Outlook Item error:", inner_e)

        except Exception as e:
            print("Outlook fetch error:", e)

        return flagged

    def _import_outlook_flags(self):
        flagged = self._get_flagged_emails()
        if not flagged:
            messagebox.showinfo("Outlook", "No active tasks or flagged emails found.")
            return

        cur = self.db.conn.cursor()
        new_tasks = []
        for task in flagged:
            cur.execute("SELECT 1 FROM tasks WHERE outlook_id=?", (task["outlook_id"],))
            if not cur.fetchone():  # only add if not already imported
                new_tasks.append(task)

        if new_tasks:
            self.db.bulk_add(new_tasks)
            self._populate(); self._populate_kanban()
            messagebox.showinfo("Outlook", f"Imported {len(new_tasks)} new tasks.")
        else:
            messagebox.showinfo("Outlook", "No new Outlook tasks to import.")

    def _refresh_outlook_flags(self):
        """Two-way sync between Outlook To-Do List and the tool."""
        flagged = self._get_flagged_emails()
        if not flagged:
            messagebox.showinfo("Outlook", "No active tasks or flagged emails found.")
            return

        cur = self.db.conn.cursor()
        new_count, updated_count, synced_count = 0, 0, 0

        # 1. Import new or update changed from Outlook -> Tool
        for task in flagged:
            cur.execute("SELECT * FROM tasks WHERE outlook_id=?", (task["outlook_id"],))
            row = cur.fetchone()
            if not row:
                self.db.bulk_add([task])
                new_count += 1
            else:
                if (row["title"] != task["title"] or
                    row["description"] != task["description"] or
                    row["due_date"] != task["due_date"] or
                    row["status"] != task["status"]):
                    self.db.update(row["id"], task["title"], task["description"],
                                   task["due_date"], task["priority"], task["status"])
                    updated_count += 1

        # 2. Push updates from Tool -> Outlook
        cur.execute("SELECT * FROM tasks WHERE outlook_id IS NOT NULL")
        for row in cur.fetchall():
            self._sync_outlook_task(row["id"], {
                "title": row["title"],
                "desc": row["description"],
                "due": row["due_date"],
                "status": row["status"]
            }, action="update")
            synced_count += 1

        self._populate(); self._populate_kanban()

        messagebox.showinfo("Outlook Sync",
                            f"New Imported: {new_count}\n"
                            f"Updated from Outlook: {updated_count}\n"
                            f"Pushed back to Outlook: {synced_count}")

    def _schedule_outlook_refresh(self, minutes):
        self.after(minutes*60*1000, self._refresh_outlook_flags)

    def _sync_outlook_task(self, task_id, data, action="update"):
        """Push updates back to Outlook To-Do List."""
        if not HAS_OUTLOOK: return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            cur = self.db.conn.cursor()
            cur.execute("SELECT outlook_id FROM tasks WHERE id=?", (task_id,))
            row = cur.fetchone()
            if not row or not row["outlook_id"]: return
            entryid = row["outlook_id"]

            item = None
            try:
                item = outlook.GetItemFromID(entryid)
            except Exception as e:
                print(f"DEBUG: GetItemFromID failed for {entryid} - {e}")
                return

            if not item: return

            if action=="update":
                if hasattr(item,"Subject"): item.Subject = data.get("title", item.Subject)
                if hasattr(item,"Body"): item.Body = data.get("desc", item.Body)
                if data.get("due"):
                    try: item.DueDate = datetime.strptime(data["due"], "%Y-%m-%d")
                    except: pass
                item.Save()
            elif action=="delete":
                item.Delete()
            elif action=="done":
                if hasattr(item,"MarkComplete"):
                    item.MarkComplete()
                elif hasattr(item,"FlagStatus"):
                    item.FlagStatus = 1  # olFlagComplete
                item.Save()
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
                row = {k.strip().lower(): (v.strip() if v else "") for k,v in r.items()}
                if not row.get("title"): continue
                rows.append({
                    "title": row.get("title","Untitled Task"),
                    "description": row.get("description",""),
                    "due_date": row.get("due_date") or None,
                    "priority": row.get("priority","Medium"),
                    "status": row.get("status","Pending")
                })
        if rows:
            self.db.bulk_add(rows)
            self._populate(); self._populate_kanban()
            messagebox.showinfo("CSV Import", f"Imported {len(rows)} tasks.")
        else:
            messagebox.showwarning("CSV Import","No valid tasks found.")

    def _export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv")])
        if not path: return
        rows = self.db.fetch()
        with open(path,"w",newline="",encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["title","description","due_date","priority","status"])
            for r in rows:
                writer.writerow([r["title"], r["description"], r["due_date"], r["priority"], r["status"]])
        messagebox.showinfo("CSV Export", f"Exported {len(rows)} tasks.")

    # -------------------- Settings --------------------
    def _open_settings(self):
        win = tk.Toplevel(self); win.title("Settings")
        tk.Label(win,text="Outlook Refresh Minutes").grid(row=0,column=0,sticky="w")
        refresh_var = tk.IntVar(value=self.settings.get("outlook_refresh_minutes",30))
        tk.Entry(win,textvariable=refresh_var).grid(row=0,column=1)
        show_desc_var = tk.BooleanVar(value=self.settings.get("show_description",False))
        tk.Checkbutton(win,text="Show Description in Task List",variable=show_desc_var).grid(row=1,column=0,columnspan=2,sticky="w")

        def save_and_close():
            self.settings["outlook_refresh_minutes"]=refresh_var.get()
            self.settings["show_description"]=show_desc_var.get()
            save_settings(self.settings)
            messagebox.showinfo("Settings","Saved. Restart app to apply fully.")
            win.destroy()
        ttk.Button(win,text="Save",command=save_and_close).grid(row=2,column=0,columnspan=2,pady=6)

    # -------------------- Reminders --------------------
    def _check_reminders(self):
        due_today = self.db.fetch_due_today()
        if due_today and HAS_NOTIFY:
            notification.notify(title="Tasks Due Today", message=f"{len(due_today)} tasks due today", timeout=5)
        self.after(3600*1000, self._check_reminders)

    # -------------------- Popups --------------------
    def _show_overdue_popup(self):
        rows = self.db.fetch_overdue()
        msg = "\n".join([f"{r['title']} (Due {r['due_date']})" for r in rows]) or "No overdue tasks."
        messagebox.showinfo("Overdue Tasks", msg)

    def _show_today_popup(self):
        rows = self.db.fetch_due_today()
        msg = "\n".join([f"{r['title']} (Due {r['due_date']})" for r in rows]) or "No tasks due today."
        messagebox.showinfo("Today's Tasks", msg)

# -------------------- Main --------------------
def main():
    app=TaskApp()
    app.mainloop()

if __name__=="__main__":
    main()