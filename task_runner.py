import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import schedule
import time
import threading
import json
from datetime import datetime, timedelta
import logging
import webbrowser
import subprocess
import shutil
import getpass
from croniter import croniter
import re
import glob

STARTUP_MARKER = '.task_runner_first_run'

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)
    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("Segoe UI", 10))
        label.pack(ipadx=1)
    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

class TaskRunner:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Task Runner")
        self.root.geometry("700x500")
        self.set_theme()
        self.setup_logging()
        self.tasks = self.load_tasks()
        # Add a lock to prevent overlapping runs
        self.job_running_lock = threading.Lock()
        # Recalculate next_run for all tasks if missing or invalid
        for task in self.tasks:
            if not task.get('next_run') or task.get('next_run') in (None, '-', ''):
                try:
                    base = datetime.now()
                    itr = croniter(task["cron_expr"], base)
                    task["next_run"] = itr.get_next(datetime).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    task["next_run"] = "-"
        self.save_tasks()
        self.create_gui()
        self.update_task_list()
        self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduler_thread.start()
        # Minimize if --minimized is passed
        if '--minimized' in sys.argv:
            self.root.iconify()

    def set_theme(self):
        style = ttk.Style(self.root)
        # Use a modern theme if available
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11), padding=6)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)

    def setup_logging(self):
        logging.basicConfig(
            filename='task_runner.log',
            level=logging.INFO,
            format='%(asctime)s - %(message)s'
        )

    def create_gui(self):
        # Task Configuration Frame
        file_frame = ttk.LabelFrame(self.root, text="Task Configuration", padding="15 10 15 10")
        file_frame.pack(fill="x", padx=15, pady=10)
        file_frame.columnconfigure(1, weight=1)

        # Name
        ttk.Label(file_frame, text="Name:", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(file_frame, textvariable=self.name_var, width=25)
        name_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ToolTip(name_entry, "Name for this task (defaults to filename)")

        # Script/File
        ttk.Label(file_frame, text="Script/File:", font=("Segoe UI", 11, "bold")).grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.file_path = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        file_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        ToolTip(browse_btn, "Browse for a script or executable file to schedule.")

        # Cron Expression
        ttk.Label(file_frame, text="Cron Expression:", font=("Segoe UI", 11, "bold")).grid(row=2, column=0, sticky="w", pady=5, padx=5)
        self.cron_var = tk.StringVar(value="* * * * *")
        cron_entry = ttk.Entry(file_frame, textvariable=self.cron_var, width=15)
        cron_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        ToolTip(cron_entry, "Enter a cron expression (min hour day month weekday)")
        # Place crontab.guru link to the right of the input
        url_label = tk.Label(file_frame, text="crontab.guru", fg="blue", cursor="hand2", font=("Segoe UI", 11, "underline"))
        url_label.grid(row=2, column=2, sticky="w", padx=(8, 0))
        url_label.bind("<Button-1>", lambda e: self.open_crontab_guru())
        ToolTip(url_label, "Open crontab.guru to help build cron expressions.")
        # Place the label below the cron input
        ttk.Label(file_frame, text="(min hour day month weekday)").grid(row=3, column=1, padx=5, sticky="w")

        # Log Retention and Log Executions on the same row
        ttk.Label(file_frame, text="Log Retention (days):", font=("Segoe UI", 11, "bold")).grid(row=4, column=0, sticky="w", pady=5, padx=5)
        self.log_retention = tk.StringVar(value="1")
        retention_entry = ttk.Entry(file_frame, textvariable=self.log_retention, width=10)
        retention_entry.grid(row=4, column=1, sticky="w", padx=(5, 20), pady=5)
        ToolTip(retention_entry, "How many days to keep logs for this task.")

        ttk.Label(file_frame, text="Log Executions to Keep:", font=("Segoe UI", 11, "bold")).grid(row=4, column=2, sticky="w", pady=5, padx=5)
        self.log_executions = tk.StringVar(value="1")
        executions_entry = ttk.Entry(file_frame, textvariable=self.log_executions, width=10)
        executions_entry.grid(row=4, column=3, sticky="w", padx=5, pady=5)
        ToolTip(executions_entry, "How many executions to keep in the log for this task.")

        # Add progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(file_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.progress_bar.grid_remove()  # Hide initially

        # Save Task Button
        button_frame_top = ttk.Frame(file_frame)
        button_frame_top.grid(row=8, column=0, columnspan=3, pady=12, sticky="w")
        self.add_update_button = ttk.Button(button_frame_top, text="Save Task", command=self.add_or_update_task)
        self.add_update_button.pack(side="left", padx=(0, 8))
        ToolTip(self.add_update_button, "Save the configured task.")
        new_btn = ttk.Button(button_frame_top, text="New Task", command=self.clear_form)
        new_btn.pack(side="left", padx=(0, 8))
        ToolTip(new_btn, "Clear the form to add a new task.")
        self.run_btn = ttk.Button(button_frame_top, text="Run Task", command=self.run_selected_task)
        self.run_btn.pack(side="left", padx=(0, 8))
        ToolTip(self.run_btn, "Run the selected task immediately and show terminal output.")
        self.show_log_button = ttk.Button(button_frame_top, text="Show Log", command=self.show_log)
        self.show_log_button.pack(side="left", padx=(0, 8))
        ToolTip(self.show_log_button, "View the log file for the selected task.")
        self.delete_btn = ttk.Button(button_frame_top, text="Delete Task", command=self.remove_task)
        self.delete_btn.pack(side="left")
        ToolTip(self.delete_btn, "Delete the selected task.")
        # Initially hide action buttons
        self.run_btn.pack_forget()
        self.show_log_button.pack_forget()
        self.delete_btn.pack_forget()

        # Scheduled Tasks Frame
        list_frame = ttk.LabelFrame(self.root, text="Scheduled Tasks", padding="10 10 10 10")
        list_frame.pack(fill="both", expand=True, padx=15, pady=5)

        # Task List with alternate row colors
        self.task_tree = ttk.Treeview(list_frame, columns=("Name", "Schedule", "Last Execution", "Next Execution", "Status"), show="headings")
        self.task_tree.heading("Name", text="Name")
        self.task_tree.heading("Schedule", text="Schedule")
        self.task_tree.heading("Last Execution", text="Last Execution")
        self.task_tree.heading("Next Execution", text="Next Execution")
        self.task_tree.heading("Status", text="Status")
        self.task_tree.column("Name", width=140)
        self.task_tree.column("Schedule", width=100)
        self.task_tree.column("Last Execution", width=120)
        self.task_tree.column("Next Execution", width=120)
        self.task_tree.column("Status", width=80)
        self.task_tree.pack(fill="both", expand=True, pady=(0, 10))
        self.task_tree.tag_configure('oddrow', background='#f0f4ff')
        self.task_tree.tag_configure('evenrow', background='#ffffff')
        self.task_tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        self.selected_task_index = None

        self.update_task_list()

    def open_crontab_guru(self):
        webbrowser.open_new("https://crontab.guru/")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("All Files", "*.*"), ("Python Files", "*.py"), ("Batch Files", "*.bat")]
        )
        if filename:
            self.file_path.set(filename)
            if not self.name_var.get().strip():
                self.name_var.set(os.path.basename(filename))

    def add_or_update_task(self):
        file_path = self.file_path.get()
        cron_expr = self.cron_var.get()
        retention = self.log_retention.get()
        name = self.name_var.get().strip()
        log_executions = self.log_executions.get().strip()
        if not file_path or not cron_expr:
            messagebox.showerror("Error", "Please fill in all fields")
            return
        try:
            retention = int(retention)
            if retention < 1:
                raise ValueError
            log_executions = int(log_executions)
            if log_executions < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Log retention and executions must be positive numbers")
            return
        if not name:
            name = os.path.basename(file_path)
        task = {
            "name": name,
            "file_path": file_path,
            "cron_expr": cron_expr,
            "retention": retention,
            "status": "Active",
            "log_executions": log_executions
        }
        # Calculate next_run for new/updated task
        try:
            base = datetime.now()
            itr = croniter(cron_expr, base)
            task["next_run"] = itr.get_next(datetime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            task["next_run"] = "-"
        if self.selected_task_index is not None:
            # Update existing task
            old_task = self.tasks[self.selected_task_index]
            # Preserve last_execution if present
            if "last_execution" in old_task:
                task["last_execution"] = old_task["last_execution"]
            self.tasks[self.selected_task_index] = task
            self.selected_task_index = None
            self.add_update_button.config(text="Save Task")
            self.save_tasks()
            self.update_task_list()
            self.root.after(100, lambda: messagebox.showinfo("Task Saved", "Task updated successfully."))
            self.root.after(150, self.clear_form)
            return
        else:
            # Add new task
            self.tasks.append(task)
            self.schedule_task(task)
        self.save_tasks()
        self.update_task_list()
        self.clear_form()

    def clear_form(self):
        self.clearing_form = True
        self.name_var.set("")
        self.file_path.set("")
        self.cron_var.set("* * * * *")
        self.log_retention.set("1")
        self.log_executions.set("1")
        self.selected_task_index = None
        self.add_update_button.config(text="Save Task")
        self.task_tree.selection_remove(self.task_tree.selection())
        # Hide action buttons
        self.run_btn.pack_forget()
        self.show_log_button.pack_forget()
        self.delete_btn.pack_forget()
        self.clearing_form = False

    def remove_task(self):
        selected = self.task_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a task to delete")
            return
        index = self.task_tree.index(selected[0])
        del self.tasks[index]
        self.save_tasks()
        self.update_task_list()
        self.clear_form()

    def update_task_list(self):
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        for idx, task in enumerate(self.tasks):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            last_exec = task.get("last_execution", "-")
            next_exec = task.get("next_run", self.get_next_execution(task))
            self.task_tree.insert("", "end", values=(
                task.get("name", os.path.basename(task["file_path"])),
                task["cron_expr"],
                last_exec,
                next_exec,
                task["status"]
            ), tags=(tag,))

    def schedule_task(self, task):
        # Calculate and store the next_run time using croniter
        try:
            base = datetime.now()
            itr = croniter(task["cron_expr"], base)
            task["next_run"] = itr.get_next(datetime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            task["next_run"] = "-"
        self.save_tasks()
        self.update_task_list()

    def run_scheduler(self):
        while True:
            # Skip if a job is running
            if self.job_running_lock.locked():
                time.sleep(1)
                continue
            now = datetime.now()
            updated = False
            for task in self.tasks:
                # Only run if next_run is set and due
                next_run_str = task.get("next_run")
                print(f"[TaskRunner] Checking task: {task.get('name')} next_run={next_run_str} now={now.strftime('%Y-%m-%d %H:%M:%S')}")
                if not next_run_str or next_run_str == "-":
                    continue
                try:
                    next_run = datetime.strptime(next_run_str, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
                if now >= next_run:
                    # Run the task
                    self._run_scheduled_task(task)
                    # Update last_execution
                    task["last_execution"] = now.strftime("%Y-%m-%d %H:%M:%S")
                    # Calculate next_run
                    try:
                        itr = croniter(task["cron_expr"], now)
                        task["next_run"] = itr.get_next(datetime).strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        task["next_run"] = "-"
                    updated = True
            if updated:
                self.save_tasks()
                self.update_task_list()
            time.sleep(30)

    def _run_scheduled_task(self, task):
        def task_thread():
            if not self.job_running_lock.acquire(blocking=False):
                print(f"[TaskRunner] [SCHEDULED] Skipping: Another job is running.")
                return
            try:
                # Store original working directory
                original_dir = os.getcwd()
                # Get the directory of the task file
                task_dir = os.path.dirname(os.path.abspath(task['file_path']))
                # Change to that directory
                os.chdir(task_dir)
                
                ext = os.path.splitext(task['file_path'])[1].lower()
                if ext in ['.bat', '.cmd']:
                    cmd = f'cmd /c "{task["file_path"]}"'
                elif ext == '.py':
                    cmd = f'python "{task["file_path"]}"'
                else:
                    cmd = f'"{task["file_path"]}"'
                print(f"[TaskRunner] [SCHEDULED] Would run: {cmd}")
                logging.info(f"[SCHEDULED] Would run: {cmd}")
                # Run and capture output
                start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace')
                output, _ = proc.communicate()
                # Restore original working directory before writing log
                os.chdir(original_dir)
                # Log block with task name and file path
                log_block = (
                    f"==== EXECUTION {start_time} ====" + "\n"
                    f"Task: {task.get('name', '')}\n"
                    f"File: {task.get('file_path', '')}\n\n"
                    + output + "\n"
                )
                self._append_and_trim_log(task, log_block)
                logging.info(f"[SCHEDULED] Task completed: {task['file_path']}")
                print(f"[TaskRunner] [SCHEDULED] Completed: {task['file_path']}")
            except Exception as e:
                try:
                    os.chdir(original_dir)
                except:
                    pass
                logging.error(f"[SCHEDULED] Error running task {task['file_path']}: {str(e)}")
                print(f"[TaskRunner] [SCHEDULED] Error: {e}")
            finally:
                self.job_running_lock.release()
        threading.Thread(target=task_thread, daemon=True).start()

    def _append_and_trim_log(self, task, log_block):
        # Sanitize name for filename and convert to lowercase
        safe_name = re.sub(r'[^\w\-_]', '_', task.get('name', 'task')).lower()
        log_file = f"{safe_name}.log"
        abs_log_file = os.path.abspath(log_file)
        print(f"[TaskRunner] Writing log to: {log_file} (absolute: {abs_log_file})")
        # Append new block
        with open(log_file, 'a', encoding='utf-8', errors='replace') as f:
            f.write(log_block)
        # Trim to last N executions
        try:
            with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            blocks = content.split('==== EXECUTION ')
            if blocks[0].strip() == '':
                blocks = blocks[1:]
            else:
                blocks[0] = '==== EXECUTION ' + blocks[0]
            n = int(task.get('log_executions', 1))
            trimmed = ['==== EXECUTION ' + b for b in blocks[-n:]]
            with open(log_file, 'w', encoding='utf-8', errors='replace') as f:
                f.write(''.join(trimmed))
        except Exception as e:
            print(f"[TaskRunner] Error trimming log file: {e}")

    def load_tasks(self):
        try:
            with open("tasks.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return []

    def save_tasks(self):
        with open("tasks.json", "w") as f:
            json.dump(self.tasks, f)

    def on_tree_select(self, event):
        if getattr(self, 'clearing_form', False):
            return
        selected = self.task_tree.selection()
        if not selected:
            return
        index = self.task_tree.index(selected[0])
        task = self.tasks[index]
        self.name_var.set(task.get("name", os.path.basename(task["file_path"])))
        self.file_path.set(task["file_path"])
        self.cron_var.set(task["cron_expr"])
        self.log_retention.set(str(task["retention"]))
        self.log_executions.set(str(task.get("log_executions", 1)))
        self.selected_task_index = index
        self.add_update_button.config(text="Save Task")
        # Show action buttons
        self.run_btn.pack(side="left", padx=(0, 8))
        self.show_log_button.pack(side="left", padx=(0, 8))
        self.delete_btn.pack(side="left")

    def get_next_execution(self, task):
        # Calculate next execution time based on cron expression
        try:
            from croniter import croniter
        except ImportError:
            return "(croniter req.)"
        base = datetime.now()
        try:
            itr = croniter(task["cron_expr"], base)
            return itr.get_next(datetime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "-"

    def show_log(self):
        if self.selected_task_index is None:
            messagebox.showwarning("No Task Selected", "Please select a task to view its log.")
            return
        task = self.tasks[self.selected_task_index]
        # Find all log files that could match this task (by name or previous names)
        safe_name = re.sub(r'[^\w\-_]', '_', task.get('name', 'task')).lower()
        log_pattern = f"{safe_name}*.log"
        log_files = glob.glob(log_pattern)
        if not log_files:
            # Try to find any log file that contains the script filename
            script_base = os.path.splitext(os.path.basename(task['file_path']))[0]
            log_files = glob.glob(f"*{script_base}*.log")
        if not log_files:
            messagebox.showinfo("No Log", "No log file exists for this task yet.")
            return
        # If multiple logs, let user pick
        log_file = log_files[0]
        if len(log_files) > 1:
            import tkinter.simpledialog
            log_file = tkinter.simpledialog.askstring(
                "Multiple Logs Found",
                "Multiple log files found. Enter the number to view:\n" +
                "\n".join(f"{i+1}: {os.path.basename(f)}" for i, f in enumerate(log_files)),
                initialvalue="1"
            )
            try:
                log_file = log_files[int(log_file)-1]
            except Exception:
                messagebox.showerror("Error", "Invalid selection.")
                return
        # Show log content in a popup window
        try:
            with open(log_file, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            log_win = tk.Toplevel(self.root)
            log_win.title(f"Log: {os.path.basename(log_file)}")
            log_win.geometry("700x500")

            # Create a frame for proper layout
            frame = ttk.Frame(log_win)
            frame.pack(fill="both", expand=True)

            text = tk.Text(frame, wrap="none", font=("Consolas", 10))
            yscroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
            xscroll = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)

            text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
            text.grid(row=0, column=0, sticky="nsew")
            yscroll.grid(row=0, column=1, sticky="ns")
            xscroll.grid(row=1, column=0, sticky="ew")

            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)

            text.insert("1.0", content)
            text.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open log file: {e}")

    def run_selected_task(self):
        if self.selected_task_index is None:
            messagebox.showwarning("No Task Selected", "Please select a task to run.")
            return
        # Prevent manual run if a job is running
        if self.job_running_lock.locked():
            messagebox.showwarning("Job Running", "A task is already running. Please wait until it finishes.")
            return
        task = self.tasks[self.selected_task_index]
        def run_task_with_progress():
            self.job_running_lock.acquire()
            try:
                # Set last_execution to 'running...' and update UI
                task['last_execution'] = 'running...'
                self.update_task_list()
                # Print a log message similar to scheduled run
                now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                next_run_str = task.get('next_run', '-')
                print(f"[TaskRunner] [MANUAL] Checking task: {task.get('name')} next_run={next_run_str} now={now_str}")
                # Store original working directory
                original_dir = os.getcwd()
                # Get the directory of the task file
                task_dir = os.path.dirname(os.path.abspath(task['file_path']))
                # Change to that directory
                os.chdir(task_dir)
                ext = os.path.splitext(task['file_path'])[1].lower()
                if ext in ['.bat', '.cmd']:
                    cmd = f"cmd /c \"{task['file_path']}\""
                elif ext == '.py':
                    cmd = f"python \"{task['file_path']}\""
                else:
                    cmd = f"\"{task['file_path']}\""
                # Run and capture output
                start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace')
                output, _ = proc.communicate()
                # Restore original working directory before writing log
                os.chdir(original_dir)
                # Log block with task name and file path
                log_block = (
                    f"==== MANUAL EXECUTION {start_time} ====" + "\n"
                    f"Task: {task.get('name', '')}\n"
                    f"File: {task.get('file_path', '')}\n\n"
                    + output + "\n"
                )
                self._append_and_trim_log(task, log_block)
                # Update last_execution to completion time and update UI
                task['last_execution'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.update_task_list()
            except Exception as e:
                try:
                    os.chdir(original_dir)
                except:
                    pass
                task['last_execution'] = f'Error: {e}'
                self.update_task_list()
                messagebox.showerror("Error", f"Could not run task: {e}")
            finally:
                self.job_running_lock.release()
                self.save_tasks()
        threading.Thread(target=run_task_with_progress, daemon=True).start()

    def run(self):
        self.root.mainloop()

# --- First run logic ---
def ensure_startup_shortcut():
    print('[TaskRunner Startup] Checking if startup shortcut needs to be created...')
    # Only run on first start
    if os.path.exists(STARTUP_MARKER):
        print(f'[TaskRunner Startup] Marker file {STARTUP_MARKER} exists. Skipping shortcut creation.')
        return
    # Path to this script or batch file
    script_path = os.path.abspath(sys.argv[0])
    print(f'[TaskRunner Startup] Script path: {script_path}')
    # Use pythonw.exe for no console, or python.exe for console
    pythonw = sys.executable.replace('python.exe', 'pythonw.exe') if sys.executable.endswith('python.exe') else sys.executable
    print(f'[TaskRunner Startup] Pythonw path: {pythonw}')
    # Shortcut target
    target = f'"{pythonw}" "{script_path}" --minimized'
    print(f'[TaskRunner Startup] Shortcut target: {target}')
    # Name for the shortcut
    shortcut_name = 'TaskRunner.lnk'
    # Windows Startup folder
    startup_dir = os.path.join(os.environ['APPDATA'], r'Microsoft\Windows\Start Menu\Programs\Startup')
    print(f'[TaskRunner Startup] Startup folder: {startup_dir}')
    shortcut_path = os.path.join(startup_dir, shortcut_name)
    print(f'[TaskRunner Startup] Shortcut path: {shortcut_path}')
    try:
        import pythoncom
        from win32com.shell import shell, shellcon
        from win32com.client import Dispatch
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = pythonw
        shortcut.Arguments = f'"{script_path}" --minimized'
        shortcut.WorkingDirectory = os.path.dirname(script_path)
        shortcut.IconLocation = script_path
        shortcut.save()
        print(f'[TaskRunner Startup] Shortcut created successfully at {shortcut_path}')
    except Exception as e:
        print(f'[TaskRunner Startup] Failed to create shortcut, falling back to batch file. Error: {e}')
        # Fallback: create a batch file in Startup
        batch_path = os.path.join(startup_dir, 'TaskRunnerStartup.bat')
        with open(batch_path, 'w') as f:
            f.write(f'start "" {target}\n')
        print(f'[TaskRunner Startup] Batch file created at {batch_path}')
    with open(STARTUP_MARKER, 'w') as f:
        f.write('This file marks that Task Runner has set up startup shortcut.')

# --- Main entry point ---
if __name__ == "__main__":
    ensure_startup_shortcut()
    app = TaskRunner()
    app.run() 