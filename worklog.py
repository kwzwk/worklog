#!/usr/bin/env python3
"""
Work-log timer — a background system-tray utility for logging what you work on.

Runs quietly as a tray icon (down by the clock). Every INTERVAL_MINUTES a popup
asks "What have you worked on?", lets you tick the project(s) and task(s) you
worked on, and records the time in a monthly CSV. Right-click the tray icon to
log on demand, pause, manage tasks, create a monthly report, or quit.

REPORTS: "Create report…" builds an Excel workbook for a month: totals for the
1st–15th, the 16th–end and the whole month, hours per project and per task, and
a sheet per day and per week.

DAILY LOG (optional, see WRITE_DAILY_LOG): a readable Markdown .txt diary per
day with a work-day summary at the end.
    Entry format:   * HH:MM–HH:MM - Project Name - Entry text

NEW DAY: when the date changes, the day is closed automatically at your last
activity (so a forgotten evening isn't counted as work) and a new day's log is
started at your first activity the next morning.

AWAY: after IDLE_MINUTES without keyboard/mouse input you count as away (no
prompts). When you're back you're asked whether the time was a break.

TASKS: a to-do list kept in tasks.csv. Add, edit and complete tasks from the
tray's "Tasks" menu; open tasks appear in every prompt so you can tick the ones
you worked on (their time is tracked) or mark them done. Open tasks stay in the
list from day to day until they're done.

PATHS: everything lives next to the program itself (the .py file, or the .exe if
you build one) — no hard-coded user path. Monthly CSVs, reports and daily logs
(if switched on) go in a "log" subfolder; projects.txt and tasks.csv sit in the
main folder next to the program.

REQUIREMENTS (beyond the Python standard library):
    pip install pystray pillow openpyxl
All bundle fine into a PyInstaller .exe. Build it WINDOWLESS (--noconsole) —
see build_exe.bat.
"""

import csv
import os
import queue
import sys
import threading
from datetime import datetime, timedelta, time as dtime

# ── Make print() safe when running windowless ───────────────────────────────
# Under PyInstaller --noconsole there is no console, so sys.stdout/stderr are
# None and a stray print() would crash. Route them to a harmless sink.
class _NullWriter:
    def write(self, *_a, **_k): pass
    def flush(self, *_a, **_k): pass

if sys.stdout is None:
    sys.stdout = _NullWriter()
if sys.stderr is None:
    sys.stderr = _NullWriter()

# ── GUI (standard library) ───────────────────────────────────────────────────
import tkinter as tk
from tkinter import ttk, messagebox

# ── Tray icon (third-party) ──────────────────────────────────────────────────
try:
    import pystray
    from pystray import Menu, MenuItem
    from PIL import Image, ImageDraw
except Exception as _import_err:
    # Show the REAL error and WHICH Python is running. If the packages look
    # installed but you still land here, it almost always means this
    # interpreter isn't the one you pip-installed into — compare the path below
    # with the output of `where python` / your pip's target.
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror(
        "Work-log — import failed",
        "Couldn't import pystray / Pillow.\n\n"
        f"Actual error:\n    {type(_import_err).__name__}: {_import_err}\n\n"
        f"Python running this script:\n    {sys.executable}\n\n"
        f"Version:\n    {sys.version.splitlines()[0]}\n\n"
        "Fix: install into THIS interpreter, e.g.\n"
        f'    "{sys.executable}" -m pip install pystray pillow',
    )
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# Where the program lives
# ─────────────────────────────────────────────────────────────────────────────

# BASE_DIR = the folder containing this program. When frozen by PyInstaller we
# use the .exe's folder; otherwise the .py file's folder. Everything is relative
# to this, so you can move the whole folder anywhere and it still works.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS — tweak these
# ─────────────────────────────────────────────────────────────────────────────

# How often to automatically show the prompt (in minutes). This is the DEFAULT
# at startup; you can change it live from the tray's "Interval" submenu. Set it
# to a small number (e.g. 1) while testing.
INTERVAL_MINUTES = 30

# The choices offered in the tray "Interval" submenu (minutes).
INTERVAL_CHOICES = [15, 30, 60]

# If True, a MANUAL log (tray "Log now") restarts the countdown — the next auto
# prompt becomes INTERVAL_MINUTES from when you finish logging. If False, manual
# logs leave the regular schedule untouched.
RESET_TIMER_ON_MANUAL_LOG = True

# Minutes without keyboard/mouse input after which you count as away. While
# away there are no prompts; when you're back you're asked whether the time
# was a break. (Uses Windows' last-input time; elsewhere it's switched off.)
IDLE_MINUTES = 30

# Monthly CSVs, reports and daily logs go in a "log" subfolder next to the
# program.
LOG_DIR = os.path.join(BASE_DIR, "log")

# Also write a readable Markdown diary per day (worklog_YYYY-MM-DD.txt) with a
# work-day summary at the end. Everything is in the monthly CSV and the report
# either way.
WRITE_DAILY_LOG = False

# Project list, in the main folder next to the program. One project per line.
# Blank lines and lines starting with '#' are ignored. Re-read live — edits take
# effect at the next prompt, no restart needed.
PROJECTS_FILE = os.path.join(BASE_DIR, "projects.txt")

# Task list, in the main folder next to the program. Managed from the tray's
# "Tasks" menu (the app keeps it up to date — no need to edit it by hand).
TASKS_FILE = os.path.join(BASE_DIR, "tasks.csv")
TASK_HEADERS = ["ID", "Created", "Project", "Task", "Due", "Status", "Done",
                "Minutes"]

# The question you'll be asked.
PROMPT_QUESTION = "What have you worked on?"

# How often the main loop checks the clock / tray commands (milliseconds).
POLL_MS = 250

# ── CSV output options (for opening in Excel and importing to SharePoint) ─────
# Delimiter between fields. "," is the universal / SharePoint-friendly default.
# If you ONLY care about double-clicking the file open in a European Excel
# (German, French, etc., where Excel expects ";"), set this to ";". Note that a
# CSV-parsing Power Automate flow would then need to use the same delimiter;
# SharePoint's "From Excel" import is unaffected either way.
CSV_DELIMITER = ";"

# Write a UTF-8 byte-order mark at the start of each new CSV. This makes Excel
# reliably detect UTF-8 so accented characters (ä ö ü ß, etc.) render correctly
# on double-click. Leave True unless a downstream tool dislikes the BOM.
CSV_WRITE_BOM = True

# Decimal separator for the Hours column. "," suits a German/European Excel
# (which would read 7.5 as a date); use "." for an English Excel.
CSV_DECIMAL = ","

# CSV column headers. These become the SharePoint list's column names on import,
# so they're kept clean (no spaces/punctuation) for tidy internal names.
# One row per project/task share of each logged block, so summing Hours by
# any column (project, task, week, ...) always adds up.
CSV_HEADERS = ["Date", "Week", "Weekday", "Start", "End", "Hours", "Type",
               "Project", "Task_ID", "Task", "Note"]

# Label used for time worked without a project ticked.
NONPROJECT = "(non-project work)"


# ─────────────────────────────────────────────────────────────────────────────
# Internal state
# ─────────────────────────────────────────────────────────────────────────────

LOG_PATH = None             # full path to today's markdown file
entries_count = 0            # how many entries logged this session
root = None                  # hidden Tk root window
icon = None                  # the tray icon

session_start = None
session_active = False       # False between a day's close and your first
                             # activity of the next day
next_due = None              # datetime of the next scheduled prompt
interval = None              # timedelta of the current interval
interval_minutes = None      # current interval in minutes (changeable at runtime)
paused = False               # auto-prompts suspended while True
dialog_open = False          # guard so prompts never stack
active_dialog = None         # the open prompt / away dialog, so it can be
                             # closed when you go away, at midnight or on quit
draft_text = ""              # typed text kept when an open prompt was closed
quitting = False
showing_error = False        # guard: one error message at a time
rolling_over = False         # guard: closing the day is in progress

# Away detection.
last_active = None           # last moment you were known to be at the PC
away_since = None            # start of the current away period, or None
last_poll = None             # previous poll time, to notice the PC sleeping

# Commands from the tray thread → main (Tk) thread. The tray menu callbacks run
# on pystray's own thread, so they must NOT touch Tk directly; they enqueue a
# command here and the main-thread poll() handles it.
cmd_q = queue.Queue()

# Time accounting (for the end-of-day summary). Time falls into one of three
# buckets: project work (per named project), non-project work (worked, but no
# project ticked), and breaks (explicitly marked not-work). "Worked" time =
# project + non-project; breaks are excluded from it.
project_minutes = {}         # project name -> minutes
nonproject_minutes = 0.0     # worked, but no project ticked
break_minutes = 0.0          # explicitly marked as a break (not work)
last_entry_time = None       # start of the current (uncredited) block

# Tasks, this session (the running totals live in tasks.csv).
task_minutes = {}            # task ID -> minutes
tasks_done_today = []        # task IDs completed this session
task_names = {}              # task ID -> label, so the summary can still
                             # name a task deleted during the day
task_window = None           # the "Manage tasks" window, if open

# "Same task as before" memory.
last_projects = []
last_answer = ""
last_tasks = []


# ─────────────────────────────────────────────────────────────────────────────
# Small helpers
# ─────────────────────────────────────────────────────────────────────────────

def now():
    return datetime.now()


def hhmm(dt=None):
    return (dt or now()).strftime("%H:%M")


def fmt_hm(minutes):
    m = int(round(minutes))
    return f"{m // 60}h {m % 60:02d}m"


def log_path_for(dt):
    return os.path.join(LOG_DIR, "worklog_" + dt.strftime("%Y-%m-%d") + ".txt")


def csv_path_for(dt):
    """Monthly CSV path, e.g. ...\\log\\worklog_2026-10.csv"""
    return os.path.join(LOG_DIR, "worklog_" + dt.strftime("%Y-%m") + ".csv")


def append_to_log(text):
    """Append to today's daily log (only if WRITE_DAILY_LOG is on)."""
    if not WRITE_DAILY_LOG:
        return
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(text)


def save_with_retry(what, write, *args):
    """Run write(*args). If the file can't be written (most often the CSV is
    open in Excel, which locks it), ask to close it and retry instead of
    failing. Returns True if it was saved."""
    while True:
        try:
            write(*args)
            return True
        except OSError as e:
            if not messagebox.askretrycancel(
                    "Work-log",
                    f"Couldn't save the {what}:\n{e}\n\n"
                    "If the file is open (e.g. the CSV in Excel), close it "
                    "and click Retry."):
                return False


def load_projects():
    try:
        with open(PROJECTS_FILE, "r", encoding="utf-8-sig", errors="replace") as f:
            lines = f.readlines()
    except FileNotFoundError:
        return []
    return [ln.strip() for ln in lines
            if ln.strip() and not ln.strip().startswith("#")]


def open_path(path):
    """Open a file or folder in Explorer / default app. Windows only."""
    try:
        os.startfile(path)          # noqa: provided on Windows
    except Exception as e:
        try:
            messagebox.showwarning("Work-log", f"Couldn't open:\n{path}\n\n{e}")
        except Exception:
            pass


def idle_seconds():
    """Seconds since the last keyboard/mouse input. Windows only; elsewhere
    returns 0, which switches away detection off."""
    if os.name != "nt":
        return 0.0
    try:
        import ctypes

        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

        info = LASTINPUTINFO()
        info.cbSize = ctypes.sizeof(info)
        if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info)):
            return 0.0
        # Both are millisecond tick counts that wrap around every ~49 days.
        ticks = ctypes.windll.kernel32.GetTickCount()
        return ((ticks - info.dwTime) & 0xFFFFFFFF) / 1000.0
    except Exception:
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Tasks (tasks.csv)
# ─────────────────────────────────────────────────────────────────────────────

def load_tasks():
    """All tasks, as dicts with every TASK_HEADERS key."""
    try:
        with open(TASKS_FILE, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f, delimiter=CSV_DELIMITER))
    except FileNotFoundError:
        return []
    tasks = []
    for r in rows:
        t = {h: (r.get(h) or "").strip() for h in TASK_HEADERS}
        if t["ID"] and t["Task"]:
            tasks.append(t)
    return tasks


def save_tasks(tasks):
    """Write the whole task list. Goes through a temp file so tasks.csv is
    never left half-written."""
    tmp = TASKS_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            if CSV_WRITE_BOM:
                f.write("﻿")
            writer = csv.DictWriter(f, fieldnames=TASK_HEADERS,
                                    delimiter=CSV_DELIMITER)
            writer.writeheader()
            writer.writerows(tasks)
        os.replace(tmp, TASKS_FILE)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def change_tasks(change):
    """Re-read the task list, apply change(tasks) and save it. Returns
    (saved, whatever change() returned)."""
    tasks = load_tasks()
    result = change(tasks)
    return save_with_retry("task list", save_tasks, tasks), result


def open_task_list(tasks=None):
    """Open tasks, soonest due first (no due date last)."""
    if tasks is None:
        tasks = load_tasks()
    return sorted((t for t in tasks if t["Status"] != "done"),
                  key=lambda t: (t["Due"] or "9999-99-99", t["ID"]))


def next_task_id(tasks):
    nums = [int(t["ID"][2:]) for t in tasks
            if t["ID"].startswith("T-") and t["ID"][2:].isdigit()]
    return f"T-{max(nums, default=0) + 1:03d}"


def task_label(t):
    label = t["ID"] + " " + t["Task"]
    if t["Project"]:
        label += f" [{t['Project']}]"
    task_names[t["ID"]] = label
    return label


def task_total_minutes(t):
    try:
        return int(float(t["Minutes"] or 0))
    except ValueError:
        return 0


def log_task_event(event, t, extra=""):
    """Record a task change (ADDED, DONE, ...) in the daily log. (tasks.csv
    itself keeps each task's created and done dates.)"""
    ensure_session()
    save_with_retry("daily log", append_to_log,
                    f"* {hhmm()} - TASK {event} - {task_label(t)}{extra}\n")


# ─────────────────────────────────────────────────────────────────────────────
# Writing to the markdown log
# ─────────────────────────────────────────────────────────────────────────────

def write_session_header():
    file_existed = os.path.exists(LOG_PATH)
    chunk = ""
    if not file_existed:
        chunk += "# Worklog — " + session_start.strftime("%A, %Y-%m-%d") + "\n\n"
    else:
        chunk += "\n"
    chunk += "**Session started:** " + hhmm(session_start) + "\n\n"
    open_count = len(open_task_list())
    if open_count:
        chunk += f"Open tasks: {open_count}\n\n"
    append_to_log(chunk)


def write_entry(start, end, answer_text, projects, kind, tasks_text=""):
    """Write one markdown line. `kind` is 'work', 'break', or 'skip'.
    Label is the project list, 'BREAK' for a break, or '(non-project work)' for
    worked-but-untracked time (which is also what a Skip with no projects logs).
    Tasks worked on go on an indented line below."""
    span = hhmm(start) + "–" + hhmm(end)
    if kind == "break":
        label = "BREAK"
    elif projects:
        label = ", ".join(projects)
    else:
        label = "(non-project work)"
    lines = [ln.rstrip() for ln in answer_text.splitlines()]
    while lines and lines[-1] == "":
        lines.pop()
    if not lines:
        entry = "(no entry)"
    else:
        entry = lines[0]
        for extra in lines[1:]:
            entry += "\n  " + extra
    if tasks_text:
        entry += "\n  Tasks: " + tasks_text
    append_to_log("* " + span + " - " + label + " - " + entry + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# Splitting a block's time, and writing it to the monthly CSV
# ─────────────────────────────────────────────────────────────────────────────

def allocate(minutes, projects, tasks):
    """Split a block's minutes into (project, task or None, minutes) shares.

    Every ticked task, and every ticked project without a ticked task of its
    own, gets an equal share. A task's share goes to its project when that
    project is ticked; otherwise (no project, or you unticked it because the
    task was done for another project) it goes to the ticked projects that have
    no task of their own (or to all ticked projects), or to non-project work if
    no project is ticked.

    Example, 60 min: task T-001 [Alpha] + projects Alpha and Beta ticked
    → Alpha 30 (T-001), Beta 30.
    """
    projects = list(projects)
    own = [t for t in tasks if t["Project"] and t["Project"] in projects]
    floating = [t for t in tasks if t not in own]
    have_task = {t["Project"] for t in own}
    bare = [p for p in projects if p not in have_task]

    units = [([t["Project"]], t) for t in own]
    if floating:
        targets = bare or projects or [NONPROJECT]
        units += [(targets, t) for t in floating]
    else:
        units += [([p], None) for p in bare]
    if not units:
        units = [([NONPROJECT], None)]

    share = minutes / len(units)
    return [(p, t, share / len(targets))
            for targets, t in units for p in targets]


def fmt_hours(minutes):
    """Hours with 2 decimals and the configured decimal separator."""
    return f"{minutes / 60:.2f}".replace(".", CSV_DECIMAL)


def write_csv_rows(start, end, answer_text, kind, shares):
    """Append a block to the monthly CSV, optimized for Excel and SharePoint.

    One row per (project, task) share from allocate(), so summing Hours by any
    column adds up. Type is Work, Non-project or Break.
    - Date is ISO YYYY-MM-DD (the format SharePoint date columns expect).
    - Week is the ISO week (e.g. 2026-W40); Weekday is Mon…Sun.
    - The csv module quotes any field containing the delimiter, quotes or newlines
      (RFC 4180), so semicolons in your notes never break the columns.
    """
    path = csv_path_for(start)
    file_existed = os.path.exists(path)

    # Flatten the note to a single line (no embedded newlines) for clean cells.
    note = answer_text.replace("\r\n", "\n").replace("\n", " | ").strip()
    iso_year, iso_week, _ = start.isocalendar()
    fixed = [start.strftime("%Y-%m-%d"), f"{iso_year}-W{iso_week:02d}",
             start.strftime("%a"), hhmm(start), hhmm(end)]

    # newline="" prevents blank rows; csv.writer then emits RFC-4180 \r\n.
    with open(path, "a", encoding="utf-8", newline="") as f:
        if not file_existed and CSV_WRITE_BOM:
            f.write("\ufeff")                       # UTF-8 BOM → Excel encoding
        writer = csv.writer(f, delimiter=CSV_DELIMITER,
                            quoting=csv.QUOTE_MINIMAL)
        if not file_existed:
            writer.writerow(CSV_HEADERS)
        for project, task, minutes in shares:
            if kind == "break":
                row_type, project = "Break", ""
            elif project == NONPROJECT:
                row_type = "Non-project"
            else:
                row_type = "Work"
            writer.writerow(fixed + [
                fmt_hours(minutes), row_type, project,
                task["ID"] if task else "", task["Task"] if task else "",
                note])


# ─────────────────────────────────────────────────────────────────────────────
# Dialogs (Tkinter) — always run on the main thread
# ─────────────────────────────────────────────────────────────────────────────

def run_dialog(dlg, cancel, focus=None):
    """Centre dlg on screen, bring it to the front and wait until it closes.
    While open it's the active dialog, which close_active_dialog() can close
    by calling cancel()."""
    global active_dialog
    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw - w) // 2}+{(sh - h) // 3}")
    dlg.focus_force()
    if focus is not None:
        focus.focus_set()
    active_dialog = {"dlg": dlg, "cancel": cancel}
    try:
        root.wait_window(dlg)
    finally:
        active_dialog = None


def close_active_dialog():
    """Close an open prompt / away dialog without logging anything."""
    if active_dialog is not None:
        try:
            active_dialog["cancel"]()
        except Exception:
            pass


def raise_active_dialog():
    if active_dialog is not None:
        try:
            active_dialog["dlg"].lift()
            active_dialog["dlg"].focus_force()
        except Exception:
            pass


def scroll_box(parent, rows, max_rows=8):
    """A frame to put rows of check boxes in; it scrolls when there are more
    than max_rows."""
    outer = tk.Frame(parent)
    outer.pack(anchor="w", padx=22, pady=(2, 10), fill="x")
    if rows <= max_rows:
        return outer
    canvas = tk.Canvas(outer, height=max_rows * 26, highlightthickness=0)
    bar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(
        scrollregion=canvas.bbox("all"), width=inner.winfo_reqwidth()))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=bar.set)
    canvas.pack(side="left", fill="both", expand=True)
    bar.pack(side="right", fill="y")
    return inner


def show_prompt_dialog(reason, end=None):
    """Question + multi-line text box + project and task tick boxes.
    `end` fixes the end of the block (e.g. when you went on a break); by
    default the block ends when you answer.
    Returns (answer_text, selected_projects, kind, task_ids, done_ids) where
    kind is:
      'work'   — Save: log what's typed/ticked
      'break'  — Break: time off
      'same'   — Same as before: reuse the last work task (handled by caller)
      'skip'   — Skip: nothing to repeat, just dismiss (logged as non-project)
      'cancel' — closed by the program (you went away, a new day, quit):
                 nothing is logged
    """
    global draft_text
    projects = load_projects()
    tasks = open_task_list()
    has_previous = bool(last_projects or last_answer or last_tasks)

    dlg = tk.Toplevel(root)
    dlg.title("Work-log")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.lift()

    shown_end = end or now()
    if reason == "break":
        question = (f"What did you work on before your break "
                    f"({hhmm(last_entry_time)}–{hhmm(shown_end)})?")
    elif reason == "day_end":
        question = (f"{PROMPT_QUESTION}   (end of day · "
                    f"{hhmm(last_entry_time)}–{hhmm(shown_end)})")
    else:
        tag = "auto" if reason == "scheduled" else "manual"
        question = f"{PROMPT_QUESTION}   ({tag} · {hhmm(shown_end)})"
    tk.Label(dlg, text=question,
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 6))

    txt = tk.Text(dlg, width=58, height=4, wrap="word", font=("Segoe UI", 10))
    txt.pack(padx=14, pady=(0, 6), fill="x")
    if draft_text:
        txt.insert("1.0", draft_text)     # what you'd typed in a closed prompt
        draft_text = ""

    vars_by_name = {}
    if projects:
        tk.Label(dlg, text="Project(s) worked on:"
                 ).pack(anchor="w", padx=14)
        box = tk.Frame(dlg)
        box.pack(anchor="w", padx=22, pady=(2, 10))
        for name in projects:
            var = tk.BooleanVar(value=False)
            vars_by_name[name] = var
            ttk.Checkbutton(box, text=name, variable=var).pack(anchor="w")
    else:
        tk.Label(dlg,
                 text="(no projects found — add a projects.txt next to the "
                      "program,\n one project per line)",
                 fg="grey", justify="left").pack(anchor="w", padx=14,
                                                  pady=(0, 10))

    # Open tasks: tick the ones you worked on (their time is tracked, and their
    # project gets ticked too — untick it if the task was for another project)
    # and/or mark them done.
    task_vars = {}
    if tasks:
        tk.Label(dlg, text="Task(s) worked on:"
                 ).pack(anchor="w", padx=14)
        box = scroll_box(dlg, len(tasks))
        today = now().strftime("%Y-%m-%d")
        for t in tasks:
            row = tk.Frame(box)
            row.pack(anchor="w", fill="x")
            worked = tk.BooleanVar(value=False)
            done = tk.BooleanVar(value=False)
            task_vars[t["ID"]] = (worked, done)
            text = task_label(t)
            if t["Due"]:
                text += f"  (due {t['Due']})"
            ttk.Checkbutton(row, text=text, variable=worked,
                            style="Overdue.TCheckbutton"
                            if t["Due"] and t["Due"] < today else
                            "TCheckbutton").pack(side="left")
            ttk.Checkbutton(row, text="done", variable=done
                            ).pack(side="right", padx=(12, 0))

    # Live preview of how the time will be recorded (see allocate()), and a
    # hint. "covers the last N min" is the REAL gap since the previous entry
    # (so a manual log shows the true elapsed time, not the fixed interval).
    elapsed_min = max(0, int(round(
        (shown_end - last_entry_time).total_seconds() / 60)))
    split_label = tk.Label(dlg, text="", justify="left", wraplength=520)
    split_label.pack(anchor="w", padx=14, pady=(0, 2))
    tk.Label(dlg,
             text=f"Tip: covers the last {elapsed_min} min · "
                  "no project ticked = non-project work · "
                  "use Break for time off.",
             fg="grey").pack(anchor="w", padx=14, pady=(0, 6))

    def update_preview(*_):
        ticked = [n for n, v in vars_by_name.items() if v.get()]
        ticked_tasks = [t for t in tasks if task_vars[t["ID"]][0].get()]
        parts = []
        for project, t, minutes in allocate(elapsed_min, ticked, ticked_tasks):
            part = f"{project} {fmt_hm(minutes)}"
            if t is not None:
                part += f" ({t['ID']})"
            parts.append(part)
        split_label.config(text="Recorded as: " + " · ".join(parts))

    def task_ticked(t, worked):
        # Ticking a task ticks its project (you can untick it again).
        if worked.get() and t["Project"] in vars_by_name:
            vars_by_name[t["Project"]].set(True)
        update_preview()

    for var in vars_by_name.values():
        var.trace_add("write", update_preview)
    for t in tasks:
        worked = task_vars[t["ID"]][0]
        worked.trace_add("write",
                         lambda *_a, t=t, worked=worked: task_ticked(t, worked))
    update_preview()

    result = {"answer": "", "projects": [], "kind": "work",
              "tasks": [], "done": []}

    def do_save():
        result["answer"] = txt.get("1.0", "end").strip()
        result["projects"] = [n for n, v in vars_by_name.items() if v.get()]
        result["tasks"] = [i for i, (w, _d) in task_vars.items() if w.get()]
        result["done"] = [i for i, (_w, d) in task_vars.items() if d.get()]
        result["kind"] = "work"
        dlg.destroy()

    def do_break():
        # A break: record any typed note, but ignore project/task ticks.
        result["answer"] = txt.get("1.0", "end").strip()
        result["projects"] = []
        result["kind"] = "break"
        dlg.destroy()

    def do_same_or_skip():
        # If there's a previous work task, one click repeats it. Otherwise this
        # is a plain Skip (dismiss → logged as non-project work with no entry).
        result["kind"] = "same" if has_previous else "skip"
        dlg.destroy()

    def do_cancel():
        # Closed by the program: keep what was typed for the next prompt.
        global draft_text
        draft_text = txt.get("1.0", "end").strip()
        result["kind"] = "cancel"
        dlg.destroy()

    # Third button is adaptive: "Same as before (…)" when there's something to
    # repeat, otherwise "Skip".
    if has_previous:
        if last_projects:
            preview = ", ".join(last_projects)
        elif last_tasks:
            preview = ", ".join(last_tasks)
        else:
            preview = "previous note"
        third_label = f"↺ Same as before ({preview})"
    else:
        third_label = "Skip"

    # Buttons, left → right: [Skip / Same as before] · Break · Save.
    btns = tk.Frame(dlg)
    btns.pack(padx=14, pady=(0, 14), anchor="e")
    ttk.Button(btns, text=third_label,
               command=do_same_or_skip).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="Break", command=do_break).pack(side="left", padx=(0, 6))
    ttk.Button(btns, text="Save", command=do_save).pack(side="left")

    dlg.bind("<Control-Return>", lambda e: do_save())       # Ctrl+Enter = Save
    dlg.bind("<Alt-b>", lambda e: do_break())               # Alt+B = Break
    dlg.bind("<Alt-s>", lambda e: do_same_or_skip())        # Alt+S = Same/Skip
    dlg.bind("<Escape>", lambda e: do_same_or_skip())       # Esc   = Same/Skip
    dlg.protocol("WM_DELETE_WINDOW", do_same_or_skip)

    run_dialog(dlg, do_cancel, focus=txt)
    return (result["answer"], result["projects"], result["kind"],
            result["tasks"], result["done"])


def show_away_dialog(start, back):
    """Ask how the time you were away should count.
    Returns 'break', 'work' or 'cancel'."""
    dlg = tk.Toplevel(root)
    dlg.title("Work-log")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    minutes = int(round((back - start).total_seconds() / 60))
    tk.Label(dlg, text="Welcome back!",
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 4))
    tk.Label(dlg, text=f"You were away from {hhmm(start)} to {hhmm(back)} "
                       f"({fmt_hm(minutes)}).\nHow should this time count?",
             justify="left").pack(anchor="w", padx=14, pady=(0, 10))
    result = {"choice": "work"}

    def choose(choice):
        result["choice"] = choice
        dlg.destroy()

    btns = tk.Frame(dlg)
    btns.pack(padx=14, pady=(0, 14), anchor="e")
    ttk.Button(btns, text="It was work (e.g. a meeting)",
               command=lambda: choose("work")).pack(side="left", padx=(0, 6))
    brk = ttk.Button(btns, text="Break", command=lambda: choose("break"))
    brk.pack(side="left")
    dlg.bind("<Return>", lambda e: choose("break"))
    dlg.bind("<Escape>", lambda e: choose("work"))
    dlg.protocol("WM_DELETE_WINDOW", lambda: choose("work"))

    run_dialog(dlg, lambda: choose("cancel"), focus=brk)
    return result["choice"]


# ─────────────────────────────────────────────────────────────────────────────
# Logging a block of time
# ─────────────────────────────────────────────────────────────────────────────

def do_prompt(reason, end=None):
    """Show the popup, account the time into the right bucket, write to markdown
    + CSV. Guarded so two prompts never overlap. Returns True if something was
    logged."""
    global last_projects, last_answer, last_tasks, dialog_open

    if dialog_open:
        return False
    ensure_session()
    dialog_open = True
    try:
        answer, projects, kind, task_ids, done_ids = \
            show_prompt_dialog(reason, end)
    finally:
        dialog_open = False
    if kind == "cancel":
        return False

    # "Same as before" → reuse the remembered last work task, then treat it as
    # an ordinary work entry from here on.
    if kind == "same":
        projects = list(last_projects)
        answer = last_answer
        still_open = {t["ID"] for t in open_task_list()}
        task_ids = [i for i in last_tasks if i in still_open]
        done_ids = []
        kind = "work"

    log_block(last_entry_time, end or now(), answer, projects, kind,
              task_ids, done_ids)

    # Remember the last real WORK task for "Same as before" (not breaks).
    if kind == "work" and (projects or answer.strip() or task_ids):
        last_projects = list(projects)
        last_answer = answer
        last_tasks = list(task_ids)
    return True


def log_block(start, end, answer, projects, kind, task_ids=(), done_ids=()):
    """Account the time from start to end, update the tasks worked on / done,
    and write the entry to the monthly CSV (and the daily log, if on)."""
    global entries_count, last_entry_time, nonproject_minutes, break_minutes

    end = max(end, start)
    delta_min = (end - start).total_seconds() / 60.0

    # Split the block into project/task shares (see allocate()), add each
    # task's share to its running total in tasks.csv, and mark the ones ticked
    # "done".
    worked, done = [], []
    if kind == "work" and (task_ids or done_ids):
        def change(tasks):
            by_id = {t["ID"]: t for t in tasks}
            w = [by_id[i] for i in task_ids if i in by_id]
            d = [by_id[i] for i in done_ids
                 if i in by_id and by_id[i]["Status"] != "done"]
            shares = allocate(delta_min, projects, w)
            for _p, t, minutes in shares:
                if t is not None:
                    t["Minutes"] = str(task_total_minutes(t)
                                       + int(round(minutes)))
                    task_minutes[t["ID"]] = (task_minutes.get(t["ID"], 0.0)
                                             + minutes)
            for t in d:
                t["Status"] = "done"
                t["Done"] = end.strftime("%Y-%m-%d")
                tasks_done_today.append(t["ID"])
            return w, d, shares
        _saved, (worked, done, shares) = change_tasks(change)
    elif kind == "break":
        shares = [("", None, delta_min)]
    else:
        shares = allocate(delta_min, projects, [])

    for project, _t, minutes in shares:
        if kind == "break":
            break_minutes += minutes
        elif project == NONPROJECT:
            nonproject_minutes += minutes      # worked, just not on a project
        else:
            project_minutes[project] = project_minutes.get(project, 0.0) + minutes
    last_entry_time = end

    labels = [task_label(t) + (" ✓" if t in done else "") for t in worked]
    labels += [task_label(t) + " ✓" for t in done if t not in worked]
    tasks_text = ", ".join(labels)
    log_projects = []
    for project, _t, _m in shares:
        if project and project != NONPROJECT and project not in log_projects:
            log_projects.append(project)

    saved_csv = save_with_retry("monthly CSV", write_csv_rows,
                                start, end, answer, kind, shares)
    save_with_retry("daily log", write_entry,
                    start, end, answer, log_projects, kind, tasks_text)
    if not saved_csv:
        # Don't lose what was typed. (Ctrl+C copies a message box's text.)
        messagebox.showwarning(
            "Work-log", "This entry was NOT saved:\n\n"
                        f"{hhmm(start)}\u2013{hhmm(end)}  "
                        f"{', '.join(log_projects)}  {answer}")
    for t in done:
        log_task_event("DONE", t, f" (total {fmt_hm(task_total_minutes(t))})")
    refresh_task_window()

    entries_count += 1


# ─────────────────────────────────────────────────────────────────────────────
# Sessions: one per day
# ─────────────────────────────────────────────────────────────────────────────

def start_session(start):
    """Start a day's session: its own log file, fresh totals and countdown."""
    global LOG_PATH, session_start, session_active, last_entry_time, last_active
    global nonproject_minutes, break_minutes, entries_count, next_due, away_since
    global last_poll

    LOG_PATH = log_path_for(start)
    session_start = last_entry_time = last_active = last_poll = start
    project_minutes.clear()
    task_minutes.clear()
    del tasks_done_today[:]
    nonproject_minutes = 0.0
    break_minutes = 0.0
    entries_count = 0
    away_since = None
    next_due = start + interval
    session_active = True
    save_with_retry("daily log", write_session_header)


def ensure_session():
    """Start a new day's session if there isn't one (e.g. you log or add a
    task before the first automatic start of the day)."""
    if not session_active and not quitting:
        start_session(now())


def build_summary(end):
    present_min = (end - session_start).total_seconds() / 60.0
    worked_min = sum(project_minutes.values()) + nonproject_minutes
    lines = ["## Summary", "",
             f"* Present: {fmt_hm(present_min)} "
             f"({hhmm(session_start)}–{hhmm(end)})",
             f"* Worked (excl. breaks): {fmt_hm(worked_min)}",
             f"* Breaks: {fmt_hm(break_minutes)}",
             f"* Entries logged: {entries_count}",
             "* Hours:"]
    # Named projects, biggest first, then non-project work as a catch-all line.
    rows = [(name, project_minutes[name])
            for name in sorted(project_minutes, key=project_minutes.get,
                               reverse=True)]
    if nonproject_minutes > 0.05:
        rows.append(("(non-project work)", nonproject_minutes))
    if rows:
        for name, mins in rows:
            lines.append(f"    * {name}: {fmt_hm(mins)}")
    else:
        lines.append("    * (none recorded)")
    # Tasks: time today (biggest first) and which were finished.
    tasks = load_tasks()
    by_id = {t["ID"]: t for t in tasks}
    task_ids = sorted(set(task_minutes) | set(tasks_done_today),
                      key=lambda i: -task_minutes.get(i, 0.0))
    if task_ids:
        lines.append("* Tasks:")
        for tid in task_ids:
            name = (task_label(by_id[tid]) if tid in by_id
                    else task_names.get(tid, tid))
            done = "  ✓ done" if tid in tasks_done_today else ""
            lines.append(f"    * {name}: {fmt_hm(task_minutes.get(tid, 0.0))}"
                         f"{done}")
    lines.append(f"* Open tasks: {len(open_task_list(tasks))}")
    return lines


def end_session(end, note=""):
    """Write the 'Session ended' line and the work-day summary. Time since the
    last entry is treated as non-project work (it was after your final log)."""
    global nonproject_minutes, session_active
    end = max(end, last_entry_time)
    trailing = (end - last_entry_time).total_seconds() / 60.0
    if trailing > 0:
        nonproject_minutes += trailing
    ended = "\n**Session ended:** " + hhmm(end)
    if note:
        ended += f" ({note})"
    save_with_retry("work-day summary", append_to_log,
                    ended + "\n\n" + "\n".join(build_summary(end)) + "\n")
    session_active = False


def roll_over_day():
    """A new calendar day has started: close yesterday's session at your last
    activity (so a forgotten evening isn't counted as work). The next session
    starts at your first activity today."""
    global rolling_over, away_since
    if rolling_over:
        return
    if dialog_open and away_since is None:
        return                  # you're answering a prompt: finish that first
    rolling_over = True
    try:
        midnight = datetime.combine(session_start.date() + timedelta(days=1),
                                    dtime.min)
        end = away_since if away_since is not None else last_active
        end = min(max(end, last_entry_time), midnight)
        close_active_dialog()
        present = away_since is None and idle_seconds() < 120
        away_since = None
        if present and end - last_entry_time >= timedelta(minutes=1):
            # Still at the PC at midnight: ask about the last block first.
            do_prompt("day_end", end)
        end_session(end, "day closed automatically")
    finally:
        rolling_over = False


def finalize_and_quit():
    """Write the summary, stop the tray icon, end the Tk loop."""
    global quitting
    quitting = True
    close_active_dialog()
    if session_active:
        end_session(now())
    try:
        if icon is not None:
            icon.stop()
    except Exception:
        pass
    if root is not None:
        root.quit()


# ─────────────────────────────────────────────────────────────────────────────
# Away detection
# ─────────────────────────────────────────────────────────────────────────────

def update_activity():
    """Track the last moment you were at the PC, and notice when you've gone
    away: no input for IDLE_MINUTES, or the PC was asleep at least that long."""
    global last_active, away_since, last_poll
    t = now()
    idle = idle_seconds()
    slept = (last_poll is not None
             and t - last_poll > timedelta(minutes=IDLE_MINUTES))
    last_poll = t
    if away_since is not None:
        return
    # While paused you're never counted as away, but your last activity is
    # still tracked so the day closes at the right time.
    if slept:
        # You left before the PC went to sleep; keep last_active from then.
        if not paused:
            away_since = last_active
            close_active_dialog()       # nobody was there to answer it
    elif idle < IDLE_MINUTES * 60:
        last_active = t - timedelta(seconds=idle)
    elif not paused:
        away_since = t - timedelta(seconds=idle)
        close_active_dialog()


def handle_return():
    """You're back after being away: ask whether that time was a break. A
    break is logged as its own entry, after asking what you worked on before
    it. Work time is simply included in the next normal entry."""
    global away_since, last_active, next_due, dialog_open
    if dialog_open:
        return
    back = now()
    start, away_since = away_since, None
    last_active = back
    dialog_open = True
    try:
        choice = show_away_dialog(start, back)
    finally:
        dialog_open = False
    if choice != "break":
        return
    start = max(start, last_entry_time)
    if start - last_entry_time >= timedelta(minutes=1):
        if not do_prompt("break", start):
            # Not answered: count the time before the break as unspecified.
            log_block(last_entry_time, start, "", [], "skip")
    log_block(start, back, "away from the PC", [], "break")
    next_due = now() + interval


# ─────────────────────────────────────────────────────────────────────────────
# Task windows
# ─────────────────────────────────────────────────────────────────────────────

def task_form(task=None, parent=None):
    """Form to create a task (task=None) or edit one. Returns True if saved."""
    projects = load_projects()
    dlg = tk.Toplevel(parent or root)
    dlg.title("Work-log — " + ("Edit task" if task else "New task"))
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)

    tk.Label(dlg, text="Task:").grid(row=0, column=0, sticky="w",
                                     padx=(14, 6), pady=(14, 4))
    e_task = ttk.Entry(dlg, width=50)
    e_task.grid(row=0, column=1, sticky="we", padx=(0, 14), pady=(14, 4))
    tk.Label(dlg, text="Project:").grid(row=1, column=0, sticky="w",
                                        padx=(14, 6), pady=4)
    cb_project = ttk.Combobox(dlg, values=[""] + projects, width=47)
    cb_project.grid(row=1, column=1, sticky="we", padx=(0, 14), pady=4)
    tk.Label(dlg, text="Due:").grid(row=2, column=0, sticky="w",
                                    padx=(14, 6), pady=4)
    e_due = ttk.Entry(dlg, width=12)
    e_due.grid(row=2, column=1, sticky="w", padx=(0, 14), pady=4)
    tk.Label(dlg, text="YYYY-MM-DD, optional", fg="grey").grid(
        row=3, column=1, sticky="w", padx=(0, 14))
    if task:
        e_task.insert(0, task["Task"])
        cb_project.set(task["Project"])
        e_due.insert(0, task["Due"])

    result = {}

    def save():
        text = " ".join(e_task.get().split())
        due = e_due.get().strip()
        if not text:
            messagebox.showwarning("Work-log", "Please enter the task.",
                                   parent=dlg)
            return
        if due:
            try:
                due = datetime.strptime(due, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                messagebox.showwarning(
                    "Work-log", "Due date must look like 2026-10-31 "
                                "(or leave it empty).", parent=dlg)
                return
        result.update(text=text, project=cb_project.get().strip(), due=due)
        dlg.destroy()

    btns = tk.Frame(dlg)
    btns.grid(row=4, column=0, columnspan=2, sticky="e", padx=14, pady=14)
    ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left",
                                                              padx=(0, 6))
    ttk.Button(btns, text="Save", command=save).pack(side="left")
    dlg.bind("<Return>", lambda e: save())
    dlg.bind("<Escape>", lambda e: dlg.destroy())

    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw - w) // 2}+{(sh - h) // 3}")
    dlg.focus_force()
    e_task.focus_set()
    root.wait_window(dlg)
    if not result:
        return False

    due_note = f" (due {result['due']})" if result["due"] else ""
    if task is None:
        def add(tasks):
            t = {"ID": next_task_id(tasks),
                 "Created": now().strftime("%Y-%m-%d"),
                 "Project": result["project"], "Task": result["text"],
                 "Due": result["due"], "Status": "open", "Done": "",
                 "Minutes": "0"}
            tasks.append(t)
            return t
        saved, t = change_tasks(add)
        if saved:
            log_task_event("ADDED", t, due_note)
    else:
        def edit(tasks):
            for t in tasks:
                if t["ID"] == task["ID"]:
                    t.update(Task=result["text"], Project=result["project"],
                             Due=result["due"])
                    return t
            return None
        saved, t = change_tasks(edit)
        if t is None:
            messagebox.showwarning("Work-log", "That task no longer exists.")
            saved = False
        elif saved:
            log_task_event("EDITED", t, due_note)
    refresh_task_window()
    return saved


def set_task_done(task_id, done):
    """Mark a task done, or open it again."""
    def change(tasks):
        for t in tasks:
            if t["ID"] == task_id:
                t["Status"] = "done" if done else "open"
                t["Done"] = now().strftime("%Y-%m-%d") if done else ""
                return t
        return None
    saved, t = change_tasks(change)
    if saved and t is not None:
        if done:
            if task_id not in tasks_done_today:
                tasks_done_today.append(task_id)
            log_task_event("DONE", t,
                           f" (total {fmt_hm(task_total_minutes(t))})")
        else:
            if task_id in tasks_done_today:
                tasks_done_today.remove(task_id)
            log_task_event("REOPENED", t)
    refresh_task_window()


def delete_task(task_id):
    def change(tasks):
        for i, t in enumerate(tasks):
            if t["ID"] == task_id:
                return tasks.pop(i)
        return None
    saved, t = change_tasks(change)
    if saved and t is not None:
        log_task_event("DELETED", t)
    refresh_task_window()


def open_tasks_in_excel():
    """Open a read-only COPY of the task list, so Excel never locks the real
    file while Work-log needs to update it."""
    import shutil
    import stat
    import tempfile
    if not os.path.exists(TASKS_FILE):
        messagebox.showinfo("Work-log", "There are no tasks yet.")
        return
    copy = os.path.join(tempfile.gettempdir(),
                        "worklog_tasks_" + now().strftime("%Y%m%d_%H%M%S")
                        + ".csv")
    try:
        shutil.copyfile(TASKS_FILE, copy)
        os.chmod(copy, stat.S_IREAD)
    except OSError as e:
        messagebox.showwarning("Work-log", f"Couldn't copy the task list:\n{e}")
        return
    open_path(copy)


def refresh_task_window():
    if task_window is not None:
        try:
            task_window.refresh()
        except Exception:
            pass


def show_task_window():
    """The "Manage tasks" window: a list you can keep open beside your work."""
    global task_window
    if task_window is not None and task_window.winfo_exists():
        task_window.deiconify()
        task_window.lift()
        task_window.focus_force()
        return

    win = tk.Toplevel(root)
    win.title("Work-log — Tasks")
    win.minsize(560, 300)

    columns = ("id", "task", "project", "due", "time", "status")
    tree = ttk.Treeview(win, columns=columns, show="headings",
                        selectmode="browse", height=14)
    for col, heading, width, stretch in (
            ("id", "ID", 60, False), ("task", "Task", 300, True),
            ("project", "Project", 120, False), ("due", "Due", 90, False),
            ("time", "Time", 70, False), ("status", "Status", 70, False)):
        tree.heading(col, text=heading)
        tree.column(col, width=width, stretch=stretch)
    tree.tag_configure("overdue", foreground="#c00000")
    tree.tag_configure("done", foreground="grey")
    bar = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=bar.set)

    show_done = tk.BooleanVar(value=False)
    btns = tk.Frame(win)

    def selected_id():
        sel = tree.selection()
        return sel[0] if sel else None

    def refresh():
        keep = selected_id()
        tree.delete(*tree.get_children())
        tasks = load_tasks()
        today = now().strftime("%Y-%m-%d")
        shown = open_task_list(tasks)
        if show_done.get():
            shown += [t for t in tasks if t["Status"] == "done"]
        for t in shown:
            if t["Status"] == "done":
                tags = ("done",)
            elif t["Due"] and t["Due"] < today:
                tags = ("overdue",)
            else:
                tags = ()
            tree.insert("", "end", iid=t["ID"], tags=tags, values=(
                t["ID"], t["Task"], t["Project"], t["Due"],
                fmt_hm(task_total_minutes(t)), t["Status"]))
        if keep and tree.exists(keep):
            tree.selection_set(keep)

    def new():
        task_form(parent=win)

    def edit(_event=None):
        tid = selected_id()
        task = next((t for t in load_tasks() if t["ID"] == tid), None)
        if task:
            task_form(task, parent=win)

    def toggle_done():
        tid = selected_id()
        task = next((t for t in load_tasks() if t["ID"] == tid), None)
        if task:
            set_task_done(tid, task["Status"] != "done")

    def delete():
        tid = selected_id()
        if tid and messagebox.askyesno("Work-log", f"Delete task {tid}?",
                                       parent=win):
            delete_task(tid)

    def close():
        global task_window
        task_window = None
        win.destroy()

    for text, cmd in (("New…", new), ("Edit…", edit),
                      ("Done / reopen", toggle_done), ("Delete", delete),
                      ("Open in Excel", open_tasks_in_excel)):
        ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 6))
    ttk.Checkbutton(btns, text="Show done tasks", variable=show_done,
                    command=refresh).pack(side="right")

    btns.pack(side="bottom", fill="x", padx=10, pady=10)
    bar.pack(side="right", fill="y", pady=(10, 0))
    tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(10, 0))
    tree.bind("<Double-1>", edit)
    win.protocol("WM_DELETE_WINDOW", close)

    win.refresh = refresh
    task_window = win
    refresh()


# ─────────────────────────────────────────────────────────────────────────────
# Monthly report (Excel)
# ─────────────────────────────────────────────────────────────────────────────

def read_month(year, month):
    """The rows of a month's CSV, with Hours as a number and Day as a date."""
    path = csv_path_for(datetime(year, month, 1))
    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            raw = list(csv.DictReader(f, delimiter=CSV_DELIMITER))
    except FileNotFoundError:
        return []
    rows = []
    for r in raw:
        try:
            hours = float((r.get("Hours") or "0").replace(",", "."))
            day = datetime.strptime(r.get("Date") or "", "%Y-%m-%d").date()
        except ValueError:
            continue                    # skip a damaged line
        rows.append(dict(r, Hours=hours, Day=day))
    return rows


def report_months():
    """Months that have a CSV, newest first, as (year, month)."""
    try:
        names = os.listdir(LOG_DIR)
    except OSError:
        return []
    months = set()
    for name in names:
        if (name.startswith("worklog_") and name.endswith(".csv")
                and len(name) == len("worklog_2026-10.csv")):
            try:
                d = datetime.strptime(name[8:15], "%Y-%m")
            except ValueError:
                continue
            months.add((d.year, d.month))
    return sorted(months, reverse=True)


def create_report(year, month):
    """Build the Excel report for a month and return its path, or None if the
    month has no entries. Sheets: Summary (1st–15th, 16th–end and the whole
    month: time, hours per project and per task), Days, Weeks and Data."""
    import calendar
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    rows = read_month(year, month)
    if not rows:
        return None
    first = datetime(year, month, 1)
    last_day = calendar.monthrange(year, month)[1]
    mon = first.strftime("%b")
    periods = [(f"1–15 {mon}", lambda d: d.day <= 15),
               (f"16–{last_day} {mon}", lambda d: d.day > 15),
               (first.strftime("%B %Y"), lambda d: True)]
    tasks_by_id = {t["ID"]: t for t in load_tasks()}
    worked_types = ("Work", "Non-project")

    bold, title = Font(bold=True), Font(bold=True, size=14)
    fill = PatternFill("solid", fgColor="DDE6F4")
    HOURS = "0.00"

    def header(ws, r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(r, c, v)
            cell.font, cell.fill = bold, fill

    def put_row(ws, r, values, font=None):
        for c, v in enumerate(values, 1):
            cell = ws.cell(r, c, v)
            if isinstance(v, float):
                cell.number_format = HOURS
            if font:
                cell.font = font

    def hours(match, period):
        return sum(x["Hours"] for x in rows if match(x) and period(x["Day"]))

    def widths(ws, values):
        for c, w in enumerate(values, 1):
            ws.column_dimensions[get_column_letter(c)].width = w

    wb = Workbook()

    # ── Summary ──
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Work log — " + first.strftime("%B %Y")
    ws["A1"].font = title
    r = 3
    header(ws, r, ["Time (hours)"] + [p for p, _ in periods])
    lines = [
        ("Worked (excl. breaks)", lambda x: x["Type"] in worked_types),
        ("   on projects", lambda x: x["Type"] == "Work"),
        ("   non-project work", lambda x: x["Type"] == "Non-project"),
        ("Breaks", lambda x: x["Type"] == "Break"),
    ]
    for label, match in lines:
        r += 1
        put_row(ws, r, [label] + [hours(match, p) for _, p in periods],
                bold if r == 4 else None)
    r += 1
    put_row(ws, r, ["Days worked"] + [
        len({x["Day"] for x in rows if x["Type"] in worked_types and p(x["Day"])})
        for _, p in periods])

    r += 2
    header(ws, r, ["Hours per project"] + [p for p, _ in periods])
    projects = sorted({x["Project"] for x in rows if x["Type"] in worked_types},
                      key=lambda name: -hours(
                          lambda x, n=name: x["Project"] == n, lambda d: True))
    for name in projects:
        r += 1
        put_row(ws, r, [name] + [
            hours(lambda x, n=name: x["Project"] == n
                  and x["Type"] in worked_types, p) for _, p in periods])
    r += 1
    put_row(ws, r, ["Total"] + [
        hours(lambda x: x["Type"] in worked_types, p) for _, p in periods], bold)

    r += 2
    header(ws, r, ["Hours per task", "Project"] + [p for p, _ in periods]
           + ["Status"])
    task_ids = sorted({x["Task_ID"] for x in rows if x.get("Task_ID")},
                      key=lambda i: -hours(lambda x, i=i: x["Task_ID"] == i,
                                           lambda d: True))
    for tid in task_ids:
        sample = next(x for x in rows if x["Task_ID"] == tid)
        t = tasks_by_id.get(tid)
        if t is None:
            status = "deleted"
        elif t["Status"] == "done":
            status = f"done {t['Done']}"
        else:
            status = f"open (due {t['Due']})" if t["Due"] else "open"
        r += 1
        put_row(ws, r, [f"{tid} {sample['Task']}",
                        t["Project"] if t else sample["Project"]]
                + [hours(lambda x, i=tid: x["Task_ID"] == i, p)
                   for _, p in periods] + [status])
    if not task_ids:
        r += 1
        ws.cell(r, 1, "(no task time this month)")
    widths(ws, [34, 18, 16, 16, 18, 22])

    # ── Days ──
    ws = wb.create_sheet("Days")
    header(ws, 1, ["Date", "Weekday", "Start", "End", "Worked h", "Breaks h",
                   "Projects"])
    for i, day in enumerate(sorted({x["Day"] for x in rows}), 2):
        of_day = [x for x in rows if x["Day"] == day]
        names = []
        for x in of_day:
            if x["Type"] == "Work" and x["Project"] not in names:
                names.append(x["Project"])
        put_row(ws, i, [day, day.strftime("%a"),
                        min(x["Start"] for x in of_day),
                        max(x["End"] for x in of_day),
                        sum(x["Hours"] for x in of_day
                            if x["Type"] in worked_types),
                        sum(x["Hours"] for x in of_day if x["Type"] == "Break"),
                        ", ".join(names)])
        ws.cell(i, 1).number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    widths(ws, [12, 9, 7, 7, 10, 10, 40])

    # ── Weeks: hours per project per ISO week ──
    ws = wb.create_sheet("Weeks")
    weeks = sorted({x["Week"] for x in rows})
    header(ws, 1, ["Project"] + weeks + ["Total"])
    for i, name in enumerate(projects, 2):
        mine = [x for x in rows if x["Project"] == name
                and x["Type"] in worked_types]
        put_row(ws, i, [name] + [sum(x["Hours"] for x in mine
                                     if x["Week"] == w) for w in weeks]
                + [sum(x["Hours"] for x in mine)])
    put_row(ws, len(projects) + 2, ["Total"] + [
        sum(x["Hours"] for x in rows if x["Week"] == w
            and x["Type"] in worked_types) for w in weeks]
        + [sum(x["Hours"] for x in rows if x["Type"] in worked_types)], bold)
    ws.freeze_panes = "B2"
    widths(ws, [30] + [11] * (len(weeks) + 1))

    # ── Data: every row, filterable ──
    ws = wb.create_sheet("Data")
    header(ws, 1, CSV_HEADERS)
    for i, x in enumerate(rows, 2):
        put_row(ws, i, [x["Day"], x.get("Week", ""), x.get("Weekday", ""),
                        x.get("Start", ""), x.get("End", ""), x["Hours"],
                        x.get("Type", ""), x.get("Project", ""),
                        x.get("Task_ID", ""), x.get("Task", ""),
                        x.get("Note", "")])
        ws.cell(i, 1).number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CSV_HEADERS))}{len(rows) + 1}"
    widths(ws, [12, 10, 8, 7, 7, 8, 12, 24, 8, 30, 50])

    path = os.path.join(LOG_DIR, f"report_{year}-{month:02d}.xlsx")
    try:
        wb.save(path)
    except PermissionError:
        # Last report still open in Excel: save next to it under a new name.
        path = path[:-5] + now().strftime("_%H%M%S") + ".xlsx"
        wb.save(path)
    return path


def report_dialog():
    """Ask which month to report on, then build the report and open it."""
    months = report_months()
    if not months:
        messagebox.showinfo("Work-log", "There's nothing to report yet.")
        return
    labels = [datetime(y, m, 1).strftime("%B %Y") for y, m in months]

    dlg = tk.Toplevel(root)
    dlg.title("Work-log — Report")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    tk.Label(dlg, text="Create a report for:").grid(
        row=0, column=0, sticky="w", padx=(14, 6), pady=14)
    cb = ttk.Combobox(dlg, values=labels, state="readonly", width=18)
    cb.current(0)
    cb.grid(row=0, column=1, sticky="w", padx=(0, 14), pady=14)
    chosen = {}

    def create():
        chosen["month"] = months[cb.current()]
        dlg.destroy()

    btns = tk.Frame(dlg)
    btns.grid(row=1, column=0, columnspan=2, sticky="e", padx=14, pady=(0, 14))
    ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left",
                                                              padx=(0, 6))
    ttk.Button(btns, text="Create", command=create).pack(side="left")
    dlg.bind("<Return>", lambda e: create())
    dlg.bind("<Escape>", lambda e: dlg.destroy())
    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw - w) // 2}+{(sh - h) // 3}")
    dlg.focus_force()
    root.wait_window(dlg)
    if not chosen:
        return

    try:
        path = create_report(*chosen["month"])
    except ImportError:
        messagebox.showerror("Work-log", "Reports need the openpyxl package:\n"
                                         "    pip install openpyxl")
        return
    except OSError as e:
        messagebox.showerror("Work-log", f"Couldn't create the report:\n{e}")
        return
    if path is None:
        messagebox.showinfo("Work-log", "That month has no entries.")
    else:
        open_path(path)


# ─────────────────────────────────────────────────────────────────────────────
# Tray menu callbacks — these run on pystray's THREAD, so they only enqueue.
# ─────────────────────────────────────────────────────────────────────────────

def tray_log_now(_icon=None, _item=None):
    cmd_q.put(("log",))


def tray_toggle_pause(_icon=None, _item=None):
    cmd_q.put(("toggle_pause",))


def tray_open_log(_icon=None, _item=None):
    cmd_q.put(("open", LOG_PATH))


def tray_open_csv(_icon=None, _item=None):
    cmd_q.put(("open", csv_path_for(now())))


def tray_open_folder(_icon=None, _item=None):
    cmd_q.put(("open", LOG_DIR))


def tray_quit(_icon=None, _item=None):
    cmd_q.put(("quit",))


def tray_command(name):
    """Returns a callback that just queues `name` for the main thread."""
    def _cb(_icon=None, _item=None):
        cmd_q.put((name,))
    return _cb


def tray_set_interval(minutes):
    """Returns a callback that asks the main thread to switch the interval."""
    def _cb(_icon=None, _item=None):
        cmd_q.put(("set_interval", minutes))
    return _cb


def build_tray_icon():
    """Create the tray icon with its right-click menu."""
    # A simple drawn clock face so we don't need an external .ico file.
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse([3, 3, 61, 61], fill=(45, 120, 220, 255))
    d.line([32, 32, 32, 15], fill="white", width=4)     # minute hand
    d.line([32, 32, 46, 39], fill="white", width=4)     # hour hand
    d.ellipse([29, 29, 35, 35], fill="white")

    # "Interval" submenu — one radio-style item per choice, checked when active.
    interval_items = [
        MenuItem(
            f"{m} minutes",
            tray_set_interval(m),
            checked=(lambda mm: (lambda item: interval_minutes == mm))(m),
            radio=True,
        )
        for m in INTERVAL_CHOICES
    ]

    menu = Menu(
        MenuItem("Log now", tray_log_now, default=True),
        MenuItem("Pause auto-prompts", tray_toggle_pause,
                 checked=lambda item: paused),
        MenuItem("Interval", Menu(*interval_items)),
        Menu.SEPARATOR,
        MenuItem("Tasks", Menu(
            MenuItem("New task…", tray_command("new_task")),
            MenuItem("Manage tasks…", tray_command("manage_tasks")),
            MenuItem("Open task list in Excel", tray_command("tasks_excel")),
        )),
        Menu.SEPARATOR,
        MenuItem("Create report…", tray_command("report")),
        MenuItem("Open this month's CSV", tray_open_csv),
        MenuItem("Open today's log", tray_open_log, visible=WRITE_DAILY_LOG),
        MenuItem("Open log folder", tray_open_folder),
        Menu.SEPARATOR,
        MenuItem("Quit", tray_quit),
    )
    return pystray.Icon("worklog", img, "Work-log", menu=menu)


# ─────────────────────────────────────────────────────────────────────────────
# Main-thread poll loop (handles the timer + tray commands)
# ─────────────────────────────────────────────────────────────────────────────

def poll():
    """Run one poll step and schedule the next. An unexpected error is shown
    instead of silently stopping the timer and the tray menu."""
    global next_due, showing_error
    if quitting:
        return
    # Schedule the next check FIRST: while a dialog is open, Tk waits in a
    # nested event loop, and this keeps the timer, tray commands and the
    # away / new-day checks running during it.
    root.after(POLL_MS, poll)
    try:
        poll_step()
    except Exception as e:
        # Restart the countdown so a failing prompt doesn't re-fire at once.
        next_due = now() + interval
        if showing_error:
            return
        showing_error = True
        try:
            messagebox.showerror("Work-log",
                                 f"Unexpected error:\n{type(e).__name__}: {e}"
                                 "\n\nWork-log keeps running.")
        except Exception:
            pass
        finally:
            showing_error = False


def poll_step():
    """Handle queued tray commands, the day/away bookkeeping and the scheduled
    prompt. Returns False when the program is quitting.

    Note: while a dialog is open, Tk keeps calling poll() (the dialog waits in
    a nested event loop), so this must cope with being re-entered."""
    global paused, next_due, interval, interval_minutes

    # 1) Handle any queued tray commands first.
    while True:
        try:
            cmd = cmd_q.get_nowait()
        except queue.Empty:
            break

        name = cmd[0]
        if name == "log":
            if dialog_open:
                raise_active_dialog()        # a prompt is already waiting
                continue
            do_prompt("manual")
            if RESET_TIMER_ON_MANUAL_LOG:
                next_due = now() + interval
        elif name == "toggle_pause":
            paused = not paused
            if not paused:
                # Resuming: restart the countdown so it doesn't fire instantly.
                next_due = now() + interval
        elif name == "set_interval":
            interval_minutes = cmd[1]
            interval = timedelta(minutes=interval_minutes)
            # Re-base the countdown from now so the new cadence starts cleanly.
            next_due = now() + interval
        elif name == "new_task":
            task_form()
        elif name == "manage_tasks":
            show_task_window()
        elif name == "tasks_excel":
            open_tasks_in_excel()
        elif name == "report":
            report_dialog()
        elif name == "open":
            open_path(cmd[1])
        elif name == "quit":
            finalize_and_quit()
            return False

    # 2) After a day was closed: start the new day at your first activity.
    if not session_active:
        if idle_seconds() < 60:
            start_session(now())
        return True

    # 3) Away detection, and closing the day when the date changes.
    update_activity()
    if now().date() != session_start.date():
        roll_over_day()
        return True
    if away_since is not None:
        if idle_seconds() < 10:
            handle_return()
        return True                      # no prompts while you're away

    # 4) Scheduled auto-prompt (skipped while paused).
    if not paused and now() >= next_due:
        do_prompt("scheduled")
        while next_due <= now():         # catch up if a long answer overran
            next_due += interval

    return True


# ─────────────────────────────────────────────────────────────────────────────
# Startup
# ─────────────────────────────────────────────────────────────────────────────

def main():
    global root, icon, interval, interval_minutes

    os.makedirs(LOG_DIR, exist_ok=True)

    # Hidden Tk root drives the popups and the timer; it never shows itself.
    root = tk.Tk()
    root.withdraw()
    # Overdue tasks are shown in red in the prompt.
    ttk.Style(root).configure("Overdue.TCheckbutton", foreground="#c00000")

    interval_minutes = INTERVAL_MINUTES
    interval = timedelta(minutes=interval_minutes)
    start_session(now())

    # Start the tray icon on its own thread (its menu callbacks just enqueue).
    icon = build_tray_icon()
    threading.Thread(target=icon.run, daemon=True).start()

    # Begin the poll loop and hand control to Tk.
    root.after(POLL_MS, poll)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        finalize_and_quit()
    finally:
        try:
            root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        _r = tk.Tk(); _r.withdraw()
        messagebox.showerror("Work-log",
                             f"Work-log couldn't start:\n"
                             f"{type(e).__name__}: {e}")
        sys.exit(1)
