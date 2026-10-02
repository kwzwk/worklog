#!/usr/bin/env python3
"""
Work-log timer — a background system-tray utility for logging what you work on.

Runs quietly as a tray icon (down by the clock). Every INTERVAL_MINUTES a popup
asks "What have you worked on?", lets you tick the project(s) you worked on, and
appends a line to a daily Markdown .txt log and a weekly CSV. Right-click the
tray icon to log on demand, pause, open the files, or quit.

Markdown entry format:   * HH:MM–HH:MM - Project Name - Entry text
On quit it appends a work-day summary (hours per project + total) to the log.

PATHS: everything lives next to the program itself (the .py file, or the .exe if
you build one) — no hard-coded user path. Daily logs and weekly CSVs go in a
"log" subfolder; projects.txt sits in the main folder next to the program.

REQUIREMENTS (beyond the Python standard library):
    pip install pystray pillow
Both bundle fine into a PyInstaller .exe. Build it WINDOWLESS (--noconsole) —
see build_exe.bat.
"""

import csv
import os
import queue
import sys
import threading
from datetime import datetime, timedelta

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

# Daily log files (and weekly CSVs) go in a "log" subfolder next to the program.
LOG_DIR = os.path.join(BASE_DIR, "log")

# Project list, in the main folder next to the program. One project per line.
# Blank lines and lines starting with '#' are ignored. Re-read live — edits take
# effect at the next prompt, no restart needed.
PROJECTS_FILE = os.path.join(BASE_DIR, "projects.txt")

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

# CSV column headers. These become the SharePoint list's column names on import,
# so they're kept clean (no spaces/punctuation) for tidy internal names.
CSV_HEADERS = ["Date", "Start", "End", "Duration_min", "Project", "Entry"]


# ─────────────────────────────────────────────────────────────────────────────
# Internal state
# ─────────────────────────────────────────────────────────────────────────────

LOG_PATH = None             # full path to today's markdown file
entries_count = 0            # how many entries logged this session
root = None                  # hidden Tk root window
icon = None                  # the tray icon

session_start = None
next_due = None              # datetime of the next scheduled prompt
interval = None              # timedelta of the current interval
interval_minutes = None      # current interval in minutes (changeable at runtime)
paused = False               # auto-prompts suspended while True
dialog_open = False          # guard so prompts never stack

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

# "Same task as before" memory (this session only).
last_projects = []
last_answer = ""


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


def log_path_for_today():
    return os.path.join(LOG_DIR, "worklog_" + now().strftime("%Y-%m-%d") + ".txt")


def csv_path_for(dt):
    """Weekly CSV path, e.g. ...\\log\\worklog_2026-W22.csv (ISO weeks)."""
    iso_year, iso_week, _ = dt.isocalendar()
    return os.path.join(LOG_DIR, f"worklog_{iso_year}-W{iso_week:02d}.csv")


def append_to_log(text):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(text)


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


# ─────────────────────────────────────────────────────────────────────────────
# Writing to the markdown log
# ─────────────────────────────────────────────────────────────────────────────

def write_session_header():
    file_existed = os.path.exists(LOG_PATH)
    chunk = ""
    if not file_existed:
        chunk += "# Worklog — " + now().strftime("%A, %Y-%m-%d") + "\n\n"
    else:
        chunk += "\n"
    chunk += "**Session started:** " + hhmm() + "\n\n"
    append_to_log(chunk)


def write_entry(start, end, answer_text, projects, kind):
    """Write one markdown line. `kind` is 'work', 'break', or 'skip'.
    Label is the project list, 'BREAK' for a break, or '(non-project work)' for
    worked-but-untracked time (which is also what a Skip with no projects logs)."""
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
    append_to_log("* " + span + " - " + label + " - " + entry + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# Writing to the weekly CSV
# ─────────────────────────────────────────────────────────────────────────────

def write_csv_rows(start, end, answer_text, projects, kind):
    """Append rows to the weekly CSV, optimized for Excel and SharePoint import.

    Columns (see CSV_HEADERS): Date, Start, End, Duration_min, Project, Entry.
    - Date is ISO YYYY-MM-DD (the format SharePoint date columns expect).
    - Duration_min is a plain number with a '.' decimal (invariant), so it imports
      as a real Number, not text. Split evenly across projects so the column sums
      to total time.
    - One row per project; breaks → 'BREAK', untracked work → '(non-project work)',
      so a pivot / grouping on Project cleanly separates all three categories.
    - The csv module quotes any field containing the delimiter, quotes or newlines
      (RFC 4180), so commas in your notes never break the columns.
    """
    path = csv_path_for(end)
    file_existed = os.path.exists(path)

    # Flatten the note to a single line (no embedded newlines) for clean cells.
    cell = answer_text.replace("\r\n", "\n").replace("\n", " | ").strip() \
        or "(no entry)"

    total_min = max(0.0, (end - start).total_seconds() / 60.0)
    if kind == "break":
        label_list = ["BREAK"]
    elif projects:
        label_list = projects
    else:
        label_list = ["(non-project work)"]
    share = total_min / len(label_list)
    # Whole numbers as "60", fractional as "30.5" — tidy and numeric either way.
    dur = f"{share:.2f}".rstrip("0").rstrip(".")

    # newline="" prevents blank rows; csv.writer then emits RFC-4180 \r\n.
    with open(path, "a", encoding="utf-8", newline="") as f:
        if not file_existed and CSV_WRITE_BOM:
            f.write("\ufeff")                       # UTF-8 BOM → Excel encoding
        writer = csv.writer(f, delimiter=CSV_DELIMITER,
                            quoting=csv.QUOTE_MINIMAL)
        if not file_existed:
            writer.writerow(CSV_HEADERS)
        date_str = end.strftime("%Y-%m-%d")         # ISO date
        for p in label_list:
            writer.writerow([date_str, hhmm(start), hhmm(end), dur, p, cell])


# ─────────────────────────────────────────────────────────────────────────────
# The popup prompt (Tkinter) — always runs on the main thread
# ─────────────────────────────────────────────────────────────────────────────

def show_prompt_dialog(reason):
    """Question + multi-line text box + project tick boxes.
    Returns (answer_text, selected_projects, kind) where kind is:
      'work'  — Save: log what's typed/ticked
      'break' — Break: time off
      'same'  — Same as before: reuse the last work task (handled by caller)
      'skip'  — Skip: nothing to repeat, just dismiss (logged as non-project)
    """
    projects = load_projects()
    has_previous = bool(last_projects or last_answer)

    dlg = tk.Toplevel(root)
    dlg.title("Work-log")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.lift()

    tag = "auto" if reason == "scheduled" else "manual"
    tk.Label(dlg, text=f"{PROMPT_QUESTION}   ({tag} · {hhmm()})",
             font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 6))

    txt = tk.Text(dlg, width=58, height=4, wrap="word", font=("Segoe UI", 10))
    txt.pack(padx=14, pady=(0, 6), fill="x")

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

    # Hint so it's clear what happens with no project ticked. "covers the last
    # N min" is the REAL gap since the previous entry (so a manual log shows the
    # true elapsed time, not the fixed interval).
    elapsed_min = max(0, int(round((now() - last_entry_time).total_seconds() / 60)))
    tk.Label(dlg,
             text=f"Tip: covers the last {elapsed_min} min · "
                  "no project ticked = non-project work · "
                  "use Break for time off.",
             fg="grey").pack(anchor="w", padx=14, pady=(0, 6))

    result = {"answer": "", "projects": [], "kind": "work"}

    def do_save():
        result["answer"] = txt.get("1.0", "end").strip()
        result["projects"] = [n for n, v in vars_by_name.items() if v.get()]
        result["kind"] = "work"
        dlg.destroy()

    def do_break():
        # A break: record any typed note, but ignore project ticks entirely.
        result["answer"] = txt.get("1.0", "end").strip()
        result["projects"] = []
        result["kind"] = "break"
        dlg.destroy()

    def do_same_or_skip():
        # If there's a previous work task, one click repeats it. Otherwise this
        # is a plain Skip (dismiss → logged as non-project work with no entry).
        result["kind"] = "same" if has_previous else "skip"
        dlg.destroy()

    # Third button is adaptive: "Same as before (…)" when there's something to
    # repeat, otherwise "Skip".
    if has_previous:
        preview = ", ".join(last_projects) if last_projects else "previous note"
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

    dlg.update_idletasks()
    w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw - w) // 2}+{(sh - h) // 3}")
    dlg.focus_force()
    txt.focus_set()

    root.wait_window(dlg)
    return result["answer"], result["projects"], result["kind"]


def do_prompt(reason):
    """Show the popup, account the time into the right bucket, write to markdown
    + CSV. Guarded so two prompts never overlap."""
    global entries_count, last_entry_time, nonproject_minutes, break_minutes
    global last_projects, last_answer, dialog_open

    if dialog_open:
        return
    dialog_open = True
    try:
        answer, projects, kind = show_prompt_dialog(reason)
    finally:
        dialog_open = False

    # "Same as before" → reuse the remembered last work task, then treat it as
    # an ordinary work entry from here on.
    if kind == "same":
        projects = list(last_projects)
        answer = last_answer
        kind = "work"

    start = last_entry_time
    end = now()
    delta_min = max(0.0, (end - start).total_seconds() / 60.0)
    if kind == "break":
        break_minutes += delta_min
    elif projects:
        share = delta_min / len(projects)
        for p in projects:
            project_minutes[p] = project_minutes.get(p, 0.0) + share
    else:
        nonproject_minutes += delta_min        # worked, just not on a project
    last_entry_time = end

    write_entry(start, end, answer, projects, kind)
    write_csv_rows(start, end, answer, projects, kind)

    # Remember the last real WORK task for "Same as before" (not breaks).
    if kind == "work" and (projects or answer.strip()):
        last_projects = list(projects)
        last_answer = answer

    entries_count += 1


# ─────────────────────────────────────────────────────────────────────────────
# Work-day summary (written on quit)
# ─────────────────────────────────────────────────────────────────────────────

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
    return lines


def finalize_and_quit():
    """Write the summary, stop the tray icon, end the Tk loop. Time since the
    last entry is treated as non-project work (it was after your final log)."""
    global nonproject_minutes
    end = now()
    trailing = (end - last_entry_time).total_seconds() / 60.0
    if trailing > 0:
        nonproject_minutes += trailing
    append_to_log("\n**Session ended:** " + hhmm(end) + "\n\n"
                  + "\n".join(build_summary(end)) + "\n")
    try:
        if icon is not None:
            icon.stop()
    except Exception:
        pass
    if root is not None:
        root.quit()


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
        MenuItem("Open today's log", tray_open_log),
        MenuItem("Open this week's CSV", tray_open_csv),
        MenuItem("Open log folder", tray_open_folder),
        Menu.SEPARATOR,
        MenuItem("Quit", tray_quit),
    )
    return pystray.Icon("worklog", img, "Work-log", menu=menu)


# ─────────────────────────────────────────────────────────────────────────────
# Main-thread poll loop (handles the timer + tray commands)
# ─────────────────────────────────────────────────────────────────────────────

def poll():
    global paused, next_due, interval, interval_minutes

    # 1) Handle any queued tray commands first.
    while True:
        try:
            cmd = cmd_q.get_nowait()
        except queue.Empty:
            break

        name = cmd[0]
        if name == "log":
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
        elif name == "open":
            open_path(cmd[1])
        elif name == "quit":
            finalize_and_quit()
            return                       # stop the poll chain

    # 2) Scheduled auto-prompt (skipped while paused).
    if not paused and now() >= next_due:
        do_prompt("scheduled")
        while next_due <= now():         # catch up if a long answer overran
            next_due += interval

    # 3) Schedule the next check.
    root.after(POLL_MS, poll)


# ─────────────────────────────────────────────────────────────────────────────
# Startup
# ─────────────────────────────────────────────────────────────────────────────

def main():
    global LOG_PATH, root, icon, session_start, next_due, interval
    global interval_minutes, last_entry_time, nonproject_minutes, break_minutes

    os.makedirs(LOG_DIR, exist_ok=True)
    LOG_PATH = log_path_for_today()

    # Hidden Tk root drives the popups and the timer; it never shows itself.
    root = tk.Tk()
    root.withdraw()

    session_start = now()
    project_minutes.clear()
    nonproject_minutes = 0.0
    break_minutes = 0.0
    last_entry_time = session_start
    write_session_header()

    interval_minutes = INTERVAL_MINUTES
    interval = timedelta(minutes=interval_minutes)
    next_due = session_start + interval

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
    main()
