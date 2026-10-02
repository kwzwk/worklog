#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
worklog.py — a tiny background work-logging tool.

Run it at the start of your workday and leave it running. Every
INTERVAL_MINUTES it asks "What have you worked on?" and writes your answer
to a per-day Markdown .txt file. You can also log on demand at any time, and
quit cleanly when you're done.

See the bottom of this file (or the instructions you were given) for how to
run it and how to build a standalone .exe.
"""

import os
import sys
import time
import queue
import threading
from datetime import datetime

# ==========================================================================
# CONFIG  —  the things you'll most likely want to change live right here.
# ==========================================================================

# How often the program automatically asks what you've worked on (in minutes).
INTERVAL_MINUTES = 60

# Folder where the daily log files are written. It is created if missing.
# (The leading r"" means backslashes are taken literally — handy on Windows.)
LOG_DIR = r"C:\Users\KZ1317\Desktop\TimeLogger"

# The question you get asked.
PROMPT_TEXT = "What have you worked on?"

# ==========================================================================
# Nothing below here normally needs editing.
# ==========================================================================

INTERVAL_SECONDS = INTERVAL_MINUTES * 60

# The log file path is decided once, when the program starts, so a single
# run always writes to the file for the day you started on.
LOG_PATH = None


# --------------------------------------------------------------------------
# Small time / file helpers
# --------------------------------------------------------------------------

def now_hhmm():
    """Current wall-clock time as 24-hour HH:MM."""
    return datetime.now().strftime("%H:%M")


def todays_log_path():
    """Full path to today's log file, e.g. .../worklog_2026-05-27.txt"""
    filename = "worklog_" + datetime.now().strftime("%Y-%m-%d") + ".txt"
    return os.path.join(LOG_DIR, filename)


def append_to_log(text):
    """Append text to the log file (UTF-8). Opened per write so the file is
    flushed to disk immediately and isn't held open while you work."""
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(text)


# --------------------------------------------------------------------------
# Writing the different kinds of lines into the Markdown file
# --------------------------------------------------------------------------

def write_session_header():
    """Write the start-of-session marker.

    - If today's file is brand new (or empty), write the top-level
      '# Worklog — ...' heading followed by a 'Session started' line.
    - If the file already exists (an earlier session happened today), just
      add another 'Session started' line instead of a second heading.
    """
    file_is_new = (not os.path.exists(LOG_PATH)) or os.path.getsize(LOG_PATH) == 0

    if file_is_new:
        # e.g. "Monday, 2026-05-27"
        nice_date = datetime.now().strftime("%A, %Y-%m-%d")
        append_to_log("# Worklog \u2014 " + nice_date + "\n\n")
        append_to_log("**Session started:** " + now_hhmm() + "\n\n")
    else:
        # Blank line first to separate this session from earlier content.
        append_to_log("\n**Session started:** " + now_hhmm() + "\n\n")


def write_entry(answer_lines):
    """Write one log entry as a timestamped Markdown bullet and return the
    timestamp string. `answer_lines` is a list of text lines (may be empty)."""
    ts = now_hhmm()

    # Trim blank lines from the top and bottom of the answer.
    lines = [ln.rstrip() for ln in answer_lines]
    while lines and lines[0] == "":
        lines.pop(0)
    while lines and lines[-1] == "":
        lines.pop()

    if not lines:
        # Empty answer -> recorded as "(no entry)".
        append_to_log("- **" + ts + "** \u2014 (no entry)\n")
    else:
        # First line goes right after the timestamp; any further lines are
        # indented two spaces so they stay part of the same Markdown bullet.
        append_to_log("- **" + ts + "** \u2014 " + lines[0] + "\n")
        for extra in lines[1:]:
            append_to_log("  " + extra + "\n")

    return ts


def write_session_end():
    """Append the end-of-session marker."""
    append_to_log("\n**Session ended:** " + now_hhmm() + "\n")


# --------------------------------------------------------------------------
# Reading keyboard input
#
# A dedicated background thread reads lines from stdin and drops them onto a
# queue. This is the ONLY place that reads the keyboard, so the main loop can
# watch both the queue and the timer without two threads fighting over input.
# --------------------------------------------------------------------------

def stdin_reader(line_queue):
    """Read lines forever and put them on the queue. Put None on EOF."""
    while True:
        line = sys.stdin.readline()
        if line == "":            # End-of-file (e.g. input stream closed)
            line_queue.put(None)
            return
        # Strip the trailing newline / carriage return only.
        line_queue.put(line.rstrip("\n").rstrip("\r"))


def collect_answer(line_queue):
    """Print the question and gather a (possibly multi-line) answer.

    The user types one or more lines and finishes by pressing Enter on an
    empty line. An immediately-empty answer comes back as an empty list.
    Returns the list of answer lines.
    """
    print()
    print(PROMPT_TEXT)
    print("(Type your answer. Multiple lines are OK. "
          "Press Enter on an empty line to save.)")

    answer = []
    while True:
        try:
            # Short timeout so Ctrl+C stays responsive while you type.
            line = line_queue.get(timeout=0.5)
        except queue.Empty:
            continue
        if line is None:          # EOF
            break
        if line.strip() == "":    # blank line ends the answer
            break
        answer.append(line)
    return answer


def do_log(line_queue, kind):
    """Ask the question, write the entry, and confirm on screen.
    `kind` is just a label, e.g. 'manual' or 'automatic'."""
    ts = write_entry(collect_answer(line_queue))
    print("  Logged at " + ts + " (" + kind + ").")


# --------------------------------------------------------------------------
# Startup banner
# --------------------------------------------------------------------------

def print_banner():
    print("=" * 62)
    print("  Worklog is running.")
    print("  Logging to : " + LOG_PATH)
    print("  Auto-prompt: every " + str(INTERVAL_MINUTES) + " minute(s).")
    print("-" * 62)
    print("  Commands:")
    print("    [Enter] or 'log'  ->  log an entry right now")
    print("    'quit'            ->  stop and save")
    print("    (Ctrl+C also quits cleanly)")
    print("=" * 62)


# --------------------------------------------------------------------------
# Main loop
# --------------------------------------------------------------------------

def main():
    global LOG_PATH

    # Make sure the target folder exists, then fix today's file path.
    os.makedirs(LOG_DIR, exist_ok=True)
    LOG_PATH = todays_log_path()

    write_session_header()

    # Start the keyboard-reading thread (daemon = dies when the program exits).
    line_queue = queue.Queue()
    threading.Thread(target=stdin_reader, args=(line_queue,), daemon=True).start()

    # When the next automatic prompt is due. We use a monotonic clock so it
    # isn't thrown off by the system clock changing.
    next_due = time.monotonic() + INTERVAL_SECONDS

    print_banner()

    try:
        while True:
            # Draw the command prompt once, then wait for either a typed
            # command or the timer to come due.
            sys.stdout.write("> ")
            sys.stdout.flush()

            command = None
            got_command = False
            while not got_command:
                # 1) Has the automatic timer come due?
                if time.monotonic() >= next_due:
                    print()                      # finish the "> " line
                    do_log(line_queue, "automatic")
                    # Reschedule the NEXT automatic prompt. (Only the timer
                    # itself moves this — manual logs never touch it.)
                    next_due = time.monotonic() + INTERVAL_SECONDS
                    sys.stdout.write("> ")       # redraw the prompt
                    sys.stdout.flush()
                    continue

                # 2) Did the user type a command? (Short wait so we keep
                #    checking the timer and stay responsive to Ctrl+C.)
                try:
                    command = line_queue.get(timeout=0.5)
                except queue.Empty:
                    continue
                got_command = True

            # ---- We have a command from the keyboard ----
            if command is None:                  # EOF -> treat like quit
                break

            cmd = command.strip().lower()
            if cmd in ("quit", "exit", "q"):
                break
            elif cmd in ("", "log"):
                # Manual trigger. NOTE: we deliberately do NOT change
                # next_due here, so the regular timer keeps its schedule.
                do_log(line_queue, "manual")
            else:
                print("  Unknown command. Press Enter or type 'log' to log; "
                      "type 'quit' to exit.")

    except KeyboardInterrupt:
        print()  # move to a fresh line after the ^C

    finally:
        # Always record the end of the session, however we got here.
        write_session_end()
        print("Session ended at " + now_hhmm() + ". Log saved to:")
        print("  " + LOG_PATH)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        # If something unexpected goes wrong, keep the window open long
        # enough to read the message when launched by double-click.
        print("Unexpected error: " + repr(exc))
        try:
            input("Press Enter to close...")
        except Exception:
            pass
