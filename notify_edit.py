#!/usr/bin/env python3
"""
Notification editor for macOS.

  python3 notify_edit.py           edits every notification; one line per edit
  python3 notify_edit.py --debug   same, plus details: the notification as received,
                                   the edited version, the raw database record, and
                                   every script that is executed

Needs Full Disk Access (reads the notification database) and Accessibility
(reads alerts from the screen, e.g. Tractive) for the app running it.
Run only this script, not older versions, or you will get duplicates.

You normally only change edit() below.
"""
import os
import plistlib
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import traceback
from collections import Counter
from datetime import datetime

DB_DIR = os.path.expanduser("~/Library/Group Containers/group.com.apple.usernoted/db2")
POLL_SECONDS = 1.0
MARK = "[Edited] "          # prefix on every edited title; prevents feedback loops
DEFAULT_SOUND = "Frog"
DEBUG = "--debug" in sys.argv

LOG_FILE = "/tmp/notifier-interceptor.log"


def timestamp():
    """Current time as "<day> H:MM AM/PM", for example "Wed 3:45 PM" (hour not zero-padded)."""
    now = datetime.now()
    day = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")[now.weekday()]
    return f"{day} {now.hour % 12 or 12}:{now.minute:02d} {'AM' if now.hour < 12 else 'PM'}"


def write(msg):
    """
    Append one line to the log file, stamped "<day> H:MM AM/PM": the three-letter
    day name, then the 12-hour time with no zero padding on the hour, for example:
        Wed 9:05 AM  your message
        Wed 3:45 PM  your message
    Can be called from anywhere, including edit(). A logging problem never
    stops the script.
    """
    stamp = timestamp()
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{stamp}  {msg}\n")
    except OSError as e:
        print(f"(could not write to {LOG_FILE}: {e})")

# Friendly names for database app ids. Run with --debug to see the id of any app
# not listed here, then add it. Unlisted apps appear in n.app under their raw id.
APP_NAMES = {
    "com.apple.reminders": "reminders",
    "com.apple.news": "news",
    "org.mozilla.firefox": "firefox",
}

# Notifications containing any of these strings are ignored completely: edit() is
# not called and nothing is posted. Matching is case-insensitive and looks for the
# string anywhere in the title, subtitle, body or app id (database app id, e.g.
# "com.apple.tmhelperagent"). Empty strings are ignored.
BLACKLIST = [
    # "tmhelperagent",      # Time Machine
    # "diskarbitration",    # disk notices
    # "some text in a body",
]

# Label lines macOS shows above a title on screen (lowercase). Add other languages here.
LABELS = {"time sensitive", "critical"}


class Notification:
    """What edit() receives. See the comment above edit() for the fields."""

    def __init__(self, source, app_id, app, site, title, subtitle, body, raw):
        self.source, self.app_id, self.app, self.site, self.raw = source, app_id, app, site, raw
        self.title, self.subtitle, self.body = title, subtitle, body
        self.sound = DEFAULT_SOUND
        self.send = None

    @property
    def text(self):
        return " ".join((self.title, self.subtitle, self.body)).lower()

    def matches(self, pattern):
        return re.search(pattern, " ".join((self.title, self.subtitle, self.body)), re.I) is not None

    def state(self):
        return (self.title, self.subtitle, self.body, self.sound or DEFAULT_SOUND)


# ===========================================================================
# EDIT HERE.  edit(n) is called once for every notification.
#
# Change fields on n; that is all. If anything changed (title, subtitle, body or
# sound), an edited notification is posted. If nothing changed, nothing is posted.
#
# Read and write:
#   n.title, n.subtitle, n.body   the texts ("" when absent)
#   n.sound                       sound of the edited notification (default "Frog")
#   n.send                        None = automatic (post only if something changed)
#                                 False = never post, True = always post
# Read only:
#   n.app      "reminders", "news", "firefox", "safari" (website notification),
#              "applescript" (osascript / Script Editor), "system"; "unknown" for
#              alerts read from the screen, because the screen does not say which
#              app posted them; any other database app appears under its raw id
#              (give it a name in APP_NAMES).
#   n.site     host name of a Safari website notification ("localhost",
#              "my.tractive.com"), otherwise "".
#   n.source   "db" (found in the notification database) or "screen" (read from
#              the screen, because the database never recorded it)
#   n.text     title, subtitle and body joined and lowercased, for simple tests
#   n.matches(regex)   True if the regex is found, ignoring case, in the title,
#              subtitle or body
#   n.app_id, n.raw    the raw database id and the complete database record
#              (title/subtitle/body are three entries taken from it; the rest are
#              extra fields). Rarely needed: run with --debug to see what raw holds.
#
# Examples:
#   if n.app == "reminders":                    n.title = "Task: " + n.title
#   if n.app == "safari" and n.site == "localhost":  n.sound = "Glass"
#   if n.app == "firefox":                      n.send = False   # leave alone
#   if n.app == "applescript":                  n.body = n.body.upper()
#   if n.app == "system":                       n.send = False
#   if n.source == "screen":                    ...  # only alerts the database never saw
#   if "keyword" in n.text:                     ...
#   if n.matches(r"\bout\b"):                   ...  # whole word
#   write("text")                               append a timestamped line to the log file
#   timestamp()                                 the current time as text, e.g. "Wed 3:45 PM"
# ===========================================================================
def edit(n):
    # Tractive: not in the database, so it arrives from the screen. Match by text.
    if n.matches(r"safe zone"):
        stamp = timestamp()
        msg = None
        if n.matches(r"back in|entered"):
            msg = stamp + ": Shénah is in"
            n.sound, n.subtitle = "shenah-in", msg
        elif n.matches(r"\b(out|left|exited)\b"):
            msg = stamp + ": Shénah is out"
            n.sound, n.subtitle = "shenah-out", msg
        if msg:
            write(msg)
        return

    # Everything else: prepend "Body:  " to the body
    n.body = "Body:  " + n.body


# ===========================================================================
# Below: plumbing. Normally no need to change anything.
# ===========================================================================
SEP = "|||"
GROUP = "###"
SCREEN_DELAY = 2.0      # screen alerts wait this long so the database source goes first
DEDUP_SECONDS = 30      # the same notification is not handled by both sources within this time
INVISIBLE = re.compile("[\u200e\u200f\u202a-\u202e\u2066-\u2069]")


def dbg(*args):
    if DEBUG:
        print(*args)


def clean(s):
    return INVISIBLE.sub("", s).strip()


# ---- pre-edit: turn raw database / screen data into a simple Notification ----
def classify(app_id):
    """(friendly app name, site) for a database app id."""
    low = app_id.lower()
    if "scripteditor" in low or "osascript" in low:
        return "applescript", ""
    if low.startswith("_web_center_:web."):
        host = app_id.split("web.", 1)[1]
        return "safari", "" if host == "(null)" else host
    if low.startswith("_system_center_:"):
        return "system", ""
    return APP_NAMES.get(low, app_id), ""


def pre_edit(source, app_id, title, subtitle, body, raw):
    """Build the Notification for edit(), or return None if it is blacklisted."""
    haystack = "\n".join((app_id or "", title, subtitle, body)).lower()
    for entry in BLACKLIST:
        if entry and entry.lower() in haystack:
            dbg(f"skip (blacklisted by {entry!r}): {title!r} | {body!r}")
            return None
    if source == "db":
        app, site = classify(app_id)
    else:
        app, site = "unknown", ""
    return Notification(source, app_id, app, site, title, subtitle, body, raw)


# ---- duplicate protection between the two sources ----
def norm(s):
    return re.sub(r"\s+", " ", clean(s)).rstrip("….").strip().lower()


def same(a, b):
    a, b = norm(a), norm(b)          # tolerate banner truncation ("..." / "…")
    return a == b or (a and b and (a.startswith(b) or b.startswith(a)))


def recently(handled, title, body):
    now = time.time()
    return any(now - ts < DEDUP_SECONDS and same(t, title) and same(b, body)
               for (t, b), ts in handled.items())


def prune(handled):
    now = time.time()
    for k in [k for k, ts in handled.items() if now - ts >= DEDUP_SECONDS]:
        del handled[k]


# ---- database source ----
def read_rows():
    tmp = tempfile.mkdtemp()
    try:
        for name in ("db", "db-wal", "db-shm"):
            src = os.path.join(DB_DIR, name)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(tmp, name))
        con = sqlite3.connect(os.path.join(tmp, "db"))
        rows = con.execute(
            "SELECT r.rec_id, a.identifier, r.data "
            "FROM record r JOIN app a ON a.app_id = r.app_id").fetchall()
        con.close()
        return rows
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---- screen source: each alert is a run of static texts between AXGroup elements ----
READ_SCREEN_SCRIPT = '''
tell application "System Events"
	if not (exists process "NotificationCenter") then return ""
	tell process "NotificationCenter"
		set out to ""
		repeat with w in windows
			try
				if (name of w) is "Notification Center" then
					set els to entire contents of w
					repeat with e in els
						try
							set r to (role of e)
							if r is "AXStaticText" then
								set v to (value of e) as text
								set out to out & v & "|||"
							else if r is "AXGroup" then
								set out to out & "###|||"
							end if
						end try
					end repeat
				end if
			end try
		end repeat
		return out
	end tell
end tell
'''


def get_screen_alerts():
    """List of alerts on screen; each alert is a list of its texts."""
    r = subprocess.run(["osascript", "-e", READ_SCREEN_SCRIPT], capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError(r.stderr.strip())
    alerts, run = [], []
    for tok in (clean(x) for x in r.stdout.split(SEP)):
        if tok == GROUP:
            if run:
                alerts.append(run)
                run = []
        elif tok:
            run.append(tok)
    if run:
        alerts.append(run)
    return alerts


def parse_alert(texts):
    """(title, subtitle, body) from an alert's texts, or None if it has no body."""
    texts = list(texts)
    while texts and texts[0].lower() in LABELS:
        texts.pop(0)
    if len(texts) < 2:
        return None
    return texts[0], " ".join(texts[1:-1]), texts[-1]


def screen_candidates(alerts):
    found = Counter()
    for texts in alerts:
        parsed = parse_alert(texts)
        if parsed and not parsed[0].startswith(MARK):
            found[parsed] += 1
    return found


# ---- posting and dispatch ----
def post(n):
    q = lambda s: s.replace("\\", "\\\\").replace('"', '\\"')
    script = (f'display notification "{q(n.body)}" with title "{q(MARK + n.title)}" '
              f'sound name "{q(n.sound or DEFAULT_SOUND)}"')
    if n.subtitle:
        script += f' subtitle "{q(n.subtitle)}"'
    dbg("  executing: osascript -e", repr(script))
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if r.returncode:
        print("  osascript error:", r.stderr.strip())


def process(source, app_id, title, sub, body, raw):
    """Run edit() on one notification. True if an edited notification was posted."""
    n = pre_edit(source, app_id, title, sub, body, raw)
    if n is None:                      # blacklisted
        return False
    before = n.state()
    dbg(f"RECEIVED  source={source} app_id={app_id!r} -> app={n.app!r} site={n.site!r}")
    dbg(f"          title={title!r} subtitle={sub!r} body={body!r}")
    if raw is not None:
        dbg(f"          raw={dict(raw)!r}")
    try:
        edit(n)
    except Exception:
        print("edit() raised an error:")
        traceback.print_exc()
        return False
    send = n.send if n.send is not None else n.state() != before
    dbg(f"EDITED    title={n.title!r} subtitle={n.subtitle!r} body={n.body!r} "
        f"sound={n.sound!r} send={send}")
    if not send:
        dbg("          nothing changed -> not posting")
        return False
    print(f"edited [{n.app}] {title!r} | {sub!r} | {body!r}  =>  "
          f"{n.title!r} | {n.subtitle!r} | {n.body!r}")
    post(n)
    return True


# ---- main loop ----
def main():
    seen = set()
    db_seen = {}           # (title, body) -> time the database source saw it
    screen_handled = {}    # (title, body) -> time the screen source posted an edit
    use_db = use_screen = True

    try:
        for rec_id, _app, blob in read_rows():
            seen.add((rec_id, hash(blob)))
    except Exception as e:
        print("Cannot read the notification database:", e)
        print("Database source disabled. Check Full Disk Access for this app.")
        use_db = False

    active, last_alerts, pending = Counter(), [], []
    try:
        last_alerts = get_screen_alerts()
        active = screen_candidates(last_alerts)
    except RuntimeError as e:
        print("Cannot read Notification Center UI:", e)
        print("Screen source disabled. Check Accessibility permission for this app.")
        use_screen = False

    if not use_db and not use_screen:
        sys.exit(1)

    print(f"Watching (db {'on' if use_db else 'off'}, screen {'on' if use_screen else 'off'}); "
          f"ignoring {len(seen)} existing records and {sum(active.values())} alert(s) on screen. "
          f"Ctrl+C to stop.")
    if use_screen:
        dbg("screen reading script (runs every second):" + READ_SCREEN_SCRIPT)

    screen_warned = False
    while True:
        prune(db_seen)
        prune(screen_handled)

        if use_db:
            try:
                rows = read_rows()
            except Exception:
                print("db read error:")
                traceback.print_exc()
                rows = []
            for rec_id, app_id, blob in rows:
                key = (rec_id, hash(blob))
                if key in seen:
                    continue
                seen.add(key)
                try:
                    raw = plistlib.loads(blob).get("req", {})
                    title = str(raw.get("titl") or "")
                    sub = str(raw.get("subt") or "")
                    body = str(raw.get("body") or "")
                    if title.startswith(MARK):
                        dbg(f"skip (our own edited notification): {title!r}")
                        continue
                    db_seen[(title, body)] = time.time()
                    if recently(screen_handled, title, body):
                        dbg(f"skip (already handled from the screen): {title!r}")
                        continue
                    process("db", app_id, title, sub, body, raw)
                except Exception:
                    print("error handling a database record:")
                    traceback.print_exc()

        if use_screen:
            try:
                alerts = get_screen_alerts()
                screen_warned = False
                if alerts != last_alerts:
                    dbg("on screen:", alerts)
                last_alerts = alerts
                current = screen_candidates(alerts)
                now = time.time()
                for sig, n in current.items():
                    for _ in range(max(0, n - active.get(sig, 0))):
                        pending.append((now, sig))
                active = current

                still = []
                for t0, sig in pending:
                    if now - t0 < SCREEN_DELAY:
                        still.append((t0, sig))
                        continue
                    title, sub, body = sig
                    if recently(db_seen, title, body):
                        dbg(f"skip (handled via the database): {title!r}")
                        continue
                    if process("screen", None, title, sub, body, None):
                        screen_handled[(title, body)] = now
                pending = still
            except RuntimeError as e:
                if not screen_warned:
                    print("screen read error:", e)
                    screen_warned = True
            except Exception:
                print("error handling screen alerts:")
                traceback.print_exc()

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
