#!/usr/bin/env python3
"""
dump_notifications.py

Dumps notification information into a spreadsheet and opens it in Numbers:
  sheet "Database"   the newest records in the Notification Center database
  sheet "UI alerts"  the alerts currently shown in the Notification Center window

  python3 dump_notifications.py        newest 20 database records
  python3 dump_notifications.py 50     newest 50 database records

Needs, for the app you run it from (Terminal or VS Code):
  Full Disk Access   (database sheet)   System Settings > Privacy & Security
  Accessibility      (UI sheet)         System Settings > Privacy & Security
If one of them is missing, that sheet explains the problem and the other still works.

No extra Python packages are needed: the .xlsx file is written with the standard library.
"""
import datetime
import os
import plistlib
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

DB_DIR = os.path.expanduser("~/Library/Group Containers/group.com.apple.usernoted/db2")
OUTPUT_DIR = "/tmp"                 # the spreadsheet is saved here, with a time-stamped name
CORE_DATA_EPOCH = 978307200         # database dates count seconds from 2001-01-01
MARK = "[Edited] "                  # prefix notify_edit.py puts on its own notifications
LABELS = {"time sensitive", "critical"}   # label lines macOS shows above a title on screen

APP_NAMES = {
    "com.apple.reminders": "Reminders",
    "com.apple.news": "News",
    "org.mozilla.firefox": "Firefox",
}

INVISIBLE = re.compile("[\u200e\u200f\u202a-\u202e\u2066-\u2069]")


def clean(s):
    return INVISIBLE.sub("", s).strip()


# ---------------------------------------------------------------------------
# Database sheet
# ---------------------------------------------------------------------------
def classify(app_id):
    """(friendly app name, website host) for a database app id."""
    low = (app_id or "").lower()
    if "scripteditor" in low or "osascript" in low:
        return "AppleScript", ""
    if low.startswith("_web_center_:web."):
        host = app_id.split("web.", 1)[1]
        return "Safari website", "" if host == "(null)" else host
    if low.startswith("_system_center_:"):
        return "System", ""
    return APP_NAMES.get(low, app_id or ""), ""


def to_datetime(value):
    if not value:
        return None
    try:
        return datetime.datetime.fromtimestamp(value + CORE_DATA_EPOCH)
    except (OverflowError, OSError, ValueError, TypeError):
        return None


def short(v):
    s = repr(v)
    return s if len(s) <= 200 else s[:200] + "..."


def field_lines(d, skip):
    """'key: value' lines for every entry of a dict except the skipped keys."""
    if not isinstance(d, dict):
        return ""
    return "\n".join(f"{k}: {short(d[k])}" for k in sorted(d, key=str) if k not in skip)


def db_columns(date_headers):
    return (["Record #"] + date_headers +
            ["App", "Title", "Subtitle", "Body", "Website", "App ID",
             "Other request fields", "Other record fields"])


def read_database(limit):
    """(columns, rows, note). note is set when there is nothing to show."""
    tmp = tempfile.mkdtemp()
    columns = db_columns(["Delivered"])
    try:
        for name in ("db", "db-wal", "db-shm"):
            src = os.path.join(DB_DIR, name)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(tmp, name))
        con = sqlite3.connect(os.path.join(tmp, "db"))
        have = {r[1] for r in con.execute("PRAGMA table_info(record)")}
        date_cols = [(c, h) for c, h in (("delivered_date", "Delivered"),
                                         ("request_date", "Requested")) if c in have]
        columns = db_columns([h for _, h in date_cols])
        select = "r.rec_id, a.identifier, r.data" + "".join(f", r.{c}" for c, _ in date_cols)
        fetched = con.execute(
            f"SELECT {select} FROM record r JOIN app a ON a.app_id = r.app_id "
            "ORDER BY r.rec_id DESC LIMIT ?", (limit,)).fetchall()
        con.close()
    except Exception as e:
        return columns, [], (f"Could not read the notification database: {e}. "
                             "Check Full Disk Access for this app.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    rows = []
    for rec_id, ident, blob, *dates in fetched:
        app, site = classify(ident)
        try:
            plist = plistlib.loads(blob)
            req = plist.get("req", {})
            if not isinstance(req, dict):
                req = {}
            title, sub, body = (str(req.get(k) or "") for k in ("titl", "subt", "body"))
            other_req = field_lines(req, ("titl", "subt", "body"))
            other_rec = field_lines(plist, ("req",))
        except Exception as e:
            title = sub = other_req = other_rec = ""
            body = f"(cannot parse this record's data: {e})"
        rows.append([rec_id] + [to_datetime(d) for d in dates] +
                    [app, title, sub, body, site, ident, other_req, other_rec])
    return columns, rows, "" if rows else "(the database has no records)"


# ---------------------------------------------------------------------------
# UI sheet: each alert is a run of static texts between AXGroup elements
# ---------------------------------------------------------------------------
UI_COLUMNS = ["Alert #", "Label", "Title", "Subtitle", "Body", "Notes",
              "All texts for this alert"]

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


def read_ui():
    """(rows, note). note is set when there is nothing to show."""
    r = subprocess.run(["osascript", "-e", READ_SCREEN_SCRIPT], capture_output=True, text=True)
    if r.returncode:
        return [], (f"Could not read the Notification Center UI: {r.stderr.strip()} "
                    "Check Accessibility permission for this app.")
    alerts, cur = [], []
    for tok in (clean(x) for x in r.stdout.split("|||")):
        if tok == "###":
            if cur:
                alerts.append(cur)
                cur = []
        elif tok:
            cur.append(tok)
    if cur:
        alerts.append(cur)

    rows = []
    for i, texts in enumerate(alerts, 1):
        rest, labels = list(texts), []
        while rest and rest[0].lower() in LABELS:
            labels.append(rest.pop(0))
        title = sub = body = notes = ""
        if len(rest) >= 2:
            title, sub, body = rest[0], " ".join(rest[1:-1]), rest[-1]
        elif rest:
            title = rest[0]
        if len(rest) < 2:
            notes = "only one text: notify_edit.py ignores it"
        elif title.startswith(MARK):
            notes = "edited by notify_edit.py: ignored"
        rows.append([i, ", ".join(labels), title, sub, body, notes, " | ".join(texts)])
    return rows, "" if rows else "(no alerts on screen)"


# ---------------------------------------------------------------------------
# Minimal .xlsx writer (standard library only)
# ---------------------------------------------------------------------------
WIDTHS = {   # column widths, in characters
    "Record #": 10, "Delivered": 26, "Requested": 26, "App": 16, "Title": 30,
    "Subtitle": 24, "Body": 50, "Website": 20, "App ID": 36,
    "Other request fields": 44, "Other record fields": 36,
    "Alert #": 10, "Label": 16, "Notes": 34, "All texts for this alert": 60,
}

STYLE_HEADER, STYLE_TEXT, STYLE_DATE, STYLE_INT = 1, 2, 3, 4

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="ddd mmm d, yyyy h:mm AM/PM"/></numFmts>
<fonts count="2">
<font><sz val="10"/><name val="Arial"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2F4F6F"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border>
<left style="thin"><color rgb="FFC8C8C8"/></left><right style="thin"><color rgb="FFC8C8C8"/></right>
<top style="thin"><color rgb="FFC8C8C8"/></top><bottom style="thin"><color rgb="FFC8C8C8"/></bottom><diagonal/>
</border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>
<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def xml_text(s):
    s = s.encode("utf-8", "ignore").decode("utf-8")      # drop lone surrogates
    return escape(BAD_XML.sub("", s))


def col_letter(i):
    i += 1
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def row_height(values, widths):
    """Estimated height so wrapped text is fully visible in Numbers and Excel."""
    lines = 1
    for v, w in zip(values, widths):
        if isinstance(v, str) and v:
            per_line = max(1, int(w) - 2)
            lines = max(lines, sum(max(1, -(-len(part) // per_line)) for part in v.split("\n")))
    return min(400, 13 * lines + 5)


def write_xlsx(path, sheets):
    """sheets: list of (name, columns, rows). Row values: str, int, datetime or None."""
    strings, index = [], {}

    def sst(s):
        if s not in index:
            index[s] = len(strings)
            strings.append(s)
        return index[s]

    sheet_xml = []
    for n, (name, columns, rows) in enumerate(sheets):
        widths = [WIDTHS.get(c, 20) for c in columns]
        selected = ' tabSelected="1"' if n == 0 else ""
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               f'<worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"{selected}>'
               '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
               '<selection pane="bottomLeft"/></sheetView></sheetViews>'
               '<sheetFormatPr defaultRowHeight="15"/><cols>']
        for i, w in enumerate(widths, 1):
            out.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
        out.append('</cols><sheetData>')
        out.append('<row r="1" ht="28" customHeight="1">' + "".join(
            f'<c r="{col_letter(i)}1" s="{STYLE_HEADER}" t="s"><v>{sst(xml_text(c))}</v></c>'
            for i, c in enumerate(columns)) + '</row>')
        for r, values in enumerate(rows, 2):
            cells = []
            for i, v in enumerate(values):
                ref = f"{col_letter(i)}{r}"
                if v is None or v == "":
                    cells.append(f'<c r="{ref}" s="{STYLE_TEXT}"/>')
                elif isinstance(v, datetime.datetime):
                    serial = (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
                    cells.append(f'<c r="{ref}" s="{STYLE_DATE}"><v>{serial:.10f}</v></c>')
                elif isinstance(v, int) and not isinstance(v, bool):
                    cells.append(f'<c r="{ref}" s="{STYLE_INT}"><v>{v}</v></c>')
                else:
                    cells.append(f'<c r="{ref}" s="{STYLE_TEXT}" t="s"><v>{sst(xml_text(str(v)))}</v></c>')
            out.append(f'<row r="{r}" ht="{row_height(values, widths)}" customHeight="1">'
                       + "".join(cells) + '</row>')
        out.append('</sheetData></worksheet>')
        sheet_xml.append("".join(out))

    count = len(sheets)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                  for i in range(1, count + 1)) +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" '
        f'Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS}" xmlns:r="{REL_NS}"><sheets>'
        + "".join(f'<sheet name="{xml_text(name)}" sheetId="{i}" r:id="rId{i}"/>'
                  for i, (name, _, _) in enumerate(sheets, 1)) + '</sheets></workbook>')
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL}">'
        + "".join(f'<Relationship Id="rId{i}" Type="{REL_NS}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, count + 1)) +
        f'<Relationship Id="rId{count + 1}" Type="{REL_NS}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId{count + 2}" Type="{REL_NS}/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>')
    shared = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS}" count="{len(strings)}" uniqueCount="{len(strings)}">'
        + "".join(f'<si><t xml:space="preserve">{s}</t></si>' for s in strings) + '</sst>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", STYLES_XML)
        z.writestr("xl/sharedStrings.xml", shared)
        for i, xml in enumerate(sheet_xml, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)


# ---------------------------------------------------------------------------
def with_note(columns, rows, note, note_column="Body"):
    """If there is a note (error or 'nothing found'), show it as a row under the headers."""
    if not note:
        return rows
    row = [""] * len(columns)
    row[columns.index(note_column)] = note
    return [row]


def main():
    limit = 20
    if len(sys.argv) > 1:
        try:
            limit = max(1, int(sys.argv[1]))
        except ValueError:
            print(f"Ignoring '{sys.argv[1]}': expected a number of records, using {limit}.")

    db_cols, db_rows, db_note = read_database(limit)
    ui_rows, ui_note = read_ui()

    path = os.path.join(OUTPUT_DIR,
                        f"notification-dump-{datetime.datetime.now():%Y%m%d-%H%M%S}.xlsx")
    write_xlsx(path, [
        ("Database", db_cols, with_note(db_cols, db_rows, db_note)),
        ("UI alerts", UI_COLUMNS, with_note(UI_COLUMNS, ui_rows, ui_note)),
    ])

    print(f"Saved {path}")
    print(f"  Database sheet: {len(db_rows)} record(s)" + (f"  [{db_note}]" if db_note else ""))
    print(f"  UI alerts sheet: {len(ui_rows)} alert(s)" + (f"  [{ui_note}]" if ui_note else ""))
    if subprocess.run(["open", "-a", "Numbers", path]).returncode != 0:
        print("Could not open Numbers; opening with the default app instead.")
        subprocess.run(["open", path])


if __name__ == "__main__":
    main()
