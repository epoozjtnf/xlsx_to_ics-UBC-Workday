#!/usr/bin/env python3
import sys, os, csv, argparse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

CRLF = "\r\n"
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
SECTION_COL, DATE_COL, FORMAT_COL, DELIVERY_COL, INSTRUCTOR_COL = 'G', 'K', 'I', 'J', 'L'
SKIP_BEFORE = 4 # Skip rows before actual data
TEMP_FILE, XLSX_FILE, ICS_FILE = "temp_calendar.csv", "View_My_Courses.xlsx", "My Enrolled Courses.ics"
DEFAULT_TZ = "America/Vancouver"
_DOW = dict(MO=0, TU=1, WE=2, TH=3, FR=4, SA=5, SU=6)

def to_24h(t):
    return datetime.strptime(t.replace('.', '').replace('  ', ' ').upper().strip(), "%I:%M %p").strftime("%H:%M")

def reformat(item):
    month, day, year = item.split('/')
    return year + "-" + month + "-" + day

def parse_date_location(s):
    temp = s.split('|')
    date_part = temp[0]
    days_part = temp[1]
    time_part = temp[2]
    loc = " ".join(temp[3:])
    start_date, end_date = (reformat(d.strip()) for d in date_part.split(' - '))
    start_time, end_time = (t.strip() for t in time_part.split(' - '))
    start_dt = f"{start_date}T{to_24h(start_time)}"
    end_dt   = f"{start_date}T{to_24h(end_time)}"
    day_map = {'Mon':'MO','Tue':'TU','Wed':'WE','Thu':'TH','Fri':'FR','Sat':'SA','Sun':'SU'}
    rrule = f"FREQ=WEEKLY;UNTIL={end_date.replace('-','')}T235959Z;BYDAY={','.join(day_map[d] for d in days_part.split())}"
    return start_dt, end_dt, loc, rrule

def make_event(cat, start, end, summary, desc, loc, rrule, color=""):
    q = lambda x: f'"{x}"'
    return {"category": q(cat), "start": q(start), "end": q(end), "summary": q(summary),
            "description": q(desc), "location": q(loc), "rrule": q(rrule), "color": q(color)}

def read_shared_strings(zf):
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [si.find("main:t", NS).text for si in root.iterfind("main:si", NS)]

def read_sheet_data(zf, sst):
    root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    data = {}
    for c in root.iterfind('.//main:c', NS):
        if (v := c.find('main:v', NS)) is None: continue
        data[c.attrib['r']] = sst[int(v.text)] if c.attrib.get('t')=='s' else v.text
    return data

def build_calendar(data):
    events = [{"category":"category","start":"start","end":"end","summary":"summary",
               "description":"description","location":"location","rrule":"rrule","color":"color"}]
    for row in sorted({cell[1:] for cell in data}):
        if not row.isdigit() or int(row)<SKIP_BEFORE: continue
        raw = data.get(f"{DATE_COL}{row}","").strip()
        if not raw: continue
        summary = data.get(f"{SECTION_COL}{row}", "")
        cat = data.get('A1', 'Courses')
        desc = ' '.join(filter(None, [data.get(f"{FORMAT_COL}{row}", ""),
                                      data.get(f"{DELIVERY_COL}{row}", ""),
                                      data.get(f"{INSTRUCTOR_COL}{row}", "")]))
        for part in raw.split('\n'):
            part = part.strip()
            if not part: continue
            try:
                start, end, loc, rrule = parse_date_location(part)
            except Exception:
                continue
            events.append(make_event(cat, start, end, summary, desc, loc, rrule))
    return events

def write_csv(events, fname):
    with open(fname, 'w', newline='') as f:
        for ev in events:
            row = [ev[k] for k in ['category','start','end','summary','description','location','rrule','color']]
            f.write(','.join(row)+'\n')

def format_dt(dt):
    return dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

def ical_escape(text):
    return text.replace("\\", "\\\\").replace(",", r"\\,").replace(";", r"\\;")\
               .replace("\r\n", r"\\n").replace("\n", r"\\n")

def fold_lines(lines):
    FOLD_BYTES = 75
    folded = []
    for line in lines:
        while len(line.encode('utf-8')) > FOLD_BYTES:
            cut = FOLD_BYTES
            while len(line[:cut].encode('utf-8')) > FOLD_BYTES:
                cut -= 1
            folded.append(line[:cut])
            line = ' ' + line[cut:]
        folded.append(line)
    return folded

def build_vevent(ev, default_tz=timezone.utc):
    start, end = ev["start"], ev["end"]
    if ev.get("rrule") and "BYDAY=" in ev["rrule"]:
        offset = 7
        for day in [day for day in ev['rrule'].split("BYDAY=")[1].split(";")[0].split(",")]:
            offset = min(offset, (_DOW[day] - start.weekday()) % 7)
        start, end = start + timedelta(days=offset), end + timedelta(days=offset)
    uid = ev.get("uid") or f"{start.timestamp()}-{id(ev)}@unified_ics"
    tzid = ev.get("zone") or default_tz.tzname(None)
    lines = [
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{format_dt(datetime.utcnow())}",
        f"DTSTART;TZID={tzid}:{start.strftime('%Y%m%dT%H%M%S')}",
        f"DTEND;TZID={tzid}:{end.strftime('%Y%m%dT%H%M%S')}",
        f"SUMMARY:{ical_escape(ev.get('summary',''))}"
    ]
    if ev.get("description"):
        lines.append(f"DESCRIPTION:{ical_escape(ev['description'])}")
    if ev.get("location"):
        lines.append(f"LOCATION:{ical_escape(ev['location'])}")
    if ev.get("rrule"):
        lines.append(f"RRULE:{ev['rrule']}")
    if ev.get("color"):
        lines.append(f"X-APPLE-CALENDAR-COLOR:{ev['color']}")
    lines.append("END:VEVENT")
    return fold_lines(lines)

def _load_events_from_csv(src, default_tz):
    events = []
    with src.open() as f:
        for row in csv.DictReader(f):
            events.append({
                "category": row.get("category","Default"),
                "start": datetime.fromisoformat(row["start"].replace('/', '-')).replace(tzinfo=default_tz),
                "end": datetime.fromisoformat(row["end"].replace('/', '-')).replace(tzinfo=default_tz),
                "summary": row.get("summary",""),
                "description": row.get("description",""),
                "location": row.get("location",""),
                "rrule": row.get("rrule"),
                "color": row.get("color"),
                "zone": row.get("zone"),
            })
    return events

def main(argv=None):
    p = argparse.ArgumentParser(description="Generate RFC 5545 .ics calendars.")
    p.add_argument("--csv", default=TEMP_FILE, help="CSV input file")
    p.add_argument("--outfile", default=ICS_FILE, help="Output file path")
    p.add_argument("--input", default=XLSX_FILE, help="XLSX file path")
    args = p.parse_args(argv)

    if os.path.exists(args.outfile):
        if input(f"\"{args.outfile}\" exists. Type 'yes' to overwrite or anything else to cancel: ").strip().lower() != 'yes':
            return

    default_tz = datetime.now().astimezone().tzinfo

    if not os.path.exists(args.input):
        sys.exit(f"File not found: {args.input}")
    with ZipFile(args.input) as zf:
        sst = read_shared_strings(zf)
        data = read_sheet_data(zf, sst)

    calendar = build_calendar(data)
    target = args.csv
    if os.path.exists(target):
        if input(f"\"{TEMP_FILE}\" exists. Type 'yes' to overwrite or anything else to cancel: ").strip().lower() != 'yes':
            return
    write_csv(calendar, target)
    print(f"Written {len(calendar)} events to '{target}'")

    src = Path(target)
    if not src.exists():
        sys.exit(f"File not found: {src}")
    events = _load_events_from_csv(src, default_tz)
    calendars = {}
    for ev in events:
        calendars.setdefault(ev.pop("category"), []).append(build_vevent(ev, default_tz))
    for cat, vevents in calendars.items():
        cal = ["BEGIN:VCALENDAR","VERSION:2.0","CALSCALE:GREGORIAN",
               "PRODID:-//UnifiedICS//EN",f"X-WR-CALNAME:{ical_escape(cat)}",f"X-WR-TIMEZONE:{DEFAULT_TZ}"]
        for v in vevents: cal.extend(v)
        cal.append("END:VCALENDAR")
        fname = f"{cat}.ics"
        with open(fname, "w", newline="") as f: f.write(CRLF.join(cal))
        print(f"Wrote {fname} ({Path(fname).stat().st_size} bytes)")
    if os.path.exists(TEMP_FILE):
        os.remove(TEMP_FILE)
        print(f"Removed temporary file: {TEMP_FILE}")

if __name__ == "__main__":
    main()
