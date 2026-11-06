"""Build Silas's weekly activity log (Saturday to Friday) as an Excel file
laid out like the Activity log on the "Digital Marketer" tab of the
HRCC H2 2026 Marketing Tracker (headers in row 63, data from row 64).

Inputs
  --buffer    Buffer list_posts response (JSON with "edges"), includeMetrics on
  --week      week JSON: {"start","end","posts":[...], "manual":[...]}
              posts:  {"match": caption prefix, "desc", "type", "unit",
                       "audience", "org"(optional, default Organic), "notes"}
              manual: {"date": "YYYY-MM-DD", "desc", "channel", "type", "unit",
                       "org", "audience", "impressions", "engagements",
                       "clicks", "leads", "spend", "notes"}
  --out       output .xlsx

Metric mapping from Buffer
  Impressions  Facebook and LinkedIn "impressions"; Instagram "views"
  Engagements  reactions + comments + shares (+ saves on Instagram)
  Clicks       Facebook "clicks" only (Buffer gives none for IG or LinkedIn)
"""
import argparse
import datetime as dt
import json
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ACCRA = ZoneInfo("Africa/Accra")
SERVICE = {"facebook": "Facebook", "instagram": "Instagram", "linkedin": "LinkedIn"}
CH_ORDER = {"Facebook": 0, "Instagram": 1, "LinkedIn": 2}

HEADERS = ["Date", "Activity / content description", "Channel / Platform",
           "Content type", "Unit", "Organic / Paid", "Target audience",
           "Impressions", "Engagements", "Clicks", "CTR %", "Leads",
           "Ad Spend ($)", "Cost / Lead", "Notes", "Kind (auto)"]
WIDTHS = [11.4, 46, 18, 16, 18, 14, 26, 12, 12, 9, 9, 8, 11, 10, 40, 11]
FORMULA_COLS = {11, 14, 16}  # K, N, P live as formulas in the tracker

LISTS = {
    "C": ["Facebook", "Instagram", "LinkedIn", "TikTok", "X (Twitter)", "YouTube",
          "WhatsApp", "SMS", "Email", "Website / Blog", "Google Ads", "Other"],
    "D": ["Flyer", "Educational flyer", "Video", "Carousel", "Countdown post",
          "Text / message", "Email blast", "Story", "Picture", "Other"],
    "E": ["Main Certification", "Admin", "ACHR", "Mentorship", "Consulting", "Other"],
    "F": ["Organic", "Paid / Boosted", "Ad"],
    "G": ["HR professionals", "People managers & supervisors",
          "Functional heads / leaders", "Students & early‑career",
          "Public sector HR", "Senior HR & analysts", "Alumni",
          "SME owners / entrepreneurs", "Existing cohort", "Warm leads",
          "General public", "Leads from webinar", "Leads from programme launch",
          "Mass Target"],
}

NAVY = "1B2A4A"
CREAM = "FFF8E7"
GREY = "E7E9EE"


def metrics(node):
    m = {x["type"]: x["value"] for x in node.get("metrics") or []}
    svc = node["channelService"]
    imp = m.get("views") if svc == "instagram" else m.get("impressions")
    eng = sum(m.get(k) or 0 for k in ("reactions", "comments", "shares", "saves"))
    clicks = m.get("clicks") if svc == "facebook" else None
    return imp, eng, clicks, m


def load_posts(path):
    raw = open(path, encoding="utf-8").read()
    data = json.loads(raw)
    if isinstance(data, list):  # tool wrapper: [{"type":"text","text":"{...}"}]
        data = json.JSONDecoder().raw_decode(data[0]["text"])[0]
    return [e["node"] for e in data["edges"]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--buffer", required=True)
    ap.add_argument("--week", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    week = json.load(open(a.week, encoding="utf-8"))
    start = dt.date.fromisoformat(week["start"])
    end = dt.date.fromisoformat(week["end"])
    rows, check, unmatched = [], [], []

    for n in load_posts(a.buffer):
        if n["status"] != "sent":
            continue
        when = dt.datetime.fromisoformat(n["sentAt"].replace("Z", "+00:00")).astimezone(ACCRA)
        if not start <= when.date() <= end:
            continue
        text = (n.get("text") or "").strip()
        rule = next((p for p in week["posts"] if p["match"].lower() in text.lower()), None)
        if rule is None:
            unmatched.append(text[:80])
            continue
        ch = SERVICE[n["channelService"]]
        imp, eng, clicks, m = metrics(n)
        rows.append(dict(date=when.date(), time=when, desc=rule["desc"], channel=ch,
                         type=rule["type"], unit=rule["unit"],
                         org=rule.get("org", "Organic"), audience=rule["audience"],
                         impressions=imp, engagements=eng, clicks=clicks,
                         leads=None, spend=None, notes=rule.get("notes")))
        check.append(dict(when=when, channel=ch, desc=rule["desc"], via=n.get("via"),
                          assets=len(n.get("assets") or []), raw=m,
                          text=text[:120].replace("\n", " "),
                          updated=n.get("metricsUpdatedAt")))

    if unmatched:
        raise SystemExit("No classification for:\n  " + "\n  ".join(unmatched))

    for x in week.get("manual", []):
        rows.append(dict(date=dt.date.fromisoformat(x["date"]), time=None,
                         desc=x["desc"], channel=x.get("channel", "Other"),
                         type=x.get("type", "Other"), unit=x.get("unit", "Other"),
                         org=x.get("org"), audience=x.get("audience"),
                         impressions=x.get("impressions"), engagements=x.get("engagements"),
                         clicks=x.get("clicks"), leads=x.get("leads"),
                         spend=x.get("spend"), notes=x.get("notes")))

    # Order: date, then first time a post went out that day, then FB/IG/LI
    first_time = {}
    for r in rows:
        k = (r["date"], r["desc"])
        t = r["time"] or dt.datetime.combine(r["date"], dt.time(23, 59), ACCRA)
        first_time[k] = min(first_time.get(k, t), t)
    rows.sort(key=lambda r: (r["date"], first_time[(r["date"], r["desc"])],
                             r["desc"], CH_ORDER.get(r["channel"], 9)))

    wb = Workbook()
    ws = wb.active
    ws.title = "Activity log"
    thin = Side(style="thin", color="C9CED8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = f"Weekly activity log  -  Silas  -  {start:%a %d %b} to {end:%a %d %b %Y}"
    ws["A1"].font = Font(bold=True, size=13, color=NAVY)
    ws["A2"] = ("Paste into the Digital Marketer tab from the first empty row of the Activity log. "
                "Paste in three blocks so the tracker's formulas stay: A to J, then L to M, then O. "
                "Grey columns (K, N, P) are formulas in the live sheet; do not paste them.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:P2")
    ws.row_dimensions[2].height = 32

    hr = 4
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(hr, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY if i not in FORMULA_COLS else "6B7280")
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS[i - 1]
    ws.row_dimensions[hr].height = 30

    first = hr + 1
    for j, r in enumerate(rows):
        rr = first + j
        vals = [r["date"], r["desc"], r["channel"], r["type"], r["unit"], r["org"],
                r["audience"], r["impressions"], r["engagements"], r["clicks"],
                None, r["leads"], r["spend"], None, r["notes"], None]
        for i, v in enumerate(vals, 1):
            c = ws.cell(rr, i, v)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=(i in (2, 15)))
            c.fill = PatternFill("solid", fgColor=GREY if i in FORMULA_COLS else CREAM)
        ws.cell(rr, 1).number_format = "d/m/yyyy"
        # Same formulas as the tracker, so this file reads the same on its own
        ws.cell(rr, 11, f'=IFERROR(J{rr}/H{rr},"")').number_format = "0.00%"
        ws.cell(rr, 14, f'=IFERROR(M{rr}/L{rr},"")').number_format = "0.00"
        ws.cell(rr, 16, f'=IF(C{rr}="","",IF(OR(C{rr}="Email",C{rr}="SMS",C{rr}="WhatsApp"),"Message","Post"))')
    last = first + max(len(rows), 1) - 1 + 25  # spare rows for additions

    for col, items in LISTS.items():
        dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
        dv.add(f"{col}{first}:{col}{last}")
        ws.add_data_validation(dv)
    ws.freeze_panes = ws.cell(first, 3)

    # Summary under the table
    srow = first + len(rows) + 2
    posts = [r for r in rows if r["channel"] in SERVICE.values()]
    other = [r for r in rows if r not in posts]
    summary = [("Buffer posts (FB, IG, LinkedIn)", len(posts)),
               ("Other activities", len(other)),
               ("Total rows", len(rows))]
    for ch in ("Facebook", "Instagram", "LinkedIn"):
        summary.append((f"  {ch} posts", sum(1 for r in posts if r["channel"] == ch)))
    for k, (label, v) in enumerate(summary):
        ws.cell(srow + k, 2, label).font = Font(bold=(k < 3))
        ws.cell(srow + k, 3, v)

    # Check sheet: where every number came from
    cs = wb.create_sheet("Buffer check")
    ch_headers = ["Sent (Accra)", "Channel", "Log description", "Posted via",
                  "Media files", "Reactions", "Comments", "Shares", "Saves",
                  "Impressions", "Views", "Reach", "Clicks", "Eng. rate %",
                  "Caption start", "Metrics updated"]
    for i, h in enumerate(ch_headers, 1):
        c = cs.cell(1, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
    for j, x in enumerate(sorted(check, key=lambda x: x["when"]), 2):
        m = x["raw"]
        vals = [x["when"].strftime("%a %d/%m %H:%M"), x["channel"], x["desc"],
                "Buffer" if x["via"] == "buffer" else "Native (imported)", x["assets"],
                m.get("reactions"), m.get("comments"), m.get("shares"), m.get("saves"),
                m.get("impressions"), m.get("views"), m.get("reach"), m.get("clicks"),
                round(m["engagementRate"], 2) if m.get("engagementRate") is not None else None,
                x["text"], x["updated"]]
        for i, v in enumerate(vals, 1):
            cs.cell(j, i, v)
    for i, w in enumerate([16, 11, 40, 16, 8, 9, 9, 8, 7, 11, 9, 8, 8, 10, 60, 24], 1):
        cs.column_dimensions[get_column_letter(i)].width = w
    cs.freeze_panes = "A2"

    wb.save(a.out)
    print(f"{len(posts)} Buffer rows, {len(other)} other rows -> {a.out}")


if __name__ == "__main__":
    main()
