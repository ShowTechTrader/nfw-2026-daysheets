#!/usr/bin/env python3
"""
NFW 2026 Day Sheet Generator
Usage:  python generate_daysheet.py
        python generate_daysheet.py --excel schedule.xlsx --notes day_notes.json --output daysheet.html

Reads the production schedule Excel, fetches live weather from Open-Meteo,
reads editorial content from day_notes.json, and outputs a fully pre-rendered
HTML file that works in OneDrive without any JavaScript rendering.

Requirements:
    pip install openpyxl requests Pillow numpy
"""

import argparse, base64, html, io, json, sys
from datetime import date, timedelta, datetime as dt
from pathlib import Path

NOOSA_LAT = -26.393
NOOSA_LON = 153.073

DAYS_META = [
    {"day":1,"sheet":"Monday 8th",    "date":"Monday 8th June",    "short":"Mon 8",  "type":"Bump-In",      "iso":"2026-06-08"},
    {"day":2,"sheet":"Tuesday 9th",   "date":"Tuesday 9th June",   "short":"Tue 9",  "type":"Bump-In",      "iso":"2026-06-09"},
    {"day":3,"sheet":"Wednesday 10th","date":"Wednesday 10th June", "short":"Wed 10", "type":"Bump-In",      "iso":"2026-06-10"},
    {"day":4,"sheet":"Thursday 11th", "date":"Thursday 11th June",  "short":"Thu 11", "type":"Bump-In",      "iso":"2026-06-11"},
    {"day":5,"sheet":"Friday 12th",   "date":"Friday 12th June",    "short":"Fri 12", "type":"Bump-In",      "iso":"2026-06-12"},
    {"day":6,"sheet":"Saturday 13th", "date":"Saturday 13th June",  "short":"Sat 13", "type":"Festival Day", "iso":"2026-06-13"},
    {"day":7,"sheet":"Sunday 14th",   "date":"Sunday 14th June",    "short":"Sun 14", "type":"Wrap Day",     "iso":"2026-06-14"},
    {"day":8,"sheet":"Monday 15th",   "date":"Monday 15th June",    "short":"Mon 15", "type":"Bump-Out",     "iso":"2026-06-15"},
    {"day":9,"sheet":"Tuesday 16th",  "date":"Tuesday 16th June",   "short":"Tue 16", "type":"Bump-Out",     "iso":"2026-06-16"},
]

WMO_DESC = {0:"Clear sky",1:"Mainly clear",2:"Partly cloudy",3:"Overcast",
    45:"Foggy",48:"Icy fog",51:"Light drizzle",53:"Moderate drizzle",55:"Dense drizzle",
    61:"Slight rain",63:"Moderate rain",65:"Heavy rain",71:"Slight snow",73:"Moderate snow",
    75:"Heavy snow",77:"Snow grains",80:"Light showers",81:"Moderate showers",82:"Heavy showers",
    95:"Thunderstorm",96:"Thunderstorm + hail",99:"Thunderstorm + heavy hail"}
WMO_ICON = {0:"&#x2600;&#xFE0F;",1:"&#x1F324;",2:"&#x26C5;",3:"&#x2601;&#xFE0F;",
    45:"&#x1F32B;",51:"&#x1F326;",53:"&#x1F326;",55:"&#x1F327;",61:"&#x1F327;",
    63:"&#x1F327;",65:"&#x1F327;",80:"&#x1F326;",81:"&#x1F327;",82:"&#x26C8;",
    95:"&#x26C8;",96:"&#x26C8;",99:"&#x26C8;"}

# ── LOGO ──────────────────────────────────────────────────────────────────────
def process_logo(path):
    try:
        import numpy as np
        from PIL import Image
        img = Image.open(path).convert("RGBA")
        arr = np.array(img)
        bright = arr[:,:,:3].max(axis=2)
        out = np.zeros_like(arr)
        out[bright > 15] = [8, 32, 56, 255]
        result = Image.fromarray(out).resize((120, 120))
        buf = io.BytesIO()
        result.save(buf, "PNG", optimize=True)
        return base64.b64encode(buf.getvalue()).decode()
    except Exception as ex:
        print(f"  Warning: logo processing failed ({ex})")
        return ""

# ── EXCEL ─────────────────────────────────────────────────────────────────────
def parse_schedule(wb, sheet_name):
    ws = wb[sheet_name]
    events, last_time = [], None
    for row in ws.iter_rows(min_row=5, values_only=True):
        cols = (list(row) + [None]*8)[:8]
        t, dept, activity, resp, contact, company, number, notes = cols
        if t is not None:
            last_time = t
        if activity:
            events.append({
                "time":    last_time.strftime("%H:%M") if last_time else "",
                "dept":    str(dept    or "").strip(),
                "act":     str(activity).strip(),
                "resp":    str(resp    or "").strip(),
                "contact": str(contact or "").strip(),
                "company": str(company or "").strip(),
                "number":  str(number  or "").strip(),
                "notes":   str(notes   or "").strip(),
            })
    return events

# ── WEATHER ───────────────────────────────────────────────────────────────────
def fetch_weather():
    try:
        import requests
        r = requests.get("https://api.open-meteo.com/v1/forecast", params={
            "latitude": NOOSA_LAT, "longitude": NOOSA_LON,
            "daily": "temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum,weathercode,windspeed_10m_max",
            "timezone": "Australia/Brisbane",
            "start_date": "2026-06-08", "end_date": "2026-06-16"}, timeout=15)
        dd = r.json().get("daily", {})
        if not dd:
            return {}
        result = {}
        for i, iso in enumerate(dd["time"]):
            result[iso] = {
                "max":  dd["temperature_2m_max"][i],
                "min":  dd["temperature_2m_min"][i],
                "rain": dd["precipitation_probability_max"][i] or 0,
                "mm":   dd["precipitation_sum"][i] or 0,
                "code": dd["weathercode"][i],
                "wind": dd["windspeed_10m_max"][i],
            }
        print(f"  Weather: fetched {len(result)} days OK.")
        return result
    except Exception as ex:
        print(f"  Warning: weather fetch failed ({ex}). Sections will show placeholder.")
        return {}

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
def e(v):
    return "" if v is None else html.escape(str(v))

def dept_css(dept):
    l = (dept or "").lower()
    if "vendor" in l: return "dept-Vendor"
    if "security" in l: return "dept-Security"
    if "technical" in l: return "dept-Technical"
    if "theming" in l: return "dept-Theming"
    if "crew" in l: return "dept-Crew"
    if "event" in l: return "dept-Event"
    return "dept-Site"

def bullets(items):
    if not items:
        return '<p class="empty-section">No entries.</p>'
    rows = "\n".join(f"<li>{e(i)}</li>" for i in items)
    return f'<ul class="bullet-list">{rows}</ul>'

def render_schedule(events):
    if not events:
        return '<p class="empty-section">No schedule entries for this day.</p>'
    rows = []
    for ev in events:
        parts = [p for p in [e(ev["contact"]), e(ev["company"])] if p]
        contact_disp = " / ".join(parts)
        rows.append(
            f'<tr><td class="td-time">{e(ev["time"])}</td>'
            f'<td><span class="dept-badge {dept_css(ev["dept"])}">{e(ev["dept"]) or "&mdash;"}</span></td>'
            f'<td>{e(ev["act"])}</td><td>{e(ev["resp"])}</td><td>{contact_disp}</td>'
            f'<td class="td-num">{e(ev["number"])}</td><td class="td-notes">{e(ev["notes"])}</td></tr>'
        )
    body = "\n".join(rows)
    return ('<div class="schedule-wrap"><table class="schedule-table">'
            '<thead><tr><th>Time</th><th>Dept</th><th>Activity</th><th>Responsibility</th>'
            '<th>Contact / Company</th><th>Number</th><th>Notes</th></tr></thead>'
            f'<tbody>{body}</tbody></table></div>')

def render_weather(wx, iso, idx):
    try:
        next_iso = (date.fromisoformat(iso) + timedelta(days=1)).isoformat()
    except Exception:
        next_iso = iso
    if not wx:
        wbody = '<p class="empty-section">Weather unavailable. Run script with internet access, or add manually to day_notes.json.</p>'
    else:
        code = wx.get("code", 0)
        icon = WMO_ICON.get(code, "")
        desc = WMO_DESC.get(code, "Variable")
        max_t = round(wx["max"]); min_t = round(wx["min"])
        rain = int(wx["rain"]); mm = wx["mm"]; wind = round(wx["wind"])
        mm_str = f'<span class="rain-mm">{mm:.1f}mm</span>' if mm > 0 else ""
        wbody = (
            f'<div class="weather-temps"><span class="temp-min">Min {min_t}&deg;</span>'
            f'<span class="temp-max">{max_t}</span><span class="temp-unit">&deg;C</span></div>'
            f'<div class="weather-summary">{icon} {e(desc)}</div>'
            f'<div class="rain-row">Chance of rain: <strong>{rain}%</strong>'
            f'<div class="rain-bar"><div class="rain-fill" style="width:{rain}%"></div></div>{mm_str}</div>'
            f'<div class="weather-extras"><span>Wind: <strong>{wind}&nbsp;km/h</strong></span>'
            f'<span>Open-Meteo &middot; Noosa Beach &middot; {iso}</span></div>')
    map_wrap = (
        f'<div class="weather-map-wrap" id="wmap-{idx}">'
        f'<img id="wimg-{idx}" src="" alt="Rainfall forecast" data-nextdate="{next_iso}">'
        f'<div class="map-label">metvuw &middot; QLD Rainfall</div>'
        f'<div class="map-caption" id="wcap-{idx}">Loading forecast map&hellip;</div></div>')
    return (
        f'<div class="weather-card">'
        f'<div class="weather-date-row">{iso}</div>'
        f'<div class="weather-left">{wbody}</div>'
        f'{map_wrap}</div>')

def render_radios(r):
    if not r: return '<p class="empty-section">Radio assignments not yet set.</p>'
    assignments = r.get("assignments", [])
    channels    = r.get("channels", [])
    half = (len(assignments)+1)//2
    tr = lambda items: "".join(f'<tr><td>{e(a["num"])}</td><td>{e(a["name"])}</td></tr>' for a in items)
    tc = lambda items: "".join(f'<tr><td>{e(c["num"])}</td><td>{e(c["use"])}</td></tr>' for c in items)
    return ('<div class="radio-grid">'
            f'<table class="radio-table"><thead><tr><th>No.</th><th>Name</th></tr></thead><tbody>{tr(assignments[:half])}</tbody></table>'
            f'<div><table class="radio-table" style="margin-bottom:16px"><thead><tr><th>No.</th><th>Name</th></tr></thead><tbody>{tr(assignments[half:])}</tbody></table>'
            f'<table class="radio-table"><thead><tr><th>Ch</th><th>Use</th></tr></thead><tbody>{tc(channels)}</tbody></table></div></div>')

def render_wifi(w):
    if not w: return ""
    return ('<div class="wifi-box"><h4>WiFi Credentials</h4>'
            f'<div class="wifi-row"><div class="wl">Vendor Username</div><div class="wv">{e(w.get("vendor_user",""))}</div></div>'
            f'<div class="wifi-row"><div class="wl">Vendor Password</div><div class="wv">{e(w.get("vendor_pass",""))}</div></div>'
            f'<div class="wifi-row"><div class="wl">Internal Username</div><div class="wv int">{e(w.get("internal_user",""))}</div></div>'
            f'<div class="wifi-row"><div class="wl">Internal Password</div><div class="wv int">{e(w.get("internal_pass",""))}</div></div></div>')

def sec(heading, content):
    return f'<div class="ds-section"><h2 class="section-heading">{heading}</h2>{content}</div>'

# ── DAY PANEL ─────────────────────────────────────────────────────────────────
def render_panel(meta, idx, events, notes, weather_by_date):
    iso = meta["iso"]
    wx  = weather_by_date.get(iso)
    n   = notes
    ticker = n.get("ticker") or {}
    ta = "true" if ticker.get("active") else "false"
    tt = e(ticker.get("text",""))
    tu = e(ticker.get("urgency","green"))
    active_cls = "active" if idx == 0 else ""

    sitemap = (
        '<div class="sitemap-toggle" onclick="toggleSiteMap(this)">'
        '<span>&#x1F4CD; Site Map</span><span class="sitemap-arrow">&#x25BC;</span></div>'
        '<div class="sitemap-body">'
        '<embed src="./site-map.pdf" type="application/pdf">'
        '<p class="sitemap-fallback">Ensure <code>site-map.pdf</code> is in the same folder as this HTML file.</p>'
        '</div>')

    parts = []
    parts.append(f'<div class="ds-section"><h2 class="section-heading">Overview / Key Notes</h2>{bullets(n.get("overview"))}{sitemap}</div>')
    if n.get("changes"):   parts.append(sec("Changes", bullets(n["changes"])))
    if n.get("flags"):     parts.append(sec("Flags",   bullets(n["flags"])))
    parts.append(sec("Schedule", render_schedule(events)))
    parts.append(sec("Before End of Day", bullets(n.get('before_eod'))))
    parts.append(sec("Weather",  render_weather(wx, iso, idx)))
    if n.get("to_sort"):   parts.append(sec("To Sort", bullets(n["to_sort"])))
    if n.get("wifi"):      parts.append(sec("WiFi",    render_wifi(n["wifi"])))
    if n.get("radios"):    parts.append(sec("Radio Assignments / Channels", render_radios(n["radios"])))

    body = "".join(parts)
    return (
        f'<div class="day-panel {active_cls}" id="panel-{idx}"'
        f'     data-ticker-active="{ta}" data-ticker-text="{tt}" data-ticker-urgency="{tu}">'
        f'<div class="day-header">'
        f'<div class="lbl">Noosa Food &amp; Wine Festival 2026 &middot; Festival Village</div>'
        f'<h1>Day Sheet</h1><div class="dd">Day {meta["day"]} &mdash; {e(meta["date"])}</div></div>'
        f'<div class="day-body">{body}</div></div>')

# ── CSS + JS (embedded in generated HTML) ────────────────────────────────────
CSS = "\n:root{--blue:#d5f0fe;--blue-light:rgba(213,240,254,.30);--blue-mid:#1a9bc6;--dark:#082038;--text:rgba(0,0,0,.82);--muted:rgba(0,0,0,.45);--white:#fff;--nav-w:220px;--r:10px;--c-event:#F97316;--c-site:#16A34A;--c-tech:#2563EB;--c-vendor:#D97706;--c-security:#DC2626;--c-theming:#7C3AED;--c-crew:#525252}\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:'DM Sans',sans-serif;font-size:14px;color:var(--text);background:var(--white);letter-spacing:.003em;line-height:1.55}\n.ticker{position:fixed;top:0;left:var(--nav-w);right:0;z-index:500;padding:10px 48px 10px 20px;display:flex;align-items:center;gap:12px;font-size:13px;font-weight:500;transition:transform .3s ease}\n.ticker.hidden{transform:translateY(-110%)}\n.ticker-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}\n.ticker-close{position:absolute;right:16px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;font-size:20px;line-height:1;opacity:.6;padding:4px}\n.ticker-close:hover{opacity:1}\n.ticker.green{background:#dcfce7;color:#14532d;border-bottom:1px solid #86efac}.ticker.green .ticker-dot{background:#16a34a}\n.ticker.orange{background:#fff7ed;color:#7c2d12;border-bottom:1px solid #fdba74}.ticker.orange .ticker-dot{background:#f97316}\n.ticker.red{background:#fef2f2;color:#7f1d1d;border-bottom:1px solid #fca5a5}.ticker.red .ticker-dot{background:#dc2626}\nbody.has-ticker .main{padding-top:42px}\n.sidebar{position:fixed;top:0;left:0;width:var(--nav-w);height:100vh;background:var(--blue);display:flex;flex-direction:column;overflow-y:auto;overflow-x:hidden;z-index:100;scrollbar-width:none}\n.sidebar::-webkit-scrollbar{display:none}\n.sidebar-logo{padding:22px 18px 14px;flex-shrink:0}\n.sidebar-logo img{width:62px;height:62px;display:block}\n.sidebar-meta{padding:0 18px 16px;border-bottom:1px solid rgba(8,32,56,.12);flex-shrink:0}\n.sidebar-meta .fn{font-family:'Outfit',sans-serif;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dark);line-height:1.3;margin-bottom:2px}\n.sidebar-meta .fs{font-size:11px;color:rgba(8,32,56,.55)}\n.day-nav{list-style:none;padding:10px 0 20px;flex:1}\n.day-nav li a{display:block;padding:9px 18px;text-decoration:none;color:rgba(8,32,56,.7);font-size:12px;font-weight:500;line-height:1.3;transition:background .15s;border-left:3px solid transparent;cursor:pointer}\n.day-nav li a .dn{display:block;font-family:'Outfit',sans-serif;font-weight:700;font-size:13px;color:var(--dark);line-height:1.2}\n.day-nav li a .dt{font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:rgba(8,32,56,.45)}\n.day-nav li a:hover{background:rgba(8,32,56,.07)}\n.day-nav li a.active{background:rgba(8,32,56,.1);border-left-color:var(--dark)}\n.sidebar-footer{padding:14px 18px;border-top:1px solid rgba(8,32,56,.1);font-size:11px;color:rgba(8,32,56,.45);flex-shrink:0}\n.main{margin-left:var(--nav-w);min-height:100vh}\n.day-panel{display:none}.day-panel.active{display:block}\n.day-header{background:var(--blue);padding:36px 48px 28px;border-bottom:1px solid rgba(8,32,56,.08)}\n.day-header .lbl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:rgba(8,32,56,.5);margin-bottom:6px}\n.day-header h1{font-family:'Outfit',sans-serif;font-size:52px;font-weight:300;letter-spacing:-.03em;color:var(--dark);line-height:1;margin-bottom:6px}\n.day-header .dd{font-family:'Outfit',sans-serif;font-size:18px;font-weight:400;color:rgba(8,32,56,.6);letter-spacing:-.01em}\n.day-body{padding:0 48px 60px}\n.ds-section{padding:32px 0 28px;border-bottom:1px solid rgba(8,32,56,.07)}\n.ds-section:last-child{border-bottom:none}\n.section-heading{font-family:'Outfit',sans-serif;font-size:28px;font-weight:300;color:var(--blue-mid);letter-spacing:-.02em;margin-bottom:16px;line-height:1.1}\n.empty-section{color:var(--muted);font-size:13px;font-style:italic}\n.sitemap-toggle{display:flex;align-items:center;gap:10px;cursor:pointer;user-select:none;background:var(--blue-light);border:1px solid rgba(8,32,56,.1);border-radius:var(--r);padding:12px 16px;margin-top:16px;font-family:'Outfit',sans-serif;font-size:14px;font-weight:600;color:var(--dark);transition:background .2s}\n.sitemap-toggle:hover{background:rgba(213,240,254,.6)}\n.sitemap-arrow{font-size:11px;transition:transform .25s;margin-left:auto;color:var(--muted)}\n.sitemap-toggle.open .sitemap-arrow{transform:rotate(180deg)}\n.sitemap-body{display:none;margin-top:8px;border-radius:var(--r);overflow:hidden;border:1px solid rgba(8,32,56,.1)}\n.sitemap-body.open{display:block}\n.sitemap-body embed{display:block;width:100%;height:72vh;border:none}\n.sitemap-fallback{padding:16px;color:var(--muted);font-size:12px;background:#f8f9fa}\n.bullet-list{list-style:none;display:flex;flex-direction:column;gap:8px}\n.bullet-list li{padding-left:22px;position:relative;font-size:14px;line-height:1.55}\n.bullet-list li::before{content:'';position:absolute;left:0;top:9px;width:8px;height:8px;border:1.5px solid var(--dark);border-radius:50%}\n.schedule-wrap{overflow-x:auto;margin-top:4px}\n.schedule-table{width:100%;border-collapse:collapse;font-size:12.5px;min-width:780px}\n.schedule-table thead th{background:var(--dark);color:var(--blue);padding:9px 12px;text-align:left;font-family:'Outfit',sans-serif;font-weight:600;font-size:10.5px;text-transform:uppercase;letter-spacing:.1em;white-space:nowrap}\n.schedule-table tbody tr{border-bottom:1px solid rgba(8,32,56,.07)}\n.schedule-table tbody tr:nth-child(even){background:rgba(213,240,254,.18)}\n.schedule-table tbody tr:hover{background:rgba(213,240,254,.35)}\n.schedule-table td{padding:8px 12px;vertical-align:top}\n.td-time{font-family:'Outfit',sans-serif;font-weight:600;font-size:13px;color:var(--dark);white-space:nowrap;width:70px}\n.td-num{white-space:nowrap}\n.td-notes{font-size:11.5px;color:var(--muted);font-style:italic}\n.dept-badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;color:white}\n.dept-Event{background:var(--c-event)}.dept-Site{background:var(--c-site)}.dept-Technical{background:var(--c-tech)}\n.dept-Vendor{background:var(--c-vendor)}.dept-Security{background:var(--c-security)}.dept-Theming{background:var(--c-theming)}.dept-Crew{background:var(--c-crew)}\n.weather-card{background:var(--blue-light);border:1px solid rgba(8,32,56,.1);border-radius:var(--r);padding:20px 24px;display:grid;grid-template-columns:1fr auto;gap:16px 24px}\n.weather-date-row{grid-column:1/-1;font-family:'Outfit',sans-serif;font-weight:600;font-size:13px;color:rgba(8,32,56,.6);text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid rgba(8,32,56,.1);padding-bottom:10px}\n.weather-temps{display:flex;align-items:baseline;gap:12px;margin-bottom:6px}\n.temp-min{font-size:15px;color:var(--muted)}.temp-max{font-family:'Outfit',sans-serif;font-size:36px;font-weight:300;color:var(--dark);letter-spacing:-.02em}.temp-unit{font-size:16px;color:var(--muted)}\n.weather-summary{font-weight:600;font-size:13px;color:var(--dark);margin-bottom:8px}\n.rain-row{display:flex;align-items:center;gap:10px;font-size:12.5px;margin-bottom:4px}\n.rain-bar{flex:1;max-width:100px;height:5px;background:rgba(8,32,56,.12);border-radius:3px;overflow:hidden}\n.rain-fill{height:100%;background:var(--blue-mid);border-radius:3px}\n.rain-mm{font-size:11px;color:var(--muted)}\n.weather-extras{display:flex;gap:20px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin-top:8px;padding-top:8px;border-top:1px solid rgba(8,32,56,.08)}\n.weather-extras strong{color:var(--text);font-weight:600}\n.weather-left{}\n.weather-map-wrap{flex-shrink:0;text-align:center}\n.weather-map-wrap img{width:520px;max-width:100%;border-radius:6px;border:1px solid rgba(8,32,56,.12);display:block}\n.map-label{font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-top:6px;font-weight:600}\n.map-caption{font-size:11px;color:rgba(8,32,56,.45);margin-top:3px}\n.radio-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}\n.radio-table{width:100%;border-collapse:collapse;font-size:12.5px}\n.radio-table thead th{background:rgba(8,32,56,.06);color:var(--dark);padding:7px 10px;text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;font-weight:700}\n.radio-table td{padding:6px 10px;border-bottom:1px solid rgba(8,32,56,.06)}\n.radio-table td:first-child{font-family:'Outfit',sans-serif;font-weight:600;color:var(--dark);width:30px;text-align:center}\n.wifi-box{background:var(--dark);color:white;border-radius:var(--r);padding:20px 24px;display:grid;grid-template-columns:1fr 1fr;gap:12px 32px}\n.wifi-box h4{grid-column:1/-1;font-family:'Outfit',sans-serif;font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--blue);margin-bottom:4px}\n.wifi-row .wl{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:rgba(213,240,254,.6);margin-bottom:2px}\n.wifi-row .wv{font-family:'Outfit',sans-serif;font-weight:600;color:white}\n.wifi-row .wv.int{color:var(--blue)}\n.contacts-btn{position:fixed;top:18px;right:24px;z-index:600;background:var(--dark);color:var(--blue);border:none;border-radius:8px;padding:9px 18px;font-family:'Outfit',sans-serif;font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;transition:background .2s,transform .15s;box-shadow:0 2px 8px rgba(8,32,56,.25)}.contacts-btn:hover{background:#0d3055;transform:translateY(-1px)}.contacts-backdrop{position:fixed;inset:0;background:rgba(8,32,56,.55);z-index:700;opacity:0;pointer-events:none;transition:opacity .3s ease}.contacts-backdrop.open{opacity:1;pointer-events:all}.contacts-panel{position:fixed;top:0;right:0;width:460px;max-width:100vw;height:100vh;background:var(--white);z-index:800;display:flex;flex-direction:column;transform:translateX(100%);transition:transform .35s cubic-bezier(.4,0,.2,1);box-shadow:-4px 0 32px rgba(8,32,56,.18)}.contacts-panel.open{transform:translateX(0)}.contacts-panel-head{padding:24px 24px 20px;background:var(--dark);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}.contacts-panel-head h2{font-family:'Outfit',sans-serif;font-size:18px;font-weight:600;color:var(--blue);letter-spacing:.02em}.contacts-close{background:none;border:none;color:rgba(213,240,254,.6);font-size:26px;line-height:1;cursor:pointer;padding:4px;transition:color .15s}.contacts-close:hover{color:var(--blue)}.contacts-search{padding:14px 16px;border-bottom:1px solid rgba(8,32,56,.08);flex-shrink:0}.contacts-search input{width:100%;border:1px solid rgba(8,32,56,.15);border-radius:6px;padding:8px 12px;font-size:13px;font-family:'DM Sans',sans-serif;outline:none;color:var(--dark);transition:border-color .15s}.contacts-search input:focus{border-color:var(--blue-mid)}.contacts-list{flex:1;overflow-y:auto;padding:12px 0}.contact-card{display:flex;align-items:center;gap:14px;padding:12px 20px;border-bottom:1px solid rgba(8,32,56,.06);transition:background .15s}.contact-card:hover{background:rgba(213,240,254,.35)}.contact-avatar{width:38px;height:38px;border-radius:50%;background:var(--blue);display:flex;align-items:center;justify-content:center;flex-shrink:0;font-family:'Outfit',sans-serif;font-weight:700;font-size:13px;color:var(--dark)}.contact-info{flex:1;min-width:0}.contact-company{font-family:'Outfit',sans-serif;font-size:13px;font-weight:700;color:var(--dark);line-height:1.3}.contact-name{font-size:12px;color:var(--muted);margin-top:1px}.contact-number{flex-shrink:0;text-align:right}.contact-number a{font-family:'Outfit',sans-serif;font-size:13px;font-weight:600;color:var(--blue-mid);text-decoration:none;white-space:nowrap}.contact-number a:hover{text-decoration:underline}.contact-no-number{font-size:11px;color:rgba(8,32,56,.25);font-style:italic}.contacts-empty{padding:40px 24px;text-align:center;color:var(--muted);font-size:13px}@media(max-width:900px){:root{--nav-w:0px}.sidebar{transform:translateX(-220px);transition:transform .3s ease;width:220px}.sidebar.open{transform:translateX(0)}.ticker{left:0}.mobile-toggle{display:flex;position:fixed;top:12px;left:12px;z-index:200;width:40px;height:40px;background:var(--blue);border-radius:6px;align-items:center;justify-content:center;cursor:pointer;border:none;flex-direction:column;gap:4px;padding:10px}.mobile-toggle span{display:block;width:20px;height:2px;background:var(--dark);border-radius:2px}.day-header,.day-body{padding-left:20px;padding-right:20px}.day-header{padding-top:60px}.weather-card{grid-template-columns:1fr}.weather-map-wrap img{width:100%;max-width:340px}.radio-grid,.wifi-box{grid-template-columns:1fr}}\n@media(min-width:901px){.mobile-toggle{display:none}}\n"
JS  = '\nfunction showTicker(text, urgency) {\n  var el = document.getElementById(\'ticker\');\n  if (!text) { hideTicker(); return; }\n  el.className = \'ticker \' + (urgency || \'green\');\n  el.innerHTML = \'<span class="ticker-dot"></span><span class="ticker-text">\' + text + \'</span>\'\n    + \'<button class="ticker-close" onclick="hideTicker()" aria-label="Dismiss">&times;</button>\';\n  document.body.classList.add(\'has-ticker\');\n}\nfunction hideTicker() {\n  document.getElementById(\'ticker\').className = \'ticker hidden\';\n  document.body.classList.remove(\'has-ticker\');\n}\nfunction showDay(idx) {\n  document.querySelectorAll(\'.day-panel\').forEach(function(p){ p.classList.remove(\'active\'); });\n  document.querySelectorAll(\'#day-nav a\').forEach(function(a){ a.classList.remove(\'active\'); });\n  var panel = document.getElementById(\'panel-\' + idx);\n  panel.classList.add(\'active\');\n  document.querySelector(\'#day-nav a[data-day="\' + idx + \'"]\').classList.add(\'active\');\n  window.scrollTo(0,0);\n  document.querySelector(\'.sidebar\').classList.remove(\'open\');\n  if (panel.dataset.tickerActive === \'true\') {\n    showTicker(panel.dataset.tickerText, panel.dataset.tickerUrgency);\n  } else { hideTicker(); }\n}\nfunction toggleSiteMap(btn) {\n  btn.classList.toggle(\'open\');\n  btn.nextElementSibling.classList.toggle(\'open\');\n}\nfunction metVuwUrls(nextDayIso) {\n  var p2=function(n){return String(n).padStart(2,\'0\');};\n  var p3=function(n){return String(n).padStart(3,\'0\');};\n  function urlsForCut(cut){\n    var rh=[0,6,12,18],ch=cut.getUTCHours(),bestH=18;\n    for(var i=rh.length-1;i>=0;i--){if(rh[i]<=ch){bestH=rh[i];break;}}\n    var run=new Date(cut); run.setUTCHours(bestH,0,0,0);\n    if(bestH>ch) run.setUTCDate(run.getUTCDate()-1);\n    var target=new Date(nextDayIso+\'T02:00:00Z\');\n    var baseHrs=Math.round((target-run)/(6*3600000))*6;\n    var runStr=run.getUTCFullYear()+p2(run.getUTCMonth()+1)+p2(run.getUTCDate())+p2(bestH);\n    var base=\'https://www.metvuw.com/forecast/\'+runStr+\'/rain-queensland-\'+runStr+\'-\';\n    var urls=[];\n    [0,-6,6,-12,12,18,-18].forEach(function(offset){var h=baseHrs+offset;if(h>=0&&h<=240)urls.push(base+p3(h)+\'.gif\');});\n    return urls;\n  }\n  var now=new Date();\n  var u1=urlsForCut(new Date(now.getTime()-3*3600000));\n  var u2=urlsForCut(new Date(now.getTime()-9*3600000));\n  return u1.concat(u2.filter(function(u){return u1.indexOf(u)<0;}));\n}\nfunction loadAllMaps() {\n  document.querySelectorAll(\'.weather-map-wrap img\').forEach(function(img) {\n    var nd=img.dataset.nextdate;\n    if(!nd){img.closest(\'.weather-map-wrap\').style.display=\'none\';return;}\n    var urls=metVuwUrls(nd),i=0;\n    var capEl=img.parentElement.querySelector(\'.map-caption\');\n    function tryNext(){\n      if(i>=urls.length){img.closest(\'.weather-map-wrap\').style.display=\'none\';return;}\n      img.onerror=tryNext;\n      img.onload=function(){if(capEl)capEl.textContent=\'Queensland rainfall — next day forecast\';};\n      img.src=urls[i++];\n    }\n    tryNext();\n  });\n}\ndocument.addEventListener(\'DOMContentLoaded\',function(){\n  var first=document.getElementById(\'panel-0\');\n  if(first&&first.dataset.tickerActive===\'true\'){\n    showTicker(first.dataset.tickerText,first.dataset.tickerUrgency);\n  }\n  loadAllMaps();\n});\n'

# ── GENERATE HTML ─────────────────────────────────────────────────────────────
def generate_html(days_data, logo_b64, generated_at, contacts=None):
    nav = "".join(
        f'<li><a href="#" data-day="{i}" class="{"active" if i==0 else ""}"'
        f'    onclick="showDay({i});return false;">'
        f'<span class="dn">Day {m["day"]} &middot; {e(m["short"])}</span>'
        f'<span class="dt">{e(m["type"])}</span></a></li>'
        for i, (m, panel) in enumerate(days_data))
    logo_tag = (f'<img src="data:image/png;base64,{logo_b64}" alt="NF&amp;W">'
                if logo_b64 else
                '<div style="width:62px;height:62px;background:var(--dark);border-radius:6px;'
                'display:flex;align-items:center;justify-content:center;color:white;'
                'font-family:Outfit,sans-serif;font-weight:700;font-size:14px">NF&amp;W</div>')
    panels = "".join(panel for _, panel in days_data)
    import json as _json
    contacts_json = _json.dumps(contacts or [], ensure_ascii=False)
    contacts_js = (
        "var CONTACTS_DATA = " + contacts_json + ";"
        """
function openContacts() {
  document.getElementById('contacts-backdrop').classList.add('open');
  document.getElementById('contacts-panel').classList.add('open');
  document.getElementById('contacts-search').value = '';
  renderContacts('');
  setTimeout(function(){ document.getElementById('contacts-search').focus(); }, 350);
}
function closeContacts() {
  document.getElementById('contacts-backdrop').classList.remove('open');
  document.getElementById('contacts-panel').classList.remove('open');
}
function renderContacts(filter) {
  var list = document.getElementById('contacts-list');
  var f = (filter || '').toLowerCase();
  var items = CONTACTS_DATA.filter(function(c) {
    return !f || c.company.toLowerCase().indexOf(f) !== -1 ||
                 c.name.toLowerCase().indexOf(f) !== -1 ||
                 c.number.toLowerCase().indexOf(f) !== -1;
  });
  if (!items.length) {
    list.innerHTML = '<div class="contacts-empty">No contacts found.</div>';
    return;
  }
  list.innerHTML = items.map(function(c) {
    var words = c.company.split(' ').filter(function(w){ return w.length > 0; });
    var initials = words.slice(0,2).map(function(w){ return w[0].toUpperCase(); }).join('');
    var numHtml = c.number
      ? '<a href="tel:' + c.number.replace(/\s/g,'') + '">' + c.number + '</a>'
      : '<span class="contact-no-number">No number</span>';
    var nameHtml = c.name ? '<div class="contact-name">' + c.name + '</div>' : '';
    return '<div class="contact-card">'
      + '<div class="contact-avatar">' + initials + '</div>'
      + '<div class="contact-info"><div class="contact-company">' + c.company + '</div>' + nameHtml + '</div>'
      + '<div class="contact-number">' + numHtml + '</div>'
      + '</div>';
  }).join('');
}
document.addEventListener('DOMContentLoaded', function() {
  renderContacts('');
  document.getElementById('contacts-search').addEventListener('input', function() {
    renderContacts(this.value);
  });
  document.getElementById('contacts-backdrop').addEventListener('click', closeContacts);
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') closeContacts();
  });
});
"""
    )
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NFW 2026 — Day Sheets</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<button class="mobile-toggle" onclick="document.querySelector('.sidebar').classList.toggle('open')" aria-label="Menu"><span></span><span></span><span></span></button>
<div id="ticker" class="ticker hidden"></div>
<nav class="sidebar">
  <div class="sidebar-logo">{logo_tag}</div>
  <div class="sidebar-meta"><div class="fn">Noosa Food &amp; Wine Festival</div><div class="fs">2026 &middot; Festival Village</div></div>
  <ul class="day-nav" id="day-nav">{nav}</ul>
  <div class="sidebar-footer">Production Day Sheets<br>Generated {generated_at}</div>
</nav>
<main class="main">{panels}</main>
<button class="contacts-btn" onclick="openContacts()">Contacts</button>
<div class="contacts-backdrop" id="contacts-backdrop"></div>
<div class="contacts-panel" id="contacts-panel">
  <div class="contacts-panel-head">
    <h2>Contacts</h2>
    <button class="contacts-close" onclick="closeContacts()" aria-label="Close">&times;</button>
  </div>
  <div class="contacts-search"><input type="search" id="contacts-search" placeholder="Search company, name or number&hellip;"></div>
  <div class="contacts-list" id="contacts-list"></div>
</div>
<script>{JS}</script>
<script>{contacts_js}</script>
</body>
</html>"""


# ── EXCEL NOTES READER ────────────────────────────────────────────────────────
# Standing data (shared across all days)
WIFI_ROW_VENDOR   = 10   # col B = vendor username, col D = vendor password
WIFI_ROW_INTERNAL = 11   # col B = internal username, col D = internal password
RADIO_FIRST_ROW   = 15   # col A = no., col B = name  (24 rows)
CHANNEL_FIRST_ROW = 15   # col D = ch no., col E = use (5 rows)

# Per-day layout
FIRST_DAY_ROW  = 42
SECTION_HEIGHT = 93
OFF = {
    'ticker_active':   3,  'ticker_text':      4,  'ticker_urgency':   5,
    'overview_hdr':    7,  'overview_count':  15,
    'changes_hdr':    24,  'changes_count':   15,
    'flags_hdr':      41,  'flags_count':     15,
    'to_sort_hdr':    58,  'to_sort_count':   15,
    'before_eod_hdr': 75,  'before_eod_count':15,
}


def read_contacts_from_excel(wb):
    """Read company contacts from Data Validation Sheet col C/D/E."""
    if 'Data Validation Sheet' not in wb.sheetnames:
        return []
    ws = wb['Data Validation Sheet']
    contacts = []
    for r in range(2, 200):
        company = str(ws.cell(r, 4).value or '').strip()
        if not company:
            continue
        name   = str(ws.cell(r, 3).value or '').strip()
        number = str(ws.cell(r, 5).value or '').strip()
        if name in ('None','') and number in ('None',''):
            contacts.append({'company': company, 'name': '', 'number': ''})
        else:
            contacts.append({
                'company': company,
                'name':    '' if name   == 'None' else name,
                'number':  '' if number == 'None' else number,
            })
    contacts.sort(key=lambda c: c['company'].lower())
    return contacts

def read_notes_from_excel(wb):
    if 'Day Sheet Notes' not in wb.sheetnames:
        print("  Notes: 'Day Sheet Notes' sheet not found — editorial sections empty.")
        return {}
    ws = wb['Day Sheet Notes']

    # Read shared WiFi (once)
    vu = str(ws.cell(WIFI_ROW_VENDOR,   2).value or '').strip()
    vp = str(ws.cell(WIFI_ROW_VENDOR,   4).value or '').strip()
    iu = str(ws.cell(WIFI_ROW_INTERNAL, 2).value or '').strip()
    ip = str(ws.cell(WIFI_ROW_INTERNAL, 4).value or '').strip()
    shared_wifi = {'vendor_user': vu, 'vendor_pass': vp,
                   'internal_user': iu, 'internal_pass': ip} if vu else None

    # Read shared radios (once)
    assignments = []
    for i in range(24):
        r = RADIO_FIRST_ROW + i
        num  = ws.cell(r, 1).value
        name = str(ws.cell(r, 2).value or '').strip()
        if name:
            assignments.append({'num': int(num or i+1), 'name': name})
    channels = []
    for i in range(10):
        r = CHANNEL_FIRST_ROW + i
        ch_n = ws.cell(r, 4).value
        ch_u = str(ws.cell(r, 5).value or '').strip()
        if ch_n and ch_u:
            channels.append({'num': int(ch_n), 'use': ch_u})
    shared_radios = {'assignments': assignments, 'channels': channels} if assignments else None

    def read_bullets(start_row, count):
        return [str(ws.cell(start_row + i, 2).value or '').strip()
                for i in range(count) if ws.cell(start_row + i, 2).value]

    result = {}
    for day_idx in range(9):
        S = FIRST_DAY_ROW + day_idx * SECTION_HEIGHT
        day_num = day_idx + 1

        ta  = str(ws.cell(S + OFF['ticker_active'],  2).value or '').strip()
        tt  = str(ws.cell(S + OFF['ticker_text'],    2).value or '').strip()
        tu  = str(ws.cell(S + OFF['ticker_urgency'], 2).value or 'Green').strip()
        ticker = {'active': True, 'text': tt, 'urgency': tu.lower()}                  if ta.lower() == 'yes' and tt else None

        result[day_num] = {
            'day':        day_num,
            'ticker':     ticker,
            'overview':   read_bullets(S + OFF['overview_hdr']   + 1, OFF['overview_count']),
            'changes':    read_bullets(S + OFF['changes_hdr']    + 1, OFF['changes_count']),
            'flags':      read_bullets(S + OFF['flags_hdr']      + 1, OFF['flags_count']),
            'to_sort':    read_bullets(S + OFF['to_sort_hdr']    + 1, OFF['to_sort_count']),
            'before_eod': read_bullets(S + OFF['before_eod_hdr'] + 1, OFF['before_eod_count']),
            'wifi':       shared_wifi,
            'radios':     shared_radios,
        }

    print(f"  Notes: read from 'Day Sheet Notes' sheet.")
    return result

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="NFW 2026 Day Sheet Generator")
    parser.add_argument("--excel",  default="NFW_Production_Schedule_2026_-_Master.xlsx")
    parser.add_argument("--logo",   default="NFWF_Monogram_Positive.png")
    parser.add_argument("--output", default="daysheet.html")
    args = parser.parse_args()

    print("NFW 2026 Day Sheet Generator")
    print("-" * 40)

    logo_b64 = process_logo(args.logo) if Path(args.logo).exists() else ""
    logo_status = "OK" if logo_b64 else "not found, using text fallback"
    print(f"  Logo: {logo_status}")

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"  ERROR: Excel file not found: {excel_path}"); sys.exit(1)
    import openpyxl
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    print(f"  Excel: loaded {excel_path.name}")

    contacts     = read_contacts_from_excel(wb)
    notes_by_day = read_notes_from_excel(wb)

    print("  Weather: fetching from Open-Meteo...")
    weather_by_date = fetch_weather()

    days_data = []
    for meta in DAYS_META:
        events = parse_schedule(wb, meta["sheet"]) if meta["sheet"] in wb.sheetnames else []
        notes  = notes_by_day.get(meta["day"], {})
        idx    = len(days_data)
        panel  = render_panel(meta, idx, events, notes, weather_by_date)
        days_data.append((meta, panel))
        print(f'  Day {meta["day"]}: {meta["date"]:25s}  {len(events):2d} schedule events')

    html_out = generate_html(days_data, logo_b64, dt.now().strftime("%d %b %Y %H:%M"), contacts)
    Path(args.output).write_text(html_out, encoding="utf-8")
    size_kb = Path(args.output).stat().st_size // 1024
    print(f"\n  Output: {args.output}  ({size_kb} KB)")
    print("\nPut these files together in your OneDrive folder:")
    print(f"   {args.output}")
    print(f"   site-map.pdf")
    print(f"   day_notes.json  (edit this to update editorial content)")

if __name__ == "__main__":
    main()