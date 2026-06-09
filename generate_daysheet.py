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

import argparse, base64, html, io, json, os, re, sys
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
        cols = (list(row) + [None]*9)[:9]
        t, dept, activity, resp, contact, company, number, notes, done_val = cols
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
                "done":    bool(done_val and str(done_val).strip() not in ("", "False", "FALSE")),
            })
    return events

# ── WEATHER ───────────────────────────────────────────────────────────────────
def fetch_weather():
    try:
        import requests
        r = requests.get("https://api.open-meteo.com/v1/forecast", params={
            "latitude": NOOSA_LAT, "longitude": NOOSA_LON,
            "daily": ("temperature_2m_max,temperature_2m_min,precipitation_probability_max,"
                      "precipitation_sum,weathercode,windspeed_10m_max,windgusts_10m_max,"
                      "winddirection_10m_dominant,sunrise,sunset,uv_index_max"),
            "hourly": ("temperature_2m,apparent_temperature,precipitation_probability,"
                       "precipitation,windspeed_10m,winddirection_10m,"
                       "relative_humidity_2m,cloudcover,uv_index"),
            "timezone": "Australia/Brisbane",
            "start_date": "2026-06-08", "end_date": "2026-06-16"
        }, timeout=20)
        data = r.json()
        dd = data.get("daily", {})
        hd = data.get("hourly", {})
        if not dd:
            return {}, {}

        def dv(key, i):
            a = dd.get(key, [])
            return a[i] if i < len(a) else None

        daily = {}
        for i, iso in enumerate(dd.get("time", [])):
            daily[iso] = {
                "max":     dv("temperature_2m_max", i),
                "min":     dv("temperature_2m_min", i),
                "rain":    dv("precipitation_probability_max", i) or 0,
                "mm":      dv("precipitation_sum", i) or 0,
                "code":    dv("weathercode", i) or 0,
                "wind":    dv("windspeed_10m_max", i) or 0,
                "gusts":   dv("windgusts_10m_max", i) or 0,
                "winddir": dv("winddirection_10m_dominant", i) or 0,
                "sunrise": dv("sunrise", i) or "",
                "sunset":  dv("sunset",  i) or "",
                "uv_max":  dv("uv_index_max", i) or 0,
            }

        hourly = {}
        for i, ts in enumerate(hd.get("time", [])):
            day_iso = ts[:10]
            hour = int(ts[11:13])
            if 6 <= hour <= 23 and day_iso in daily:
                def hv(key, fallback=0):
                    a = hd.get(key, [])
                    v = a[i] if i < len(a) else None
                    return v if v is not None else fallback
                hourly.setdefault(day_iso, []).append({
                    "hour":          hour,
                    "temp":          hv("temperature_2m", None),
                    "apparent_temp": hv("apparent_temperature", None),
                    "prob":          int(hv("precipitation_probability")),
                    "precip":        hv("precipitation"),
                    "windspeed":     hv("windspeed_10m"),
                    "winddir":       hv("winddirection_10m"),
                    "humidity":      hv("relative_humidity_2m"),
                    "cloud":         hv("cloudcover"),
                    "uv":            hv("uv_index"),
                })

        for iso, hrs in hourly.items():
            hum = [h["humidity"] for h in hrs if h["humidity"]]
            if hum:
                daily[iso]["humidity"] = round(max(hum))

        print(f"  Weather: fetched {len(daily)} days OK.")
        return daily, hourly
    except Exception as ex:
        print(f"  Warning: weather fetch failed ({ex}). Sections will show placeholder.")
        return {}, {}

def fetch_metvuw_maps(days_meta):
    """Download Metvuw QLD rainfall maps at generation time; return {iso: data_uri or None}."""
    try:
        import requests as _req
    except ImportError:
        print("  Maps: requests not installed, skipping.")
        return {m['iso']: None for m in days_meta}

    def get_urls(next_iso, cut):
        rh = [0, 6, 12, 18]
        ch = cut.hour
        best_h = 18
        for h in reversed(rh):
            if h <= ch:
                best_h = h
                break
        run = cut.replace(hour=best_h, minute=0, second=0, microsecond=0)
        tgt = dt.strptime(next_iso + 'T02:00:00', '%Y-%m-%dT%H:%M:%S')
        base_hrs = round((tgt - run).total_seconds() / (6 * 3600)) * 6
        run_str = run.strftime('%Y%m%d%H')
        base_url = f'https://www.metvuw.com/forecast/{run_str}/rain-queensland-{run_str}-'
        urls = []
        for offset in [0, -6, 6, -12, 12, 18, -18]:
            h = base_hrs + offset
            if 0 <= h <= 240:
                urls.append(f'{base_url}{str(int(h)).zfill(3)}.gif')
        return urls

    now = dt.utcnow()
    maps = {}
    for meta in days_meta:
        iso = meta['iso']
        try:
            next_iso = (date.fromisoformat(iso) + timedelta(days=1)).isoformat()
        except Exception:
            next_iso = iso
        seen = set()
        urls = []
        for u in get_urls(next_iso, now - timedelta(hours=3)) + get_urls(next_iso, now - timedelta(hours=9)):
            if u not in seen:
                seen.add(u)
                urls.append(u)
        maps[iso] = None
        for url in urls:
            try:
                resp = _req.get(url, timeout=10)
                if resp.status_code == 200 and resp.content:
                    maps[iso] = 'data:image/gif;base64,' + base64.b64encode(resp.content).decode('ascii')
                    break
            except Exception:
                continue
    fetched = sum(1 for v in maps.values() if v)
    print(f"  Maps: {fetched}/{len(days_meta)} Metvuw images embedded.")
    return maps

# ── WEATHER HELPERS ───────────────────────────────────────────────────────────
def wind_dir_label(deg):
    dirs = ['N','NNE','NE','ENE','E','ESE','SE','SSE','S','SSW','SW','WSW','W','WNW','NW','NNW']
    return dirs[int((float(deg) + 11.25) / 22.5) % 16]

def uv_label_color(uv):
    uv = uv or 0
    if uv < 3:  return 'Low',       '#22c55e'
    if uv < 6:  return 'Moderate',  '#eab308'
    if uv < 8:  return 'High',      '#f97316'
    if uv < 11: return 'Very High', '#ef4444'
    return 'Extreme', '#9333ea'

def fmt_sunrise(ts):
    try:
        t = dt.strptime(ts[-5:], '%H:%M')
        return t.strftime('%I:%M%p').lstrip('0').lower()
    except Exception:
        return ts

def rain_advisory(hourly):
    if not hourly:
        return 'green', ''
    max_prob  = max(h['prob'] for h in hourly)
    total_mm  = sum(h['precip'] for h in hourly)
    windows   = []
    start = None
    for h in hourly:
        if h['prob'] >= 40 and start is None:
            start = h['hour']
        elif h['prob'] < 40 and start is not None:
            windows.append((start, h['hour']))
            start = None
    if start is not None:
        windows.append((start, hourly[-1]['hour'] + 1))

    def fh(h):
        if h == 12: return '12pm'
        return (f'{h-12}pm' if h > 12 else f'{h}am')

    if max_prob >= 70 or total_mm >= 5:
        level = 'red'
        win   = f' ({fh(windows[0][0])}–{fh(windows[0][1])})' if windows else ''
        msg   = f'High rain risk{win}. {total_mm:.1f}mm possible. Have contingency ready.'
    elif max_prob >= 40:
        level = 'orange'
        win   = f' {fh(windows[0][0])}–{fh(windows[0][1])}' if windows else ''
        msg   = f'Rain possible{win}. Monitor conditions.'
    else:
        level = 'green'
        msg   = 'Dry conditions expected.'
    return level, msg

def render_hourly_chart(hourly):
    if not hourly:
        return ''
    W, HC, HL = 720, 100, 22
    PL, PR    = 6, 6
    CW        = W - PL - PR
    n         = len(hourly)
    cw        = CW / n

    probs  = [h['prob']   for h in hourly]
    precip = [h['precip'] for h in hourly]
    temps  = [h.get('temp') for h in hourly]

    max_p  = max(max(precip) if precip else 0, 2)
    vt     = [t for t in temps if t is not None]
    t_lo   = (min(vt) - 2) if vt else 10
    t_hi   = (max(vt) + 2) if vt else 35
    t_rng  = max(t_hi - t_lo, 4)

    def xc(i):   return PL + (i + 0.5) * cw
    def yp(p):   return HC * (1.0 - p / 100.0 * 0.88)
    def ymm(mm): return HC * (1.0 - min(mm, max_p) / max_p * 0.65)
    def yt(t):   return HC * (0.08 + (1 - (t - t_lo) / t_rng) * 0.55)

    def smooth(pts):
        d = f'M {pts[0][0]:.1f},{pts[0][1]:.1f}'
        for k in range(1, len(pts)):
            x0, y0 = pts[k-1]; x1, y1 = pts[k]
            cx = (x0 + x1) / 2
            d += f' C {cx:.1f},{y0:.1f} {cx:.1f},{y1:.1f} {x1:.1f},{y1:.1f}'
        return d

    o = []
    o.append(f'<svg viewBox="0 0 {W} {HC+HL}" xmlns="http://www.w3.org/2000/svg" style="width:100%;display:block">')
    o.append('<defs>'
             '<linearGradient id="wxpg" x1="0" y1="0" x2="0" y2="1">'
             '<stop offset="0%" stop-color="#1a9bc6" stop-opacity="0.45"/>'
             '<stop offset="100%" stop-color="#1a9bc6" stop-opacity="0.04"/>'
             '</linearGradient>'
             '<linearGradient id="wxbg" x1="0" y1="0" x2="0" y2="1">'
             '<stop offset="0%" stop-color="#0c4a6e"/>'
             '<stop offset="100%" stop-color="#0369a1"/>'
             '</linearGradient>'
             '</defs>')
    o.append(f'<rect x="{PL}" y="0" width="{CW}" height="{HC}" fill="rgba(213,240,254,0.18)" rx="4"/>')

    for pct in [25, 50, 75]:
        gy  = yp(pct)
        col = 'rgba(8,32,56,0.12)' if pct == 50 else 'rgba(8,32,56,0.06)'
        o.append(f'<line x1="{PL}" y1="{gy:.1f}" x2="{W-PR}" y2="{gy:.1f}" stroke="{col}" stroke-width="1"/>')
        o.append(f'<text x="{PL+3}" y="{gy-2:.1f}" font-size="8" fill="rgba(8,32,56,0.28)" font-family="DM Sans,sans-serif">{pct}%</text>')

    # Probability area
    pts = [(xc(i), yp(probs[i])) for i in range(n)]
    area_d = smooth(pts) + f' L {pts[-1][0]:.1f},{HC} L {PL},{HC} Z'
    o.append(f'<path d="{area_d}" fill="url(#wxpg)"/>')
    o.append(f'<path d="{smooth(pts)}" fill="none" stroke="#1a9bc6" stroke-width="1.5" stroke-linejoin="round"/>')

    # Precipitation bars
    bw = cw * 0.42
    for i, h in enumerate(hourly):
        if h['precip'] >= 0.1:
            bt  = ymm(h['precip'])
            bh  = HC - bt
            bx  = xc(i) - bw / 2
            o.append(f'<rect x="{bx:.1f}" y="{bt:.1f}" width="{bw:.1f}" height="{bh:.1f}" fill="url(#wxbg)" rx="1.5" opacity="0.85"/>')
            if h['precip'] >= 0.4:
                lv = h['precip']
                o.append(f'<text x="{xc(i):.1f}" y="{bt-2.5:.1f}" text-anchor="middle" font-size="7.5" fill="#0c4a6e" font-weight="700" font-family="DM Sans,sans-serif">{lv:.1f}</text>')

    # Temperature line
    tpts = [(xc(i), yt(t)) for i, t in enumerate(temps) if t is not None]
    if len(tpts) > 1:
        o.append(f'<path d="{smooth(tpts)}" fill="none" stroke="#f97316" stroke-width="2" stroke-linecap="round"/>')
        for j, (px, py) in enumerate(tpts):
            if j % 3 == 0:
                o.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="2.5" fill="#f97316" stroke="white" stroke-width="1"/>')
                tv = temps[[k for k, t in enumerate(temps) if t is not None][j]]
                o.append(f'<text x="{px:.1f}" y="{py-5:.1f}" text-anchor="middle" font-size="8" fill="#ea580c" font-weight="600" font-family="DM Sans,sans-serif">{round(tv)}°</text>')

    # Time labels
    for i, h in enumerate(hourly):
        hr = h['hour']
        if hr in [6, 9, 12, 15, 18, 21]:
            lx  = xc(i)
            lbl = '12pm' if hr == 12 else (f'{hr-12}pm' if hr > 12 else f'{hr}am')
            o.append(f'<line x1="{lx:.1f}" y1="0" x2="{lx:.1f}" y2="{HC}" stroke="rgba(8,32,56,0.05)" stroke-width="1"/>')
            o.append(f'<line x1="{lx:.1f}" y1="{HC}" x2="{lx:.1f}" y2="{HC+4}" stroke="rgba(8,32,56,0.25)" stroke-width="1"/>')
            o.append(f'<text x="{lx:.1f}" y="{HC+15}" text-anchor="middle" font-size="9" fill="rgba(8,32,56,0.5)" font-family="DM Sans,sans-serif">{lbl}</text>')

    # Peak probability callout
    mp = max(probs)
    if mp >= 20:
        mpi = probs.index(mp)
        px, py = xc(mpi), yp(mp)
        o.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="3.5" fill="#1a9bc6"/>')
        o.append(f'<text x="{px:.1f}" y="{py-7:.1f}" text-anchor="middle" font-size="9.5" fill="#0c4a6e" font-weight="700" font-family="Outfit,sans-serif">{mp}%</text>')

    # Legend
    lx0 = PL + 4
    o.append(f'<rect x="{lx0}" y="5" width="9" height="7" fill="#1a9bc6" opacity="0.45" rx="1"/>')
    o.append(f'<text x="{lx0+12}" y="12" font-size="7.5" fill="rgba(8,32,56,0.4)" font-family="DM Sans,sans-serif">Rain prob.</text>')
    o.append(f'<rect x="{lx0+65}" y="5" width="9" height="7" fill="#0c4a6e" opacity="0.85" rx="1"/>')
    o.append(f'<text x="{lx0+77}" y="12" font-size="7.5" fill="rgba(8,32,56,0.4)" font-family="DM Sans,sans-serif">Precip mm</text>')
    o.append(f'<line x1="{lx0+140}" y1="8.5" x2="{lx0+152}" y2="8.5" stroke="#f97316" stroke-width="2"/>')
    o.append(f'<text x="{lx0+155}" y="12" font-size="7.5" fill="rgba(8,32,56,0.4)" font-family="DM Sans,sans-serif">Temp °C</text>')

    o.append('</svg>')
    return ''.join(o)

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
def e(v):
    return "" if v is None else html.escape(str(v))

def dept_css(dept):
    l = (dept or "").lower()
    if "vendor" in l:     return "dept-Vendor"
    if "security" in l:   return "dept-Security"
    if "technical" in l:  return "dept-Technical"
    if "theming" in l:    return "dept-Theming"
    if "crew" in l:       return "dept-Crew"
    if "deliver" in l:    return "dept-Deliveries"
    if "clean" in l:      return "dept-Cleaning"
    if "event" in l:      return "dept-Event"
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
        done = ev.get("done", False)
        tr_style = ' class="schedule-done"' if done else ''
        rows.append(
            f'<tr{tr_style}><td class="td-time">{e(ev["time"])}</td>'
            f'<td><span class="dept-badge {dept_css(ev["dept"])}">{e(ev["dept"]) or "&mdash;"}</span></td>'
            f'<td>{e(ev["act"])}</td><td>{e(ev["resp"])}</td><td>{contact_disp}</td>'
            f'<td class="td-num">{e(ev["number"])}</td><td class="td-notes">{e(ev["notes"])}</td></tr>'
        )
    body = "\n".join(rows)
    return ('<div class="schedule-wrap"><table class="schedule-table">'
            '<thead><tr><th>Time</th><th>Dept</th><th>Activity</th><th>Responsibility</th>'
            '<th>Contact / Company</th><th>Number</th><th>Notes</th></tr></thead>'
            f'<tbody>{body}</tbody></table></div>')

def render_weather(wx, iso, idx, hourly=None):
    if not wx:
        return '<p class="empty-section">Weather unavailable. Run script with internet access.</p>'

    code    = wx.get("code", 0)
    icon    = WMO_ICON.get(code, "")
    desc    = WMO_DESC.get(code, "Variable")
    max_t   = round(wx["max"])
    min_t   = round(wx["min"])
    rain    = int(wx["rain"])
    total   = wx["mm"]
    wind    = round(wx["wind"])
    gusts   = round(wx.get("gusts", 0) or 0)
    winddir = wx.get("winddir", 0) or 0
    sunrise = wx.get("sunrise", "")
    sunset  = wx.get("sunset",  "")
    uv_max  = wx.get("uv_max",  0) or 0
    humidity = wx.get("humidity", 0) or 0

    uv_lbl, uv_col = uv_label_color(uv_max)
    wdir_lbl = wind_dir_label(winddir)

    feels = None
    if hourly:
        at = [h.get("apparent_temp") for h in hourly if h.get("apparent_temp") is not None]
        if at:
            feels = round(max(at))


    # Headline
    feels_str = f'&nbsp;&nbsp;Feels {feels}&deg;' if feels else ''
    mm_disp   = f'{total:.1f} mm' if total > 0 else 'No rain'
    headline  = (
        f'<div class="wx-headline">'
        f'<span class="wx-icon">{icon}</span>'
        f'<div class="wx-desc-wrap">'
        f'<span class="wx-desc">{e(desc)}</span>'
        f'<span class="wx-temps"><span class="wx-tmax">&uarr;{max_t}&deg;</span>'
        f'<span class="wx-tmin">&darr;{min_t}&deg;C</span>{feels_str}</span>'
        f'</div>'
        f'<span class="wx-mm-badge">{mm_disp}</span>'
        f'</div>')

    # Chart
    chart_html = ''
    if hourly:
        chart_html = '<div class="wx-chart-wrap">' + render_hourly_chart(hourly) + '</div>'

    # Stats grid
    def stat(lbl, val):
        return f'<div class="wx-stat"><div class="wx-stat-lbl">{lbl}</div><div class="wx-stat-val">{val}</div></div>'

    uv_val = f'<span style="color:{uv_col};font-weight:700">{int(uv_max)} &mdash; {uv_lbl}</span>' if uv_max else '&mdash;'
    stats  = []
    if sunrise: stats.append(stat('Sunrise',     fmt_sunrise(sunrise)))
    if sunset:  stats.append(stat('Sunset',      fmt_sunrise(sunset)))
    if uv_max:  stats.append(stat('UV Index',    uv_val))
    stats.append(stat('Wind',        f'{wind} km/h <span class="wx-stat-sub">{wdir_lbl}</span>'))
    if gusts:   stats.append(stat('Gusts',       f'{gusts} km/h'))
    if humidity: stats.append(stat('Humidity',   f'{int(humidity)}%'))
    stats.append(stat('Rain Chance', f'{rain}%'))
    if total > 0: stats.append(stat('Total Rain', f'{total:.1f} mm'))
    stats.append(stat('Source', 'Open-Meteo &middot; Noosa Beach'))

    stats_html = '<div class="wx-stats">' + ''.join(stats) + '</div>'

    return (
        f'<div class="wx-card">'
        f'{headline}'
        f'{chart_html}'
        f'{stats_html}'
        f'</div>'
    )

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
def _day_video_key(sheet_name):
    """'Monday 8th' → 'Monday_8'"""
    parts = sheet_name.split()
    num = re.sub(r'\D+$', '', parts[1])
    return f"{parts[0]}_{num}"

def find_videos(videos_dir):
    """Scan videos/ folder → dict of {key: filename}, e.g. {'Monday_8': 'Monday_8.mp4'}"""
    result = {}
    if not os.path.isdir(videos_dir):
        return result
    for f in os.listdir(videos_dir):
        ext = os.path.splitext(f)[1].lower()
        if ext in ('.mp4', '.mov', '.m4v'):
            key = os.path.splitext(f)[0]
            result[key] = f
    return result

def render_panel(meta, idx, events, notes, weather_by_date, hourly_by_date=None, video_file=None):
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

    if video_file:
        ext = os.path.splitext(video_file)[1].lower()
        mime = 'video/quicktime' if ext == '.mov' else 'video/mp4'
        vid_src = 'videos/' + video_file
        site_video = (
            '<div class="sitemap-toggle" onclick="toggleSiteMap(this)" style="margin-top:8px">'
            '<span>&#x1F3A5; Site Video</span><span class="sitemap-arrow">&#x25BC;</span></div>'
            '<div class="sitemap-body">'
            f'<video controls style="width:100%;display:block;max-height:72vh;background:#000">'
            f'<source src="{vid_src}" type="{mime}">'
            'Your browser does not support video playback.'
            '</video>'
            '</div>')
    else:
        site_video = ''

    parts = []
    parts.append(f'<div class="ds-section"><h2 class="section-heading">Overview / Key Notes</h2>{bullets(n.get("overview"))}{sitemap}{site_video}</div>')
    if n.get("changes"):   parts.append(sec("Changes", bullets(n["changes"])))
    if n.get("flags"):     parts.append(sec("Flags",   bullets(n["flags"])))
    parts.append(sec("Schedule", render_schedule(events)))
    parts.append(sec("Before End of Day", bullets(n.get('before_eod'))))
    hourly = (hourly_by_date or {}).get(iso)
    parts.append(sec("Weather",  render_weather(wx, iso, idx, hourly)))
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
CSS = "\n:root{--blue:#d5f0fe;--blue-light:rgba(213,240,254,.30);--blue-mid:#1a9bc6;--dark:#082038;--text:rgba(0,0,0,.82);--muted:rgba(0,0,0,.45);--white:#fff;--nav-w:220px;--r:10px;--c-event:#F97316;--c-site:#16A34A;--c-tech:#2563EB;--c-vendor:#D97706;--c-security:#DC2626;--c-theming:#7C3AED;--c-crew:#525252;--c-deliveries:#0891B2;--c-cleaning:#65A30D}\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:'DM Sans',sans-serif;font-size:14px;color:var(--text);background:var(--white);letter-spacing:.003em;line-height:1.55}\n.ticker{position:fixed;top:0;left:var(--nav-w);right:0;z-index:500;padding:10px 48px 10px 20px;display:flex;align-items:center;gap:12px;font-size:13px;font-weight:500;transition:transform .3s ease}\n.ticker.hidden{transform:translateY(-110%)}\n.ticker-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}\n.ticker-close{position:absolute;right:16px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;font-size:20px;line-height:1;opacity:.6;padding:4px}\n.ticker-close:hover{opacity:1}\n.ticker.green{background:#dcfce7;color:#14532d;border-bottom:1px solid #86efac}.ticker.green .ticker-dot{background:#16a34a}\n.ticker.orange{background:#fff7ed;color:#7c2d12;border-bottom:1px solid #fdba74}.ticker.orange .ticker-dot{background:#f97316}\n.ticker.red{background:#fef2f2;color:#7f1d1d;border-bottom:1px solid #fca5a5}.ticker.red .ticker-dot{background:#dc2626}\nbody.has-ticker .main{padding-top:42px}\n.sidebar{position:fixed;top:0;left:0;width:var(--nav-w);height:100vh;background:var(--blue);display:flex;flex-direction:column;overflow-y:auto;overflow-x:hidden;z-index:100;scrollbar-width:none}\n.sidebar::-webkit-scrollbar{display:none}\n.sidebar-logo{padding:22px 18px 14px;flex-shrink:0}\n.sidebar-logo img{width:62px;height:62px;display:block}\n.sidebar-meta{padding:0 18px 16px;border-bottom:1px solid rgba(8,32,56,.12);flex-shrink:0}\n.sidebar-meta .fn{font-family:'Outfit',sans-serif;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dark);line-height:1.3;margin-bottom:2px}\n.sidebar-meta .fs{font-size:11px;color:rgba(8,32,56,.55)}\n.day-nav{list-style:none;padding:10px 0 20px;flex:1}\n.day-nav li a{display:block;padding:9px 18px;text-decoration:none;color:rgba(8,32,56,.7);font-size:12px;font-weight:500;line-height:1.3;transition:background .15s;border-left:3px solid transparent;cursor:pointer}\n.day-nav li a .dn{display:block;font-family:'Outfit',sans-serif;font-weight:700;font-size:13px;color:var(--dark);line-height:1.2}\n.day-nav li a .dt{font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:rgba(8,32,56,.45)}\n.day-nav li a:hover{background:rgba(8,32,56,.07)}\n.day-nav li a.active{background:rgba(8,32,56,.1);border-left-color:var(--dark)}\n.sidebar-footer{padding:14px 18px;border-top:1px solid rgba(8,32,56,.1);font-size:11px;color:rgba(8,32,56,.45);flex-shrink:0}\n.main{margin-left:var(--nav-w);min-height:100vh}\n.day-panel{display:none}.day-panel.active{display:block}\n.day-header{background:var(--blue);padding:36px 48px 28px;border-bottom:1px solid rgba(8,32,56,.08)}\n.day-header .lbl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:rgba(8,32,56,.5);margin-bottom:6px}\n.day-header h1{font-family:'Outfit',sans-serif;font-size:52px;font-weight:300;letter-spacing:-.03em;color:var(--dark);line-height:1;margin-bottom:6px}\n.day-header .dd{font-family:'Outfit',sans-serif;font-size:18px;font-weight:400;color:rgba(8,32,56,.6);letter-spacing:-.01em}\n.day-body{padding:0 48px 60px}\n.ds-section{padding:32px 0 28px;border-bottom:1px solid rgba(8,32,56,.07)}\n.ds-section:last-child{border-bottom:none}\n.section-heading{font-family:'Outfit',sans-serif;font-size:28px;font-weight:300;color:var(--blue-mid);letter-spacing:-.02em;margin-bottom:16px;line-height:1.1}\n.empty-section{color:var(--muted);font-size:13px;font-style:italic}\n.sitemap-toggle{display:flex;align-items:center;gap:10px;cursor:pointer;user-select:none;background:var(--blue-light);border:1px solid rgba(8,32,56,.1);border-radius:var(--r);padding:12px 16px;margin-top:16px;font-family:'Outfit',sans-serif;font-size:14px;font-weight:600;color:var(--dark);transition:background .2s}\n.sitemap-toggle:hover{background:rgba(213,240,254,.6)}\n.sitemap-arrow{font-size:11px;transition:transform .25s;margin-left:auto;color:var(--muted)}\n.sitemap-toggle.open .sitemap-arrow{transform:rotate(180deg)}\n.sitemap-body{display:none;margin-top:8px;border-radius:var(--r);overflow:hidden;border:1px solid rgba(8,32,56,.1)}\n.sitemap-body.open{display:block}\n.sitemap-body embed{display:block;width:100%;height:72vh;border:none}\n.sitemap-fallback{padding:16px;color:var(--muted);font-size:12px;background:#f8f9fa}\n.bullet-list{list-style:none;display:flex;flex-direction:column;gap:8px}\n.bullet-list li{padding-left:22px;position:relative;font-size:14px;line-height:1.55}\n.bullet-list li::before{content:'';position:absolute;left:0;top:9px;width:8px;height:8px;border:1.5px solid var(--dark);border-radius:50%}\n.schedule-wrap{overflow-x:auto;margin-top:4px}\n.schedule-table{width:100%;border-collapse:collapse;font-size:12.5px;min-width:780px}\n.schedule-table thead th{background:var(--dark);color:var(--blue);padding:9px 12px;text-align:left;font-family:'Outfit',sans-serif;font-weight:600;font-size:10.5px;text-transform:uppercase;letter-spacing:.1em;white-space:nowrap}\n.schedule-table tbody tr{border-bottom:1px solid rgba(8,32,56,.07)}\n.schedule-table tbody tr:nth-child(even){background:rgba(213,240,254,.18)}\n.schedule-table tbody tr:hover{background:rgba(213,240,254,.35)}\n.schedule-table td{padding:8px 12px;vertical-align:top}\n.td-time{font-family:'Outfit',sans-serif;font-weight:600;font-size:13px;color:var(--dark);white-space:nowrap;width:70px}\n.td-num{white-space:nowrap}\n.td-notes{font-size:11.5px;color:var(--muted);font-style:italic}\n.dept-badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;color:white}\n.dept-Event{background:var(--c-event)}.dept-Site{background:var(--c-site)}.dept-Technical{background:var(--c-tech)}\n.dept-Vendor{background:var(--c-vendor)}.dept-Security{background:var(--c-security)}.dept-Theming{background:var(--c-theming)}.dept-Crew{background:var(--c-crew)}.dept-Deliveries{background:var(--c-deliveries)}.dept-Cleaning{background:var(--c-cleaning)}\n.weather-card{background:var(--blue-light);border:1px solid rgba(8,32,56,.1);border-radius:var(--r);padding:20px 24px;display:grid;grid-template-columns:1fr auto;gap:16px 24px}\n.weather-date-row{grid-column:1/-1;font-family:'Outfit',sans-serif;font-weight:600;font-size:13px;color:rgba(8,32,56,.6);text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid rgba(8,32,56,.1);padding-bottom:10px}\n.weather-temps{display:flex;align-items:baseline;gap:12px;margin-bottom:6px}\n.temp-min{font-size:15px;color:var(--muted)}.temp-max{font-family:'Outfit',sans-serif;font-size:36px;font-weight:300;color:var(--dark);letter-spacing:-.02em}.temp-unit{font-size:16px;color:var(--muted)}\n.weather-summary{font-weight:600;font-size:13px;color:var(--dark);margin-bottom:8px}\n.rain-row{display:flex;align-items:center;gap:10px;font-size:12.5px;margin-bottom:4px}\n.rain-bar{flex:1;max-width:100px;height:5px;background:rgba(8,32,56,.12);border-radius:3px;overflow:hidden}\n.rain-fill{height:100%;background:var(--blue-mid);border-radius:3px}\n.rain-mm{font-size:11px;color:var(--muted)}\n.weather-extras{display:flex;gap:20px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin-top:8px;padding-top:8px;border-top:1px solid rgba(8,32,56,.08)}\n.weather-extras strong{color:var(--text);font-weight:600}\n.weather-left{}\n.weather-map-wrap{flex-shrink:0;text-align:center}\n.weather-map-wrap img{width:520px;max-width:100%;border-radius:6px;border:1px solid rgba(8,32,56,.12);display:block}\n.map-label{font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);margin-top:6px;font-weight:600}\n.map-caption{font-size:11px;color:rgba(8,32,56,.45);margin-top:3px}\n.radio-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}\n.radio-table{width:100%;border-collapse:collapse;font-size:12.5px}\n.radio-table thead th{background:rgba(8,32,56,.06);color:var(--dark);padding:7px 10px;text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;font-weight:700}\n.radio-table td{padding:6px 10px;border-bottom:1px solid rgba(8,32,56,.06)}\n.radio-table td:first-child{font-family:'Outfit',sans-serif;font-weight:600;color:var(--dark);width:30px;text-align:center}\n.wifi-box{background:var(--dark);color:white;border-radius:var(--r);padding:20px 24px;display:grid;grid-template-columns:1fr 1fr;gap:12px 32px}\n.wifi-box h4{grid-column:1/-1;font-family:'Outfit',sans-serif;font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:var(--blue);margin-bottom:4px}\n.wifi-row .wl{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:rgba(213,240,254,.6);margin-bottom:2px}\n.wifi-row .wv{font-family:'Outfit',sans-serif;font-weight:600;color:white}\n.wifi-row .wv.int{color:var(--blue)}\n.contacts-btn{position:fixed;top:18px;right:24px;z-index:600;background:var(--dark);color:var(--blue);border:none;border-radius:8px;padding:9px 18px;font-family:'Outfit',sans-serif;font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;transition:background .2s,transform .15s,top .3s ease;box-shadow:0 2px 8px rgba(8,32,56,.25)}.contacts-btn:hover{background:#0d3055;transform:translateY(-1px)}body.has-ticker .contacts-btn{top:58px}body.has-ticker .day-badge{top:100px}.contacts-backdrop{position:fixed;inset:0;background:rgba(8,32,56,.55);z-index:700;opacity:0;pointer-events:none;transition:opacity .3s ease}.contacts-backdrop.open{opacity:1;pointer-events:all}.contacts-panel{position:fixed;top:0;right:0;width:460px;max-width:100vw;height:100vh;background:var(--white);z-index:800;display:flex;flex-direction:column;transform:translateX(100%);transition:transform .35s cubic-bezier(.4,0,.2,1);box-shadow:-4px 0 32px rgba(8,32,56,.18)}.contacts-panel.open{transform:translateX(0)}.contacts-panel-head{padding:24px 24px 20px;background:var(--dark);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}.contacts-panel-head h2{font-family:'Outfit',sans-serif;font-size:18px;font-weight:600;color:var(--blue);letter-spacing:.02em}.contacts-close{background:none;border:none;color:rgba(213,240,254,.6);font-size:26px;line-height:1;cursor:pointer;padding:4px;transition:color .15s}.contacts-close:hover{color:var(--blue)}.contacts-search{padding:14px 16px;border-bottom:1px solid rgba(8,32,56,.08);flex-shrink:0}.contacts-search input{width:100%;border:1px solid rgba(8,32,56,.15);border-radius:6px;padding:8px 12px;font-size:13px;font-family:'DM Sans',sans-serif;outline:none;color:var(--dark);transition:border-color .15s}.contacts-search input:focus{border-color:var(--blue-mid)}.contacts-list{flex:1;overflow-y:auto;padding:12px 0}.contact-card{display:flex;align-items:center;gap:14px;padding:12px 20px;border-bottom:1px solid rgba(8,32,56,.06);transition:background .15s}.contact-card:hover{background:rgba(213,240,254,.35)}.contact-avatar{width:38px;height:38px;border-radius:50%;background:var(--blue);display:flex;align-items:center;justify-content:center;flex-shrink:0;font-family:'Outfit',sans-serif;font-weight:700;font-size:13px;color:var(--dark)}.contact-info{flex:1;min-width:0}.contact-company{font-family:'Outfit',sans-serif;font-size:13px;font-weight:700;color:var(--dark);line-height:1.3}.contact-name{font-size:12px;color:var(--muted);margin-top:1px}.contact-number{flex-shrink:0;text-align:right}.contact-number a{font-family:'Outfit',sans-serif;font-size:13px;font-weight:600;color:var(--blue-mid);text-decoration:none;white-space:nowrap}.contact-number a:hover{text-decoration:underline}.contact-no-number{font-size:11px;color:rgba(8,32,56,.25);font-style:italic}.contacts-empty{padding:40px 24px;text-align:center;color:var(--muted);font-size:13px}@media(max-width:900px){:root{--nav-w:0px}.sidebar{transform:translateX(-220px);transition:transform .3s ease;width:220px}.sidebar.open{transform:translateX(0)}.ticker{left:0}.mobile-toggle{display:flex;position:fixed;top:12px;left:12px;z-index:600;width:40px;height:40px;background:var(--blue);border-radius:6px;align-items:center;justify-content:center;cursor:pointer;border:none;flex-direction:column;gap:4px;padding:10px;transition:top .3s ease}.mobile-toggle span{display:block;width:20px;height:2px;background:var(--dark);border-radius:2px}body.has-ticker .mobile-toggle{top:52px}.day-header,.day-body{padding-left:20px;padding-right:20px}.day-header{padding-top:60px}.weather-card{grid-template-columns:1fr}.weather-map-wrap img{width:100%;max-width:340px}.radio-grid,.wifi-box{grid-template-columns:1fr}}\n@media(min-width:901px){.mobile-toggle{display:none}}\n.schedule-done td{color:rgba(8,32,56,0.3)!important;text-decoration:line-through}\n.schedule-done .dept-badge{opacity:0.35}\n.wx-card{display:flex;flex-direction:column;gap:14px}\n.wx-headline{display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding-bottom:14px;border-bottom:1px solid rgba(8,32,56,.08)}\n.wx-icon{font-size:36px;line-height:1;flex-shrink:0}\n.wx-desc-wrap{flex:1;min-width:0}\n.wx-desc{display:block;font-family:'Outfit',sans-serif;font-size:19px;font-weight:500;color:var(--dark);line-height:1.2}\n.wx-temps{display:flex;align-items:baseline;gap:8px;margin-top:3px;font-size:14px}\n.wx-tmax{font-family:'Outfit',sans-serif;font-weight:700;font-size:22px;color:var(--dark)}\n.wx-tmin{color:var(--muted);font-size:15px}\n.wx-mm-badge{flex-shrink:0;background:var(--dark);color:var(--blue);font-family:'Outfit',sans-serif;font-size:13px;font-weight:700;padding:5px 12px;border-radius:999px}\n.wx-chart-wrap{border-radius:8px;overflow:hidden;padding:10px 6px 4px;background:rgba(213,240,254,.18);border:1px solid rgba(8,32,56,.06)}\n.wx-advisory{padding:9px 14px;border-radius:6px;font-size:12.5px;font-weight:500;display:flex;align-items:center;gap:8px;line-height:1.4}\n.wx-adv-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}\n.wx-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:2px 16px}\n.wx-stat{padding:9px 0;border-bottom:1px solid rgba(8,32,56,.05)}\n.wx-stat-lbl{font-size:9.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);font-weight:700;margin-bottom:3px}\n.wx-stat-val{font-family:'Outfit',sans-serif;font-size:14px;font-weight:600;color:var(--dark);line-height:1.3}\n.wx-stat-sub{font-family:'DM Sans',sans-serif;font-weight:400;font-size:12px;color:var(--muted);margin-left:3px}\n@media(max-width:900px){.wx-stats{grid-template-columns:repeat(2,1fr)}}\n.day-badge{position:fixed;top:58px;right:24px;z-index:600;background:var(--white);border:2px solid var(--dark);border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;padding:9px 18px;box-shadow:0 2px 8px rgba(8,32,56,.15);transition:top .3s ease}\n.day-badge-lbl{font-family:'DM Sans',sans-serif;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:rgba(8,32,56,.5);line-height:1}\n.day-badge-num{font-family:'Outfit',sans-serif;font-size:26px;font-weight:700;color:var(--dark);line-height:1.2}\n.day-nav li.day-archived a{color:rgba(8,32,56,.3)}\n.day-nav li.day-archived a .dn{color:rgba(8,32,56,.3)}\n.day-nav li.day-archived a .dt{color:rgba(8,32,56,.2)}\n.day-nav li.day-archived a:hover{background:rgba(8,32,56,.04)}\n.day-nav li.day-archived a.active{color:rgba(8,32,56,.5);border-left-color:rgba(8,32,56,.3)}\n.day-nav li.day-archived a.active .dn{color:rgba(8,32,56,.5)}\n.shopping-list-btn{position:fixed;top:18px;right:136px;z-index:600;background:var(--blue-mid);color:white;border:none;border-radius:8px;padding:9px 18px;font-family:'Outfit',sans-serif;font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;transition:background .2s,transform .15s,top .3s ease;box-shadow:0 2px 8px rgba(8,32,56,.25)}.shopping-list-btn:hover{background:#1589b0;transform:translateY(-1px)}body.has-ticker .shopping-list-btn{top:58px}.shopping-backdrop{position:fixed;inset:0;background:rgba(8,32,56,.55);z-index:700;opacity:0;pointer-events:none;transition:opacity .3s ease}.shopping-backdrop.open{opacity:1;pointer-events:all}.shopping-panel{position:fixed;top:0;right:0;width:min(920px,95vw);height:100vh;background:var(--white);z-index:800;display:flex;flex-direction:column;transform:translateX(100%);transition:transform .35s cubic-bezier(.4,0,.2,1);box-shadow:-4px 0 32px rgba(8,32,56,.18)}.shopping-panel.open{transform:translateX(0)}.shopping-panel-head{padding:24px 24px 20px;background:var(--dark);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}.shopping-panel-head h2{font-family:'Outfit',sans-serif;font-size:18px;font-weight:600;color:var(--blue);letter-spacing:.02em}.shopping-close{background:none;border:none;color:rgba(213,240,254,.6);font-size:26px;line-height:1;cursor:pointer;padding:4px;transition:color .15s}.shopping-close:hover{color:var(--blue)}.shopping-filters{padding:12px 16px;border-bottom:1px solid rgba(8,32,56,.08);flex-shrink:0;display:flex;gap:10px;flex-wrap:wrap;align-items:center}.shopping-filters input,.shopping-filters select{border:1px solid rgba(8,32,56,.15);border-radius:6px;padding:7px 11px;font-size:13px;font-family:'DM Sans',sans-serif;outline:none;color:var(--dark);background:white;transition:border-color .15s}.shopping-filters input{flex:1;min-width:160px}.shopping-filters input:focus,.shopping-filters select:focus{border-color:var(--blue-mid)}.shopping-table-wrap{flex:1;overflow-y:auto;overflow-x:auto}.shopping-table{width:100%;border-collapse:collapse;font-size:12.5px;min-width:700px}.shopping-table thead th{background:var(--dark);color:var(--blue);padding:9px 12px;text-align:left;font-family:'Outfit',sans-serif;font-weight:600;font-size:10.5px;text-transform:uppercase;letter-spacing:.1em;white-space:nowrap;cursor:pointer;user-select:none}.shopping-table thead th:hover{background:#0d3055}.shopping-table thead th.sort-asc::after{content:' ↑'}.shopping-table thead th.sort-desc::after{content:' ↓'}.shopping-table tbody tr{border-bottom:1px solid rgba(8,32,56,.07)}.shopping-table tbody tr:nth-child(even){background:rgba(213,240,254,.15)}.shopping-table tbody tr:hover{background:rgba(213,240,254,.35)}.shopping-table td{padding:8px 12px;vertical-align:top}.sl-item-name{font-weight:600;color:var(--dark)}.sl-notes{font-size:11.5px;color:var(--muted);font-style:italic;max-width:180px}.sl-type{display:inline-block;padding:2px 7px;border-radius:999px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap}.sl-type-consumable{background:#fef3c7;color:#92400e}.sl-type-asset{background:#dbeafe;color:#1e40af}.sl-status{display:inline-block;padding:2px 7px;border-radius:999px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}.sl-status-pending{background:#fef3c7;color:#92400e}.sl-status-ordered{background:#e5e7eb;color:#6b7280}.sl-status-checking-pallet{background:#ede9fe;color:#6d28d9}.sl-status-received,.sl-status-done{background:#d1fae5;color:#065f46}.sl-count{font-size:11px;color:rgba(8,32,56,.45);padding:8px 16px;flex-shrink:0;border-top:1px solid rgba(8,32,56,.07)}.sl-row-ordered td{color:rgba(8,32,56,.3)!important}.sl-row-ordered .sl-item-name{color:rgba(8,32,56,.3)!important}.sl-row-ordered .sl-type,.sl-row-ordered .sl-status{opacity:0.45}.sl-row-ordered{background:rgba(0,0,0,.015)!important}.sl-divider td{padding:6px 12px;background:#f3f4f6;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:rgba(8,32,56,.35);border-bottom:2px solid #e5e7eb}.sl-link-btn{display:inline-block;padding:2px 8px;background:var(--blue-mid);color:#fff!important;border-radius:4px;font-size:10px;font-weight:700;text-decoration:none!important;letter-spacing:.04em;text-transform:uppercase;white-space:nowrap;transition:background .15s}.sl-link-btn:hover{background:#1589b0}\n"
JS  = '\nfunction showTicker(text, urgency) {\n  var el = document.getElementById(\'ticker\');\n  if (!text) { hideTicker(); return; }\n  el.className = \'ticker \' + (urgency || \'green\');\n  el.innerHTML = \'<span class="ticker-dot"></span><span class="ticker-text">\' + text + \'</span>\'\n    + \'<button class="ticker-close" onclick="hideTicker()" aria-label="Dismiss">&times;</button>\';\n  document.body.classList.add(\'has-ticker\');\n}\nfunction hideTicker() {\n  document.getElementById(\'ticker\').className = \'ticker hidden\';\n  document.body.classList.remove(\'has-ticker\');\n}\nfunction showDay(idx) {\n  document.querySelectorAll(\'.day-panel\').forEach(function(p){ p.classList.remove(\'active\'); });\n  document.querySelectorAll(\'#day-nav a\').forEach(function(a){ a.classList.remove(\'active\'); });\n  var panel = document.getElementById(\'panel-\' + idx);\n  panel.classList.add(\'active\');\n  document.querySelector(\'#day-nav a[data-day="\' + idx + \'"]\').classList.add(\'active\');\n  var badge = document.getElementById(\'day-badge-num\');\n  if (badge && DAY_NUMS) badge.textContent = DAY_NUMS[idx] || (idx + 1);\n  document.title = \'NFW 2026 — Day \' + (DAY_NUMS[idx] || (idx + 1));\n  window.scrollTo(0,0);\n  document.querySelector(\'.sidebar\').classList.remove(\'open\');\n  if (panel.dataset.tickerActive === \'true\') {\n    showTicker(panel.dataset.tickerText, panel.dataset.tickerUrgency);\n  } else { hideTicker(); }\n}\nfunction toggleSiteMap(btn) {\n  btn.classList.toggle(\'open\');\n  btn.nextElementSibling.classList.toggle(\'open\');\n}\nfunction brisbaneNow() {\n  var now = new Date();\n  var utc = now.getTime() + now.getTimezoneOffset() * 60000;\n  return new Date(utc + 10 * 3600000);\n}\nfunction brisIso(d) {\n  return d.getFullYear() + \'-\' + (\'0\'+(d.getMonth()+1)).slice(-2) + \'-\' + (\'0\'+d.getDate()).slice(-2);\n}\ndocument.addEventListener(\'DOMContentLoaded\',function(){\n  var bris = brisbaneNow();\n  var todayIso = brisIso(bris);\n  var isPM = bris.getHours() >= 17;\n  var targetIso = isPM ? brisIso(new Date(bris.getTime() + 86400000)) : todayIso;\n  var defaultIdx = 0;\n  for (var i = 0; i < DAY_ISOS.length; i++) {\n    if (DAY_ISOS[i] === targetIso) { defaultIdx = i; break; }\n    if (DAY_ISOS[i] > targetIso && DAY_ISOS[i-1] < targetIso) { defaultIdx = i; break; }\n  }\n  showDay(defaultIdx);\n  document.querySelectorAll(\'#day-nav a[data-iso]\').forEach(function(a) {\n    if (a.dataset.iso < targetIso) a.closest(\'li\').classList.add(\'day-archived\');\n  });\n});\n'

# ── GENERATE HTML ─────────────────────────────────────────────────────────────
def generate_html(days_data, logo_b64, generated_at, contacts=None, shopping_list=None):
    nav = "".join(
        f'<li><a href="#" data-day="{i}" data-iso="{m["iso"]}" class="{"active" if i==0 else ""}"'
        f'    onclick="showDay({i});return false;">'
        f'<span class="dn">Day {m["day"]} &middot; {e(m["short"])}</span>'
        f'<span class="dt">{e(m["type"])}</span></a></li>'
        for i, (m, panel) in enumerate(days_data))
    import json as _json2
    day_isos_js = _json2.dumps([m["iso"] for m, _ in days_data])
    day_nums_js = _json2.dumps([m["day"] for m, _ in days_data])
    logo_tag = (f'<img src="data:image/png;base64,{logo_b64}" alt="NF&amp;W">'
                if logo_b64 else
                '<div style="width:62px;height:62px;background:var(--dark);border-radius:6px;'
                'display:flex;align-items:center;justify-content:center;color:white;'
                'font-family:Outfit,sans-serif;font-weight:700;font-size:14px">NF&amp;W</div>')
    panels = "".join(panel for _, panel in days_data)
    import json as _json
    contacts_json = _json.dumps(contacts or [], ensure_ascii=False)
    shopping_json = _json.dumps(shopping_list or [], ensure_ascii=False)
    shopping_js = (
        "var SHOPPING_DATA = " + shopping_json + ";"
        """
var slSortCol = -1;
var slSortAsc = true;
function openShoppingList() {
  document.getElementById('shopping-backdrop').classList.add('open');
  document.getElementById('shopping-panel').classList.add('open');
  document.getElementById('shopping-search').value = '';
  document.getElementById('sl-filter-type').value = '';
  document.getElementById('sl-filter-status').value = '';
  slSortCol = -1; slSortAsc = true;
  document.querySelectorAll('#sl-thead th').forEach(function(th){ th.classList.remove('sort-asc','sort-desc'); });
  renderShoppingList();
  setTimeout(function(){ document.getElementById('shopping-search').focus(); }, 350);
}
function closeShoppingList() {
  document.getElementById('shopping-backdrop').classList.remove('open');
  document.getElementById('shopping-panel').classList.remove('open');
}
function linkify(text) {
  if (!text) return '';
  return String(text).replace(/https?:\/\/[^\s<>"']+/g, function(url) {
    return '<a href="' + url + '" target="_blank" rel="noopener" class="sl-link-btn">Link</a>';
  });
}
function renderShoppingList() {
  var search = (document.getElementById('shopping-search').value || '').toLowerCase();
  var typeF  = (document.getElementById('sl-filter-type').value || '').toLowerCase();
  var statF  = (document.getElementById('sl-filter-status').value || '').toLowerCase();
  var all = SHOPPING_DATA.filter(function(r) {
    if (typeF  && r.type.toLowerCase()   !== typeF)  return false;
    if (statF  && r.status.toLowerCase() !== statF)  return false;
    if (search) {
      var t = [r.purpose,r.type,r.item,r.qty,r.source,r.status,r.supply_to,r.notes].join(' ').toLowerCase();
      if (t.indexOf(search) === -1) return false;
    }
    return true;
  });
  var active  = all.filter(function(r){ return r.status.toLowerCase() !== 'ordered'; });
  var ordered = all.filter(function(r){ return r.status.toLowerCase() === 'ordered'; });
  if (slSortCol >= 0) {
    var keys = ['purpose','type','item','qty','source','status','supply_to','notes'];
    var k = keys[slSortCol];
    function cmp(a, b) {
      var av = String(a[k]||'').toLowerCase(), bv = String(b[k]||'').toLowerCase();
      return slSortAsc ? av.localeCompare(bv) : bv.localeCompare(av);
    }
    active.sort(cmp);
    ordered.sort(cmp);
  }
  var tbody = document.getElementById('sl-tbody');
  var countEl = document.getElementById('sl-count');
  var total = active.length + ordered.length;
  countEl.textContent = total + ' item' + (total !== 1 ? 's' : '') + ' shown'
    + (ordered.length ? ' · ' + ordered.length + ' ordered (archived)' : '');
  if (!total) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:40px;color:rgba(0,0,0,.4);font-style:italic">No items found.</td></tr>';
    return;
  }
  function renderRow(r, cls) {
    var tCls = 'sl-type sl-type-' + r.type.toLowerCase();
    var sCls = 'sl-status sl-status-' + r.status.toLowerCase().replace(/[\s\/]+/g,'-');
    return '<tr class="' + cls + '">'
      + '<td>' + (r.purpose||'') + '</td>'
      + '<td><span class="' + tCls + '">' + (r.type||'') + '</span></td>'
      + '<td class="sl-item-name">' + (r.item||'') + '</td>'
      + '<td>' + (r.qty||'') + '</td>'
      + '<td>' + (r.source||'') + '</td>'
      + '<td><span class="' + sCls + '">' + (r.status||'') + '</span></td>'
      + '<td>' + (r.supply_to||'') + '</td>'
      + '<td class="sl-notes">' + linkify(r.notes) + '</td>'
      + '</tr>';
  }
  var rows = active.map(function(r){ return renderRow(r, ''); });
  if (ordered.length) {
    rows.push('<tr class="sl-divider"><td colspan="8">Ordered &mdash; awaiting delivery</td></tr>');
    ordered.forEach(function(r){ rows.push(renderRow(r, 'sl-row-ordered')); });
  }
  tbody.innerHTML = rows.join('');
}
function slSort(col) {
  if (slSortCol === col) { slSortAsc = !slSortAsc; } else { slSortCol = col; slSortAsc = true; }
  document.querySelectorAll('#sl-thead th').forEach(function(th,i){
    th.classList.remove('sort-asc','sort-desc');
    if (i === col) th.classList.add(slSortAsc ? 'sort-asc' : 'sort-desc');
  });
  renderShoppingList();
}
document.addEventListener('DOMContentLoaded', function() {
  renderShoppingList();
  document.getElementById('shopping-search').addEventListener('input', renderShoppingList);
  document.getElementById('sl-filter-type').addEventListener('change', renderShoppingList);
  document.getElementById('sl-filter-status').addEventListener('change', renderShoppingList);
  document.getElementById('shopping-backdrop').addEventListener('click', closeShoppingList);
  document.addEventListener('keydown', function(ev) { if (ev.key === 'Escape') closeShoppingList(); });
});
"""
    )
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
<button class="shopping-list-btn" onclick="openShoppingList()">Shopping List</button>
<button class="contacts-btn" onclick="openContacts()">Contacts</button>
<div class="day-badge" id="day-badge"><div class="day-badge-lbl">Day</div><div class="day-badge-num" id="day-badge-num">1</div></div>
<div class="shopping-backdrop" id="shopping-backdrop"></div>
<div class="shopping-panel" id="shopping-panel">
  <div class="shopping-panel-head">
    <h2>Shopping List</h2>
    <button class="shopping-close" onclick="closeShoppingList()" aria-label="Close">&times;</button>
  </div>
  <div class="shopping-filters">
    <input type="search" id="shopping-search" placeholder="Search items&hellip;">
    <select id="sl-filter-type">
      <option value="">All Types</option>
      <option value="consumable">Consumable</option>
      <option value="asset">Asset</option>
    </select>
    <select id="sl-filter-status">
      <option value="">All Statuses</option>
      <option value="pending">Pending</option>
      <option value="checking pallet">Checking Pallet</option>
      <option value="ordered">Ordered</option>
      <option value="received">Received</option>
      <option value="done">Done</option>
    </select>
  </div>
  <div class="shopping-table-wrap">
    <table class="shopping-table">
      <thead id="sl-thead">
        <tr>
          <th onclick="slSort(0)">Purpose</th>
          <th onclick="slSort(1)">Type</th>
          <th onclick="slSort(2)">Item</th>
          <th onclick="slSort(3)">Qty / Unit</th>
          <th onclick="slSort(4)">Source</th>
          <th onclick="slSort(5)">Status</th>
          <th onclick="slSort(6)">Supply To</th>
          <th onclick="slSort(7)">Notes</th>
        </tr>
      </thead>
      <tbody id="sl-tbody"></tbody>
    </table>
  </div>
  <div class="sl-count" id="sl-count"></div>
</div>
<div class="contacts-backdrop" id="contacts-backdrop"></div>
<div class="contacts-panel" id="contacts-panel">
  <div class="contacts-panel-head">
    <h2>Contacts</h2>
    <button class="contacts-close" onclick="closeContacts()" aria-label="Close">&times;</button>
  </div>
  <div class="contacts-search"><input type="search" id="contacts-search" placeholder="Search company, name or number&hellip;"></div>
  <div class="contacts-list" id="contacts-list"></div>
</div>
<script>var DAY_ISOS={day_isos_js};var DAY_NUMS={day_nums_js};</script>
<script>{JS}</script>
<script>{shopping_js}</script>
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


def read_shopping_list_from_excel(wb):
    """Read shopping list from 'Shopping List' sheet."""
    if 'Shopping List' not in wb.sheetnames:
        print("  Shopping list: 'Shopping List' sheet not found.")
        return []
    ws = wb['Shopping List']
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        cols = (list(row) + [None]*9)[:9]
        # Columns: EVENT(skip), Purpose, Type, Item, Qty/Unit, Source, Status, Supply To, Notes
        purpose   = str(cols[1] or '').strip()
        item_type = str(cols[2] or '').strip()
        item      = str(cols[3] or '').strip()
        qty       = str(cols[4] or '').strip()
        source    = str(cols[5] or '').strip()
        status    = str(cols[6] or '').strip() or 'Pending'
        supply_to = str(cols[7] or '').strip()
        notes     = str(cols[8] or '').strip()
        if item:
            items.append({
                'purpose': purpose, 'type': item_type, 'item': item,
                'qty': qty, 'source': source, 'status': status,
                'supply_to': supply_to, 'notes': notes,
            })
    print(f"  Shopping list: {len(items)} items loaded.")
    return items


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

    contacts      = read_contacts_from_excel(wb)
    shopping_list = read_shopping_list_from_excel(wb)
    notes_by_day  = read_notes_from_excel(wb)

    print("  Weather: fetching from Open-Meteo...")
    weather_by_date, hourly_by_date = fetch_weather()

    videos_dir = str(excel_path.parent / 'videos')
    all_videos = find_videos(videos_dir)
    if all_videos:
        print(f"  Videos: found {len(all_videos)} — {', '.join(sorted(all_videos))}")
    else:
        print("  Videos: none found in videos/ folder")

    days_data = []
    for i, meta in enumerate(DAYS_META):
        events = parse_schedule(wb, meta["sheet"]) if meta["sheet"] in wb.sheetnames else []
        notes  = notes_by_day.get(meta["day"], {})
        idx    = len(days_data)
        # Show previous day's video on this panel
        video_file = None
        if i > 0:
            prev_key = _day_video_key(DAYS_META[i - 1]["sheet"])
            video_file = all_videos.get(prev_key)
        panel  = render_panel(meta, idx, events, notes, weather_by_date, hourly_by_date, video_file)
        days_data.append((meta, panel))
        vid_note = f'  video: {video_file}' if video_file else ''
        print(f'  Day {meta["day"]}: {meta["date"]:25s}  {len(events):2d} schedule events{vid_note}')

    html_out = generate_html(days_data, logo_b64, dt.now().strftime("%d %b %Y %H:%M"), contacts, shopping_list)
    Path(args.output).write_text(html_out, encoding="utf-8")
    size_kb = Path(args.output).stat().st_size // 1024
    print(f"\n  Output: {args.output}  ({size_kb} KB)")
    print("\nPut these files together in your OneDrive folder:")
    print(f"   {args.output}")
    print(f"   site-map.pdf")
    print(f"   day_notes.json  (edit this to update editorial content)")

if __name__ == "__main__":
    main()