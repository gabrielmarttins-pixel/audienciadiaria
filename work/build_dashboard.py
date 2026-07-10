from __future__ import annotations

import base64
import html
import json
import math
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(r"C:\Users\gamartin\Documents\Codex\2026-07-09\que")
DATA_ROOT = ROOT / "work" / "base-diaria"
OUT = ROOT / "outputs" / "audiencia_df.html"
FONT_REG = Path(r"C:\Users\gamartin\Downloads\GlobotipoCorporativa-Regular.ttf")
FONT_BOLD = Path(r"C:\Users\gamartin\Downloads\GlobotipoCorporativa-Bold.ttf")
LOGO_ROOT = Path(r"C:\Users\gamartin\Documents\Codex\2026-05-18\preciso-criar-uma-aplica-o-em\assets")
PLIMPLIM = Path(r"C:\Users\gamartin\Downloads\Logos Jornalismo\PLIMPLIM.png")

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

CHANNELS = [
    {"key": "globo", "label": "Globo", "aliases": ["GLOBO", "Reference:"], "color": "#005cef", "logo": "GLOBO.png"},
    {"key": "record", "label": "Record", "aliases": ["Record TV", "RECORD"], "color": "#d71920", "logo": "RECORD.png"},
    {"key": "sbt", "label": "SBT", "aliases": ["SBT"], "color": "#00a651", "logo": "SBT.png"},
    {"key": "band", "label": "Band", "aliases": ["TV BAND", "BAND"], "color": "#ec4b91", "logo": "BAND.png"},
    {
        "key": "nic",
        "label": "NIC",
        "aliases": ["Conteúdo de TV/Vídeo sem referência", "Conteúdo de TV/Vídeo sem referęncia"],
        "color": "#7a4a27",
        "textLogo": "NIC",
    },
    {"key": "paytv", "label": "TV Paga", "aliases": ["Canais PayTV", "OCP"], "color": "#6cc8ff", "textLogo": "TVP"},
    {"key": "tle", "label": "TLE", "aliases": ["Total Ligados Especial"], "color": "#111111"},
]

LEADERSHIP_ORDER = ["globo", "nic", "record", "sbt", "band", "paytv"]
LEADERSHIP_CHANNELS = [next(c for c in CHANNELS if c["key"] == key) for key in LEADERSHIP_ORDER]
RANKING_CHANNELS = [c for c in CHANNELS if c["key"] in ("globo", "record", "sbt", "band")]
TARGETS = [
    "Total Domicílios",
    "Total Indivíduos",
    "Masculino",
    "Feminino",
    "AB1",
    "B2",
    "C1",
    "C2",
    "DE",
    "4-11 anos",
    "12-17 anos",
    "18-24 anos",
    "25-34 anos",
    "35-49 anos",
    "50+",
]


def col_to_idx(ref: str) -> int:
    n = 0
    for ch in "".join(c for c in ref if c.isalpha()):
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_xlsx(path: Path) -> dict[str, list[list[str]]]:
    with ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join(t.text or "" for t in si.findall(".//a:t", NS)))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        result = {}
        for sheet in wb.find("a:sheets", NS):
            sid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = relmap[sid]
            sheet_path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
            root = ET.fromstring(z.read(sheet_path))
            rows = []
            for row in root.findall(".//a:sheetData/a:row", NS):
                vals: list[str] = []
                for c in row.findall("a:c", NS):
                    idx = col_to_idx(c.attrib["r"])
                    while len(vals) <= idx:
                        vals.append("")
                    typ = c.attrib.get("t")
                    if typ == "inlineStr":
                        txt = "".join(t.text or "" for t in c.findall(".//a:t", NS))
                    else:
                        v = c.find("a:v", NS)
                        txt = "" if v is None else v.text or ""
                        if typ == "s" and txt:
                            txt = shared[int(txt)]
                    vals[idx] = txt
                rows.append(vals)
            result[sheet.attrib["name"]] = rows
        return result


def first_file(name: str) -> Path:
    return next(DATA_ROOT.rglob(name))


def fnum(value, default=None):
    if value in (None, "", "n/a"):
        return default
    try:
        return float(str(value).replace(",", "."))
    except ValueError:
        return default


def pct(value):
    return None if value is None else round(value, 2)


def channel_for(raw: str) -> dict | None:
    for ch in CHANNELS:
        if raw in ch["aliases"] or raw == ch["label"]:
            return ch
    return None


def value_by_alias(values: dict[str, float], ch: dict):
    for alias in ch["aliases"]:
        if alias in values:
            return values[alias]
    return None


def excel_date(serial: str) -> str:
    n = fnum(serial)
    if n is None:
        return ""
    return (datetime(1899, 12, 30) + timedelta(days=int(n))).strftime("%Y-%m-%d")


def frac_to_minutes(value: str) -> int | None:
    n = fnum(value)
    if n is None:
        return None
    return int(round(n * 24 * 60))


def hhmm_from_min(minutes: int) -> str:
    minutes = minutes % (24 * 60)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def parse_timeband(label: str) -> int | None:
    m = re.match(r"(\d{2}):(\d{2}):\d{2}", label or "")
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def norm_var(v: str) -> str:
    return "share" if "Shr%" in v else "aud" if "Rat%" in v else v


def read_matrix(rows, header_rows, name_col=0, allowed_vars=("aud", "share")):
    data = {}
    for row in rows[header_rows["start"] :]:
        if not row or not row[name_col] or row[name_col] == "Total":
            continue
        name = row[name_col]
        data[name] = defaultdict(lambda: defaultdict(dict))
        max_cols = max(len(r) for r in rows[: header_rows["entity"] + 1])
        for i in range(1, max_cols):
            target = rows[header_rows["target"]][i] if i < len(rows[header_rows["target"]]) else ""
            var = rows[header_rows["var"]][i] if i < len(rows[header_rows["var"]]) else ""
            entity = rows[header_rows["entity"]][i] if i < len(rows[header_rows["entity"]]) else ""
            val = row[i] if i < len(row) else ""
            ch = channel_for(entity)
            metric = norm_var(var)
            if allowed_vars is None:
                metric = var
            if target and ch and (allowed_vars is None or metric in allowed_vars):
                data[name][target][ch["key"]][metric] = fnum(val)
    return data


def load_minute_data(rows):
    headers = {"target": rows[1], "var": rows[2], "entity": rows[3]}
    minutes = []
    max_cols = max(len(headers[k]) for k in headers)
    for row in rows[5:]:
        label = row[0] if row else ""
        minute = parse_timeband(label)
        if minute is None:
            continue
        item = {"time": hhmm_from_min(minute), "minute": minute, "aud": {}, "share": {}}
        for i in range(1, max_cols):
            target = headers["target"][i] if i < len(headers["target"]) else ""
            var = norm_var(headers["var"][i] if i < len(headers["var"]) else "")
            entity = headers["entity"][i] if i < len(headers["entity"]) else ""
            if target != "Total Domicílios" or var not in ("aud", "share"):
                continue
            ch = channel_for(entity)
            if ch:
                item[var][ch["key"]] = fnum(row[i] if i < len(row) else "")
        minutes.append(item)
    return minutes


def avg(values):
    clean = [v for v in values if isinstance(v, (int, float)) and not math.isnan(v)]
    return sum(clean) / len(clean) if clean else None


def load_programs(program_rows, crosstab_rows):
    programs = []
    header_targets = program_rows[1]
    header_vars = program_rows[2]
    for row in program_rows[4:]:
        name = row[0] if row else ""
        if not name:
            continue
        p = {
            "name": name,
            "start": frac_to_minutes(row[1] if len(row) > 1 else ""),
            "end": frac_to_minutes(row[2] if len(row) > 2 else ""),
            "targets": {},
        }
        for i in range(3, len(header_targets)):
            target = header_targets[i]
            var = norm_var(header_vars[i] if i < len(header_vars) else "")
            if target in TARGETS and var in ("aud", "share"):
                p["targets"].setdefault(target, {})[var] = fnum(row[i] if i < len(row) else "")
        programs.append(p)
    competition = read_matrix(crosstab_rows, {"target": 2, "var": 3, "entity": 4, "start": 6})
    return programs, competition


def load_profile(rows):
    return read_matrix(rows, {"target": 3, "var": 2, "entity": 4, "start": 6}, allowed_vars=None)


def load_rankings(current_book, previous_book):
    sheet_map = {"globo": "Globo", "record": "Record", "sbt": "SBT", "band": "BAND"}
    rankings = {}
    for key, sheet in sheet_map.items():
        cur_rows = current_book.get(sheet, [])
        prev_rows = previous_book.get(sheet, [])
        prev_by_name = {}
        for row in prev_rows[3:]:
            if len(row) >= 4 and row[1]:
                prev_by_name[row[1]] = {"aud": fnum(row[2]), "share": fnum(row[3])}
        entries = []
        for idx, row in enumerate(cur_rows[3:13], 1):
            if len(row) < 4 or not row[1]:
                continue
            aud = fnum(row[2])
            share = fnum(row[3])
            prev = prev_by_name.get(row[1], {})

            def variation(now, before):
                if now is None or before in (None, 0):
                    return None
                return (now - before) / before * 100

            aud_var = variation(aud, prev.get("aud"))
            share_var = variation(share, prev.get("share"))
            entries.append(
                {
                    "rank": idx,
                    "program": row[1],
                    "aud": aud,
                    "audPrev": prev.get("aud"),
                    "audVar": aud_var,
                    "audStatus": "cresceu" if aud_var and aud_var > 0 else "caiu" if aud_var and aud_var < 0 else "estavel",
                    "share": share,
                    "sharePrev": prev.get("share"),
                    "shareVar": share_var,
                    "shareStatus": "cresceu"
                    if share_var and share_var > 0
                    else "caiu"
                    if share_var and share_var < 0
                    else "estavel",
                }
            )
        rankings[key] = entries
    return rankings


def build_data():
    minute_book = read_xlsx(first_file("Base_aud_minuto_DF.xlsx"))
    programs_book = read_xlsx(first_file("Base Programa Aud e shr_DF.xlsx"))
    profile_book = read_xlsx(first_file("Base Perfil_DF.xlsx"))
    turnos_book = read_xlsx(first_file("Base Turnos_DF.xlsx"))
    current_rank = read_xlsx(first_file("Base Ranking diário_REC.xlsx"))
    previous_rank = read_xlsx(first_file("Base Ranking diário_DF7d.xlsx"))

    minutes_all = load_minute_data(minute_book["Crosstab2"])
    day_minutes = [m for m in minutes_all if 7 * 60 <= m["minute"] < 24 * 60]
    programs, competition = load_programs(programs_book["Programas"], programs_book["Crosstab"])
    profile = load_profile(profile_book["Crosstab1"])

    turnos = turnos_book["Crosstab"]
    row_7_24 = next((r for r in turnos if r and "07:00-24:00" in r[0]), [])
    daily_avg = fnum(row_7_24[2] if len(row_7_24) > 2 else "")
    date_serial = ""
    for rows in programs_book.values():
        for row in rows:
            if row and re.fullmatch(r"\d{5}", row[0] or ""):
                date_serial = row[0]
                break
        if date_serial:
            break
    date_iso = excel_date(date_serial) or "2026-07-08"

    summary_bars = []
    for ch in CHANNELS:
        summary_bars.append(
            {
                "key": ch["key"],
                "label": ch["label"],
                "color": ch["color"],
                "aud": pct(avg([m["aud"].get(ch["key"]) for m in day_minutes])),
                "share": pct(avg([m["share"].get(ch["key"]) for m in day_minutes])),
            }
        )

    line = [
        {"time": m["time"], "aud": {ch["key"]: pct(m["aud"].get(ch["key"])) for ch in CHANNELS}}
        for m in day_minutes
    ]

    lead_counts = {ch["key"]: 0 for ch in LEADERSHIP_CHANNELS}
    for m in day_minutes:
        vals = [(ch["key"], m["aud"].get(ch["key"])) for ch in LEADERSHIP_CHANNELS]
        vals = [(k, v) for k, v in vals if v is not None]
        if vals:
            leader = max(vals, key=lambda x: x[1])[0]
            lead_counts[leader] += 1
    total_minutes = len(day_minutes)
    leadership = []
    for ch in LEADERSHIP_CHANNELS:
        mins = lead_counts[ch["key"]]
        leadership.append(
            {
                "key": ch["key"],
                "label": ch["label"],
                "color": ch["color"],
                "minutes": mins,
                "percent": pct(mins / total_minutes * 100 if total_minutes else 0),
                "hours": f"{mins // 60}h{mins % 60:02d}",
            }
        )

    def adh(name, target, key):
        return (profile.get(name, {}).get(target, {}).get(key, {}) or {}).get("Adh% [Total Indivíduos | ORG]")

    profile_data = {}
    for p in programs:
        name = p["name"]
        pcomp = competition.get(name, {})
        profile_data[name] = {}
        for ch in LEADERSHIP_CHANNELS:
            metrics = pcomp.get("Total Domicílios", {}).get(ch["key"], {})
            profile_data[name][ch["key"]] = {
                "aud": metrics.get("aud"),
                "share": metrics.get("share"),
                "gender": {
                    "Homem": adh(name, "Masculino", ch["key"]),
                    "Mulher": adh(name, "Feminino", ch["key"]),
                },
                "classes": {
                    "AB1": adh(name, "AB1", ch["key"]),
                    "B2": adh(name, "B2", ch["key"]),
                    "C1": adh(name, "C1", ch["key"]),
                    "C2": adh(name, "C2", ch["key"]),
                    "DE": adh(name, "DE", ch["key"]),
                },
                "ages": {
                    "4-11": adh(name, "4-11 anos", ch["key"]),
                    "12-17": adh(name, "12-17 anos", ch["key"]),
                    "18-24": adh(name, "18-24 anos", ch["key"]),
                    "25-34": adh(name, "25-34 anos", ch["key"]),
                    "35-49": adh(name, "35-49 anos", ch["key"]),
                    "50+": adh(name, "50+", ch["key"]),
                },
            }

    return {
        "meta": {
            "date": date_iso,
            "weekday": ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"][
                datetime.fromisoformat(date_iso).weekday()
            ],
            "dailyAvg": daily_avg,
        },
        "channels": CHANNELS,
        "leadershipChannels": LEADERSHIP_CHANNELS,
        "rankingChannels": RANKING_CHANNELS,
        "targets": TARGETS,
        "summaryBars": summary_bars,
        "line": line,
        "leadership": leadership,
        "programs": programs,
        "programCompetition": competition,
        "minuteAll": minutes_all,
        "profile": profile_data,
        "rankings": load_rankings(current_rank, previous_rank),
    }


def font_data(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def image_data(path: Path) -> str:
    suffix = path.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in ("jpg", "jpeg") else suffix
    return f"data:image/{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def render_html(data: dict) -> str:
    for ch in data["channels"]:
        if ch.get("logo"):
            ch["logoData"] = image_data(LOGO_ROOT / ch["logo"])
    data["assets"] = {"plim": image_data(PLIMPLIM)}
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    font_reg = font_data(FONT_REG)
    font_bold = font_data(FONT_BOLD)
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Audiência Diária DF</title>
<style>
@font-face{{font-family:Globotipo;src:url(data:font/ttf;base64,{font_reg}) format("truetype");font-weight:400}}
@font-face{{font-family:Globotipo;src:url(data:font/ttf;base64,{font_bold}) format("truetype");font-weight:700}}
:root{{--blue:#005cef;--blue2:#2b7cff;--ink:#101521;--muted:#607086;--line:#d9e0ec;--soft:#f2f4f7;--bg:#eef2f7;--card:#fff;--green:#07934a;--red:#d62839;--male:#2f80ed;--female:#e94d9a;--class1:#005cef;--class2:#00a651;--class3:#ffb000;--class4:#ec4b91;--class5:#7a4a27}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Globotipo,Arial,sans-serif;background:linear-gradient(180deg,#f8fafc 0%,var(--bg) 44%,#f7f8fa 100%);color:var(--ink);letter-spacing:0}}body.awaiting-bases main,body.awaiting-bases .tabs{{display:none}}body.awaiting-bases header.app{{min-height:100vh;display:grid;place-items:center;padding:24px}}body.awaiting-bases .top{{display:flex;flex-direction:column;justify-content:center;align-items:center;max-width:760px;width:min(92vw,760px);min-height:520px;text-align:center;padding:56px 42px}}body.awaiting-bases .brand{{flex-direction:column;gap:20px;font-size:58px}}body.awaiting-bases .brand-logo{{width:130px;height:130px}}body.awaiting-bases .header-right{{justify-items:center;width:100%}}body.awaiting-bases .actions{{min-width:0;width:auto;justify-content:center;margin-top:6px}}body.awaiting-bases .file-status{{display:none}}
header.app{{background:linear-gradient(90deg,rgba(0,92,239,.04) 1px,transparent 1px),linear-gradient(180deg,rgba(0,92,239,.04) 1px,transparent 1px),#eef2f7;background-size:72px 72px;position:relative;z-index:20;padding:18px 20px 0}}
.top{{max-width:1440px;margin:auto;background:#fff;border:1px solid var(--line);border-top:8px solid transparent;border-image:linear-gradient(90deg,#0046ff,#00a3ff) 1;border-radius:8px;padding:28px 38px;display:grid;grid-template-columns:1fr auto;gap:22px;align-items:center;box-shadow:0 10px 26px rgba(21,36,66,.08)}}
.brand{{display:flex;align-items:center;gap:22px;font-weight:700;font-size:54px;line-height:1;color:var(--blue);letter-spacing:0;text-transform:uppercase}}.brand-logo{{width:78px;height:78px;object-fit:contain;filter:none}}.header-right{{display:grid;gap:12px;justify-items:end}}.actions{{display:flex;gap:10px;align-items:center;flex-wrap:wrap;justify-content:flex-end;background:linear-gradient(135deg,#004cff,#05a9f5);border-radius:8px;padding:18px 20px;min-width:360px}}.action-btn{{border:1px solid rgba(255,255,255,.45);background:#fff;color:#005cef;border-radius:8px;padding:10px 12px;font:inherit;font-weight:700;cursor:pointer}}.action-btn.primary{{background:rgba(255,255,255,.16);color:#fff;border-color:rgba(255,255,255,.45)}}.file-status{{font-size:12px;color:#fff;max-width:330px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;grid-column:1/-1;justify-self:start}}
.tabs{{display:flex;gap:6px;background:#e8edf5;padding:4px;border-radius:8px}}.tab{{border:0;background:transparent;padding:10px 16px;border-radius:6px;font-weight:700;cursor:pointer;color:#40506a}}.tab.active{{background:linear-gradient(135deg,var(--blue),#3e8bff);color:#fff}}
main{{max-width:1440px;margin:0 auto;padding:24px;display:grid;gap:18px}}.panel{{display:none}}.panel.active{{display:grid;gap:18px}}
.hero{{background:linear-gradient(135deg,#0046c8 0%,#005cef 48%,#45a3ff 100%);color:#fff;padding:26px;border-radius:8px;display:grid;grid-template-columns:1fr auto;gap:20px;align-items:center;box-shadow:0 18px 45px rgba(0,92,239,.22)}}.hero-head{{display:flex;align-items:center;gap:18px;flex-wrap:wrap}}.hero h1{{margin:0;font-size:34px;line-height:1.05}}.date-pill{{display:inline-grid;gap:2px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.35);border-radius:8px;padding:12px 16px;min-width:190px}}.date-pill strong{{font-size:24px}}.date-pill span{{font-size:13px;text-transform:uppercase;letter-spacing:.08em}}.kpi{{background:#fff;color:var(--ink);border-radius:8px;padding:16px 20px;min-width:220px;box-shadow:0 12px 28px rgba(0,0,0,.12)}}.kpi small{{display:block;color:var(--muted);font-size:12px;text-transform:uppercase}}.kpi strong{{font-size:42px;line-height:1}}
.toolbar{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px;display:grid;grid-template-columns:260px 260px 1fr;gap:16px;align-items:center}}.toolbar.one{{grid-template-columns:280px 1fr}}label{{font-size:12px;color:var(--muted);font-weight:700;text-transform:uppercase;display:grid;gap:6px}}select{{font:inherit;border:1px solid var(--line);border-radius:6px;padding:10px 12px;background:#fff;min-width:0}}.selected-title{{text-align:center;font-weight:700;font-size:26px;color:var(--blue);background:var(--soft);border-radius:8px;padding:14px 18px}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}.card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:18px;min-width:0;box-shadow:0 10px 28px rgba(21,36,66,.06)}}.card h2{{margin:0 0 14px;font-size:18px}}.card h3{{margin:0 0 12px;font-size:16px}}.chart{{width:100%;height:290px;display:block}}.line-chart{{height:380px}}.legend{{display:flex;gap:12px;flex-wrap:wrap;margin-top:10px;font-size:12px;color:#3e4d63}}.legend span,.channel-name{{display:inline-flex;align-items:center;gap:6px}}.dot{{width:10px;height:10px;border-radius:50%;display:inline-block}}.logo-chip{{width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:9px;line-height:1;overflow:hidden;box-shadow:inset 0 0 0 1px rgba(255,255,255,.28);text-align:center}}.logo-chip img{{display:block;width:72%;height:72%;object-fit:contain;object-position:center;margin:auto}}.leader .logo-chip{{color:#fff}}
.leader-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}}.leader{{border:1px solid var(--line);border-radius:8px;padding:14px;background:linear-gradient(180deg,#fff,#f8fafc)}}.leader strong{{font-size:28px;display:block}}.leader span{{color:var(--muted);font-size:12px}}.leader-head{{display:flex;align-items:center;gap:8px;margin-bottom:8px}}
.rank-grid{{display:grid;grid-template-columns:repeat(2, minmax(0,1fr));gap:18px}}table{{width:100%;border-collapse:collapse;font-size:13px}}th,td{{padding:10px 8px;border-bottom:1px solid #edf1f7;text-align:left;vertical-align:middle}}th{{color:var(--muted);font-size:11px;text-transform:uppercase;background:#fafbfe}}td.num,th.num{{text-align:right}}.arrow{{font-size:18px;font-weight:700}}.up{{color:var(--green)}}.down{{color:var(--red)}}.empty{{color:var(--muted);padding:24px;text-align:center;background:#f8fafc;border-radius:8px}}
.profile-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}}.profile-card{{padding:20px}}.profile-head{{display:flex;justify-content:space-between;gap:16px;align-items:center;margin-bottom:16px}}.profile-title{{display:flex;align-items:center;gap:10px}}.metric-pair{{display:flex;gap:8px}}.mini{{border:1px solid var(--line);border-radius:8px;padding:8px 10px;min-width:78px;text-align:right;background:#f9fbfe}}.mini small{{display:block;color:var(--muted);font-size:11px}}.mini strong{{font-size:22px}}.mini-chart{{height:140px;width:100%;display:block}}.profile-block{{display:grid;grid-template-columns:1fr 1fr;gap:18px 14px}}.profile-block h4{{margin:0 0 8px;font-size:12px;color:var(--muted);text-transform:uppercase}}.program-table-wrap{{max-height:470px;overflow:auto;border:1px solid var(--line);border-radius:8px}}.program-table-wrap table th{{position:sticky;top:0;z-index:1}}.minute-table th,.minute-table td{{white-space:nowrap}}.footer-note{{text-align:center;color:#66758c;font-size:13px;padding:18px 0 4px}}.final-actions{{display:flex;justify-content:center;gap:12px;flex-wrap:wrap;padding:8px 0 20px}}.final-actions .action-btn{{border-color:#c9d4e5;box-shadow:0 8px 20px rgba(21,36,66,.08)}}.final-actions .action-btn.primary{{background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;border-color:transparent}}
@media(max-width:980px){{.top,.hero,.toolbar,.toolbar.one,.grid2,.rank-grid,.profile-grid{{grid-template-columns:1fr}}.brand{{font-size:34px}}.brand-logo{{width:54px;height:54px}}body.awaiting-bases .brand{{font-size:40px}}body.awaiting-bases .brand-logo{{width:108px;height:108px}}.header-right{{justify-items:stretch}}body.awaiting-bases .header-right{{justify-items:center}}.actions{{min-width:0}}.leader-grid{{grid-template-columns:repeat(2,1fr)}}.selected-title{{text-align:left}}}}
</style>
</head>
<body class="awaiting-bases">
<header class="app"><div class="top"><div class="brand"><img class="brand-logo" id="plimLogo" alt="Plim plim"><span>Audiência Diária</span></div><div class="header-right"><div class="actions"><input id="baseUpload" type="file" accept=".zip,.xlsx" multiple hidden><button class="action-btn primary" id="uploadBtn">Enviar base</button><span class="file-status" id="fileStatus"></span></div><nav class="tabs"><button class="tab active" data-tab="resumo">RESUMO</button><button class="tab" data-tab="programas">PROGRAMAS</button><button class="tab" data-tab="perfil">PERFIL</button></nav></div></div></header>
<main>
<section id="resumo" class="panel active">
  <div class="hero"><div class="hero-head"><span id="dateLabel"></span><h1>Resumo do Distrito Federal</h1></div><div class="kpi"><small>Audiência média 07h-24h</small><strong id="dailyAvg"></strong></div></div>
  <div class="grid2"><div class="card"><h2>Audiência 07:00 às 24:00</h2><svg id="audBars" class="chart"></svg><div id="legend1" class="legend"></div></div><div class="card"><h2>Share% 07:00 às 24:00</h2><svg id="shareBars" class="chart"></svg><div id="legend2" class="legend"></div></div></div>
  <div class="card"><h2>Minuto a minuto do dia</h2><svg id="dayLine" class="chart line-chart"></svg><div id="legend3" class="legend"></div></div>
  <div class="card"><h2>Minutos na liderança</h2><div id="leadership" class="leader-grid"></div></div>
  <div class="card"><h2>Ranking de audiência por emissora</h2><div id="rankings" class="rank-grid"></div></div>
</section>
<section id="programas" class="panel">
  <div class="toolbar"><label>Programa<select id="programSelect"></select></label><label>Target<select id="targetSelect"></select></label><div class="selected-title" id="programTitle"></div></div>
  <div class="grid2"><div class="card"><h2>Audiência por emissora</h2><svg id="progAud" class="chart"></svg></div><div class="card"><h2>Share% por emissora</h2><svg id="progShare" class="chart"></svg></div></div>
  <div class="card"><h2>Minuto a minuto do programa</h2><svg id="progLine" class="chart line-chart"></svg><div id="legend4" class="legend"></div></div>
  <div class="card"><h2>Tabela minuto a minuto</h2><div id="minuteTable" class="program-table-wrap"></div></div>
</section>
<section id="perfil" class="panel">
  <div class="toolbar one"><label>Programa<select id="profileSelect"></select></label><div class="selected-title" id="profileTitle"></div></div>
  <div id="profileGrid" class="profile-grid"></div>
</section>
<div class="footer-note">Modelo desenvolvido pela área de Programação da TV Globo DF.</div>
<div class="final-actions"><button class="action-btn primary" id="exportBtn">Compartilhar</button><button class="action-btn" id="imageBtn">Gerar imagem</button></div>
</main>
<script id="data" type="application/json">{data_json.replace("</", "<\\/")}</script>
<script>
const DATA = JSON.parse(document.getElementById('data').textContent);
const fmt = (v,d=2) => v === null || v === undefined || Number.isNaN(v) ? '-' : Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:d, maximumFractionDigits:d}});
const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
document.getElementById('plimLogo').src = DATA.assets.plim;
const channelMap = Object.fromEntries(DATA.channels.map(c=>[c.key,c]));
const visible = DATA.channels.filter(c=>c.key!=='tle');
const shareChannels = visible;
const lineChannels = visible;
const leadChannels = DATA.leadershipChannels;
const rankingChannels = DATA.rankingChannels;
function logoHtml(c){{return c.logoData ? `<span class="logo-chip" style="background:${{c.color}}"><img src="${{c.logoData}}" alt="${{c.label}}"></span>` : `<span class="logo-chip" style="background:${{c.color}}">${{c.textLogo || c.label.slice(0,3).toUpperCase()}}</span>`}}
function legend(el, channels=visible){{document.getElementById(el).innerHTML=channels.map(c=>`<span>${{logoHtml(c)}}${{c.label}}</span>`).join('')}}
function svgEl(tag, attrs={{}}){{const e=document.createElementNS('http://www.w3.org/2000/svg',tag);Object.entries(attrs).forEach(([k,v])=>e.setAttribute(k,v));return e}}
function clear(svg){{while(svg.firstChild)svg.removeChild(svg.firstChild)}}
function svgLogo(svg,c,cx,cy,size=24){{svg.appendChild(svgEl('circle',{{cx,cy,r:size/2,fill:c.color}}));if(c.logoData){{const inner=size*.72;const im=svgEl('image',{{href:c.logoData,x:cx-inner/2,y:cy-inner/2,width:inner,height:inner,preserveAspectRatio:'xMidYMid meet'}});svg.appendChild(im)}}else{{svg.appendChild(svgEl('text',{{x:cx,y:cy+3,'text-anchor':'middle','font-size':8,'font-weight':700,fill:'#fff'}})).textContent=c.textLogo||c.label.slice(0,3).toUpperCase()}}}}
function barChart(id, data, metric, channels=visible){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||700,h=svg.clientHeight||290,ml=44,mr=16,mt=20,mb=58;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const vals=channels.map(c=>data.find(d=>d.key===c.key)?.[metric]??0);const max=Math.max(...vals,1)*1.15;svg.appendChild(svgEl('line',{{x1:ml,y1:h-mb,x2:w-mr,y2:h-mb,stroke:'#d9e0ec'}}));const slot=(w-ml-mr)/channels.length,bw=slot*.58;channels.forEach((c,i)=>{{const v=vals[i];const x=ml+i*slot+(slot-bw)/2;const y=mt+(h-mt-mb)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:4,fill:c.color}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:y-7,'text-anchor':'middle','font-size':12,'font-weight':700,fill:'#111'}})).textContent=fmt(v,2);svgLogo(svg,c,x+bw/2,h-30,24);svg.appendChild(svgEl('text',{{x:x+bw/2,y:h-6,'text-anchor':'middle','font-size':10,fill:'#526178'}})).textContent=c.textLogo||c.label;}})}}
function lineChart(id, rows, channels=lineChannels, accessor=r=>r.aud, highlightKey='globo'){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||1000,h=svg.clientHeight||380,ml=48,mr=18,mt=20,mb=42;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const defs=svgEl('defs');const grad=svgEl('linearGradient',{{id:`${{id}}-badgeGrad`,x1:'0%',y1:'0%',x2:'100%',y2:'0%'}});grad.appendChild(svgEl('stop',{{offset:'0%','stop-color':'#004cff'}}));grad.appendChild(svgEl('stop',{{offset:'100%','stop-color':'#05a9f5'}}));defs.appendChild(grad);svg.appendChild(defs);let max=1,highlight=[];rows.forEach((r,i)=>channels.forEach(c=>{{const v=accessor(r)[c.key];if(v!==null&&v!==undefined){{if(v>max)max=v;if(c.key===highlightKey)highlight.push({{v,i,r,c}})}}}}));max*=1.12;for(let g=0;g<=4;g++){{const y=mt+(h-mt-mb)*g/4;svg.appendChild(svgEl('line',{{x1:ml,y1:y,x2:w-mr,y2:y,stroke:'#edf1f7'}}));svg.appendChild(svgEl('text',{{x:8,y:y+4,'font-size':11,fill:'#647084'}})).textContent=fmt(max*(1-g/4),1)}}channels.forEach(c=>{{let d='';rows.forEach((r,i)=>{{const v=accessor(r)[c.key];if(v===null||v===undefined)return;const x=ml+(w-ml-mr)*(i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-v/max);d+=(d?'L':'M')+x.toFixed(1)+' '+y.toFixed(1)}});svg.appendChild(svgEl('path',{{d,fill:'none',stroke:c.color,'stroke-width':c.key===highlightKey?3:2.1,'stroke-linejoin':'round','stroke-linecap':'round'}}));}});if(highlight.length){{const hi=highlight.reduce((a,b)=>b.v>a.v?b:a,highlight[0]);const lo=highlight.reduce((a,b)=>b.v<a.v?b:a,highlight[0]);[{{p:hi,label:'Maior'}},{{p:lo,label:'Menor'}}].forEach(({{p,label}})=>{{const x=ml+(w-ml-mr)*(p.i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-p.v/max);svg.appendChild(svgEl('circle',{{cx:x,cy:y,r:5.5,fill:'#005cef',stroke:'#fff','stroke-width':2}}));const tx=Math.min(w-102,Math.max(54,x+8));const ty=Math.max(24,y-12);svg.appendChild(svgEl('rect',{{x:tx-6,y:ty-17,width:94,height:24,rx:6,fill:`url(#${{id}}-badgeGrad)`}}));svg.appendChild(svgEl('text',{{x:tx,y:ty,'font-size':11,'font-weight':700,fill:'#fff'}})).textContent=`${{label}} ${{fmt(p.v,1)}}`;}})}}const tickCount=Math.min(15,Math.max(6,Math.ceil(rows.length/70)+8));for(let t=0;t<tickCount;t++){{const p=t/(tickCount-1);const i=Math.min(rows.length-1,Math.round((rows.length-1)*p));const x=ml+(w-ml-mr)*p;svg.appendChild(svgEl('text',{{x,y:h-14,'text-anchor':'middle','font-size':11,fill:'#647084'}})).textContent=rows[i]?.time||''}}}}
function miniBars(svg, obj, color){{clear(svg);const entries=Object.entries(obj||{{}});const w=svg.clientWidth||320,h=svg.clientHeight||140,ml=34,mb=30,mt=10;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const max=Math.max(...entries.map(e=>e[1]||0),1);const bw=(w-ml-8)/Math.max(entries.length,1)*.58;entries.forEach(([k,v],i)=>{{v=v||0;const fill=Array.isArray(color)?color[i%color.length]:color;const x=ml+(i+.21)*(w-ml-8)/entries.length;const y=mt+16+(h-mt-mb-16)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:3,fill}}));const val=fmt(v,0)+'%';const valueW=Math.max(34,val.length*7+14);const valueY=Math.max(8,y-24);svg.appendChild(svgEl('rect',{{x:x+bw/2-valueW/2,y:valueY,width:valueW,height:18,rx:5,fill:'#fff',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:valueY+13,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#27364e'}})).textContent=val;const labelW=Math.max(34,k.length*6+14);svg.appendChild(svgEl('rect',{{x:x+bw/2-labelW/2,y:h-24,width:labelW,height:18,rx:5,fill:'#f2f4f7',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:h-11,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#526178'}})).textContent=k;}})}}
function pie(svg, obj){{clear(svg);const entries=Object.entries(obj||{{}}).filter(e=>(e[1]||0)>0);const w=svg.clientWidth||280,h=svg.clientHeight||130,cx=62,cy=64,r=45;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const colors=['#005cef','#d71920','#00a651','#ec4b91','#7a4a27'];const total=entries.reduce((a,e)=>a+e[1],0)||1;let a0=-Math.PI/2;entries.forEach(([k,v],i)=>{{const a1=a0+(v/total)*Math.PI*2;const large=a1-a0>Math.PI?1:0;const x0=cx+r*Math.cos(a0), y0=cy+r*Math.sin(a0), x1=cx+r*Math.cos(a1), y1=cy+r*Math.sin(a1);svg.appendChild(svgEl('path',{{d:`M${{cx}} ${{cy}} L${{x0}} ${{y0}} A${{r}} ${{r}} 0 ${{large}} 1 ${{x1}} ${{y1}} Z`,fill:colors[i%colors.length]}}));a0=a1;}});entries.forEach(([k,v],i)=>{{svg.appendChild(svgEl('rect',{{x:125,y:18+i*19,width:10,height:10,fill:colors[i%colors.length]}}));svg.appendChild(svgEl('text',{{x:141,y:27+i*19,'font-size':11,fill:'#3e4d63'}})).textContent=`${{k}} ${{fmt(v,1)}}%`;}})}}
function statusArrow(status){{return status==='cresceu'?'<span class="arrow up">▲</span>':status==='caiu'?'<span class="arrow down">▼</span>':'<span class="arrow">•</span>'}}
function renderRankings(){{const box=document.getElementById('rankings');box.innerHTML=rankingChannels.map(c=>{{const rows=DATA.rankings[c.key]||[];return `<div class="card"><h3 class="channel-name">${{logoHtml(c)}}${{c.label}}</h3><table><thead><tr><th>#</th><th>Programa</th><th class="num">Aud.</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th><th class="num">Share</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th></tr></thead><tbody>${{rows.map(r=>`<tr><td>${{r.rank}}</td><td>${{esc(r.program)}}</td><td class="num">${{fmt(r.aud)}}</td><td class="num">${{fmt(r.audPrev)}}</td><td class="num">${{r.audVar==null?'-':fmt(r.audVar,1)+'%'}}</td><td>${{statusArrow(r.audStatus)}}</td><td class="num">${{fmt(r.share)}}</td><td class="num">${{fmt(r.sharePrev)}}</td><td class="num">${{r.shareVar==null?'-':fmt(r.shareVar,1)+'%'}}</td><td>${{statusArrow(r.shareStatus)}}</td></tr>`).join('')}}</tbody></table></div>`}}).join('')}}
function renderResumo(){{const d=new Date(DATA.meta.date+'T00:00:00');document.getElementById('dateLabel').className='date-pill';document.getElementById('dateLabel').innerHTML=`<span>${{DATA.meta.weekday}}</span><strong>${{d.toLocaleDateString('pt-BR')}}</strong>`;document.getElementById('dailyAvg').textContent=fmt(DATA.meta.dailyAvg,2);barChart('audBars',DATA.summaryBars,'aud',visible);barChart('shareBars',DATA.summaryBars,'share',shareChannels);lineChart('dayLine',DATA.line,lineChannels);legend('legend1',visible);legend('legend2',shareChannels);legend('legend3',lineChannels);document.getElementById('leadership').innerHTML=DATA.leadership.map(l=>{{const c=channelMap[l.key];return `<div class="leader" style="border-top:5px solid ${{l.color}}"><div class="leader-head">${{logoHtml(c)}}<span>${{l.label}}</span></div><strong>${{l.minutes}}</strong><span>${{fmt(l.percent,1)}}% | ${{l.hours}}</span></div>`}}).join('');renderRankings()}}
function fillSelect(sel, values){{sel.innerHTML=values.map(v=>`<option value="${{esc(v)}}">${{esc(v)}}</option>`).join('')}}
function programMinuteRows(program){{const start=program.start??0,end=program.end??0;return DATA.minuteAll.filter(m=>m.minute>=start && m.minute<end)}}
function renderProgramas(){{const p=DATA.programs.find(x=>x.name===programSelect.value)||DATA.programs[0];const target=targetSelect.value;programTitle.textContent=`${{p.name}} | ${{target}}`;const comp=(DATA.programCompetition[p.name]||{{}})[target]||{{}};const audBars=visible.map(c=>({{key:c.key,label:c.label,color:c.color,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));const shareBars=shareChannels.map(c=>({{key:c.key,label:c.label,color:c.color,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));barChart('progAud',audBars,'aud',visible);barChart('progShare',shareBars,'share',shareChannels);const rows=programMinuteRows(p);lineChart('progLine',rows,lineChannels,r=>r.aud);legend('legend4',lineChannels);minuteTable.innerHTML=`<table class="minute-table"><thead><tr><th>Minuto</th>${{visible.map(c=>`<th colspan="2" style="border-top:4px solid ${{c.color}}">${{logoHtml(c)}} ${{c.label}}</th>`).join('')}}</tr><tr><th></th>${{visible.map(c=>`<th class="num">Aud.</th><th class="num">Share</th>`).join('')}}</tr></thead><tbody>${{rows.map(r=>`<tr><td><strong>${{r.time}}</strong></td>${{visible.map(c=>`<td class="num" style="color:${{c.color}}">${{fmt(r.aud[c.key])}}</td><td class="num">${{fmt(r.share[c.key])}}</td>`).join('')}}</tr>`).join('')}}</tbody></table>`}}
function renderPerfil(){{const name=profileSelect.value;profileTitle.textContent=name;const pdata=DATA.profile[name]||{{}};const genderColors=['#2f80ed','#e94d9a'];const ageColors=['#005cef','#2578ff','#00a651','#ffb000','#ec4b91','#7a4a27'];profileGrid.innerHTML=leadChannels.map(c=>{{const d=pdata[c.key]||{{}};return `<div class="card profile-card" data-ch="${{c.key}}" style="border-top:5px solid ${{c.color}}"><div class="profile-head"><h3 class="profile-title">${{logoHtml(c)}}${{c.label}}</h3><div class="metric-pair"><div class="mini"><small>Aud.</small><strong>${{fmt(d.aud)}}</strong></div><div class="mini"><small>Share</small><strong>${{fmt(d.share)}}</strong></div></div></div><div class="profile-block"><div><h4>Homem e mulher</h4><svg class="mini-chart gender"></svg></div><div><h4>Classes sociais</h4><svg class="mini-chart classes"></svg></div><div style="grid-column:1/-1"><h4>Faixas etárias</h4><svg class="mini-chart ages"></svg></div></div></div>`}}).join('');document.querySelectorAll('.profile-card').forEach(card=>{{const key=card.dataset.ch, d=pdata[key]||{{}};miniBars(card.querySelector('.gender'),d.gender,genderColors);pie(card.querySelector('.classes'),d.classes);miniBars(card.querySelector('.ages'),d.ages,ageColors)}})}}
function downloadBlob(name,type,content){{const blob=new Blob([content],{{type}});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),500)}}
function exportHtml(){{const clone=document.documentElement.cloneNode(true);clone.querySelectorAll('#baseUpload').forEach(el=>el.value='');downloadBlob('audiencia-diaria-df.html','text/html;charset=utf-8','<!doctype html>\\n'+clone.outerHTML)}}
async function generateHighlightsImage(){{const d=new Date(DATA.meta.date+'T00:00:00');const globo=DATA.summaryBars.find(x=>x.key==='globo')||{{}};const topGlobo=(DATA.rankings.globo||[]).slice(0,3);const competitors=['record','sbt','band'].flatMap(k=>(DATA.rankings[k]||[]).map(r=>({{...r,channel:k}}))).sort((a,b)=>(b.aud||0)-(a.aud||0)).slice(0,3);const canvas=document.createElement('canvas');canvas.width=1400;canvas.height=900;const ctx=canvas.getContext('2d');const grad=ctx.createLinearGradient(0,0,1400,900);grad.addColorStop(0,'#003bb2');grad.addColorStop(.48,'#005cef');grad.addColorStop(1,'#65c7ff');ctx.fillStyle=grad;ctx.fillRect(0,0,1400,900);ctx.fillStyle='rgba(255,255,255,.08)';ctx.beginPath();ctx.arc(1120,90,420,0,Math.PI*2);ctx.fill();ctx.fillStyle='rgba(255,255,255,.06)';ctx.beginPath();ctx.arc(120,830,420,0,Math.PI*2);ctx.fill();const globoLogo=await new Promise(resolve=>{{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=channelMap.globo.logoData;}});if(globoLogo)ctx.drawImage(globoLogo,78,64,82,82);ctx.fillStyle='#fff';ctx.font='700 52px Globotipo, Arial';ctx.fillText('Audiência Diária DF',188,100);ctx.font='400 27px Globotipo, Arial';ctx.fillText(`${{DATA.meta.weekday}} • ${{d.toLocaleDateString('pt-BR')}}`,190,140);ctx.font='700 31px Globotipo, Arial';ctx.fillText('Principais destaques do dia',84,230);function fitText(text,maxWidth,fontSize,weight='700'){{let size=fontSize;do{{ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;if(ctx.measureText(text).width<=maxWidth)return {{text,size}};size-=1}}while(size>=14);let clipped=text;while(clipped.length>3&&ctx.measureText(clipped+'...').width>maxWidth)clipped=clipped.slice(0,-1);return {{text:clipped+'...',size}}}}function card(x,y,w,h,title,value,accent='#005cef'){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,16);ctx.fill();ctx.fillStyle=accent;ctx.fillRect(x,y,8,h);ctx.fillStyle='#607086';ctx.font='700 23px Globotipo, Arial';ctx.fillText(title,x+32,y+44);const fitted=fitText(String(value),w-64,46);ctx.fillStyle='#101521';ctx.font=`700 ${{fitted.size}}px Globotipo, Arial`;ctx.fillText(fitted.text,x+32,y+100)}}function rankCard(x,y,w,h,title,items,accent){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,16);ctx.fill();ctx.fillStyle=accent;ctx.fillRect(x,y,8,h);ctx.fillStyle='#607086';ctx.font='700 24px Globotipo, Arial';ctx.fillText(title,x+32,y+48);items.forEach((item,i)=>{{const label=item.channel?`${{channelMap[item.channel].label}} • ${{item.program}}`:item.program;const lineY=y+104+i*58;const fitted=fitText(`${{i+1}}. ${{label}}`,w-135,22,'700');ctx.fillStyle='#101521';ctx.font=`700 ${{fitted.size}}px Globotipo, Arial`;ctx.fillText(fitted.text,x+32,lineY);ctx.fillStyle=accent;ctx.font='700 24px Globotipo, Arial';ctx.fillText(fmt(item.aud,2),x+w-86,lineY)}})}}card(84,276,586,140,'Audiência 07h - 24h',fmt(globo.aud,2),'#005cef');card(730,276,586,140,'Share 07h - 24h',fmt(globo.share,2)+'%','#005cef');rankCard(84,472,586,246,'Top 3 Globo',topGlobo,'#005cef');rankCard(730,472,586,246,'Top 3 Concorrência',competitors,'#d71920');ctx.fillStyle='rgba(255,255,255,.88)';ctx.font='400 22px Globotipo, Arial';ctx.fillText('Modelo desenvolvido pela área de Programação da TV Globo DF.',84,820);const a=document.createElement('a');a.href=canvas.toDataURL('image/png');a.download='destaques-audiencia-df.png';a.click()}}
document.querySelectorAll('.tab').forEach(btn=>btn.addEventListener('click',()=>{{document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));btn.classList.add('active');document.getElementById(btn.dataset.tab).classList.add('active');setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil()}},0)}}));
const programSelect=document.getElementById('programSelect'),targetSelect=document.getElementById('targetSelect'),profileSelect=document.getElementById('profileSelect'),programTitle=document.getElementById('programTitle'),profileTitle=document.getElementById('profileTitle'),minuteTable=document.getElementById('minuteTable'),profileGrid=document.getElementById('profileGrid'),baseUpload=document.getElementById('baseUpload'),fileStatus=document.getElementById('fileStatus');
fillSelect(programSelect,DATA.programs.map(p=>p.name));fillSelect(profileSelect,DATA.programs.map(p=>p.name));fillSelect(targetSelect,DATA.targets);
programSelect.addEventListener('change',renderProgramas);targetSelect.addEventListener('change',renderProgramas);profileSelect.addEventListener('change',renderPerfil);window.addEventListener('resize',()=>{{renderResumo();renderProgramas();renderPerfil()}});
function revealDashboard(){{document.body.className=document.body.className.replace(/\\bawaiting-bases\\b/g,'').trim();document.getElementById('uploadBtn').style.display='none';fileStatus.style.display='none';setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil();window.scrollTo({{top:0,behavior:'smooth'}})}},0)}}
document.getElementById('uploadBtn').addEventListener('click',()=>baseUpload.click());baseUpload.addEventListener('change',()=>{{const files=[...baseUpload.files].map(f=>f.name);fileStatus.textContent=files.length?`${{files.length}} arquivo(s): ${{files.join(', ')}}`:'';if(files.length)revealDashboard()}});document.getElementById('exportBtn').addEventListener('click',exportHtml);document.getElementById('imageBtn').addEventListener('click',generateHighlightsImage);
renderResumo();renderProgramas();renderPerfil();
</script>
</body>
</html>"""


if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    data = build_data()
    OUT.write_text(render_html(data), encoding="utf-8")
    print(OUT)
