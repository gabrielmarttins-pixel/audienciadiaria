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
OUT = ROOT / "outputs" / "index.html"
FONT_REG = ROOT / "assets" / "GlobotipoCorporativa-Regular.ttf"
FONT_BOLD = ROOT / "assets" / "GlobotipoCorporativa-Bold.ttf"
LOGO_ROOT = ROOT / "assets"
PLIMPLIM = ROOT / "assets" / "PLIMPLIM.png"
INTELIGENCIA = ROOT / "assets" / "inteligencia.png"
JSZIP = Path(r"C:\Users\gamartin\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\.pnpm\docx@9.6.1\node_modules\jszip\dist\jszip.min.js")

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
        "aliases": ["Conteúdo de TV/Vídeo sem referência", "Conteúdo de TV/Vídeo sem referÄ™ncia"],
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
                metric = "adh"
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
    line_minutes = sorted(
        [m for m in minutes_all if m["minute"] >= 6 * 60 or m["minute"] < 6 * 60],
        key=lambda m: (m["minute"] - 6 * 60) % (24 * 60),
    )
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
        for m in line_minutes
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
        bucket = profile.get(name, {}).get(target, {}).get(key, {}) or {}
        return bucket.get("adh") if "adh" in bucket else next(iter(bucket.values()), None)

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
    if not path.exists() and OUT.exists():
        html = OUT.read_text(encoding="utf-8", errors="ignore")
        weight = "700" if "Bold" in path.name else "400"
        match = re.search(rf"@font-face\{{[^}}]+base64,([^)]*)\)[^}}]+font-weight:{weight}", html)
        if match:
            return match.group(1)
    return base64.b64encode(path.read_bytes()).decode("ascii")


def image_data(path: Path) -> str:
    if not path.exists() and OUT.exists():
        html = OUT.read_text(encoding="utf-8", errors="ignore")
        key = "plim" if "PLIMPLIM" in path.name.upper() else "inteligencia" if "inteligencia" in path.name.lower() else ""
        if key:
            match = re.search(rf'"{key}":"(data:image/[^"]+)"', html)
            if match:
                return match.group(1)
        if "PLIMPLIM" in path.name.upper():
            match = re.search(r'<link rel="icon"[^>]+href="([^"]+)"', html)
            if match:
                return match.group(1)
    suffix = path.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in ("jpg", "jpeg") else suffix
    return f"data:image/{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


def render_html(data: dict) -> str:
    for ch in data["channels"]:
        if ch.get("logo"):
            ch["logoData"] = image_data(LOGO_ROOT / ch["logo"])
    data["assets"] = {"plim": image_data(PLIMPLIM), "inteligencia": image_data(INTELIGENCIA)}
    plim_icon = data["assets"]["plim"]
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    font_reg = font_data(FONT_REG)
    font_bold = font_data(FONT_BOLD)
    jszip_code = JSZIP.read_text(encoding="utf-8").replace("</script", "<\\/script")
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Desempenho Diário DF</title>
<link rel="icon" type="image/png" href="{plim_icon}">
<style>
@font-face{{font-family:Globotipo;src:url(data:font/ttf;base64,{font_reg}) format("truetype");font-weight:400}}
@font-face{{font-family:Globotipo;src:url(data:font/ttf;base64,{font_bold}) format("truetype");font-weight:700}}
:root{{--blue:#005cef;--blue2:#2b7cff;--ink:#101521;--muted:#607086;--line:#d9e0ec;--soft:#f2f4f7;--bg:#eef2f7;--card:#fff;--green:#07934a;--red:#d62839;--male:#2f80ed;--female:#e94d9a;--class1:#005cef;--class2:#00a651;--class3:#ffb000;--class4:#ec4b91;--class5:#7a4a27}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Globotipo,Arial,sans-serif;background:linear-gradient(180deg,#f8fafc 0%,var(--bg) 44%,#f7f8fa 100%);color:var(--ink);letter-spacing:0}}body.awaiting-bases main,body.awaiting-bases .tabs{{display:none}}body.awaiting-bases header.app{{min-height:100vh;display:grid;place-items:center;padding:24px}}body.awaiting-bases .top{{display:flex;flex-direction:column;justify-content:center;align-items:center;max-width:760px;width:min(92vw,760px);min-height:520px;text-align:center;padding:56px 42px}}body.awaiting-bases .brand{{flex-direction:column;gap:20px;font-size:58px}}body.awaiting-bases .brand-logo{{width:130px;height:130px}}body.awaiting-bases .header-right{{justify-items:center;width:100%}}body.awaiting-bases .actions{{min-width:0;width:auto;justify-content:center;margin-top:6px;background:transparent;padding:0;border-radius:0}}body.awaiting-bases .action-btn.primary{{background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;border-color:transparent;box-shadow:0 12px 28px rgba(0,92,239,.22)}}body.awaiting-bases .file-status{{display:none}}
header.app{{background:linear-gradient(90deg,rgba(0,92,239,.04) 1px,transparent 1px),linear-gradient(180deg,rgba(0,92,239,.04) 1px,transparent 1px),#eef2f7;background-size:72px 72px;position:relative;z-index:20;padding:18px 20px 0}}
.top{{max-width:1440px;margin:auto;background:#fff;border:1px solid var(--line);border-top:8px solid transparent;border-image:linear-gradient(90deg,#0046ff,#00a3ff) 1;border-radius:8px;padding:28px 38px;display:grid;grid-template-columns:1fr auto;gap:22px;align-items:center;box-shadow:0 10px 26px rgba(21,36,66,.08)}}
.brand{{display:flex;align-items:center;gap:22px;font-weight:700;font-size:54px;line-height:1;color:var(--blue);letter-spacing:0;text-transform:uppercase}}.brand span{{background:linear-gradient(90deg,#0046ff 0%,#005cef 48%,#00a3ff 100%);-webkit-background-clip:text;background-clip:text;color:transparent}}.brand-logo{{width:78px;height:78px;object-fit:contain;filter:none}}.header-right{{display:grid;gap:12px;justify-items:end}}.actions{{display:flex;gap:10px;align-items:center;flex-wrap:wrap;justify-content:flex-end;background:linear-gradient(135deg,#004cff,#05a9f5);border-radius:8px;padding:18px 20px;min-width:360px}}.action-btn{{border:1px solid rgba(255,255,255,.45);background:#fff;color:#005cef;border-radius:8px;padding:10px 12px;font:inherit;font-weight:700;cursor:pointer}}.action-btn.primary{{background:rgba(255,255,255,.16);color:#fff;border-color:rgba(255,255,255,.45)}}.file-status{{font-size:12px;color:#fff;max-width:330px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;grid-column:1/-1;justify-self:start}}
.tabs{{display:flex;gap:6px;background:#e8edf5;padding:4px;border-radius:8px}}.tab{{border:0;background:transparent;padding:10px 16px;border-radius:6px;font-weight:700;cursor:pointer;color:#40506a}}.tab.active{{background:linear-gradient(135deg,var(--blue),#3e8bff);color:#fff}}
main{{max-width:1440px;margin:0 auto;padding:24px;display:grid;gap:18px}}.panel{{display:none}}.panel.active{{display:grid;gap:18px}}
.hero{{background:linear-gradient(135deg,#0046c8 0%,#005cef 48%,#45a3ff 100%);color:#fff;padding:26px;border-radius:8px;display:grid;grid-template-columns:1fr auto;gap:20px;align-items:center;box-shadow:0 18px 45px rgba(0,92,239,.22)}}.hero-head{{display:flex;align-items:center;gap:18px;flex-wrap:wrap}}.hero h1{{margin:0;font-size:34px;line-height:1.05}}.date-pill{{display:inline-grid;gap:2px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.35);border-radius:8px;padding:12px 16px;min-width:190px}}.date-pill strong{{font-size:24px}}.date-pill span{{font-size:13px;text-transform:uppercase;letter-spacing:.08em}}.kpi{{background:#fff;color:var(--ink);border-radius:8px;padding:16px 20px;min-width:220px;box-shadow:0 12px 28px rgba(0,0,0,.12)}}.kpi small{{display:block;color:var(--muted);font-size:12px;text-transform:uppercase}}.kpi strong{{font-size:42px;line-height:1}}
.toolbar{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px;display:grid;grid-template-columns:260px 260px 1fr;gap:16px;align-items:center}}.toolbar.one{{grid-template-columns:280px 1fr}}label{{font-size:12px;color:var(--muted);font-weight:700;text-transform:uppercase;display:grid;gap:6px}}select{{font:inherit;border:1px solid var(--line);border-radius:6px;padding:10px 12px;background:#fff;min-width:0}}.selected-title{{text-align:center;font-weight:700;font-size:26px;color:var(--blue);background:var(--soft);border-radius:8px;padding:14px 18px}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}.card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:18px;min-width:0;box-shadow:0 10px 28px rgba(21,36,66,.06)}}.card h2{{margin:0 0 14px;font-size:18px}}.card h3{{margin:0 0 12px;font-size:16px}}.insights-card{{background:linear-gradient(180deg,#fff,#f8fbff);border-left:6px solid var(--blue)}}.insight-grid{{display:grid;grid-template-columns:1.2fr repeat(3,1fr);gap:14px;align-items:stretch}}.insight-main{{padding:16px;border-radius:8px;background:#eef5ff;color:#18345f;font-size:17px;line-height:1.35;font-weight:700}}.insight-item{{padding:14px;border:1px solid var(--line);border-radius:8px;background:#fff;display:grid;gap:7px}}.insight-item small{{color:var(--muted);text-transform:uppercase;font-size:11px;font-weight:700}}.insight-item strong{{font-size:22px;color:var(--ink)}}.insight-item span{{color:#526178;font-size:13px;line-height:1.3}}.insight-list{{margin:14px 0 0;padding-left:18px;color:#3e4d63;line-height:1.45}}.insight-list li{{margin:6px 0}}.chart{{width:100%;height:290px;display:block}}.line-chart{{height:380px}}.legend{{display:flex;gap:12px;flex-wrap:wrap;margin-top:10px;font-size:12px;color:#3e4d63}}.legend span,.channel-name{{display:inline-flex;align-items:center;gap:6px}}.dot{{width:10px;height:10px;border-radius:50%;display:inline-block}}.logo-chip{{width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:9px;line-height:1;overflow:hidden;box-shadow:inset 0 0 0 1px rgba(255,255,255,.28);text-align:center}}.logo-chip img{{display:block;width:72%;height:72%;object-fit:contain;object-position:center;margin:auto}}.leader .logo-chip{{color:#fff}}
.leader-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}}.leader{{border:1px solid var(--line);border-radius:8px;padding:14px;background:linear-gradient(180deg,#fff,#f8fafc)}}.leader strong{{font-size:28px;display:block}}.leader span{{color:var(--muted);font-size:12px}}.leader-head{{display:flex;align-items:center;gap:8px;margin-bottom:8px}}
.rank-grid{{display:grid;grid-template-columns:repeat(2, minmax(0,1fr));gap:18px}}table{{width:100%;border-collapse:collapse;font-size:13px}}th,td{{padding:10px 8px;border-bottom:1px solid #edf1f7;text-align:left;vertical-align:middle}}th{{color:var(--muted);font-size:11px;text-transform:uppercase;background:#fafbfe}}td.num,th.num{{text-align:right}}.arrow{{font-size:18px;font-weight:700}}.up{{color:var(--green)}}.down{{color:var(--red)}}.empty{{color:var(--muted);padding:24px;text-align:center;background:#f8fafc;border-radius:8px}}
.profile-grid{{display:grid;grid-template-columns:1fr;gap:18px}}.profile-card{{padding:22px}}.profile-head{{display:flex;justify-content:space-between;gap:16px;align-items:center;margin-bottom:18px}}.profile-title{{display:flex;align-items:center;gap:10px}}.metric-pair{{display:flex;gap:10px}}.mini{{border:0;border-radius:8px;padding:10px 12px;min-width:96px;text-align:right;background:var(--metric-grad,linear-gradient(135deg,#005cef,#45a3ff));color:#fff;box-shadow:0 10px 24px rgba(21,36,66,.16)}}.mini small{{display:block;color:rgba(255,255,255,.82);font-size:11px}}.mini strong{{font-size:26px;color:#fff}}.mini-chart{{height:150px;width:100%;display:block}}.profile-block{{display:grid;grid-template-columns:minmax(230px,.9fr) minmax(360px,1.15fr) minmax(460px,1.45fr);gap:0;align-items:start}}.profile-block>div{{min-width:0;padding:0 22px}}.profile-block>div+div{{border-left:1px solid #e5ebf3}}.profile-block>div:nth-child(2){{display:grid;justify-items:center}}.profile-block>div:nth-child(2) h4{{justify-self:start;width:100%}}.profile-block h4{{margin:0 0 8px;font-size:12px;color:var(--muted);text-transform:uppercase}}.program-table-wrap{{max-height:470px;overflow:auto;border:1px solid var(--line);border-radius:8px}}.program-table-wrap table th{{position:sticky;top:0;z-index:1}}.minute-table th,.minute-table td{{white-space:nowrap}}.footer-note{{display:grid;gap:10px;color:#66758c;font-size:13px;padding:0 0 18px}}.source-note{{text-align:left}}.credit-note{{text-align:center}}.final-actions{{display:flex;justify-content:center;gap:12px;flex-wrap:wrap;padding:8px 0 12px}}.final-actions .action-btn{{border-color:#c9d4e5;box-shadow:0 8px 20px rgba(21,36,66,.08)}}.final-actions .action-btn.primary{{background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;border-color:transparent}}
@media(max-width:980px){{.top,.hero,.toolbar,.toolbar.one,.grid2,.rank-grid,.profile-grid,.profile-block,.insight-grid{{grid-template-columns:1fr}}.brand{{font-size:34px}}.brand-logo{{width:54px;height:54px}}body.awaiting-bases .brand{{font-size:40px}}body.awaiting-bases .brand-logo{{width:108px;height:108px}}.header-right{{justify-items:stretch}}body.awaiting-bases .header-right{{justify-items:center}}.actions{{min-width:0}}.leader-grid{{grid-template-columns:repeat(2,1fr)}}.selected-title{{text-align:left}}.profile-block>div{{padding:0}}.profile-block>div+div{{border-left:0;border-top:1px solid #e5ebf3;padding-top:16px}}}}
</style>
</head>
<body class="awaiting-bases">
<header class="app"><div class="top"><div class="brand"><img class="brand-logo" id="plimLogo" src="{plim_icon}" alt="Plim plim"><span>Desempenho Diário</span></div><div class="header-right"><div class="actions"><input id="baseUpload" type="file" accept=".zip,.xlsx" multiple hidden><button class="action-btn primary" id="uploadBtn">Enviar base</button><span class="file-status" id="fileStatus"></span></div><nav class="tabs"><button class="tab active" data-tab="resumo">RESUMO</button><button class="tab" data-tab="programas">PROGRAMAS</button><button class="tab" data-tab="perfil">PERFIL</button></nav></div></div></header>
<main>
<section id="resumo" class="panel active">
  <div class="hero"><div class="hero-head"><span id="dateLabel"></span><h1>Resumo do Distrito Federal</h1></div><div class="kpi"><small>Audiência média 07h-24h</small><strong id="dailyAvg"></strong></div></div>
  <div class="card insights-card"><h2>Destaques do dia</h2><div id="dailyInsights"></div></div>
  <div class="grid2"><div class="card"><h2>Audiência 07:00 às 24:00</h2><svg id="audBars" class="chart"></svg></div><div class="card"><h2>Share% 07:00 às 24:00</h2><svg id="shareBars" class="chart"></svg></div></div>
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
<div class="footer-note"><span class="source-note" id="sourceNote"></span></div>
<div class="final-actions"><button class="action-btn primary" id="exportBtn">Compartilhar</button><button class="action-btn" id="imageBtn">Gerar imagem</button></div>
<div class="footer-note"><span class="credit-note">Modelo desenvolvido pela área de Programação da TV Globo DF.</span></div>
</main>
<script>{jszip_code}</script>
<script id="data" type="application/json">{data_json.replace("</", "<\\/")}</script>
<script>
let DATA = JSON.parse(document.getElementById('data').textContent);
const fmt = (v,d=2) => v === null || v === undefined || Number.isNaN(v) ? '-' : Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:d, maximumFractionDigits:d}});
const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
document.getElementById('plimLogo').src = DATA.assets.plim;
const channelMap = Object.fromEntries(DATA.channels.map(c=>[c.key,c]));
const visible = DATA.channels.filter(c=>c.key!=='tle');
const audChannels = DATA.channels;
const shareChannels = visible;
const lineChannels = DATA.channels;
const leadChannels = DATA.leadershipChannels;
const rankingChannels = DATA.rankingChannels;
function logoHtml(c){{return c.logoData ? `<span class="logo-chip" style="background:${{c.color}}"><img src="${{c.logoData}}" alt="${{c.label}}"></span>` : `<span class="logo-chip" style="background:${{c.color}}">${{c.textLogo || c.label.slice(0,3).toUpperCase()}}</span>`}}
function legend(el, channels=visible){{document.getElementById(el).innerHTML=channels.map(c=>`<span>${{logoHtml(c)}}${{c.label}}</span>`).join('')}}
function svgEl(tag, attrs={{}}){{const e=document.createElementNS('http://www.w3.org/2000/svg',tag);Object.entries(attrs).forEach(([k,v])=>e.setAttribute(k,v));return e}}
function clear(svg){{while(svg.firstChild)svg.removeChild(svg.firstChild)}}
function svgLogo(svg,c,cx,cy,size=24){{svg.appendChild(svgEl('circle',{{cx,cy,r:size/2,fill:c.color}}));if(c.logoData){{const inner=size*.72;const im=svgEl('image',{{href:c.logoData,x:cx-inner/2,y:cy-inner/2,width:inner,height:inner,preserveAspectRatio:'xMidYMid meet'}});svg.appendChild(im)}}else{{svg.appendChild(svgEl('text',{{x:cx,y:cy+3,'text-anchor':'middle','font-size':8,'font-weight':700,fill:'#fff'}})).textContent=c.textLogo||c.label.slice(0,3).toUpperCase()}}}}
function barChart(id, data, metric, channels=visible){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||700,h=svg.clientHeight||290,ml=44,mr=16,mt=20,mb=48;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const vals=channels.map(c=>data.find(d=>d.key===c.key)?.[metric]??0);const max=Math.max(...vals,1)*1.15;svg.appendChild(svgEl('line',{{x1:ml,y1:h-mb,x2:w-mr,y2:h-mb,stroke:'#d9e0ec'}}));const slot=(w-ml-mr)/channels.length,bw=slot*.58;channels.forEach((c,i)=>{{const v=vals[i];const x=ml+i*slot+(slot-bw)/2;const y=mt+(h-mt-mb)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:4,fill:c.color}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:y-7,'text-anchor':'middle','font-size':12,'font-weight':700,fill:'#111'}})).textContent=fmt(v,2);svgLogo(svg,c,x+bw/2,h-24,24);}})}}
function lineChart(id, rows, channels=lineChannels, accessor=r=>r.aud, highlightKey='globo'){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||1000,h=svg.clientHeight||380,ml=48,mr=18,mt=20,mb=42;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const defs=svgEl('defs');const grad=svgEl('linearGradient',{{id:`${{id}}-badgeGrad`,x1:'0%',y1:'0%',x2:'100%',y2:'0%'}});grad.appendChild(svgEl('stop',{{offset:'0%','stop-color':'#004cff'}}));grad.appendChild(svgEl('stop',{{offset:'100%','stop-color':'#05a9f5'}}));defs.appendChild(grad);svg.appendChild(defs);let max=1,highlight=[];rows.forEach((r,i)=>channels.forEach(c=>{{const v=accessor(r)[c.key];if(v!==null&&v!==undefined){{if(v>max)max=v;if(c.key===highlightKey)highlight.push({{v,i,r,c}})}}}}));max*=1.12;for(let g=0;g<=4;g++){{const y=mt+(h-mt-mb)*g/4;svg.appendChild(svgEl('line',{{x1:ml,y1:y,x2:w-mr,y2:y,stroke:'#edf1f7'}}));svg.appendChild(svgEl('text',{{x:8,y:y+4,'font-size':11,fill:'#647084'}})).textContent=fmt(max*(1-g/4),1)}}channels.forEach(c=>{{let d='';rows.forEach((r,i)=>{{const v=accessor(r)[c.key];if(v===null||v===undefined)return;const x=ml+(w-ml-mr)*(i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-v/max);d+=(d?'L':'M')+x.toFixed(1)+' '+y.toFixed(1)}});svg.appendChild(svgEl('path',{{d,fill:'none',stroke:c.color,'stroke-width':c.key===highlightKey?3:2.1,'stroke-linejoin':'round','stroke-linecap':'round'}}));}});if(highlight.length){{const hi=highlight.reduce((a,b)=>b.v>a.v?b:a,highlight[0]);const lo=highlight.reduce((a,b)=>b.v<a.v?b:a,highlight[0]);[{{p:hi,label:'Maior'}},{{p:lo,label:'Menor'}}].forEach(({{p,label}})=>{{const x=ml+(w-ml-mr)*(p.i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-p.v/max);svg.appendChild(svgEl('circle',{{cx:x,cy:y,r:5.5,fill:'#005cef',stroke:'#fff','stroke-width':2}}));const text=`${{label}} ${{fmt(p.v,1)}}`;const width=Math.max(86,text.length*7+18);const tx=Math.min(w-width/2-12,Math.max(width/2+12,x+width/2));const ty=Math.max(24,y-12);svg.appendChild(svgEl('rect',{{x:tx-width/2,y:ty-17,width,height:24,rx:6,fill:`url(#${{id}}-badgeGrad)`}}));svg.appendChild(svgEl('text',{{x:tx,y:ty,'text-anchor':'middle','font-size':11,'font-weight':700,fill:'#fff'}})).textContent=text;}})}}const tickCount=Math.min(15,Math.max(6,Math.ceil(rows.length/70)+8));for(let t=0;t<tickCount;t++){{const p=t/(tickCount-1);const i=Math.min(rows.length-1,Math.round((rows.length-1)*p));const x=ml+(w-ml-mr)*p;svg.appendChild(svgEl('text',{{x,y:h-14,'text-anchor':'middle','font-size':11,fill:'#647084'}})).textContent=rows[i]?.time||''}}}}
function miniBars(svg, obj, color){{clear(svg);const entries=Object.entries(obj||{{}});const w=svg.clientWidth||320,h=svg.clientHeight||140,ml=34,mb=30,mt=10;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const max=Math.max(...entries.map(e=>e[1]||0),1);const bw=(w-ml-8)/Math.max(entries.length,1)*.58;entries.forEach(([k,v],i)=>{{v=v||0;const fill=Array.isArray(color)?color[i%color.length]:color;const x=ml+(i+.21)*(w-ml-8)/entries.length;const y=mt+16+(h-mt-mb-16)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:3,fill}}));const val=fmt(v,0)+'%';const valueW=Math.max(34,val.length*7+14);const valueY=Math.max(8,y-24);svg.appendChild(svgEl('rect',{{x:x+bw/2-valueW/2,y:valueY,width:valueW,height:18,rx:5,fill:'#fff',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:valueY+13,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#27364e'}})).textContent=val;const labelW=Math.max(34,k.length*6+14);svg.appendChild(svgEl('rect',{{x:x+bw/2-labelW/2,y:h-24,width:labelW,height:18,rx:5,fill:'#f2f4f7',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:h-11,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#526178'}})).textContent=k;}})}}
function pie(svg, obj, colors=['#f59e0b','#f97316','#fb923c','#d97706','#b45309']){{clear(svg);const entries=Object.entries(obj||{{}}).filter(e=>(e[1]||0)>0);const w=svg.clientWidth||420,h=svg.clientHeight||150,groupW=430,offset=Math.max(0,(w-groupW)/2),cx=offset+128,cy=74,r=54;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const total=entries.reduce((a,e)=>a+e[1],0)||1;let a0=-Math.PI/2;entries.forEach(([k,v],i)=>{{const a1=a0+(v/total)*Math.PI*2;const large=a1-a0>Math.PI?1:0;const x0=cx+r*Math.cos(a0), y0=cy+r*Math.sin(a0), x1=cx+r*Math.cos(a1), y1=cy+r*Math.sin(a1);svg.appendChild(svgEl('path',{{d:`M${{cx}} ${{cy}} L${{x0}} ${{y0}} A${{r}} ${{r}} 0 ${{large}} 1 ${{x1}} ${{y1}} Z`,fill:colors[i%colors.length]}}));a0=a1;}});entries.forEach(([k,v],i)=>{{const lx=offset+260;svg.appendChild(svgEl('rect',{{x:lx,y:18+i*22,width:11,height:11,rx:2,fill:colors[i%colors.length]}}));svg.appendChild(svgEl('text',{{x:lx+17,y:28+i*22,'font-size':12,fill:'#3e4d63'}})).textContent=`${{k}} ${{fmt(v,1)}}%`;}})}}
function statusArrow(status){return status==='cresceu'?'<span class="arrow up">&#9650;</span>':status==='caiu'?'<span class="arrow down">&#9660;</span>':'<span class="arrow">&bull;</span>'}
function renderRankings(){{const box=document.getElementById('rankings');box.innerHTML=rankingChannels.map(c=>{{const rows=DATA.rankings[c.key]||[];return `<div class="card"><h3 class="channel-name">${{logoHtml(c)}}${{c.label}}</h3><table><thead><tr><th>#</th><th>Programa</th><th class="num">Aud.</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th><th class="num">Share</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th></tr></thead><tbody>${{rows.map(r=>`<tr><td>${{r.rank}}</td><td>${{esc(r.program)}}</td><td class="num">${{fmt(r.aud)}}</td><td class="num">${{fmt(r.audPrev)}}</td><td class="num">${{r.audVar==null?'-':fmt(r.audVar,1)+'%'}}</td><td>${{statusArrow(r.audStatus)}}</td><td class="num">${{fmt(r.share)}}</td><td class="num">${{fmt(r.sharePrev)}}</td><td class="num">${{r.shareVar==null?'-':fmt(r.shareVar,1)+'%'}}</td><td>${{statusArrow(r.shareStatus)}}</td></tr>`).join('')}}</tbody></table></div>`}}).join('')}}
function sourceText(){{const d=new Date(DATA.meta.date+'T00:00:00');return `Fonte: Ibope. Instar Analytics. DF. Total Domicílios. Aud%, Shr% Adh%. Atividades: Live. Data: ${{d.toLocaleDateString('pt-BR')}}.`}}
function dateToken(){{const d=new Date(DATA.meta.date+'T00:00:00');return `${{String(d.getDate()).padStart(2,'0')}}${{String(d.getMonth()+1).padStart(2,'0')}}${{d.getFullYear()}}`}}
function firstValid(items, selector){{return items.find(item=>selector(item)!==null&&selector(item)!==undefined)}}
function buildDailyInsights(){{const globo=DATA.summaryBars.find(x=>x.key==='globo')||{{}};const record=DATA.summaryBars.find(x=>x.key==='record')||{{}};const sbt=DATA.summaryBars.find(x=>x.key==='sbt')||{{}};const nic=DATA.summaryBars.find(x=>x.key==='nic')||{{}};const competitors=[record,sbt,nic].filter(x=>x&&x.key).sort((a,b)=>(b.aud||0)-(a.aud||0));const mainCompetitor=competitors[0]||{{label:'concorr\u00eancia',aud:null,share:null}};const globoRank=(DATA.rankings.globo||[]).filter(r=>r.program);const byAud=[...globoRank].filter(r=>r.aud!==null&&r.aud!==undefined).sort((a,b)=>(b.aud||0)-(a.aud||0));const bestAud=byAud[0], weakest=byAud[byAud.length-1];const growth=[...globoRank].filter(r=>r.audVar!==null&&r.audVar!==undefined).sort((a,b)=>(b.audVar||0)-(a.audVar||0))[0];const fall=[...globoRank].filter(r=>r.audVar!==null&&r.audVar!==undefined).sort((a,b)=>(a.audVar||0)-(b.audVar||0))[0];const globoPoints=DATA.line.map(r=>({{time:r.time,value:r.aud?.globo}})).filter(p=>p.value!==null&&p.value!==undefined);const peak=globoPoints.length?[...globoPoints].sort((a,b)=>b.value-a.value)[0]:null;const low=globoPoints.length?[...globoPoints].sort((a,b)=>a.value-b.value)[0]:null;const leader=(DATA.leadership||[]).find(l=>l.key==='globo')||{{minutes:0,percent:0,hours:'0h00'}};const headline=`A Globo fechou o per\u00edodo 07h-24h com m\u00e9dia de ${{fmt(globo.aud,2)}} pontos e ${{fmt(globo.share,2)}}% de share. A an\u00e1lise indica ${{fall&&fall.audVar<0?'queda em parte da grade':'estabilidade na grade'}} na compara\u00e7\u00e3o com as \u00faltimas 4 semanas, com ${{leader.minutes}} minutos de lideran\u00e7a (${{fmt(leader.percent,1)}}% do per\u00edodo).`;const cards=[{{k:'Melhor programa',v:bestAud?bestAud.program:'-',s:bestAud?`${{fmt(bestAud.aud,2)}} pontos | ${{fmt(bestAud.share,2)}}% share`:'Sem dado dispon\u00edvel'}},{{k:'Maior avan\u00e7o',v:growth&&growth.audVar>0?growth.program:'-',s:growth&&growth.audVar>0?`${{fmt(growth.audVar,1)}}% vs. 4 semanas | ${{fmt(growth.aud,2)}} pontos`:'Sem alta relevante'}},{{k:'Ponto de aten\u00e7\u00e3o',v:fall&&fall.audVar<0?fall.program:(weakest?weakest.program:'-'),s:fall&&fall.audVar<0?`${{fmt(fall.audVar,1)}}% vs. 4 semanas | ${{fmt(fall.aud,2)}} pontos`:(weakest?`${{fmt(weakest.aud,2)}} pontos | ${{fmt(weakest.share,2)}}% share`:'Sem dado dispon\u00edvel')}}];const bullets=[];if(weakest)bullets.push(`A faixa de maior dificuldade entre os programas ranqueados foi ${{esc(weakest.program)}}, com ${{fmt(weakest.aud,2)}} pontos e ${{fmt(weakest.share,2)}}% de share.`);if(mainCompetitor&&mainCompetitor.key)bullets.push(`Entre os concorrentes, ${{esc(mainCompetitor.label)}} foi o maior destaque no per\u00edodo, com ${{fmt(mainCompetitor.aud,2)}} pontos e ${{fmt(mainCompetitor.share,2)}}% de share.`);if(growth&&growth.audVar>0)bullets.push(`${{esc(growth.program)}} foi o principal avan\u00e7o frente \u00e0 m\u00e9dia das 4 semanas anteriores.`);if(fall&&fall.audVar<0)bullets.push(`${{esc(fall.program)}} concentrou a principal queda na compara\u00e7\u00e3o com as 4 semanas anteriores.`);if(peak&&low)bullets.push(`No minuto a minuto, a Globo atingiu pico de ${{fmt(peak.value,1)}} \u00e0s ${{peak.time}} e menor patamar de ${{fmt(low.value,1)}} \u00e0s ${{low.time}}.`);return `<div class="insight-grid"><div class="insight-main">${{headline}}</div>${{cards.map(c=>`<div class="insight-item"><small>${{c.k}}</small><strong>${{esc(c.v)}}</strong><span>${{c.s}}</span></div>`).join('')}}</div><ul class="insight-list">${{bullets.map(b=>`<li>${{b}}</li>`).join('')}}</ul>`}}
function renderResumo(){{const d=new Date(DATA.meta.date+'T00:00:00');document.getElementById('dateLabel').className='date-pill';document.getElementById('dateLabel').innerHTML=`<span>${{DATA.meta.weekday}}</span><strong>${{d.toLocaleDateString('pt-BR')}}</strong>`;document.getElementById('dailyAvg').textContent=fmt(DATA.meta.dailyAvg,2);document.getElementById('sourceNote').textContent=sourceText();document.getElementById('dailyInsights').innerHTML=buildDailyInsights();barChart('audBars',DATA.summaryBars,'aud',audChannels);barChart('shareBars',DATA.summaryBars,'share',shareChannels);lineChart('dayLine',DATA.line,lineChannels);legend('legend3',lineChannels);document.getElementById('leadership').innerHTML=DATA.leadership.map(l=>{{const c=channelMap[l.key];return `<div class="leader" style="border-top:5px solid ${{l.color}}"><div class="leader-head">${{logoHtml(c)}}<span>${{l.label}}</span></div><strong>${{l.minutes}}</strong><span>${{fmt(l.percent,1)}}% | ${{l.hours}}</span></div>`}}).join('');renderRankings()}}
function fillSelect(sel, values){{sel.innerHTML=values.map(v=>`<option value="${{esc(v)}}">${{esc(v)}}</option>`).join('')}}
function programMinuteRows(program){{const start=program.start??0,end=program.end??0;return DATA.minuteAll.filter(m=>m.minute>=start && m.minute<end)}}
function renderProgramas(){{const p=DATA.programs.find(x=>x.name===programSelect.value)||DATA.programs[0];const target=targetSelect.value;programTitle.textContent=`${{p.name}} | ${{target}}`;const comp=(DATA.programCompetition[p.name]||{{}})[target]||{{}};const audBars=audChannels.map(c=>({{key:c.key,label:c.label,color:c.color,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));const shareBars=shareChannels.map(c=>({{key:c.key,label:c.label,color:c.color,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));barChart('progAud',audBars,'aud',audBars);barChart('progShare',shareBars,'share',shareBars);const rows=programMinuteRows(p);lineChart('progLine',rows,lineChannels,r=>r.aud);legend('legend4',lineChannels);minuteTable.innerHTML=`<table class="minute-table"><thead><tr><th>Minuto</th>${{visible.map(c=>`<th colspan="2" style="border-top:4px solid ${{c.color}}">${{logoHtml(c)}} ${{c.label}}</th>`).join('')}}</tr><tr><th></th>${{visible.map(c=>`<th class="num">Aud.</th><th class="num">Share</th>`).join('')}}</tr></thead><tbody>${{rows.map(r=>`<tr><td><strong>${{r.time}}</strong></td>${{visible.map(c=>`<td class="num" style="color:${{c.color}}">${{fmt(r.aud[c.key])}}</td><td class="num">${{fmt(r.share[c.key])}}</td>`).join('')}}</tr>`).join('')}}</tbody></table>`}}
function softenColor(hex){{const h=String(hex||'#005cef').replace('#','');const raw=h.length===3?h.split('').map(x=>x+x).join(''):h;const n=parseInt(raw,16);if(!Number.isFinite(n))return '#45a3ff';const r=(n>>16)&255,g=(n>>8)&255,b=n&255;const mix=v=>Math.round(v+(255-v)*.34);return `rgb(${{mix(r)}},${{mix(g)}},${{mix(b)}})`}}
function renderPerfil(){{const name=profileSelect.value;profileTitle.textContent=name;const pdata=DATA.profile[name]||{{}};const genderColors=['#4f8bd6','#d95f76'];const classColors=['#ffd166','#ef476f','#06d6a0','#118ab2','#8338ec'];const ageColors=['#80ffdb','#56cfe1','#4ea8de','#5e60ce','#6930c3','#1b4965'];profileGrid.innerHTML=leadChannels.map(c=>{{const d=pdata[c.key]||{{}};return `<div class="card profile-card" data-ch="${{c.key}}" style="border-top:5px solid ${{c.color}};--metric-grad:linear-gradient(135deg,${{c.color}},${{softenColor(c.color)}})"><div class="profile-head"><h3 class="profile-title">${{logoHtml(c)}}${{c.label}}</h3><div class="metric-pair"><div class="mini"><small>Aud.</small><strong>${{fmt(d.aud)}}</strong></div><div class="mini"><small>Share</small><strong>${{fmt(d.share)}}</strong></div></div></div><div class="profile-block"><div><h4>G\u00eanero</h4><svg class="mini-chart gender"></svg></div><div><h4>Classes sociais</h4><svg class="mini-chart classes"></svg></div><div><h4>Faixas etárias</h4><svg class="mini-chart ages"></svg></div></div></div>`}}).join('');document.querySelectorAll('.profile-card').forEach(card=>{{const key=card.dataset.ch, d=pdata[key]||{{}};miniBars(card.querySelector('.gender'),d.gender,genderColors);pie(card.querySelector('.classes'),d.classes,classColors);miniBars(card.querySelector('.ages'),d.ages,ageColors)}})}}
function downloadBlob(name,type,content){{const blob=new Blob([content],{{type}});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),500)}}
function exportHtml(){{const clone=document.documentElement.cloneNode(true);clone.querySelectorAll('#baseUpload').forEach(el=>el.value='');const dataScript=clone.querySelector('#data');if(dataScript)dataScript.textContent=JSON.stringify(DATA).replaceAll('</','<\\\\/');downloadBlob(`desempenho_diario_df_${{dateToken()}}.html`,'text/html;charset=utf-8','<!doctype html>\\n'+clone.outerHTML)}}
async function generateHighlightsImage(){{const d=new Date(DATA.meta.date+'T00:00:00');const metricKeys=['globo','nic','record','sbt'];const metricItems=metricKeys.map(key=>{{const c=channelMap[key]||{{key,label:key.toUpperCase(),color:'#005cef'}};const row=DATA.summaryBars.find(x=>x.key===key)||{{}};return {{...c,aud:row.aud,share:row.share}}}});const globoShares=(DATA.rankings.globo||[]).filter(r=>r.share!==null&&r.share!==undefined);const topShare=[...globoShares].sort((a,b)=>(b.share||0)-(a.share||0)).slice(0,3);const bottomShare=[...globoShares].sort((a,b)=>(a.share||0)-(b.share||0)).slice(0,3);const canvas=document.createElement('canvas');const W=1400,H=690,S=2;canvas.width=W*S;canvas.height=H*S;canvas.style.width=W+'px';canvas.style.height=H+'px';const ctx=canvas.getContext('2d');ctx.scale(S,S);ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';const grad=ctx.createLinearGradient(0,0,W,H);grad.addColorStop(0,'#003bb2');grad.addColorStop(.52,'#005cef');grad.addColorStop(1,'#65c7ff');ctx.fillStyle=grad;ctx.fillRect(0,0,W,H);ctx.fillStyle='rgba(255,255,255,.08)';ctx.beginPath();ctx.arc(1120,55,390,0,Math.PI*2);ctx.fill();ctx.fillStyle='rgba(255,255,255,.06)';ctx.beginPath();ctx.arc(120,H,360,0,Math.PI*2);ctx.fill();const loadImg=src=>new Promise(resolve=>{{if(!src)return resolve(null);const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=src;}});const intelLogo=await loadImg(DATA.assets.inteligencia);const metricLogos=Object.fromEntries(await Promise.all(metricItems.map(async item=>[item.key,await loadImg(item.logoData)])));ctx.fillStyle='#fff';ctx.font='700 50px Globotipo, Arial';ctx.fillText('DESEMPENHO DI\u00c1RIO',76,86);ctx.font='400 26px Globotipo, Arial';ctx.fillText(`${{DATA.meta.weekday}} \u2022 ${{d.toLocaleDateString('pt-BR')}}`,78,124);if(intelLogo)ctx.drawImage(intelLogo,1000,42,324,46);ctx.font='700 29px Globotipo, Arial';ctx.fillText('M\u00c9DIA | 07h \u00e0s 24h',76,182);function fitText(text,maxWidth,fontSize,weight='700'){{let size=fontSize;do{{ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;if(ctx.measureText(text).width<=maxWidth)return {{text,size}};size-=1}}while(size>=13);let clipped=text;while(clipped.length>3&&ctx.measureText(clipped+'...').width>maxWidth)clipped=clipped.slice(0,-1);return {{text:clipped+'...',size}}}}function drawChannelMark(item,x,y,size){{ctx.fillStyle=item.color||'#005cef';ctx.beginPath();ctx.arc(x+size/2,y+size/2,size/2,0,Math.PI*2);ctx.fill();const logo=metricLogos[item.key];if(logo){{ctx.save();ctx.beginPath();ctx.arc(x+size/2,y+size/2,size*.42,0,Math.PI*2);ctx.clip();ctx.drawImage(logo,x+size*.18,y+size*.18,size*.64,size*.64);ctx.restore();}}else{{ctx.fillStyle='#fff';ctx.font='700 17px Globotipo, Arial';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(item.textLogo||item.label,x+size/2,y+size/2+1);ctx.textBaseline='alphabetic';ctx.textAlign='left';}}}}function metricCard(x,y,w,h,item){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,14);ctx.fill();ctx.fillStyle=item.color;ctx.fillRect(x,y,7,h);drawChannelMark(item,x+24,y+22,44);ctx.fillStyle='#607086';ctx.font='700 18px Globotipo, Arial';ctx.fillText(item.label,x+82,y+37);ctx.fillStyle='#101521';ctx.font='700 15px Globotipo, Arial';ctx.fillText('Aud.',x+82,y+68);ctx.fillText('Share',x+178,y+68);ctx.font='700 28px Globotipo, Arial';ctx.fillText(fmt(item.aud,2),x+82,y+99);ctx.fillStyle=item.color;ctx.fillText(fmt(item.share,2)+'%',x+178,y+99);ctx.fillStyle='#101521'}}function rankCard(x,y,w,h,title,items,accent){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,14);ctx.fill();ctx.fillStyle=accent;ctx.fillRect(x,y,7,h);ctx.fillStyle='#607086';ctx.font='700 23px Globotipo, Arial';ctx.textAlign='center';ctx.fillText(title,x+w/2,y+50);ctx.textAlign='left';items.forEach((item,i)=>{{const lineY=y+102+i*46;const fitted=fitText(`${{i+1}}. ${{item.program}}`,w-150,20,'700');ctx.fillStyle='#101521';ctx.font=`700 ${{fitted.size}}px Globotipo, Arial`;ctx.fillText(fitted.text,x+30,lineY);ctx.fillStyle=accent;ctx.font='700 22px Globotipo, Arial';ctx.textAlign='right';ctx.fillText(fmt(item.share,2)+'%',x+w-30,lineY);ctx.textAlign='left'}})}}metricItems.forEach((item,i)=>metricCard(76+i*318,214,292,112,item));rankCard(76,360,610,220,'Melhores desempenhos (SHR%)',topShare,'#6b7280');rankCard(714,360,610,220,'Piores desempenhos (SHR%)',bottomShare,'#374151');ctx.fillStyle='rgba(255,255,255,.92)';ctx.font='700 22px Globotipo, Arial';ctx.fillText('Relat\u00f3rio completo em anexo.',76,626);ctx.font='400 17px Globotipo, Arial';ctx.fillText(sourceText(),76,658);const a=document.createElement('a');a.href=canvas.toDataURL('image/png');a.download=`desempenho_diario_df_${{dateToken()}}.png`;a.click()}}
function normalizeText(value){{return String(value??'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').replace(/Ä™/g,'e').toLowerCase()}}
function fnumJs(value){{if(value===null||value===undefined||value===''||value==='n/a')return null;const n=Number(String(value).replace(',','.'));return Number.isFinite(n)?n:null}}
function pctJs(value){{return value===null||value===undefined?null:Math.round(value*100)/100}}
function avgJs(values){{const clean=values.filter(v=>typeof v==='number'&&Number.isFinite(v));return clean.length?clean.reduce((a,b)=>a+b,0)/clean.length:null}}
function excelDateJs(serial){{const n=fnumJs(serial);if(n===null)return '';const d=new Date(Date.UTC(1899,11,30)+Math.floor(n)*86400000);return d.toISOString().slice(0,10)}}
function fracToMinutesJs(value){{const n=fnumJs(value);return n===null?null:Math.round(n*24*60)}}
function hhmmFromMinJs(minutes){{minutes=((minutes%(24*60))+(24*60))%(24*60);return `${{String(Math.floor(minutes/60)).padStart(2,'0')}}:${{String(minutes%60).padStart(2,'0')}}`}}
function parseTimebandJs(label){{const m=String(label||'').match(/(\\d{{2}}):(\\d{{2}}):\\d{{2}}/);return m?Number(m[1])*60+Number(m[2]):null}}
function metricFromVarJs(v){{return String(v||'').includes('Shr%')?'share':String(v||'').includes('Rat%')?'aud':v}}
function channelForRaw(raw){{const n=normalizeText(raw);if(n.includes('reference')||n==='globo')return 'globo';if(n.includes('record'))return 'record';if(n==='sbt')return 'sbt';if(n.includes('band'))return 'band';if(n.includes('conteudo')&&n.includes('refer'))return 'nic';if(n.includes('canais paytv')||n==='ocp')return 'paytv';if(n.includes('total ligados especial'))return 'tle';return null}}
function colToIdxJs(ref){{let n=0;for(const ch of String(ref).replace(/[^A-Za-z]/g,''))n=n*26+(ch.toUpperCase().charCodeAt(0)-64);return n-1}}
async function readXlsxFile(file,nameOverride){{const zip=await JSZip.loadAsync(file);const parser=new DOMParser();const xml=async p=>parser.parseFromString(await zip.file(p).async('text'),'application/xml');let shared=[];if(zip.file('xl/sharedStrings.xml')){{const sst=await xml('xl/sharedStrings.xml');shared=[...sst.getElementsByTagName('si')].map(si=>[...si.getElementsByTagName('t')].map(t=>t.textContent||'').join(''))}}const wb=await xml('xl/workbook.xml');const rels=await xml('xl/_rels/workbook.xml.rels');const relMap={{}};[...rels.getElementsByTagName('Relationship')].forEach(r=>relMap[r.getAttribute('Id')]=r.getAttribute('Target'));const sheets={{}};for(const sh of [...wb.getElementsByTagName('sheet')]){{const name=sh.getAttribute('name');const rid=sh.getAttribute('r:id')||sh.getAttribute('id');let target=relMap[rid];let path=target.startsWith('/')?target.slice(1):(target.startsWith('xl/')?target:'xl/'+target);const doc=await xml(path);const rows=[];for(const row of [...doc.getElementsByTagName('row')]){{const vals=[];for(const c of [...row.getElementsByTagName('c')]){{const idx=colToIdxJs(c.getAttribute('r')||'A1');while(vals.length<=idx)vals.push('');const typ=c.getAttribute('t');let txt='';if(typ==='inlineStr')txt=[...c.getElementsByTagName('t')].map(t=>t.textContent||'').join('');else{{const v=c.getElementsByTagName('v')[0];txt=v?v.textContent||'':'';if(typ==='s'&&txt)txt=shared[Number(txt)]||''}}vals[idx]=txt}}rows.push(vals)}}sheets[name]=rows}}return {{name:nameOverride||file.name,sheets}}}}
function readMatrixJs(rows,headerRows,allowed){{const out={{}};for(const row of rows.slice(headerRows.start)){{if(!row||!row[0]||row[0]==='Total')continue;const name=row[0];out[name]={{}};const maxCols=Math.max(...rows.slice(0,headerRows.entity+1).map(r=>r.length));for(let i=1;i<maxCols;i++){{const target=(rows[headerRows.target]||[])[i]||'';const rawVar=(rows[headerRows.var]||[])[i]||'';const entity=(rows[headerRows.entity]||[])[i]||'';const key=channelForRaw(entity);let metric=allowed===null?'adh':metricFromVarJs(rawVar);if(target&&key&&(allowed===null||allowed.includes(metric))){{out[name][target]??={{}};out[name][target][key]??={{}};out[name][target][key][metric]=fnumJs(row[i])}}}}}}return out}}
function loadMinuteDataJs(rows){{const minutes=[];const targets=rows[1]||[], vars=rows[2]||[], entities=rows[3]||[];const maxCols=Math.max(targets.length,vars.length,entities.length);for(const row of rows.slice(5)){{const minute=parseTimebandJs(row[0]);if(minute===null)continue;const item={{time:hhmmFromMinJs(minute),minute,aud:{{}},share:{{}}}};for(let i=1;i<maxCols;i++){{if(!normalizeText(targets[i]).includes('total domic'))continue;const metric=metricFromVarJs(vars[i]);if(metric!=='aud'&&metric!=='share')continue;const key=channelForRaw(entities[i]);if(key)item[metric][key]=fnumJs(row[i])}}minutes.push(item)}}return minutes}}
function loadProgramsJs(programRows,crosstabRows){{const programs=[];const targets=programRows[1]||[], vars=programRows[2]||[];for(const row of programRows.slice(4)){{if(!row||!row[0])continue;const p={{name:row[0],start:fracToMinutesJs(row[1]),end:fracToMinutesJs(row[2]),targets:{{}}}};for(let i=3;i<targets.length;i++){{const target=targets[i], metric=metricFromVarJs(vars[i]);if(metric==='aud'||metric==='share'){{p.targets[target]??={{}};p.targets[target][metric]=fnumJs(row[i])}}}}programs.push(p)}}return [programs,readMatrixJs(crosstabRows,{{target:2,var:3,entity:4,start:6}},['aud','share'])]}}
function loadRankingsJs(currentBook,previousBook){{const map={{globo:'Globo',record:'Record',sbt:'SBT',band:'BAND'}}, rankings={{}};for(const [key,sheet] of Object.entries(map)){{const cur=currentBook.sheets[sheet]||[], prev=previousBook.sheets[sheet]||[], prevBy={{}};for(const row of prev.slice(3))if(row[1])prevBy[row[1]]={{aud:fnumJs(row[2]),share:fnumJs(row[3])}};rankings[key]=cur.slice(3,13).filter(r=>r[1]).map((row,idx)=>{{const aud=fnumJs(row[2]),share=fnumJs(row[3]),p=prevBy[row[1]]||{{}};const variation=(now,before)=>now===null||before===null||before===undefined||before===0?null:(now-before)/before*100;const audVar=variation(aud,p.aud), shareVar=variation(share,p.share);return {{rank:idx+1,program:row[1],aud,audPrev:p.aud,audVar,audStatus:audVar>0?'cresceu':audVar<0?'caiu':'estavel',share,sharePrev:p.share,shareVar,shareStatus:shareVar>0?'cresceu':shareVar<0?'caiu':'estavel'}}}})}}return rankings}}
function buildDataFromBooks(books){{
  const byName=name=>books.find(b=>normalizeText(b.name).includes(normalizeText(name)));
  const minuteBook=byName('Base_aud_minuto');
  const programBook=byName('Base Programa');
  const profileBook=byName('Base Perfil');
  const turnosBook=byName('Base Turnos');
  const currentRank=byName('Base Ranking diario_REC')||byName('Base Ranking diário_REC');
  const previousRank=byName('Base Ranking diario_DF7d')||byName('Base Ranking diário_DF7d');
  if(!minuteBook||!programBook||!profileBook||!turnosBook||!currentRank||!previousRank)throw new Error('Envie as 6 bases: Perfil, Programa, Ranking DF7d, Ranking REC, Turnos e aud minuto.');
  const minutesAll=loadMinuteDataJs(minuteBook.sheets.Crosstab2||[]);
  const dayMinutes=minutesAll.filter(m=>m.minute>=7*60&&m.minute<24*60);const lineMinutes=minutesAll.filter(m=>m.minute>=6*60||m.minute<6*60).sort((a,b)=>((a.minute-6*60+24*60)%(24*60))-((b.minute-6*60+24*60)%(24*60)));
  const [programs,competition]=loadProgramsJs(programBook.sheets.Programas||[],programBook.sheets.Crosstab||[]);
  const profile=readMatrixJs(profileBook.sheets.Crosstab1||[],{{target:3,var:2,entity:4,start:6}},null);
  const turnos=turnosBook.sheets.Crosstab||[];
  const row724=turnos.find(r=>r[0]&&String(r[0]).includes('07:00-24:00'))||[];
  let dateSerial='';
  for(const rows of Object.values(programBook.sheets)){{
    for(const row of rows){{
      if(row&&/^\\d{{5}}$/.test(row[0]||'')){{dateSerial=row[0];break;}}
    }}
    if(dateSerial)break;
  }}
  const dateIso=excelDateJs(dateSerial)||new Date().toISOString().slice(0,10);
  const summaryBars=DATA.channels.map(ch=>({{key:ch.key,label:ch.label,color:ch.color,aud:pctJs(avgJs(dayMinutes.map(m=>m.aud[ch.key]))),share:pctJs(avgJs(dayMinutes.map(m=>m.share[ch.key])))}}));
  const line=lineMinutes.map(m=>({{time:m.time,aud:Object.fromEntries(DATA.channels.map(ch=>[ch.key,pctJs(m.aud[ch.key])]))}}));
  const leadership=DATA.leadershipChannels.map(ch=>{{
    const mins=dayMinutes.filter(m=>{{
      const vals=DATA.leadershipChannels.map(c=>[c.key,m.aud[c.key]]).filter(x=>x[1]!==null&&x[1]!==undefined);
      return vals.length&&vals.sort((a,b)=>b[1]-a[1])[0][0]===ch.key;
    }}).length;
    return {{key:ch.key,label:ch.label,color:ch.color,minutes:mins,percent:pctJs(dayMinutes.length?mins/dayMinutes.length*100:0),hours:`${{Math.floor(mins/60)}}h${{String(mins%60).padStart(2,'0')}}`}};
  }});
  const adh=(name,target,key)=>{{const bucket=profile[name]?.[target]?.[key]||{{}};return bucket.adh??Object.values(bucket)[0]}};
  const profileData={{}};
  for(const p of programs){{
    profileData[p.name]={{}};
    for(const ch of DATA.leadershipChannels){{
      const targetKey=Object.keys(competition[p.name]||{{}}).find(t=>normalizeText(t).includes('total domic'))||'Total Domicílios';const metrics=competition[p.name]?.[targetKey]?.[ch.key]||{{}};
      profileData[p.name][ch.key]={{aud:metrics.aud,share:metrics.share,gender:{{Homem:adh(p.name,'Masculino',ch.key),Mulher:adh(p.name,'Feminino',ch.key)}},classes:{{AB1:adh(p.name,'AB1',ch.key),B2:adh(p.name,'B2',ch.key),C1:adh(p.name,'C1',ch.key),C2:adh(p.name,'C2',ch.key),DE:adh(p.name,'DE',ch.key)}},ages:{{'4-11':adh(p.name,'4-11 anos',ch.key),'12-17':adh(p.name,'12-17 anos',ch.key),'18-24':adh(p.name,'18-24 anos',ch.key),'25-34':adh(p.name,'25-34 anos',ch.key),'35-49':adh(p.name,'35-49 anos',ch.key),'50+':adh(p.name,'50+',ch.key)}}}};
    }}
  }}
  const d=new Date(dateIso+'T00:00:00');
  const weekdays=['domingo','segunda-feira','ter\u00e7a-feira','quarta-feira','quinta-feira','sexta-feira','s\u00e1bado'];
  return {{meta:{{date:dateIso,weekday:weekdays[d.getDay()],dailyAvg:fnumJs(row724[2])}},channels:DATA.channels,assets:DATA.assets,leadershipChannels:DATA.leadershipChannels,rankingChannels:DATA.rankingChannels,targets:DATA.targets,summaryBars,line,leadership,programs,programCompetition:competition,minuteAll:minutesAll,profile:profileData,rankings:loadRankingsJs(currentRank,previousRank)}};
}}
async function handleUploadedBases(fileList){{fileStatus.style.display='block';fileStatus.textContent='Lendo bases...';const books=[];for(const file of [...fileList]){{const lower=file.name.toLowerCase();if(lower.endsWith('.xlsx'))books.push(await readXlsxFile(file));else if(lower.endsWith('.zip')){{const pack=await JSZip.loadAsync(file);for(const entry of Object.values(pack.files)){{if(!entry.dir&&entry.name.toLowerCase().endsWith('.xlsx')){{const bytes=await entry.async('uint8array');books.push(await readXlsxFile(bytes,entry.name.split(/[\\/]/).pop()))}}}}}}}}DATA=buildDataFromBooks(books);fillSelect(programSelect,DATA.programs.map(p=>p.name));fillSelect(profileSelect,DATA.programs.map(p=>p.name));fillSelect(targetSelect,DATA.targets);fileStatus.textContent='Bases carregadas';revealDashboard()}}
document.querySelectorAll('.tab').forEach(btn=>btn.addEventListener('click',()=>{{document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));btn.classList.add('active');document.getElementById(btn.dataset.tab).classList.add('active');setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil()}},0)}}));
const programSelect=document.getElementById('programSelect'),targetSelect=document.getElementById('targetSelect'),profileSelect=document.getElementById('profileSelect'),programTitle=document.getElementById('programTitle'),profileTitle=document.getElementById('profileTitle'),minuteTable=document.getElementById('minuteTable'),profileGrid=document.getElementById('profileGrid'),baseUpload=document.getElementById('baseUpload'),fileStatus=document.getElementById('fileStatus');
fillSelect(programSelect,DATA.programs.map(p=>p.name));fillSelect(profileSelect,DATA.programs.map(p=>p.name));fillSelect(targetSelect,DATA.targets);
programSelect.addEventListener('change',renderProgramas);targetSelect.addEventListener('change',renderProgramas);profileSelect.addEventListener('change',renderPerfil);window.addEventListener('resize',()=>{{renderResumo();renderProgramas();renderPerfil()}});
function revealDashboard(){{document.body.className=document.body.className.replace(/\\bawaiting-bases\\b/g,'').trim();document.querySelector('header .actions').style.display='none';setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil();window.scrollTo({{top:0,behavior:'smooth'}})}},0)}}
document.getElementById('uploadBtn').addEventListener('click',()=>baseUpload.click());baseUpload.addEventListener('change',async()=>{{try{{const files=[...baseUpload.files].map(f=>f.name);fileStatus.textContent=files.length?`${{files.length}} arquivo(s): ${{files.join(', ')}}`:'';if(files.length)await handleUploadedBases(baseUpload.files)}}catch(err){{fileStatus.style.display='block';fileStatus.textContent='Erro ao ler bases: '+err.message;console.error(err)}}}});document.getElementById('exportBtn').addEventListener('click',exportHtml);document.getElementById('imageBtn').addEventListener('click',generateHighlightsImage);
renderResumo();renderProgramas();renderPerfil();
</script>
</body>
</html>"""


if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    data = build_data()
    OUT.write_text(render_html(data), encoding="utf-8")
    print(OUT)
