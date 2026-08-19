from __future__ import annotations

import base64
import html
import json
import math
import re
from collections import defaultdict
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(r"C:\Users\gamartin\Documents\Codex\2026-07-09\que")
DATA_ROOT = ROOT / "work" / "base-diaria"
OUT = ROOT / "outputs" / "index.html"
FONT_REG = ROOT / "assets" / "GlobotipoCorporativa-Regular.ttf"
FONT_BOLD = ROOT / "assets" / "GlobotipoCorporativa-Bold.ttf"
LOGO_ROOT = ROOT / "assets"
PLIMPLIM = ROOT / "assets" / "PLIMPLIM_GD.png"
PLIMPLIM_WHITE = ROOT / "assets" / "PLIMPLIM_BRANCO.png"
INTELIGENCIA = ROOT / "assets" / "inteligencia.png"
CONFIDENTIAL = ROOT / "assets" / "conf.png"
JSZIP_CANDIDATES = [
    Path(r"C:\Users\gamartin\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\jszip\dist\jszip.min.js"),
    Path(r"C:\Users\gamartin\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\.pnpm\docx@9.6.1\node_modules\jszip\dist\jszip.min.js"),
]

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


def load_profile_total(rows):
    data = defaultdict(lambda: defaultdict(dict))
    total_row = next((row for row in rows if row and row[0] == "Total"), None)
    if not total_row:
        return data
    header_rows = {"target": 3, "var": 2, "entity": 4}
    max_cols = max(len(r) for r in rows[: header_rows["entity"] + 1])
    for i in range(1, max_cols):
        target = rows[header_rows["target"]][i] if i < len(rows[header_rows["target"]]) else ""
        entity = rows[header_rows["entity"]][i] if i < len(rows[header_rows["entity"]]) else ""
        ch = channel_for(entity)
        if target and ch:
            data[target][ch["key"]]["adh"] = fnum(total_row[i] if i < len(total_row) else "")
    return data


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

            def status_from_var(value):
                if value is None or abs(value) <= 0.4:
                    return "estavel"
                return "cresceu" if value > 0 else "caiu"

            aud_var = variation(aud, prev.get("aud"))
            share_var = variation(share, prev.get("share"))
            entries.append(
                {
                    "rank": idx,
                    "program": row[1],
                    "aud": aud,
                    "audPrev": prev.get("aud"),
                    "audVar": aud_var,
                    "audStatus": status_from_var(aud_var),
                    "share": share,
                    "sharePrev": prev.get("share"),
                    "shareVar": share_var,
                    "shareStatus": status_from_var(share_var),
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

    profile_total = load_profile_total(profile_book["Crosstab1"])

    def total_adh(target, key):
        bucket = profile_total.get(target, {}).get(key, {}) or {}
        return bucket.get("adh") if "adh" in bucket else next(iter(bucket.values()), None)

    profile_data = {}
    profile_data["07h-24h"] = {}
    for ch in LEADERSHIP_CHANNELS:
        metrics = next((row for row in summary_bars if row["key"] == ch["key"]), {})
        profile_data["07h-24h"][ch["key"]] = {
            "aud": metrics.get("aud"),
            "share": metrics.get("share"),
            "gender": {
                "Homem": total_adh("Masculino", ch["key"]),
                "Mulher": total_adh("Feminino", ch["key"]),
            },
            "classes": {
                "AB1": total_adh("AB1", ch["key"]),
                "B2": total_adh("B2", ch["key"]),
                "C1": total_adh("C1", ch["key"]),
                "C2": total_adh("C2", ch["key"]),
                "DE": total_adh("DE", ch["key"]),
            },
            "ages": {
                "4-11": total_adh("4-11 anos", ch["key"]),
                "12-17": total_adh("12-17 anos", ch["key"]),
                "18-24": total_adh("18-24 anos", ch["key"]),
                "25-34": total_adh("25-34 anos", ch["key"]),
                "35-49": total_adh("35-49 anos", ch["key"]),
                "50+": total_adh("50+", ch["key"]),
            },
        }
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


def image_data(path: Path, max_px: int | None = None) -> str:
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
    raw = path.read_bytes()
    if max_px:
        try:
            from PIL import Image

            img = Image.open(BytesIO(raw)).convert("RGBA")
            if max(img.size) > max_px:
                resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
                img.thumbnail((max_px, max_px), resampling)
            out = BytesIO()
            img.save(out, format="PNG", optimize=True)
            raw = out.getvalue()
        except Exception:
            pass
    suffix = path.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in ("jpg", "jpeg") else suffix
    return f"data:image/{mime};base64,{base64.b64encode(raw).decode('ascii')}"


def render_html(data: dict) -> str:
    for ch in data["channels"]:
        if ch.get("logo"):
            ch["logoData"] = image_data(LOGO_ROOT / ch["logo"], 180)
    data["assets"] = {
        "plim": image_data(PLIMPLIM, 220),
        "plimWhite": image_data(PLIMPLIM_WHITE, 220),
        "inteligencia": image_data(INTELIGENCIA, 900),
        "conf": image_data(CONFIDENTIAL, 180),
    }
    plim_icon = data["assets"]["plim"]
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    font_reg = font_data(FONT_REG)
    font_bold = font_data(FONT_BOLD)
    jszip_path = next((path for path in JSZIP_CANDIDATES if path.exists()), None)
    if jszip_path is None:
        raise FileNotFoundError("jszip.min.js not found in expected runtime paths")
    jszip_code = jszip_path.read_text(encoding="utf-8").replace("</script", "<\\/script")
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
*{{box-sizing:border-box}} body{{margin:0;font-family:Globotipo,Arial,sans-serif;background:linear-gradient(180deg,#f8fafc 0%,var(--bg) 44%,#f7f8fa 100%);color:var(--ink);letter-spacing:0}}body.awaiting-bases main,body.awaiting-bases .tabs{{display:none}}body.awaiting-bases header.app{{min-height:100vh;display:grid;place-items:center;padding:24px}}body.awaiting-bases .top{{display:flex;flex-direction:column;justify-content:center;align-items:center;max-width:760px;width:min(92vw,760px);min-height:520px;text-align:center;padding:56px 42px}}body.awaiting-bases .brand{{flex-direction:column;gap:20px;font-size:58px}}body.awaiting-bases .brand-logo{{width:130px;height:130px}}body.awaiting-bases .header-right{{justify-items:center;width:100%}}body.awaiting-bases .actions{{min-width:0;width:auto;justify-content:center;margin-top:6px;background:transparent;padding:0;border-radius:0}}body.awaiting-bases .action-btn.primary{{background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;border-color:transparent;box-shadow:0 12px 28px rgba(0,92,239,.22)}}body.awaiting-bases .file-status{{display:none}}body:not(.awaiting-bases) .version-info{{display:none}}.version-info{{margin-top:12px;color:#66758c;font-size:12px;font-weight:700;letter-spacing:0;text-align:center}}
header.app{{background:linear-gradient(90deg,rgba(0,92,239,.04) 1px,transparent 1px),linear-gradient(180deg,rgba(0,92,239,.04) 1px,transparent 1px),#eef2f7;background-size:72px 72px;position:relative;z-index:20;padding:18px 20px 0}}
.top{{max-width:1440px;margin:auto;background:#fff;border:1px solid var(--line);border-top:8px solid transparent;border-image:linear-gradient(90deg,#0046ff,#00a3ff) 1;border-radius:8px;padding:28px 38px;display:grid;grid-template-columns:1fr auto;gap:22px;align-items:center;box-shadow:0 10px 26px rgba(21,36,66,.08)}}
.brand{{display:flex;align-items:center;gap:22px;font-weight:700;font-size:54px;line-height:1;color:var(--blue);letter-spacing:0;text-transform:uppercase}}.brand span{{background:linear-gradient(90deg,#0046ff 0%,#005cef 48%,#00a3ff 100%);-webkit-background-clip:text;background-clip:text;color:transparent}}.brand-logo{{width:78px;height:78px;object-fit:contain;filter:none}}.header-right{{display:grid;gap:12px;justify-items:end}}.actions{{display:flex;gap:10px;align-items:center;flex-wrap:wrap;justify-content:flex-end;background:linear-gradient(135deg,#004cff,#05a9f5);border-radius:8px;padding:18px 20px;min-width:360px}}.action-btn{{border:1px solid rgba(255,255,255,.45);background:#fff;color:#005cef;border-radius:8px;padding:10px 12px;font:inherit;font-weight:700;cursor:pointer}}.action-btn.primary{{background:rgba(255,255,255,.16);color:#fff;border-color:rgba(255,255,255,.45)}}.file-status{{font-size:12px;color:#fff;max-width:330px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;grid-column:1/-1;justify-self:start}}
.tabs{{display:flex;gap:6px;background:#e8edf5;padding:4px;border-radius:8px}}.tab{{border:0;background:transparent;padding:10px 16px;border-radius:6px;font-family:Globotipo,Arial,sans-serif;font-weight:700;cursor:pointer;color:#40506a;text-decoration:none;display:inline-flex;align-items:center;justify-content:center}}.tab.active{{background:linear-gradient(135deg,var(--blue),#3e8bff);color:#fff}}
main{{max-width:1440px;margin:0 auto;padding:24px;display:grid;gap:18px}}.panel{{display:none}}.panel.active{{display:grid;gap:18px}}
.mobile-summary,.tablet-summary{{display:none}}
.hero{{background:linear-gradient(135deg,#0046c8 0%,#005cef 48%,#45a3ff 100%);color:#fff;padding:26px;border-radius:8px;display:grid;grid-template-columns:1fr auto;gap:20px;align-items:center;box-shadow:0 18px 45px rgba(0,92,239,.22)}}.hero-head{{display:flex;align-items:center;gap:18px;flex-wrap:wrap}}.hero h1{{margin:0;font-size:34px;line-height:1.05}}.date-pill{{display:inline-grid;gap:2px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.35);border-radius:8px;padding:12px 16px;min-width:190px}}.date-pill strong{{font-size:24px}}.date-pill span{{font-size:13px;text-transform:uppercase;letter-spacing:.08em}}.kpi{{background:#fff;color:var(--ink);border-radius:8px;padding:14px 20px 14px 14px;min-width:270px;box-shadow:0 12px 28px rgba(0,0,0,.12);display:flex;align-items:center;gap:14px}}.kpi-logo{{width:54px;height:54px;border-radius:50%;display:grid;place-items:center;flex:0 0 54px;background:linear-gradient(135deg,#005cef,#3e8bff);box-shadow:inset 0 0 0 1px rgba(255,255,255,.3)}}.kpi-logo img{{display:block;width:72%;height:72%;object-fit:contain}}.kpi-text{{display:grid;gap:2px}}.kpi small{{display:block;color:var(--muted);font-size:12px;text-transform:uppercase}}.kpi strong{{font-size:42px;line-height:1}}
.toolbar{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px;display:grid;grid-template-columns:260px 260px 1fr;gap:16px;align-items:center}}.toolbar.one{{grid-template-columns:280px 1fr}}label{{font-size:12px;color:var(--muted);font-weight:700;text-transform:uppercase;display:grid;gap:6px}}select{{font:inherit;border:1px solid var(--line);border-radius:6px;padding:10px 12px;background:#fff;min-width:0}}.selected-title{{text-align:center;font-weight:700;font-size:26px;color:var(--blue);background:var(--soft);border-radius:8px;padding:14px 18px}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}.card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:18px;min-width:0;box-shadow:0 10px 28px rgba(21,36,66,.06)}}.card h2{{margin:0 0 14px;font-size:18px}}.card h3{{margin:0 0 12px;font-size:16px}}.insights-card{{background:linear-gradient(180deg,#fff,#f8fbff);border-left:6px solid var(--blue)}}.insights-card .section-head{{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:14px}}.insight-tabs{{display:flex;gap:7px;flex-wrap:wrap;justify-content:flex-end}}.insight-tab{{border:1px solid var(--line);background:#fff;color:#40506a;border-radius:999px;padding:7px 10px;font:inherit;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:6px;box-shadow:0 4px 12px rgba(21,36,66,.05)}}.insight-tab.active{{color:#fff;border-color:transparent;box-shadow:0 8px 18px rgba(21,36,66,.14)}}.insight-tab .logo-chip{{width:20px;height:20px;font-size:7px}}.insight-tab.active .logo-chip{{box-shadow:inset 0 0 0 1px rgba(255,255,255,.5)}}.insight-context{{display:grid;gap:12px}}.insight-grid{{display:grid;grid-template-columns:1.2fr repeat(3,1fr);gap:14px;align-items:stretch}}.insight-main{{padding:16px;border-radius:8px;background:#eef5ff;color:#18345f;font-size:17px;line-height:1.35;font-weight:700}}.insight-item{{padding:14px;border:1px solid var(--line);border-radius:8px;background:#fff;display:grid;gap:7px}}.insight-item small{{color:var(--muted);text-transform:uppercase;font-size:11px;font-weight:700}}.insight-item strong{{font-size:22px;color:var(--ink)}}.insight-item span{{color:#526178;font-size:13px;line-height:1.3}}.insight-list{{margin:14px 0 0;padding-left:18px;color:#3e4d63;line-height:1.45}}.insight-list li{{margin:6px 0}}.chart{{width:100%;height:290px;display:block}}.line-chart{{height:380px}}.legend{{display:flex;gap:12px;flex-wrap:wrap;margin-top:10px;font-size:12px;color:#3e4d63}}.legend span,.channel-name{{display:inline-flex;align-items:center;gap:6px}}.dot{{width:10px;height:10px;border-radius:50%;display:inline-block}}.logo-chip{{width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:9px;line-height:1;overflow:hidden;box-shadow:inset 0 0 0 1px rgba(255,255,255,.28);text-align:center}}.logo-chip img{{display:block;width:72%;height:72%;object-fit:contain;object-position:center;margin:auto}}.leader .logo-chip{{color:#fff}}
.leader-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}}.leader{{border:1px solid var(--line);border-radius:8px;padding:14px;background:linear-gradient(180deg,#fff,#f8fafc)}}.leader strong{{font-size:28px;display:block}}.leader span{{color:var(--muted);font-size:12px}}.leader-head{{display:flex;align-items:center;gap:8px;margin-bottom:8px}}
.rank-grid{{display:grid;grid-template-columns:repeat(2, minmax(0,1fr));gap:18px}}table{{width:100%;border-collapse:collapse;font-size:13px}}th,td{{padding:10px 8px;border-bottom:1px solid #edf1f7;text-align:left;vertical-align:middle}}th{{color:var(--muted);font-size:11px;text-transform:uppercase;background:#fafbfe}}td.num,th.num{{text-align:right}}.arrow{{font-size:18px;font-weight:700}}.up{{color:var(--green)}}.down{{color:var(--red)}}.empty{{color:var(--muted);padding:24px;text-align:center;background:#f8fafc;border-radius:8px}}
.profile-grid{{display:grid;grid-template-columns:1fr;gap:18px}}.profile-card{{padding:22px}}.profile-head{{display:flex;justify-content:space-between;gap:16px;align-items:center;margin-bottom:18px}}.profile-title{{display:flex;align-items:center;gap:10px}}.metric-pair{{display:flex;gap:10px}}.mini{{border:0;border-radius:8px;padding:10px 12px;min-width:96px;text-align:right;background:var(--metric-grad,linear-gradient(135deg,#005cef,#45a3ff));color:#fff;box-shadow:0 10px 24px rgba(21,36,66,.16)}}.mini small{{display:block;color:rgba(255,255,255,.82);font-size:11px}}.mini strong{{font-size:26px;color:#fff}}.mini-chart{{height:150px;width:100%;display:block}}.profile-block{{display:grid;grid-template-columns:minmax(230px,.9fr) minmax(360px,1.15fr) minmax(460px,1.45fr);gap:0;align-items:start}}.profile-block>div{{min-width:0;padding:0 22px}}.profile-block>div+div{{border-left:1px solid #e5ebf3}}.profile-block>div:nth-child(2){{display:grid;justify-items:center}}.profile-block>div:nth-child(2) h4{{justify-self:start;width:100%}}.profile-block h4{{margin:0 0 8px;font-size:12px;color:var(--muted);text-transform:uppercase}}.program-table-wrap{{max-height:470px;overflow:auto;border:1px solid var(--line);border-radius:8px}}.program-table-wrap table th{{position:sticky;top:0;z-index:1}}.minute-table th,.minute-table td{{white-space:nowrap}}.footer-note{{display:grid;gap:10px;color:#66758c;font-size:13px;padding:0 0 18px}}.source-note{{text-align:left}}.credit-note{{text-align:center}}.final-actions{{display:flex;justify-content:center;gap:12px;flex-wrap:wrap;padding:8px 0 12px}}.final-actions .action-btn{{border-color:#c9d4e5;box-shadow:0 8px 20px rgba(21,36,66,.08)}}.final-actions .action-btn.primary{{background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;border-color:transparent}}.highlights-editor{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px;display:grid;gap:10px;max-width:860px;width:100%;justify-self:center;box-shadow:0 10px 28px rgba(21,36,66,.06)}}.highlights-editor h3{{margin:0;color:var(--blue);font-size:16px}}.highlight-tools{{display:flex;gap:8px;flex-wrap:wrap;align-items:center}}.format-btn,.emoji-btn{{border:1px solid #c9d4e5;background:#fff;color:#21324a;border-radius:6px;padding:7px 10px;font:inherit;font-weight:700;cursor:pointer;min-width:36px}}.emoji-btn{{font-size:16px;line-height:1}}.highlight-input{{min-height:92px;border:1px solid var(--line);border-radius:8px;padding:12px;background:#fafbfe;outline:none;line-height:1.45}}.highlight-input:empty:before{{content:attr(data-placeholder);color:#8a97aa}}
@media(max-width:1024px) and (min-width:641px),(max-width:1366px) and (min-width:641px) and (pointer:coarse),(max-width:1366px) and (min-width:641px) and (hover:none){{body:not(.awaiting-bases) .tabs,body:not(.awaiting-bases) main>section.panel,body:not(.awaiting-bases) main>.footer-note,body:not(.awaiting-bases) .final-actions,body:not(.awaiting-bases) .highlights-editor{{display:none!important}}body:not(.awaiting-bases) .top{{grid-template-columns:1fr}}body:not(.awaiting-bases) .header-right{{display:none}}main{{padding:18px;max-width:980px}}.tablet-summary{{display:grid;gap:16px}}.tablet-alert{{background:linear-gradient(135deg,#0046c8,#1c8cff);color:#fff;border-radius:8px;padding:14px 18px;font-weight:700}}.tablet-summary .mobile-hero{{background:linear-gradient(135deg,#0046c8 0%,#005cef 54%,#45a3ff 100%);color:#fff;border-radius:8px;padding:22px;display:grid;gap:6px}}.tablet-summary .mobile-hero small{{opacity:.88;font-weight:700;text-transform:uppercase}}.tablet-summary .mobile-hero h1{{font-size:32px;margin:0;line-height:1.05}}.tablet-card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:16px;box-shadow:0 8px 22px rgba(21,36,66,.06);display:grid;gap:12px}}.tablet-card h2{{margin:0;color:#17335c;font-size:20px}}.tablet-card .section-head{{display:block;margin-bottom:10px}}.tablet-card .insight-tabs{{display:none}}.tablet-card .insight-grid{{grid-template-columns:1fr!important;gap:12px!important}}.tablet-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}}.tablet-kpi{{border:1px solid #e2e8f2;border-left:5px solid var(--blue);border-radius:8px;padding:12px;background:#f8fbff;display:grid;gap:4px}}.tablet-kpi small{{color:var(--muted);font-weight:700;text-transform:uppercase}}.tablet-kpi strong{{font-size:30px}}.tablet-leaders{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}}.tablet-leaders .mobile-leader{{display:grid;grid-template-columns:auto 1fr auto;gap:8px;align-items:center;border:1px solid #e7edf6;border-radius:8px;padding:9px;background:#fbfcff}}.tablet-leaders .mobile-leader strong{{font-size:18px}}.tablet-leaders .mobile-leader span{{color:var(--muted);font-size:12px}}.tablet-leaders .logo-chip{{color:#fff}}.tablet-program-list{{display:grid;gap:10px}}.tablet-program{{border:1px solid #dbe4f0;border-radius:8px;background:#fff;overflow:hidden}}.tablet-program summary{{cursor:pointer;list-style:none;padding:13px 14px;font-weight:700;color:#17335c;background:#f6f9fd;display:flex;justify-content:space-between;gap:12px;align-items:center}}.tablet-program summary::-webkit-details-marker{{display:none}}.tablet-program summary:after{{content:'+';font-size:22px;color:var(--blue)}}.tablet-program[open] summary:after{{content:'-'}}.tablet-program-body{{padding:12px;display:grid;gap:12px}}.tablet-program-table{{width:100%;border-collapse:collapse;font-size:13px}}.tablet-program-table th,.tablet-program-table td{{padding:8px;border-bottom:1px solid #edf1f7}}.tablet-program-table th{{background:#f8fafc;color:#607086;text-transform:uppercase;font-size:11px}}.tablet-profile-stack{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}}.tablet-program .mobile-profile{{display:grid;grid-template-columns:auto 1fr;gap:10px;align-items:start;border-radius:8px;padding:11px;border:1px solid #e2e8f2;width:100%}}.tablet-program .mobile-profile-globo{{background:#eef6ff;border-color:#cfe2ff}}.tablet-program .mobile-profile-nic{{background:#f7efe8;border-color:#ead4c1}}.tablet-program .mobile-profile .logo-chip{{width:30px;height:30px;color:#fff}}.tablet-program .mobile-profile-text{{display:grid;gap:4px;min-width:0}}.tablet-program .mobile-profile b{{font-size:13px;color:#17335c}}.tablet-program .mobile-profile span{{text-align:left!important;font-size:12px;line-height:1.3;font-weight:700;color:var(--muted)}}.tablet-ranking{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}}.tablet-ranking .mobile-rank-block{{border-top:4px solid var(--blue);padding-top:10px}}.tablet-ranking .mobile-rank-title{{display:flex;align-items:center;gap:8px;font-weight:700;margin-bottom:6px}}.tablet-ranking .mobile-rank-row{{display:grid;grid-template-columns:24px 1fr auto;gap:8px;padding:8px 0;border-bottom:1px solid #edf1f7;align-items:center;font-size:13px}}.tablet-ranking .mobile-rank-row strong{{font-size:13px;overflow-wrap:anywhere}}.tablet-ranking .mobile-rank-row span{{color:var(--muted);font-size:12px;white-space:nowrap}}.tablet-source{{font-size:12px;color:#66758c;line-height:1.35}}}}
@media(max-width:980px){{.top,.hero,.toolbar,.toolbar.one,.grid2,.rank-grid,.profile-grid,.profile-block{{grid-template-columns:1fr}}.insight-grid{{grid-template-columns:1fr!important}}.brand{{font-size:34px}}.brand-logo{{width:54px;height:54px}}body.awaiting-bases .brand{{font-size:40px}}body.awaiting-bases .brand-logo{{width:108px;height:108px}}.header-right{{justify-items:stretch}}body.awaiting-bases .header-right{{justify-items:center}}.actions{{min-width:0}}.leader-grid{{grid-template-columns:repeat(2,1fr)}}.selected-title{{text-align:left}}.profile-block>div{{padding:0}}.profile-block>div+div{{border-left:0;border-top:1px solid #e5ebf3;padding-top:16px}}}}
@media(max-width:640px){{header.app{{padding:10px 10px 0;background-size:44px 44px}}main{{padding:12px;gap:12px;max-width:100vw;overflow-x:hidden}}.top{{padding:18px 14px;gap:14px;border-top-width:6px;max-width:100%;overflow:hidden}}.brand{{font-size:27px;gap:12px}}.brand-logo{{width:42px;height:42px}}body.awaiting-bases header.app{{padding:16px}}body.awaiting-bases .top{{min-height:70vh;padding:30px 18px}}body.awaiting-bases .brand{{font-size:34px}}body.awaiting-bases .brand-logo{{width:96px;height:96px}}.header-right{{gap:10px;width:100%;min-width:0}}.actions{{padding:0;background:transparent;justify-content:center}}body:not(.awaiting-bases) .tabs,body:not(.awaiting-bases) main>section.panel,body:not(.awaiting-bases) main>.footer-note,body:not(.awaiting-bases) .final-actions,body:not(.awaiting-bases) .highlights-editor{{display:none!important}}body:not(.awaiting-bases) .top{{grid-template-columns:1fr}}body:not(.awaiting-bases) .header-right{{display:none}}.tablet-summary{{display:none!important}}.mobile-summary{{display:grid;gap:12px}}.mobile-alert{{background:linear-gradient(135deg,#0046c8,#1c8cff);color:#fff;border-radius:8px;padding:14px 16px;font-weight:700;line-height:1.3;box-shadow:0 10px 24px rgba(0,92,239,.18)}}.mobile-hero{{background:linear-gradient(135deg,#0046c8 0%,#005cef 54%,#45a3ff 100%);color:#fff;border-radius:8px;padding:18px;display:grid;gap:6px}}.mobile-hero small{{opacity:.88;font-weight:700;text-transform:uppercase}}.mobile-hero h1{{font-size:25px;margin:0;line-height:1.05}}.mobile-card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px;box-shadow:0 8px 20px rgba(21,36,66,.06);display:grid;gap:10px;min-width:0}}.mobile-card h2{{font-size:17px;margin:0;color:#17335c}}.mobile-card .section-head{{display:block;margin-bottom:10px}}.mobile-card .insight-tabs{{display:none}}.mobile-card .insight-grid{{grid-template-columns:1fr!important;gap:10px!important}}.mobile-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}.mobile-kpi{{border:1px solid #e2e8f2;border-left:5px solid var(--blue);border-radius:8px;padding:10px;background:#f8fbff;display:grid;gap:4px}}.mobile-kpi small{{color:var(--muted);font-weight:700;text-transform:uppercase;font-size:10px}}.mobile-kpi strong{{font-size:24px;line-height:1}}.mobile-leaders{{display:grid;gap:8px}}.mobile-leader{{display:grid;grid-template-columns:auto 1fr auto;gap:8px;align-items:center;border:1px solid #e7edf6;border-radius:8px;padding:8px;background:#fbfcff}}.mobile-leader strong{{font-size:18px}}.mobile-leader span{{color:var(--muted);font-size:12px}}.mobile-leader .logo-chip{{color:#fff}}.mobile-ranking{{display:grid;gap:12px}}.mobile-rank-block{{border-top:4px solid var(--blue);padding-top:10px}}.mobile-rank-title{{display:flex;align-items:center;gap:8px;font-weight:700;margin-bottom:6px}}.mobile-rank-row{{display:grid;grid-template-columns:24px 1fr auto;gap:8px;padding:7px 0;border-bottom:1px solid #edf1f7;align-items:center;font-size:13px}}.mobile-rank-row strong{{font-size:13px;overflow-wrap:anywhere}}.mobile-rank-row span{{color:var(--muted);font-size:12px}}.mobile-programs{{display:grid;gap:9px}}.mobile-program{{display:grid;grid-template-columns:1fr;gap:10px;align-items:start;border:1px solid #e7edf6;border-radius:8px;padding:10px;background:#fbfcff}}.mobile-program-main{{display:grid;gap:9px;min-width:0;width:100%}}.mobile-program-main>strong{{overflow-wrap:anywhere;color:#101521;font-size:14px}}.mobile-program-profiles{{display:grid;gap:7px;width:100%}}.mobile-profile{{display:grid;grid-template-columns:auto 1fr;gap:9px;align-items:start;border-radius:7px;padding:9px;border:1px solid #e2e8f2;width:100%}}.mobile-profile-globo{{background:#eef6ff;border-color:#cfe2ff}}.mobile-profile-nic{{background:#f7efe8;border-color:#ead4c1}}.mobile-profile .logo-chip{{width:26px;height:26px;color:#fff}}.mobile-profile-text{{display:grid;gap:3px;min-width:0}}.mobile-profile b{{font-size:12px;color:#17335c}}.mobile-profile span{{text-align:left!important;font-size:11px;line-height:1.28;font-weight:700;color:var(--muted)}}.mobile-program-metrics{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}.mobile-program-metric{{border-radius:7px;padding:8px 10px;color:#fff;background:linear-gradient(135deg,#005cef,#45a3ff);font-weight:700;line-height:1.1}}.mobile-program-metric.share{{background:linear-gradient(135deg,#17335c,#607086)}}.mobile-program-metric small{{display:block;font-size:10px;opacity:.88;margin-bottom:2px}}.mobile-program-metric strong{{font-size:18px}}.mobile-source{{font-size:12px;color:#66758c;line-height:1.35;padding:6px 2px 14px}}.tab{{padding:11px 6px;white-space:nowrap;font-size:12px;min-width:0;touch-action:manipulation}}}}
</style>
</head>
<body class="awaiting-bases">
<header class="app"><div class="top"><div class="brand"><img class="brand-logo" id="plimLogo" src="{plim_icon}" alt="Plim plim"><span>Desempenho Diário</span></div><div class="header-right"><div class="actions"><input id="baseUpload" type="file" accept=".zip,.xlsx" multiple hidden><button class="action-btn primary" id="uploadBtn" type="button">Enviar base</button><span class="file-status" id="fileStatus"></span></div><div class="version-info">Versão v1.4 &bull; Última atualização: 19/08/2026</div><nav class="tabs"><button class="tab active" type="button" data-tab="resumo" onclick="activateTab(this)">RESUMO</button><button class="tab" type="button" data-tab="programas" onclick="activateTab(this)">PROGRAMAS</button><button class="tab" type="button" data-tab="perfil" onclick="activateTab(this)">PERFIL</button></nav></div></div></header>
<main>
<section id="mobileSummary" class="mobile-summary"></section>
<section id="tabletSummary" class="tablet-summary"></section>
<section id="resumo" class="panel active">
  <div class="hero"><div class="hero-head"><span id="dateLabel"></span><h1>Resumo do Distrito Federal</h1></div><div class="kpi"><span class="kpi-logo"><img id="globoKpiLogo" alt="Globo"></span><span class="kpi-text"><small>Audi&ecirc;ncia m&eacute;dia 07h-24h</small><strong id="dailyAvg"></strong></span></div></div>
  <div class="card insights-card"><div id="dailyInsights"></div></div>
  <div class="grid2"><div class="card"><h2>Audiência 07:00 às 24:00</h2><svg id="audBars" class="chart"></svg></div><div class="card"><h2>Share% 07:00 às 24:00</h2><svg id="shareBars" class="chart"></svg></div></div>
  <div class="card"><h2>Minuto a minuto do dia</h2><svg id="dayLine" class="chart line-chart"></svg><div id="legend3" class="legend"></div></div>
  <div class="card"><h2>Minutos na liderança</h2><div id="leadership" class="leader-grid"></div></div>
  <div class="card"><h2>Minutos na liderança (exceto NIC)</h2><div id="leadershipNoNic" class="leader-grid"></div></div>
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
<div class="highlights-editor" id="highlightsEditor"><h3>Destaques</h3><div class="highlight-tools"><button class="format-btn" type="button" data-command="bold">B</button><button class="format-btn" type="button" data-command="italic"><em>I</em></button><button class="emoji-btn" type="button" data-insert="🟢 " aria-label="Destaque positivo">🟢</button><button class="emoji-btn" type="button" data-insert="🔴 " aria-label="Destaque negativo">🔴</button></div><div id="highlightText" class="highlight-input" contenteditable="true" data-placeholder="Digite os destaques que devem aparecer na imagem gerada."></div></div>
<div class="footer-note"><span class="credit-note">Modelo desenvolvido pela área de Programação da TV Globo DF.</span></div>
</main>
<script>{jszip_code}</script>
<script id="data" type="application/json">{data_json.replace("</", "<\\/")}</script>
<script>
/* Fallback ES5 para celulares/visualizadores que não executam o script moderno. */
(function(){{
  function getData(){{try{{return JSON.parse(document.getElementById('data').textContent||'{{}}')}}catch(e){{return {{}}}}}}
  function escLegacy(value){{return String(value==null?'':value).replace(/[&<>"']/g,function(m){{return {{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]}})}}
  function fmtLegacy(value,d){{if(d==null)d=2;if(value===null||value===undefined||value===''||isNaN(Number(value)))return '-';try{{return Number(value).toLocaleString('pt-BR',{{minimumFractionDigits:d,maximumFractionDigits:d}})}}catch(e){{return String(Math.round(Number(value)*100)/100).replace('.',',')}}}}
  function logoLegacy(c){{if(!c)c={{}};var color=c.color||'#005cef';var label=c.textLogo||c.label||'';if(c.logoData)return '<span class="logo-chip" style="background:'+color+'"><img src="'+c.logoData+'" alt="'+escLegacy(c.label||'')+'"></span>';return '<span class="logo-chip" style="background:'+color+'">'+escLegacy(label)+'</span>'}}
  function findByKey(list,key){{list=list||[];for(var i=0;i<list.length;i++)if(list[i]&&list[i].key===key)return list[i];return null}}
  function channelList(data){{var list=data.leadershipChannels||data.channels||[];var out=[];for(var i=0;i<list.length;i++)if(list[i]&&list[i].key!=='tle')out.push(list[i]);return out}}
  function legacyInsightCard(title,value,detail){{return '<div class="insight-item"><small>'+escLegacy(title)+'</small><strong>'+escLegacy(value)+'</strong><span>'+escLegacy(detail||'')+'</span></div>'}}
  window.activateTab=function(btn){{if(!btn||!btn.getAttribute('data-tab'))return;var tabs=document.querySelectorAll('.tabs .tab');for(var i=0;i<tabs.length;i++)tabs[i].classList.remove('active');var panels=document.querySelectorAll('.panel');for(var j=0;j<panels.length;j++)panels[j].classList.remove('active');btn.classList.add('active');var panel=document.getElementById(btn.getAttribute('data-tab'));if(panel)panel.classList.add('active');if(typeof window.renderResumo==='function')setTimeout(function(){{window.renderResumo();window.renderProgramas&&window.renderProgramas();window.renderPerfil&&window.renderPerfil()}},0)}};
  window.renderDailyInsights=function(selectedKey){{
    var data=getData();var target=document.getElementById('dailyInsights');if(!target)return;
    var channels=channelList(data);var selected=findByKey(channels,selectedKey)||findByKey(channels,'globo')||channels[0]||{{key:'globo',label:'Globo',color:'#005cef'}};
    window.selectedInsightKey=selected.key;
    var tabs='';for(var i=0;i<channels.length;i++){{var c=channels[i];var active=c.key===selected.key;var style=active?'background:linear-gradient(135deg,'+(c.color||'#005cef')+',#6fb1ff);color:#fff;border-color:transparent':'';tabs+='<button class="insight-tab '+(active?'active':'')+'" type="button" data-insight="'+escLegacy(c.key)+'" onclick="renderDailyInsights(\\''+escLegacy(c.key)+'\\')" style="'+style+'">'+logoLegacy(c)+escLegacy(c.label)+'</button>'}}
    var summary=findByKey(data.summaryBars,selected.key)||{{}};var leader=findByKey(data.leadership,selected.key)||{{minutes:0,percent:0,hours:'0h00'}};var rankings=(data.rankings&&data.rankings[selected.key])||[];
    var best=null;for(var r=0;r<rankings.length;r++){{if(!best||Number(rankings[r].aud||0)>Number(best.aud||0))best=rankings[r]}}
    var headline=(selected.label||'Emissora')+' fechou 07h-24h com '+fmtLegacy(summary.aud,2)+' pontos e '+fmtLegacy(summary.share,2)+'% de share.';
    var cards='';
    cards+=legacyInsightCard('Audiência',fmtLegacy(summary.aud,2),'Média 07h-24h');
    cards+=legacyInsightCard('Share',fmtLegacy(summary.share,2)+'%','Média 07h-24h');
    cards+=legacyInsightCard('Liderança',String(leader.minutes||0)+' min',fmtLegacy(leader.percent,1)+'% | '+(leader.hours||'0h00'));
    if(best)cards+=legacyInsightCard('Melhor programa',best.program||'-',fmtLegacy(best.aud,2)+' pontos | '+fmtLegacy(best.share,2)+'% share');
    target.innerHTML='<div class="section-head"><h2>Destaques do dia</h2><div class="insight-tabs">'+tabs+'</div></div><div class="insight-context"><div class="insight-grid" style="grid-template-columns:1.2fr repeat(4,1fr)"><div class="insight-main" style="background:linear-gradient(135deg,'+(selected.color||'#005cef')+',#6fb1ff);color:#fff">'+escLegacy(headline)+'</div>'+cards+'</div></div>';
  }};
  document.addEventListener('click',function(e){{var t=e.target;if(!t)return;if(t.closest){{var tab=t.closest('.tabs .tab');if(tab){{window.activateTab(tab);return}}var insight=t.closest('.insight-tab');if(insight)window.renderDailyInsights(insight.getAttribute('data-insight'))}}}});
  setTimeout(function(){{var box=document.getElementById('dailyInsights');if(box&&!box.innerHTML)window.renderDailyInsights('globo')}},0);
}})();
</script>
<script>
let DATA = JSON.parse(document.getElementById('data').textContent);
const fmt = (v,d=2) => v === null || v === undefined || Number.isNaN(v) ? '-' : Number(v).toLocaleString('pt-BR',{{minimumFractionDigits:d, maximumFractionDigits:d}});
const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m]));
document.getElementById('plimLogo').src = DATA.assets.plim;
const channelMap = Object.fromEntries(DATA.channels.map(c=>[c.key,c]));
const visible = DATA.channels.filter(c=>c.key!=='tle');
const programTableChannels = ['globo','nic','record','sbt','band','paytv'].map(key=>DATA.channels.find(c=>c.key===key)).filter(Boolean);
const audChannels = DATA.channels;
const shareChannels = visible;
const lineChannels = DATA.channels;
const leadChannels = DATA.leadershipChannels;
const rankingChannels = DATA.rankingChannels;
function leadershipExceptNic(){{
  const channels=leadChannels.filter(c=>c.key!=='nic'&&c.key!=='tle');
  const rows=(DATA.minuteAll||[]).filter(m=>m.minute>=7*60&&m.minute<24*60);
  return channels.map(ch=>{{
    const mins=rows.filter(m=>{{
      const vals=channels.map(c=>[c.key,m.aud?.[c.key]]).filter(x=>x[1]!==null&&x[1]!==undefined);
      return vals.length&&vals.sort((a,b)=>b[1]-a[1])[0][0]===ch.key;
    }}).length;
    return {{key:ch.key,label:ch.label,color:ch.color,minutes:mins,percent:rows.length?mins/rows.length*100:0,hours:`${{Math.floor(mins/60)}}h${{String(mins%60).padStart(2,'0')}}`}};
  }});
}}
function leadershipRowsHtml(rows){{return (rows||[]).map(l=>{{const c=channelMap[l.key]||l;return `<div class="mobile-leader">${{logoHtml(c)}}<div><strong>${{c.label}}</strong><br><span>${{fmt(l.percent,1)}}% do período</span></div><strong>${{l.minutes}} min</strong></div>`}}).join('')}}
function leaderCardsHtml(rows){{return (rows||[]).map(l=>{{const c=channelMap[l.key]||l;return `<div class="leader" style="border-top:5px solid ${{l.color}}"><div class="leader-head">${{logoHtml(c)}}<span>${{l.label}}</span></div><strong>${{l.minutes}}</strong><span>${{fmt(l.percent,1)}}% | ${{l.hours}}</span></div>`}}).join('')}}function logoHtml(c){{return c.logoData ? `<span class="logo-chip" style="background:${{c.color}}"><img src="${{c.logoData}}" alt="${{c.label}}"></span>` : `<span class="logo-chip" style="background:${{c.color}}">${{c.textLogo || c.label.slice(0,3).toUpperCase()}}</span>`}}
function legend(el, channels=visible){{document.getElementById(el).innerHTML=channels.map(c=>`<span>${{logoHtml(c)}}${{c.label}}</span>`).join('')}}
function svgEl(tag, attrs={{}}){{const e=document.createElementNS('http://www.w3.org/2000/svg',tag);Object.entries(attrs).forEach(([k,v])=>e.setAttribute(k,v));return e}}
function clear(svg){{while(svg.firstChild)svg.removeChild(svg.firstChild)}}
function svgLogo(svg,c,cx,cy,size=24){{svg.appendChild(svgEl('circle',{{cx,cy,r:size/2,fill:c.color}}));if(c.logoData){{const inner=size*.72;const im=svgEl('image',{{href:c.logoData,x:cx-inner/2,y:cy-inner/2,width:inner,height:inner,preserveAspectRatio:'xMidYMid meet'}});svg.appendChild(im)}}else{{svg.appendChild(svgEl('text',{{x:cx,y:cy+3,'text-anchor':'middle','font-size':8,'font-weight':700,fill:'#fff'}})).textContent=c.textLogo||c.label.slice(0,3).toUpperCase()}}}}
function barChart(id, data, metric, channels=visible){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||700,h=svg.clientHeight||290,ml=44,mr=16,mt=20,mb=48;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const vals=channels.map(c=>data.find(d=>d.key===c.key)?.[metric]??0);const max=Math.max(...vals,1)*1.15;svg.appendChild(svgEl('line',{{x1:ml,y1:h-mb,x2:w-mr,y2:h-mb,stroke:'#d9e0ec'}}));const slot=(w-ml-mr)/channels.length,bw=slot*.58;channels.forEach((c,i)=>{{const v=vals[i];const x=ml+i*slot+(slot-bw)/2;const y=mt+(h-mt-mb)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:4,fill:c.color}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:y-7,'text-anchor':'middle','font-size':12,'font-weight':700,fill:'#111'}})).textContent=fmt(v,2);svgLogo(svg,c,x+bw/2,h-24,24);}})}}
function lineChart(id, rows, channels=lineChannels, accessor=r=>r.aud, highlightKey='globo'){{const svg=document.getElementById(id);clear(svg);const w=svg.clientWidth||1000,h=svg.clientHeight||380,ml=48,mr=18,mt=20,mb=42;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const defs=svgEl('defs');const grad=svgEl('linearGradient',{{id:`${{id}}-badgeGrad`,x1:'0%',y1:'0%',x2:'100%',y2:'0%'}});grad.appendChild(svgEl('stop',{{offset:'0%','stop-color':'#004cff'}}));grad.appendChild(svgEl('stop',{{offset:'100%','stop-color':'#05a9f5'}}));defs.appendChild(grad);svg.appendChild(defs);let max=1,highlight=[];rows.forEach((r,i)=>channels.forEach(c=>{{const v=accessor(r)[c.key];if(v!==null&&v!==undefined){{if(v>max)max=v;if(c.key===highlightKey)highlight.push({{v,i,r,c}})}}}}));max*=1.12;for(let g=0;g<=4;g++){{const y=mt+(h-mt-mb)*g/4;svg.appendChild(svgEl('line',{{x1:ml,y1:y,x2:w-mr,y2:y,stroke:'#edf1f7'}}));svg.appendChild(svgEl('text',{{x:8,y:y+4,'font-size':11,fill:'#647084'}})).textContent=fmt(max*(1-g/4),1)}}channels.forEach(c=>{{let d='';rows.forEach((r,i)=>{{const v=accessor(r)[c.key];if(v===null||v===undefined)return;const x=ml+(w-ml-mr)*(i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-v/max);d+=(d?'L':'M')+x.toFixed(1)+' '+y.toFixed(1)}});svg.appendChild(svgEl('path',{{d,fill:'none',stroke:c.color,'stroke-width':c.key===highlightKey?3:2.1,'stroke-linejoin':'round','stroke-linecap':'round'}}));}});if(highlight.length){{const hi=highlight.reduce((a,b)=>b.v>a.v?b:a,highlight[0]);const lo=highlight.reduce((a,b)=>b.v<a.v?b:a,highlight[0]);[{{p:hi,label:'Maior'}},{{p:lo,label:'Menor'}}].forEach(({{p,label}})=>{{const x=ml+(w-ml-mr)*(p.i/Math.max(rows.length-1,1));const y=mt+(h-mt-mb)*(1-p.v/max);svg.appendChild(svgEl('circle',{{cx:x,cy:y,r:5.5,fill:'#005cef',stroke:'#fff','stroke-width':2}}));const text=`${{label}} ${{fmt(p.v,1)}}`;const width=Math.max(86,text.length*7+18);const tx=Math.min(w-width/2-12,Math.max(width/2+12,x+width/2));const ty=Math.max(24,y-12);svg.appendChild(svgEl('rect',{{x:tx-width/2,y:ty-17,width,height:24,rx:6,fill:`url(#${{id}}-badgeGrad)`}}));svg.appendChild(svgEl('text',{{x:tx,y:ty,'text-anchor':'middle','font-size':11,'font-weight':700,fill:'#fff'}})).textContent=text;}})}}const tickCount=Math.min(15,Math.max(6,Math.ceil(rows.length/70)+8));for(let t=0;t<tickCount;t++){{const p=t/(tickCount-1);const i=Math.min(rows.length-1,Math.round((rows.length-1)*p));const x=ml+(w-ml-mr)*p;svg.appendChild(svgEl('text',{{x,y:h-14,'text-anchor':'middle','font-size':11,fill:'#647084'}})).textContent=rows[i]?.time||''}}}}
function miniBars(svg, obj, color){{clear(svg);const entries=Object.entries(obj||{{}});const w=svg.clientWidth||320,h=svg.clientHeight||140,ml=34,mb=30,mt=10;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const max=Math.max(...entries.map(e=>e[1]||0),1);const bw=(w-ml-8)/Math.max(entries.length,1)*.58;entries.forEach(([k,v],i)=>{{v=v||0;const fill=Array.isArray(color)?color[i%color.length]:color;const x=ml+(i+.21)*(w-ml-8)/entries.length;const y=mt+16+(h-mt-mb-16)*(1-v/max);svg.appendChild(svgEl('rect',{{x,y,width:bw,height:h-mb-y,rx:3,fill}}));const val=fmt(v,0)+'%';const valueW=Math.max(34,val.length*7+14);const valueY=Math.max(8,y-24);svg.appendChild(svgEl('rect',{{x:x+bw/2-valueW/2,y:valueY,width:valueW,height:18,rx:5,fill:'#fff',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:valueY+13,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#27364e'}})).textContent=val;const labelW=Math.max(34,k.length*6+14);svg.appendChild(svgEl('rect',{{x:x+bw/2-labelW/2,y:h-24,width:labelW,height:18,rx:5,fill:'#f2f4f7',stroke:'#d9e0ec'}}));svg.appendChild(svgEl('text',{{x:x+bw/2,y:h-11,'text-anchor':'middle','font-size':10,'font-weight':700,fill:'#526178'}})).textContent=k;}})}}
function pie(svg, obj, colors=['#f59e0b','#f97316','#fb923c','#d97706','#b45309']){{clear(svg);const entries=Object.entries(obj||{{}}).filter(e=>(e[1]||0)>0);const w=svg.clientWidth||420,h=svg.clientHeight||150,groupW=430,offset=Math.max(0,(w-groupW)/2),cx=offset+128,cy=74,r=54;svg.setAttribute('viewBox',`0 0 ${{w}} ${{h}}`);const total=entries.reduce((a,e)=>a+e[1],0)||1;let a0=-Math.PI/2;entries.forEach(([k,v],i)=>{{const a1=a0+(v/total)*Math.PI*2;const large=a1-a0>Math.PI?1:0;const x0=cx+r*Math.cos(a0), y0=cy+r*Math.sin(a0), x1=cx+r*Math.cos(a1), y1=cy+r*Math.sin(a1);svg.appendChild(svgEl('path',{{d:`M${{cx}} ${{cy}} L${{x0}} ${{y0}} A${{r}} ${{r}} 0 ${{large}} 1 ${{x1}} ${{y1}} Z`,fill:colors[i%colors.length]}}));a0=a1;}});entries.forEach(([k,v],i)=>{{const lx=offset+260;svg.appendChild(svgEl('rect',{{x:lx,y:18+i*22,width:11,height:11,rx:2,fill:colors[i%colors.length]}}));svg.appendChild(svgEl('text',{{x:lx+17,y:28+i*22,'font-size':12,fill:'#3e4d63'}})).textContent=`${{k}} ${{fmt(v,1)}}%`;}})}}
function statusFromVar(v){{return v===null||v===undefined||Math.abs(v)<=0.4?'estavel':v>0?'cresceu':'caiu'}}
function statusArrow(status,shape='circle'){{if(status==='cresceu')return '<span class="arrow up">&#9650;</span>';if(status==='caiu')return '<span class="arrow down">&#9660;</span>';return shape==='square'?'<span class="arrow" style="color:#9aa4b2">&#9632;</span>':'<span class="arrow" style="color:#9aa4b2">&#9679;</span>'}}
function renderRankings(){{const box=document.getElementById('rankings');box.innerHTML=rankingChannels.map(c=>{{const rows=DATA.rankings[c.key]||[];return `<div class="card"><h3 class="channel-name">${{logoHtml(c)}}${{c.label}}</h3><table><thead><tr><th>#</th><th>Programa</th><th class="num">Aud.</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th><th class="num">Share</th><th class="num">Méd. 4 sem.</th><th class="num">Var.</th><th></th></tr></thead><tbody>${{rows.map(r=>`<tr><td>${{r.rank}}</td><td>${{esc(r.program)}}</td><td class="num">${{fmt(r.aud)}}</td><td class="num">${{fmt(r.audPrev)}}</td><td class="num">${{r.audVar==null?'-':fmt(r.audVar,1)+'%'}}</td><td>${{statusArrow(r.audStatus,'square')}}</td><td class="num">${{fmt(r.share)}}</td><td class="num">${{fmt(r.sharePrev)}}</td><td class="num">${{r.shareVar==null?'-':fmt(r.shareVar,1)+'%'}}</td><td>${{statusArrow(r.shareStatus,'square')}}</td></tr>`).join('')}}</tbody></table></div>`}}).join('')}}
function findMobileProgram(tokens){{const list=DATA.programs||[];const wanted=tokens.map(normalizeText);return list.find(p=>{{const name=normalizeText(p.name);return wanted.every(t=>name.includes(t))}})}}
function topProfileItems(obj,count=2){{return Object.entries(obj||{{}}).filter(([,v])=>v!==null&&v!==undefined).sort((a,b)=>(b[1]||0)-(a[1]||0)).slice(0,count).map(([label,value])=>`${{label}} (${{fmt(value,0)}}%)`).join(', ')||'-'}}
function tintColor(hex,amount=.88){{const h=String(hex||'#607086').replace('#','');const raw=h.length===3?h.split('').map(x=>x+x).join(''):h;const n=parseInt(raw,16);if(!Number.isFinite(n))return '#f8fbff';const r=(n>>16)&255,g=(n>>8)&255,b=n&255;const mix=v=>Math.round(v+(255-v)*amount);return `rgb(${{mix(r)}},${{mix(g)}},${{mix(b)}})`}}
function profileSummaryHtml(key,title,profile){{const gender=topProfileItems(profile?.gender,1);const classes=topProfileItems(profile?.classes,2);const ages=topProfileItems(profile?.ages,2);const c=channelMap[key]||{{key,label:title,color:'#607086'}};return `<div class="mobile-profile mobile-profile-${{key}}" style="background:${{tintColor(c.color,.9)}};border-color:${{tintColor(c.color,.68)}}">${{logoHtml(c)}}<div class="mobile-profile-text"><b>${{title}}</b><span>Gênero: ${{gender}}</span><span>Classes: ${{classes}}</span><span>Faixas: ${{ages}}</span></div></div>`}}
function mobileProgramList(){{const d=new Date(DATA.meta.date+'T00:00:00');const day=d.getDay();const defs=day===0?[[['globo','comunidade'],'GLOBO COMUNIDADE']]:day===6?[[['df','tv','1a'],'DF TV 1A'],[['df','tv','2a'],'DF TV 2A']]:[[['bom','dia','df'],'BOM DIA DF'],[['df','tv','1a'],'DF TV 1A'],[['globo','esporte'],'GLOBO ESPORTE'],[['df','tv','2a'],'DF TV 2A']];const seen=new Set();return defs.map(([tokens,label])=>{{const p=findMobileProgram(tokens);if(!p||seen.has(p.name))return null;seen.add(p.name);const globo=DATA.profile[p.name]?.globo||{{}};const nic=DATA.profile[p.name]?.nic||{{}};return {{label,name:p.name,aud:globo.aud,share:globo.share,globoProfile:profileSummaryHtml('globo','Perfil Globo',globo),nicProfile:profileSummaryHtml('nic','Perfil NIC',nic)}}}}).filter(Boolean)}}
function renderMobileSummary(){{const box=document.getElementById('mobileSummary');if(!box)return;const d=new Date(DATA.meta.date+'T00:00:00');const mobileChannels=['globo','nic','record','sbt','band','paytv'].map(key=>DATA.summaryBars.find(x=>x.key===key)).filter(Boolean);const leaderRows=leadershipRowsHtml(DATA.leadership||[]);const leaderRowsNoNic=leadershipRowsHtml(leadershipExceptNic());const rankingHtml=rankingChannels.map(c=>{{const rows=(DATA.rankings[c.key]||[]).slice(0,5);return `<div class="mobile-rank-block" style="border-top-color:${{c.color}}"><div class="mobile-rank-title">${{logoHtml(c)}}${{c.label}}</div>${{rows.map(r=>`<div class="mobile-rank-row"><span>${{r.rank}}</span><strong>${{esc(r.program)}}</strong><span>${{fmt(r.aud)}} | ${{fmt(r.share)}}%</span></div>`).join('')||'<div class="empty">Sem dados</div>'}}</div>`}}).join('');const programs=mobileProgramList();const programHtml=programs.length?programs.map(p=>`<div class="mobile-program"><div class="mobile-program-main"><strong>${{esc(p.label)}}</strong><div class="mobile-program-metrics"><div class="mobile-program-metric"><small>Aud.</small><strong>${{fmt(p.aud)}}</strong></div><div class="mobile-program-metric share"><small>Share</small><strong>${{fmt(p.share)}}%</strong></div></div><div class="mobile-program-profiles">${{p.globoProfile}}${{p.nicProfile}}</div></div></div>`).join(''):'<div class="empty">Sem programas correspondentes nas bases enviadas.</div>';box.innerHTML=`<div class="mobile-alert">Para acessar as informações completas, recomendamos a visualização em desktop.</div><div class="mobile-hero"><small>${{DATA.meta.weekday}} • ${{d.toLocaleDateString('pt-BR')}}</small><h1>Resumo do Distrito Federal</h1></div><div class="mobile-card">${{buildDailyInsights('globo')}}</div><div class="mobile-card">${{buildDailyInsights('nic')}}</div><div class="mobile-card"><h2>Programas locais</h2><div class="mobile-programs">${{programHtml}}</div></div><div class="mobile-card"><h2>Audiência e Share 07h-24h</h2><div class="mobile-grid">${{mobileChannels.map(row=>{{const c=channelMap[row.key]||row;return `<div class="mobile-kpi" style="border-left-color:${{c.color}}"><small>${{logoHtml(c)}} ${{c.label}}</small><strong>${{fmt(row.aud)}}</strong><span>Share ${{row.share==null?'-':fmt(row.share)+'%'}}</span></div>`}}).join('')}}</div></div><div class="mobile-card"><h2>Minutos na liderança</h2><div class="mobile-leaders">${{leaderRows}}</div></div><div class="mobile-card"><h2>Minutos na liderança (exceto NIC)</h2><div class="mobile-leaders">${{leaderRowsNoNic}}</div></div><div class="mobile-card"><h2>Ranking por emissora</h2><div class="mobile-ranking">${{rankingHtml}}</div></div><div class="mobile-source">${{sourceText()}}</div>`}}
function renderTabletSummary(){{const box=document.getElementById('tabletSummary');if(!box)return;const d=new Date(DATA.meta.date+'T00:00:00');const tabletChannels=['globo','nic','record','sbt','band','paytv'].map(key=>DATA.summaryBars.find(x=>x.key===key)).filter(Boolean);const kpis=tabletChannels.map(row=>{{const c=channelMap[row.key]||row;return `<div class="tablet-kpi" style="border-left-color:${{c.color}}"><small>${{logoHtml(c)}} ${{c.label}}</small><strong>${{fmt(row.aud)}}</strong><span>Share ${{row.share==null?'-':fmt(row.share)+'%'}}</span></div>`}}).join('');const leaders=leadershipRowsHtml(DATA.leadership||[]);const leadersNoNic=leadershipRowsHtml(leadershipExceptNic());const programDetails=(DATA.programs||[]).map((p,i)=>{{const target=Object.keys(DATA.programCompetition[p.name]||{{}}).find(t=>normalizeText(t).includes('total domic'))||DATA.targets[0];const comp=(DATA.programCompetition[p.name]||{{}})[target]||{{}};const rows=programTableChannels.map(c=>`<tr><td>${{logoHtml(c)}} ${{c.label}}</td><td class="num">${{fmt(comp[c.key]?.aud)}}</td><td class="num">${{fmt(comp[c.key]?.share)}}%</td></tr>`).join('');const profileCards=programTableChannels.map(c=>profileSummaryHtml(c.key,`Perfil ${{c.label}}`,DATA.profile[p.name]?.[c.key]||{{}})).join('');return `<details class="tablet-program" ${{i<2?'open':''}}><summary><span>${{esc(p.name)}}</span></summary><div class="tablet-program-body"><table class="tablet-program-table"><thead><tr><th>Emissora</th><th class="num">Aud.</th><th class="num">Share</th></tr></thead><tbody>${{rows}}</tbody></table><div class="tablet-profile-stack">${{profileCards}}</div></div></details>`}}).join('');const rankingHtml=rankingChannels.map(c=>{{const rows=(DATA.rankings[c.key]||[]).slice(0,5);return `<div class="mobile-rank-block" style="border-top-color:${{c.color}}"><div class="mobile-rank-title">${{logoHtml(c)}}${{c.label}}</div>${{rows.map(r=>`<div class="mobile-rank-row"><span>${{r.rank}}</span><strong>${{esc(r.program)}}</strong><span>${{fmt(r.aud)}} | ${{fmt(r.share)}}%</span></div>`).join('')||'<div class="empty">Sem dados</div>'}}</div>`}}).join('');box.innerHTML=`<div class="tablet-alert">Visualização tablet: blocos compactos e expansíveis. Para gráficos completos, recomendamos desktop.</div><div class="mobile-hero"><small>${{DATA.meta.weekday}} • ${{d.toLocaleDateString('pt-BR')}}</small><h1>Resumo do Distrito Federal</h1></div><div class="tablet-card">${{buildDailyInsights('globo')}}</div><div class="tablet-card">${{buildDailyInsights('nic')}}</div><div class="tablet-card"><h2>Audiência e Share 07h-24h</h2><div class="tablet-grid">${{kpis}}</div></div><div class="tablet-card"><h2>Minutos na liderança</h2><div class="tablet-leaders">${{leaders}}</div></div><div class="tablet-card"><h2>Minutos na liderança (exceto NIC)</h2><div class="tablet-leaders">${{leadersNoNic}}</div></div><div class="tablet-card"><h2>Todos os programas</h2><div class="tablet-program-list">${{programDetails}}</div></div><div class="tablet-card"><h2>Ranking por emissora</h2><div class="tablet-ranking">${{rankingHtml}}</div></div><div class="tablet-source">${{sourceText()}}</div>`}}
function topProfileLabel(obj){{const entries=Object.entries(obj||{{}}).filter(([,v])=>v!==null&&v!==undefined);if(!entries.length)return '-';const [label,value]=entries.sort((a,b)=>(b[1]||0)-(a[1]||0))[0];return `${{label}} ${{fmt(value,0)}}%`}}
function sourceText(){{const d=new Date(DATA.meta.date+'T00:00:00');return `Fonte: Ibope. Instar Analytics. DF. Total Domicílios. Aud%, Shr% Adh%. Atividades: Live. Data: ${{d.toLocaleDateString('pt-BR')}}.`}}
function dateToken(){{const d=new Date(DATA.meta.date+'T00:00:00');return `${{String(d.getDate()).padStart(2,'0')}}${{String(d.getMonth()+1).padStart(2,'0')}}${{d.getFullYear()}}`}}
function firstValid(items, selector){{return items.find(item=>selector(item)!==null&&selector(item)!==undefined)}}
function insightChannels(){{return leadChannels.filter(c=>c.key!=='tle')}}
function renderDailyInsights(selectedKey='globo'){{
  window.selectedInsightKey=selectedKey;
  const target=document.getElementById('dailyInsights');
  if(!target)return;
  const selected=channelMap[selectedKey]||channelMap.globo||insightChannels()[0]||{{}};
  const card=target.closest('.insights-card');
  if(card){{
    card.style.borderLeftColor=selected.color||'#005cef';
    card.style.boxShadow=`0 12px 30px rgba(21,36,66,.08), inset 0 4px 0 ${{selected.color||'#005cef'}}`;
  }}
  target.innerHTML=buildDailyInsights(selectedKey);
}}
function timeToMinutesLabel(time){{const [h,m]=String(time||'0:0').split(':').map(Number);return (h||0)*60+(m||0)}}
function inRangeMinute(min,start,end){{return start<=end ? min>=start&&min<=end : min>=start||min<=end}}
function avgForRows(rows,key){{const vals=rows.map(r=>r.aud?.[key]).filter(v=>v!==null&&v!==undefined);return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null}}
function compareRows(rows,key,baseKey='globo'){{const selected=avgForRows(rows,key), base=avgForRows(rows,baseKey);return {{selected,base,diff:selected!=null&&base!=null?selected-base:null}}}}
function periodAverages(key){{const defs=[['Manh\u00e3',360,719],['Tarde',720,1079],['Noite',1080,1439],['Madrugada',0,359]];return defs.map(([label,start,end])=>{{const rows=DATA.line.filter(r=>inRangeMinute(timeToMinutesLabel(r.time),start,end));return {{label,aud:avgForRows(rows,key),rows}}}}).filter(p=>p.aud!==null&&p.aud!==undefined).sort((a,b)=>b.aud-a.aud)}}
function periodComparisons(key,baseKey='globo'){{const defs=[['Manh\u00e3',360,719],['Tarde',720,1079],['Noite',1080,1439],['Madrugada',0,359]];return defs.map(([label,start,end])=>{{const rows=DATA.line.filter(r=>inRangeMinute(timeToMinutesLabel(r.time),start,end));return {{label,...compareRows(rows,key,baseKey)}}}}).filter(p=>p.diff!==null&&p.diff!==undefined).sort((a,b)=>b.diff-a.diff)}}
function programAveragesForChannel(key){{return (DATA.programs||[]).map(p=>{{const rows=programMinuteRows(p);return {{program:p.name,aud:avgForRows(rows,key),start:p.start,end:p.end}}}}).filter(x=>x.aud!==null&&x.aud!==undefined).sort((a,b)=>b.aud-a.aud)}}
function programComparisons(key,baseKey='globo'){{return (DATA.programs||[]).map(p=>{{const rows=programMinuteRows(p);return {{program:p.name,...compareRows(rows,key,baseKey),start:p.start,end:p.end}}}}).filter(x=>x.diff!==null&&x.diff!==undefined).sort((a,b)=>b.diff-a.diff)}}
function diffText(v){{return `${{v>=0?'+':''}}${{fmt(v,2)}} ponto(s)`}}
function insightCardHtml(c,selected){{return `<div class="insight-item" style="border-top:4px solid ${{selected.color||'#005cef'}}"><small>${{c.k}}</small><strong>${{esc(c.v)}}</strong><span>${{c.s}}</span></div>`}}
function buildDailyInsights(selectedKey='globo'){{
  const channels=insightChannels();
  const selected=channelMap[selectedKey]||channelMap.globo||channels[0]||{{}};
  const globo=DATA.summaryBars.find(x=>x.key==='globo')||{{}};
  const current=DATA.summaryBars.find(x=>x.key===selected.key)||{{}};
  const rows=(DATA.rankings[selected.key]||[]).filter(r=>r.program);
  const byAud=[...rows].filter(r=>r.aud!==null&&r.aud!==undefined).sort((a,b)=>(b.aud||0)-(a.aud||0));
  const byShare=[...rows].filter(r=>r.share!==null&&r.share!==undefined).sort((a,b)=>(b.share||0)-(a.share||0));
  const bestAud=byAud[0], bestShare=byShare[0], weakest=byAud[byAud.length-1];
  const growth=[...rows].filter(r=>r.audVar!==null&&r.audVar!==undefined).sort((a,b)=>(b.audVar||0)-(a.audVar||0))[0];
  const fall=[...rows].filter(r=>r.audVar!==null&&r.audVar!==undefined).sort((a,b)=>(a.audVar||0)-(b.audVar||0))[0];
  const audVarRows=rows.filter(r=>r.audVar!==null&&r.audVar!==undefined);
  const positiveAud=audVarRows.filter(r=>r.audVar>0).length;
  const negativeAud=audVarRows.filter(r=>r.audVar<0).length;
  const stableAud=audVarRows.length-positiveAud-negativeAud;
  const globoTrendText=positiveAud>negativeAud
    ? `a maioria dos programas ranqueados avan\u00e7ou frente \u00e0s \u00faltimas 4 semanas, com ${{positiveAud}} altas${{growth?`, puxadas por ${{esc(growth.program)}}`:''}}`
    : negativeAud>positiveAud
      ? `houve press\u00e3o em parte da grade na compara\u00e7\u00e3o com as \u00faltimas 4 semanas, com ${{negativeAud}} quedas entre os programas ranqueados`
      : `o desempenho ficou equilibrado na compara\u00e7\u00e3o com as \u00faltimas 4 semanas${{audVarRows.length?`, com ${{positiveAud}} altas, ${{negativeAud}} quedas e ${{stableAud}} estabilidade(s)`:''}}`;
  const points=DATA.line.map(r=>({{time:r.time,value:r.aud?.[selected.key],base:r.aud?.globo,diff:r.aud?.[selected.key]!=null&&r.aud?.globo!=null?r.aud[selected.key]-r.aud.globo:null}})).filter(p=>p.value!==null&&p.value!==undefined);
  const peak=points.length?[...points].sort((a,b)=>b.value-a.value)[0]:null;
  const low=points.length?[...points].sort((a,b)=>a.value-b.value)[0]:null;
  const bestDiff=points.filter(p=>p.diff!==null&&p.diff!==undefined).sort((a,b)=>b.diff-a.diff)[0];
  const leader=(DATA.leadership||[]).find(l=>l.key===selected.key)||{{minutes:0,percent:0,hours:'0h00'}};
  const globoLeader=(DATA.leadership||[]).find(l=>l.key==='globo')||{{minutes:0,percent:0,hours:'0h00'}};
  const tabs=channels.map(c=>{{
    const active=c.key===selected.key;
    const style=active?'background:linear-gradient(135deg,'+c.color+','+softenColor(c.color)+')':'';
    return `<button class="insight-tab ${{active?'active':''}}" type="button" data-insight="${{c.key}}" onclick="renderDailyInsights('${{c.key}}')" style="${{style}}">${{logoHtml(c)}}${{c.label}}</button>`;
  }}).join('');
  const mainStyle=`background:linear-gradient(135deg,${{selected.color||'#005cef'}},${{softenColor(selected.color||'#005cef')}});color:#fff`;
  const cards=[];
  const bullets=[];
  let headline='';
  if(selected.key==='nic'){{
    const periodDiffs=periodComparisons('nic','globo');
    const programDiffs=programComparisons('nic','globo');
    const bestPeriod=periodDiffs[0];
    const bestProgram=programDiffs[0];
    headline=`NIC fechou o per\u00edodo 07h-24h com m\u00e9dia de ${{fmt(current.aud,2)}} pontos e ${{fmt(current.share,2)}}% de share. A leitura destaca onde a dist\u00e2ncia entre NIC e Globo ficou mais favor\u00e1vel ao NIC, por faixa hor\u00e1ria e nas janelas de programas da Globo.`;
    if(bestPeriod)cards.push({{k:'Melhor faixa vs. Globo',v:bestPeriod.label,s:`NIC ${{fmt(bestPeriod.selected,2)}} | Globo ${{fmt(bestPeriod.base,2)}} | ${{diffText(bestPeriod.diff)}}`}});
    if(periodDiffs[1])cards.push({{k:'2\u00aa melhor faixa',v:periodDiffs[1].label,s:`NIC ${{fmt(periodDiffs[1].selected,2)}} | Globo ${{fmt(periodDiffs[1].base,2)}} | ${{diffText(periodDiffs[1].diff)}}`}});
    if(bestProgram)cards.push({{k:'Programa com melhor dist\u00e2ncia',v:bestProgram.program,s:`NIC ${{fmt(bestProgram.selected,2)}} | Globo ${{fmt(bestProgram.base,2)}} | ${{diffText(bestProgram.diff)}}`}});
    if(bestDiff)cards.push({{k:'Melhor minuto vs. Globo',v:bestDiff.time,s:`NIC ${{fmt(bestDiff.value,1)}} | Globo ${{fmt(bestDiff.base,1)}} | ${{diffText(bestDiff.diff)}}`}});
    if(periodDiffs.length)bullets.push(`Por faixa hor\u00e1ria, a melhor rela\u00e7\u00e3o NIC x Globo ocorreu em ${{periodDiffs.map(p=>`${{p.label}} (${{diffText(p.diff)}})`).join(', ')}}.`);
    if(programDiffs.length)bullets.push(`Nas janelas da Globo, as melhores dist\u00e2ncias para o NIC apareceram em ${{programDiffs.slice(0,5).map(p=>`${{esc(p.program)}} (${{diffText(p.diff)}})`).join(', ')}}.`);
    if(bestDiff)bullets.push(`No minuto mais favor\u00e1vel, \u00e0s ${{bestDiff.time}}, o NIC marcou ${{fmt(bestDiff.value,1)}} ponto(s) contra ${{fmt(bestDiff.base,1)}} da Globo.`);
    bullets.push(`No total do dia, o NIC registrou ${{leader.minutes}} minutos na lideran\u00e7a (${{fmt(leader.percent,1)}}% do per\u00edodo).`);
  }}else{{
    headline=selected.key==='globo'
      ? `A Globo fechou o per\u00edodo 07h-24h com m\u00e9dia de ${{fmt(current.aud,2)}} pontos e ${{fmt(current.share,2)}}% de share. A an\u00e1lise indica que ${{globoTrendText}}, al\u00e9m de ${{leader.minutes}} minutos de lideran\u00e7a (${{fmt(leader.percent,1)}}% do per\u00edodo).`
      : `${{selected.label}} fechou o per\u00edodo 07h-24h com ${{fmt(current.aud,2)}} pontos e ${{fmt(current.share,2)}}% de share. A Globo marcou ${{fmt(globo.aud,2)}} pontos e ${{fmt(globo.share,2)}}%; no minuto a minuto, ${{selected.label}} liderou ${{leader.minutes}} minutos e a Globo liderou ${{globoLeader.minutes}} minutos.`;
    if(bestAud)cards.push({{k:'Melhor audi\u00eancia',v:bestAud.program,s:`${{fmt(bestAud.aud,2)}} pontos | ${{fmt(bestAud.share,2)}}% share`}});
    if(bestShare)cards.push({{k:'Melhor share',v:bestShare.program,s:`${{fmt(bestShare.share,2)}}% share | ${{fmt(bestShare.aud,2)}} pontos`}});
    const attention=fall&&fall.audVar<0?{{program:fall.program,s:`${{fmt(fall.audVar,1)}}% vs. 4 semanas | ${{fmt(fall.aud,2)}} pontos`}}:weakest?{{program:weakest.program,s:`${{fmt(weakest.aud,2)}} pontos | ${{fmt(weakest.share,2)}}% share`}}:null;
    if(attention)cards.push({{k:'Ponto de aten\u00e7\u00e3o',v:attention.program,s:attention.s}});
    if(!cards.length){{
      const periods=periodAverages(selected.key);
      const bestPeriod=periods[0];
      if(bestPeriod)cards.push({{k:'Faixa mais forte',v:bestPeriod.label,s:`${{fmt(bestPeriod.aud,2)}} pontos em m\u00e9dia`}});
      if(peak)cards.push({{k:'Pico no minuto',v:peak.time,s:`${{fmt(peak.value,1)}} pontos`}});
    }}
    if(weakest)bullets.push(`A faixa de maior dificuldade entre os programas ranqueados de ${{esc(selected.label)}} foi ${{esc(weakest.program)}}, com ${{fmt(weakest.aud,2)}} pontos e ${{fmt(weakest.share,2)}}% de share.`);
    if(growth&&growth.audVar>0)bullets.push(`${{esc(growth.program)}} foi o principal avan\u00e7o frente \u00e0 m\u00e9dia das 4 semanas anteriores.`);
    if(fall&&fall.audVar<0)bullets.push(`${{esc(fall.program)}} concentrou a principal queda na compara\u00e7\u00e3o com as 4 semanas anteriores.`);
    if(peak&&low)bullets.push(`No minuto a minuto, ${{esc(selected.label)}} atingiu pico de ${{fmt(peak.value,1)}} \u00e0s ${{peak.time}} e menor patamar de ${{fmt(low.value,1)}} \u00e0s ${{low.time}}.`);
    if(selected.key!=='globo')bullets.push(`A dist\u00e2ncia para a Globo no per\u00edodo foi de ${{fmt((globo.aud||0)-(current.aud||0),2)}} ponto(s) de audi\u00eancia e ${{fmt((globo.share||0)-(current.share||0),2)}} p.p. de share.`);
  }}
  if(!cards.length&&peak)cards.push({{k:'Pico no minuto',v:peak.time,s:`${{fmt(peak.value,1)}} pontos`}});
  if(!bullets.length)bullets.push('N\u00e3o h\u00e1 ranking detalhado dispon\u00edvel para esta emissora nas bases enviadas; a leitura usa audi\u00eancia, share e minuto a minuto.');
  return `<div class="section-head"><h2>Destaques do dia</h2><div class="insight-tabs">${{tabs}}</div></div><div class="insight-context"><div class="insight-grid" style="grid-template-columns:1.2fr repeat(${{Math.max(cards.length,1)}},1fr)"><div class="insight-main" style="${{mainStyle}}">${{headline}}</div>${{cards.map(c=>insightCardHtml(c,selected)).join('')}}</div><ul class="insight-list">${{bullets.map(b=>`<li>${{b}}</li>`).join('')}}</ul></div>`;
}}
function renderResumo(){{const d=new Date(DATA.meta.date+'T00:00:00');document.getElementById('dateLabel').className='date-pill';document.getElementById('dateLabel').innerHTML=`<span>${{DATA.meta.weekday}}</span><strong>${{d.toLocaleDateString('pt-BR')}}</strong>`;document.getElementById('dailyAvg').textContent=fmt(DATA.meta.dailyAvg,2);const globoLogo=document.getElementById('globoKpiLogo');if(globoLogo&&channelMap.globo?.logoData)globoLogo.src=channelMap.globo.logoData;document.getElementById('sourceNote').textContent=sourceText();renderDailyInsights(window.selectedInsightKey||'globo');barChart('audBars',DATA.summaryBars,'aud',audChannels);barChart('shareBars',DATA.summaryBars,'share',shareChannels);lineChart('dayLine',DATA.line,lineChannels);legend('legend3',lineChannels);document.getElementById('leadership').innerHTML=leaderCardsHtml(DATA.leadership);document.getElementById('leadershipNoNic').innerHTML=leaderCardsHtml(leadershipExceptNic());renderRankings();renderMobileSummary();renderTabletSummary()}}
function fillSelect(sel, values){{sel.innerHTML=values.map(v=>`<option value="${{esc(v)}}">${{esc(v)}}</option>`).join('')}}
function profileChoices(){{return ['07h-24h',...(DATA.programs||[]).map(p=>p.name)]}}
function programMinuteRows(program){{const start=program.start??0,end=program.end??0;return DATA.minuteAll.filter(m=>m.minute>=start && m.minute<end)}}
function renderProgramas(){{const ps=document.getElementById('programSelect'),ts=document.getElementById('targetSelect'),title=document.getElementById('programTitle'),table=document.getElementById('minuteTable');const selectedProgram=ps?.value||DATA.programs[0]?.name;const p=DATA.programs.find(x=>x.name===selectedProgram)||DATA.programs[0];if(!p)return;const compByTarget=DATA.programCompetition[p.name]||{{}};const target=ts?.value&&compByTarget[ts.value]?ts.value:(Object.keys(compByTarget)[0]||ts?.value||DATA.targets[0]||'');if(ts&&target&&ts.value!==target)ts.value=target;if(title)title.textContent=`${{p.name}} | ${{target}}`;const comp=compByTarget[target]||{{}};const audBars=audChannels.map(c=>({{...c,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));const shareBars=shareChannels.map(c=>({{...c,aud:comp[c.key]?.aud,share:comp[c.key]?.share}}));barChart('progAud',audBars,'aud',audBars);barChart('progShare',shareBars,'share',shareBars);const rows=programMinuteRows(p);lineChart('progLine',rows,lineChannels,r=>r.aud);legend('legend4',lineChannels);if(table)table.innerHTML=`<table class="minute-table"><thead><tr><th>Minuto</th>${{programTableChannels.map(c=>`<th colspan="2" style="border-top:4px solid ${{c.color}}">${{logoHtml(c)}} ${{c.label}}</th>`).join('')}}</tr><tr><th></th>${{programTableChannels.map(c=>`<th class="num">Aud.</th><th class="num">Share</th>`).join('')}}</tr></thead><tbody>${{rows.map(r=>`<tr><td><strong>${{r.time}}</strong></td>${{programTableChannels.map(c=>`<td class="num" style="color:${{c.color}}">${{fmt(r.aud[c.key])}}</td><td class="num">${{fmt(r.share[c.key])}}</td>`).join('')}}</tr>`).join('')}}</tbody></table>`}}
function softenColor(hex){{const h=String(hex||'#005cef').replace('#','');const raw=h.length===3?h.split('').map(x=>x+x).join(''):h;const n=parseInt(raw,16);if(!Number.isFinite(n))return '#45a3ff';const r=(n>>16)&255,g=(n>>8)&255,b=n&255;const mix=v=>Math.round(v+(255-v)*.34);return `rgb(${{mix(r)}},${{mix(g)}},${{mix(b)}})`}}
function renderPerfil(){{const sel=document.getElementById('profileSelect'),title=document.getElementById('profileTitle'),grid=document.getElementById('profileGrid');const name=sel?.value||'07h-24h';if(title)title.textContent=name;const pdata=DATA.profile[name]||{{}};const genderColors=['#4f8bd6','#d95f76'];const classColors=['#ffd166','#ef476f','#06d6a0','#118ab2','#8338ec'];const ageColors=['#80ffdb','#56cfe1','#4ea8de','#5e60ce','#6930c3','#1b4965'];if(!grid)return;grid.innerHTML=leadChannels.map(c=>{{const d=pdata[c.key]||{{}};return `<div class="card profile-card" data-ch="${{c.key}}" style="border-top:5px solid ${{c.color}};--metric-grad:linear-gradient(135deg,${{c.color}},${{softenColor(c.color)}})"><div class="profile-head"><h3 class="profile-title">${{logoHtml(c)}}${{c.label}}</h3><div class="metric-pair"><div class="mini"><small>Aud.</small><strong>${{fmt(d.aud)}}</strong></div><div class="mini"><small>Share</small><strong>${{fmt(d.share)}}</strong></div></div></div><div class="profile-block"><div><h4>G\u00eanero</h4><svg class="mini-chart gender"></svg></div><div><h4>Classes sociais</h4><svg class="mini-chart classes"></svg></div><div><h4>Faixas etárias</h4><svg class="mini-chart ages"></svg></div></div></div>`}}).join('');document.querySelectorAll('.profile-card').forEach(card=>{{const key=card.dataset.ch, d=pdata[key]||{{}};miniBars(card.querySelector('.gender'),d.gender,genderColors);pie(card.querySelector('.classes'),d.classes,classColors);miniBars(card.querySelector('.ages'),d.ages,ageColors)}})}}
function downloadBlob(name,type,content){{const blob=new Blob([content],{{type}});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),500)}}
function pdfHex(text){{const s=String(text??'');let out='FEFF';for(const ch of s){{const cp=ch.codePointAt(0);if(cp>0xffff)continue;out+=cp.toString(16).toUpperCase().padStart(4,'0')}}return out}}
function pdfWrap(text,maxChars){{const words=String(text??'').replace(/\s+/g,' ').trim().split(' ');const lines=[];let line='';words.forEach(word=>{{if(!word)return;const test=line?line+' '+word:word;if(test.length>maxChars&&line){{lines.push(line);line=word}}else line=test}});if(line)lines.push(line);return lines.length?lines:['']}}
function makePdfBlob(title,lines){{const enc=new TextEncoder();const objects=[];const pages=[];function addObject(body){{objects.push(body);return objects.length}}function escText(text){{return '<'+pdfHex(text)+'>'}}const fontObj=addObject('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');const boldObj=addObject('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>');let pageOps=[],y=800;function newPage(){{if(pageOps.length){{const content=pageOps.join('\\n');const contentObj=addObject('<< /Length '+enc.encode(content).length+' >>\\nstream\\n'+content+'\\nendstream');const pageObj=addObject('<< /Type /Page /Parent 0 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 '+fontObj+' 0 R /F2 '+boldObj+' 0 R >> >> /Contents '+contentObj+' 0 R >>');pages.push(pageObj)}}pageOps=[];y=800}}function text(x,size,value,bold=false){{pageOps.push('BT /'+(bold?'F2':'F1')+' '+size+' Tf '+x+' '+y+' Td '+escText(value)+' Tj ET');y-=size+6}}function gap(v=8){{y-=v}}function line(value,opt={{}}){{const size=opt.size||10;const x=opt.x||42;const bold=!!opt.bold;const max=opt.max||92;pdfWrap(value,max).forEach(part=>{{if(y<54)newPage();text(x,size,part,bold)}})}}newPage();line(title,{{size:20,bold:true,max:48}});gap(6);lines.forEach(item=>{{if(item===''){{gap(8);return}}if(typeof item==='object')line(item.text,item);else line(item)}});if(pageOps.length){{const content=pageOps.join('\\n');const contentObj=addObject('<< /Length '+enc.encode(content).length+' >>\\nstream\\n'+content+'\\nendstream');const pageObj=addObject('<< /Type /Page /Parent 0 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 '+fontObj+' 0 R /F2 '+boldObj+' 0 R >> >> /Contents '+contentObj+' 0 R >>');pages.push(pageObj)}}const pagesObj=addObject('<< /Type /Pages /Kids ['+pages.map(p=>p+' 0 R').join(' ')+'] /Count '+pages.length+' >>');for(const page of pages)objects[page-1]=objects[page-1].replace('/Parent 0 0 R','/Parent '+pagesObj+' 0 R');const catalogObj=addObject('<< /Type /Catalog /Pages '+pagesObj+' 0 R >>');let pdf='%PDF-1.4\\n';const offsets=[0];objects.forEach((body,i)=>{{offsets.push(enc.encode(pdf).length);pdf+=(i+1)+' 0 obj\\n'+body+'\\nendobj\\n'}});const xref=enc.encode(pdf).length;pdf+='xref\\n0 '+(objects.length+1)+'\\n0000000000 65535 f \\n';for(let i=1;i<offsets.length;i++)pdf+=String(offsets[i]).padStart(10,'0')+' 00000 n \\n';pdf+='trailer\\n<< /Size '+(objects.length+1)+' /Root '+catalogObj+' 0 R >>\\nstartxref\\n'+xref+'\\n%%EOF';return new Blob([pdf],{{type:'application/pdf'}})}}
function downloadPdf(name,lines){{const blob=makePdfBlob('DESEMPENHO DIÁRIO DF',lines);const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),800)}}
function buildPdfLines(){{const d=new Date(DATA.meta.date+'T00:00:00');const lines=[];lines.push({{text:DATA.meta.weekday+' - '+d.toLocaleDateString('pt-BR'),bold:true,size:12}});lines.push('');lines.push({{text:'Resumo 07h-24h',bold:true,size:13}});lines.push('Audiência média Globo: '+fmt(DATA.meta.dailyAvg,2));lines.push('Médias por emissora: '+DATA.summaryBars.map(r=>r.label+': Aud. '+fmt(r.aud,2)+(r.share==null?'':' | Share '+fmt(r.share,2)+'%')).join('  |  '));const insightText=document.getElementById('dailyInsights')?.innerText||'';if(insightText.trim()){{lines.push('');lines.push({{text:'Destaques do dia',bold:true,size:13}});insightText.split('\\n').filter(Boolean).slice(0,12).forEach(t=>lines.push(t))}}const visibleLeaders=rows=>(rows||[]).filter(l=>Number(l.minutes||0)>0);lines.push('');lines.push({{text:'Minutos na liderança',bold:true,size:13}});visibleLeaders(DATA.leadership).forEach(l=>lines.push(l.label+': '+l.minutes+' min | '+fmt(l.percent,1)+'% | '+l.hours));lines.push('');lines.push({{text:'Minutos na liderança (exceto NIC)',bold:true,size:13}});visibleLeaders(leadershipExceptNic()).forEach(l=>lines.push(l.label+': '+l.minutes+' min | '+fmt(l.percent,1)+'% | '+l.hours));lines.push('');lines.push({{text:'Ranking por emissora',bold:true,size:13}});rankingChannels.forEach(c=>{{lines.push({{text:c.label,bold:true,size:11}});(DATA.rankings[c.key]||[]).forEach(r=>lines.push(r.rank+'. '+r.program+' | Aud. '+fmt(r.aud,2)+' vs '+fmt(r.audPrev,2)+' ('+(r.audVar==null?'-':fmt(r.audVar,1)+'%')+') | Share '+fmt(r.share,2)+'% vs '+fmt(r.sharePrev,2)+'% ('+(r.shareVar==null?'-':fmt(r.shareVar,1)+'%')+')'))}});lines.push('');lines.push({{text:'Programas',bold:true,size:13}});(DATA.programs||[]).forEach(p=>{{const target=Object.keys(DATA.programCompetition[p.name]||{{}}).find(t=>normalizeText(t).includes('total domic'))||DATA.targets[0];const comp=(DATA.programCompetition[p.name]||{{}})[target]||{{}};lines.push({{text:p.name,bold:true,size:11}});programTableChannels.forEach(c=>{{const row=comp[c.key]||{{}};lines.push(c.label+': Aud. '+fmt(row.aud,2)+' | Share '+fmt(row.share,2)+'%')}})}});lines.push('');lines.push(sourceText());return lines}}
function dataUrlBytes(dataUrl){{const raw=atob(String(dataUrl).split(',')[1]||'');const bytes=new Uint8Array(raw.length);for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);return bytes}}
function makeImagePdfBlob(pages){{const enc=new TextEncoder();const objects=[];const offsets=[0];const chunks=[];const pageRefs=[];const addText=str=>{{chunks.push(enc.encode(str))}};const addBytes=bytes=>{{chunks.push(bytes)}};const sizeSoFar=()=>chunks.reduce((sum,c)=>sum+c.length,0);function addObjectHeader(n){{offsets[n]=sizeSoFar();addText(n+' 0 obj\\n')}}function addObject(n,body){{addObjectHeader(n);addText(body+'\\nendobj\\n')}}addText('%PDF-1.4\\n');let obj=1;const catalog=obj++,pagesObj=obj++;const imageObjects=[];for(const page of pages){{const img=obj++,content=obj++,pageObj=obj++;imageObjects.push({{img,content,pageObj,page}});pageRefs.push(pageObj+' 0 R')}}addObject(catalog,'<< /Type /Catalog /Pages '+pagesObj+' 0 R >>');addObject(pagesObj,'<< /Type /Pages /Kids ['+pageRefs.join(' ')+'] /Count '+pages.length+' >>');for(const p of imageObjects){{const bytes=dataUrlBytes(p.page.dataUrl);addObjectHeader(p.img);addText('<< /Type /XObject /Subtype /Image /Width '+p.page.width+' /Height '+p.page.height+' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length '+bytes.length+' >>\\nstream\\n');addBytes(bytes);addText('\\nendstream\\nendobj\\n');const ops='q '+p.page.pdfW+' 0 0 '+p.page.pdfH+' 0 0 cm /Im'+p.img+' Do Q';addObject(p.content,'<< /Length '+enc.encode(ops).length+' >>\\nstream\\n'+ops+'\\nendstream');addObject(p.pageObj,'<< /Type /Page /Parent '+pagesObj+' 0 R /MediaBox [0 0 '+p.page.pdfW+' '+p.page.pdfH+'] /Resources << /XObject << /Im'+p.img+' '+p.img+' 0 R >> >> /Contents '+p.content+' 0 R >>')}}const xref=sizeSoFar();addText('xref\\n0 '+obj+'\\n0000000000 65535 f \\n');for(let i=1;i<obj;i++)addText(String(offsets[i]).padStart(10,'0')+' 00000 n \\n');addText('trailer\\n<< /Size '+obj+' /Root '+catalog+' 0 R >>\\nstartxref\\n'+xref+'\\n%%EOF');return new Blob(chunks,{{type:'application/pdf'}})}}
async function downloadVisualPdf(name){{
  if(document.fonts?.ready)await document.fonts.ready;
  const W=1240,H=1754,S=1.5,pdfW=595,pdfH=842;
  const loadImg=src=>new Promise(resolve=>{{if(!src)return resolve(null);const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=src;}});
  const programChannels=['globo','nic','record','sbt','band','paytv'].map(k=>channelMap[k]).filter(Boolean);
  const metricKeys=['globo','nic','record','sbt','band','paytv'];
  const logoEntries=await Promise.all([...DATA.channels].map(async c=>[c.key,await loadImg(c.logoData)]));
  const logos=Object.fromEntries(logoEntries);
  const intelLogo=await loadImg(DATA.assets.inteligencia);
  const plimLogo=await loadImg(DATA.assets.plimWhite||DATA.assets.plim);
  const pages=[];
  function text(ctx,t,x,y,size,color='#101521',weight='400',align='left'){{ctx.fillStyle=color;ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;ctx.textAlign=align;ctx.fillText(String(t??''),x,y);ctx.textAlign='left'}}
  function fit(ctx,t,maxWidth,size=20,weight='700'){{let txt=String(t??''),s=size;while(s>11){{ctx.font=`${{weight}} ${{s}}px Globotipo, Arial`;if(ctx.measureText(txt).width<=maxWidth)return {{txt,size:s}};s--}}while(txt.length>3&&ctx.measureText(txt+'...').width>maxWidth)txt=txt.slice(0,-1);return {{txt:txt+'...',size:s}}}}
  function wrap(ctx,t,x,y,w,lh,size,color='#40506a',weight='400',maxLines=99){{ctx.fillStyle=color;ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;let line='',cy=y,count=0;for(const word of String(t??'').replace(/\\s+/g,' ').trim().split(' ')){{if(!word)continue;const test=line?line+' '+word:word;if(ctx.measureText(test).width>w&&line){{ctx.fillText(line,x,cy);cy+=lh;count++;line=word;if(count>=maxLines)break}}else line=test}}if(line&&count<maxLines){{ctx.fillText(line,x,cy);cy+=lh}}return cy}}
  function newPage(){{const canvas=document.createElement('canvas');canvas.width=W*S;canvas.height=H*S;const ctx=canvas.getContext('2d');ctx.scale(S,S);ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';const bg=ctx.createLinearGradient(0,0,W,H);bg.addColorStop(0,'#f8fbff');bg.addColorStop(1,'#eef4fb');ctx.fillStyle=bg;ctx.fillRect(0,0,W,H);ctx.fillStyle='#005cef';ctx.fillRect(0,0,W,178);ctx.fillStyle='rgba(255,255,255,.08)';ctx.beginPath();ctx.arc(W-90,10,290,0,Math.PI*2);ctx.fill();if(plimLogo)ctx.drawImage(plimLogo,54,48,74,74);text(ctx,'DESEMPENHO DIÁRIO',150,92,42,'#fff','700');const d=new Date(DATA.meta.date+'T00:00:00');text(ctx,DATA.meta.weekday+' • '+d.toLocaleDateString('pt-BR'),150,130,22,'rgba(255,255,255,.9)','400');if(intelLogo){{const iw=360,ih=iw/(intelLogo.width/intelLogo.height);ctx.drawImage(intelLogo,W-iw-54,74,iw,ih)}}return {{canvas,ctx}}}}
  function push(pg){{pages.push({{dataUrl:pg.canvas.toDataURL('image/jpeg',.9),width:pg.canvas.width,height:pg.canvas.height,pdfW,pdfH}})}}
  function card(ctx,x,y,w,h,accent='#005cef',fill='#fff'){{ctx.fillStyle=fill;ctx.beginPath();ctx.roundRect(x,y,w,h,14);ctx.fill();ctx.strokeStyle='#dbe4f0';ctx.lineWidth=1;ctx.stroke();ctx.fillStyle=accent;ctx.beginPath();ctx.roundRect(x,y,8,h,4);ctx.fill()}}
  function mark(ctx,c,x,y,s){{ctx.fillStyle=c.color||'#005cef';ctx.beginPath();ctx.arc(x+s/2,y+s/2,s/2,0,Math.PI*2);ctx.fill();const img=logos[c.key];if(img)ctx.drawImage(img,x+s*.18,y+s*.18,s*.64,s*.64);else text(ctx,c.textLogo||c.label,x+s/2,y+s*.62,15,'#fff','700','center')}}
  function section(ctx,label,y){{text(ctx,label,54,y,28,'#101521','700');ctx.fillStyle='#005cef';ctx.fillRect(54,y+11,72,5)}}
  function topItems(profile,key,n){{return Object.entries(profile?.[key]||{{}}).filter(x=>x[1]!=null).sort((a,b)=>b[1]-a[1]).slice(0,n).map(x=>x[0]+' ('+fmt(x[1],1)+'%)').join(', ')||'-'}}
  function profileLine(profile){{return 'Gênero: '+topItems(profile,'gender',1)+' | Classes: '+topItems(profile,'classes',2)+' | Faixas: '+topItems(profile,'ages',2)}}
  function metricCard(ctx,x,y,w,h,row){{const c=channelMap[row.key]||row;card(ctx,x,y,w,h,c.color,'#fff');mark(ctx,c,x+28,y+30,56);text(ctx,'Aud.',x+116,y+48,17,'#101521','700');text(ctx,'Share',x+286,y+48,17,'#101521','700');text(ctx,fmt(row.aud,2),x+116,y+90,34,'#101521','700');text(ctx,row.share==null?'-':fmt(row.share,2)+'%',x+286,y+90,34,c.color,'700')}}
  function leader(ctx,x,y,w,h,item,maxLead){{const c=channelMap[item.key]||item;card(ctx,x,y,w,h,c.color,'#fff');mark(ctx,c,x+18,y+14,38);text(ctx,c.label,x+64,y+29,15,'#607086','700');text(ctx,item.minutes+' min',x+64,y+57,22,'#101521','700');text(ctx,fmt(item.percent,1)+'% | '+item.hours,x+w-18,y+57,14,c.color,'700','right');ctx.fillStyle='#e8eef7';ctx.beginPath();ctx.roundRect(x+18,y+h-17,w-36,7,4);ctx.fill();ctx.fillStyle=c.color;ctx.beginPath();ctx.roundRect(x+18,y+h-17,(w-36)*Math.max(0,Math.min(1,item.minutes/(maxLead||1))),7,4);ctx.fill()}}
  function visibleLeaders(rows){{return (rows||[]).filter(l=>Number(l.minutes||0)>0)}}
  function leaderBlock(ctx,title,rows,y){{const items=visibleLeaders(rows);section(ctx,title,y);if(!items.length){{text(ctx,'Sem emissoras com minutos na liderança.',54,y+58,18,'#607086','700');return y+88}}const maxLead=Math.max(1,...items.map(l=>l.minutes||0));items.forEach((l,i)=>leader(ctx,54+(i%3)*388,y+50+Math.floor(i/3)*92,356,76,l,maxLead));return y+50+Math.ceil(items.length/3)*92+24}}
  function statusMarker(ctx,x,y,status,size=12){{const color=status==='estavel'?'#9aa4b2':status==='cresceu'?'#20c96b':'#ef3340';ctx.fillStyle=color;ctx.beginPath();ctx.arc(x,y,size/2,0,Math.PI*2);ctx.fill();return color}}
  function rankingMini(ctx,x,y,w,h,c,rows){{card(ctx,x,y,w,h,c.color,'#fff');text(ctx,c.label,x+34,y+42,22,c.color,'700');rows.slice(0,5).forEach((r,i)=>{{const yy=y+78+i*31;const st=r.shareStatus||statusFromVar(r.shareVar);const dot=statusMarker(ctx,x+38,yy-6,st,12);const f=fit(ctx,(i+1)+'. '+r.program,w-205,15,'700');text(ctx,f.txt,x+54,yy,f.size,'#101521','700');const varText=r.shareVar==null?'-':fmt(r.shareVar,1)+'%';text(ctx,fmt(r.share,2)+'% | '+varText,x+w-42,yy,14,'#607086','700','right')}})}}
  function rankingWide(ctx,x,y,w,h,c,rows){{card(ctx,x,y,w,h,c.color,'#fff');const left=x+52,right=x+w-58,audX=x+w-350,shareX=x+w-220; text(ctx,c.label,left,y+44,25,c.color,'700');text(ctx,'Programa',left,y+84,12,'#607086','700');text(ctx,'Aud.',audX,y+84,12,'#607086','700','right');text(ctx,'Share',shareX,y+84,12,'#607086','700','right');text(ctx,'Var. SHR',right,y+84,12,'#607086','700','right');rows.slice(0,8).forEach((r,i)=>{{const yy=y+116+i*24;const st=r.shareStatus||statusFromVar(r.shareVar);const dot=statusMarker(ctx,left,yy-7,st,13);const f=fit(ctx,(i+1)+'. '+r.program,w-500,15,'700');text(ctx,f.txt,left+22,yy,f.size,'#101521','700');text(ctx,fmt(r.aud,2),audX,yy,14,'#101521','700','right');text(ctx,fmt(r.share,2)+'%',shareX,yy,14,c.color,'700','right');text(ctx,r.shareVar==null?'-':fmt(r.shareVar,1)+'%',right,yy,14,dot,'700','right')}})}}
  let pg=newPage(),ctx=pg.ctx;
  section(ctx,'Resumo do Distrito Federal',252);
  metricKeys.map(k=>(DATA.summaryBars||[]).find(r=>r.key===k)).filter(Boolean).forEach((row,i)=>metricCard(ctx,54+(i%2)*592,306+Math.floor(i/2)*142,540,116,row));
  section(ctx,'Destaques do dia',756);
  const insight=(document.querySelector('#dailyInsights .insight-context')?.innerText||document.getElementById('dailyInsights')?.innerText||'').replace(/Destaques do dia\\s*/,'').trim();
  card(ctx,54,804,1132,220,'#005cef','#f8fbff');wrap(ctx,insight,90,852,1052,28,19,'#17335c','700',6);
  const afterLead=leaderBlock(ctx,'Minutos na liderança',DATA.leadership,1090);
  leaderBlock(ctx,'Minutos na liderança (exceto NIC)',leadershipExceptNic(),afterLead);
  wrap(ctx,sourceText(),54,1688,1132,22,15,'#607086','400',2);push(pg);
  pg=newPage();ctx=pg.ctx;section(ctx,'Ranking por emissora',238);
  rankingChannels.forEach((c,i)=>rankingWide(ctx,54,292+i*340,1132,312,c,DATA.rankings[c.key]||[]));
  wrap(ctx,sourceText(),54,1688,1132,22,15,'#607086','400',2);push(pg);
  let pIndex=0;const programs=DATA.programs||[];
  while(pIndex<programs.length){{pg=newPage();ctx=pg.ctx;section(ctx,'Detalhamento por programa',238);let y=292;while(pIndex<programs.length&&y<1250){{const p=programs[pIndex++];const target=Object.keys(DATA.programCompetition[p.name]||{{}}).find(t=>normalizeText(t).includes('total domic'))||DATA.targets[0];const comp=(DATA.programCompetition[p.name]||{{}})[target]||{{}};card(ctx,54,y,1132,340,'#005cef','#fff');const ft=fit(ctx,p.name,820,25,'700');text(ctx,ft.txt,82,y+42,ft.size,'#101521','700');text(ctx,'Target: '+(target||'Total Domicílios'),82,y+72,15,'#607086','700');programChannels.forEach((c,i)=>{{const col=i%3,row=Math.floor(i/3),x=82+col*352,cy=y+122+row*78;const r=comp[c.key]||{{}};mark(ctx,c,x,cy-31,44);text(ctx,c.label,x+56,cy-10,16,c.color,'700');text(ctx,'Aud. '+fmt(r.aud,2),x+56,cy+18,18,'#101521','700');text(ctx,'Shr. '+fmt(r.share,2)+'%',x+168,cy+18,18,c.color,'700')}});const py=y+274;programChannels.slice(0,3).forEach((c,i)=>{{ctx.fillStyle=i===0?'#eef6ff':i===1?'#f7efe8':'#f8fafc';ctx.beginPath();ctx.roundRect(82+i*356,py-28,330,58,10);ctx.fill();mark(ctx,c,94+i*356,py-18,30);wrap(ctx,profileLine(DATA.profile[p.name]?.[c.key]||{{}}),132+i*356,py-3,272,16,11,c.color,'700',2)}});y+=374}}wrap(ctx,sourceText(),54,1688,1132,22,15,'#607086','400',2);push(pg)}}
  const blob=makeImagePdfBlob(pages);const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1200)
}}
async function exportHtml(){{
  const clone=document.documentElement.cloneNode(true);
  clone.querySelectorAll('#baseUpload').forEach(el=>el.value='');
  clone.querySelectorAll('.final-actions,.highlights-editor,header .actions').forEach(el=>el.remove());
  clone.querySelectorAll('script').forEach(el=>{{const txt=(el.textContent||'').trimStart();const head=txt.slice(0,260);if(txt.startsWith('/*!')&&head.includes('JS'+'Zip'))el.remove()}});
  clone.querySelectorAll('.tabs .tab').forEach(btn=>btn.classList.toggle('active',btn.dataset.tab==='resumo'));
  clone.querySelectorAll('.panel').forEach(panel=>panel.classList.toggle('active',panel.id==='resumo'));
  const cloneBody=clone.querySelector('body');
  if(cloneBody)cloneBody.className=cloneBody.className.replace(/\\bawaiting-bases\\b/g,'').trim();
  const insight=clone.querySelector('#dailyInsights');
  if(insight)insight.innerHTML=buildDailyInsights(window.selectedInsightKey||'globo');
  const dataScript=clone.querySelector('#data');
  if(dataScript)dataScript.textContent=JSON.stringify(DATA).replace(/<\\//g,'<\\\\/');
  downloadBlob(`desempenho_diario_df_${{dateToken()}}.html`,'text/html;charset=utf-8','<!doctype html>\\n'+clone.outerHTML);
  try{{await downloadVisualPdf(`desempenho_diario_df_${{dateToken()}}.pdf`)}}catch(err){{console.error(err);downloadPdf(`desempenho_diario_df_${{dateToken()}}.pdf`,buildPdfLines())}}
}}
function normalizeCanvasColor(color,fallback='#fff'){{if(!color)return fallback;const raw=String(color).trim().toLowerCase();if(raw==='black'||raw==='#000'||raw==='#000000')return fallback;const m=raw.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)/);if(m){{const r=Number(m[1]),g=Number(m[2]),b=Number(m[3]);return (r+g+b)<180?fallback:`rgb(${{r}},${{g}},${{b}})`}}return raw}}
function highlightTokensFromNode(node,state={{bold:false,italic:false,color:'#fff'}},isRoot=false){{const tokens=[];if(node.nodeType===Node.TEXT_NODE){{tokens.push({{text:node.textContent||'',...state}});return tokens}}if(node.nodeName==='BR')return [{{newline:true}}];if(node.nodeType!==Node.ELEMENT_NODE)return tokens;const tag=node.tagName;const style=node.style||{{}};const isBlock=(tag==='DIV'||tag==='P')&&!isRoot;if(isBlock)tokens.push({{newline:true}});const next={{...state}};if(tag==='B'||tag==='STRONG'||Number(style.fontWeight)>=600)next.bold=true;if(tag==='I'||tag==='EM'||style.fontStyle==='italic')next.italic=true;if(style.color)next.color=normalizeCanvasColor(style.color,next.color);node.childNodes.forEach(child=>tokens.push(...highlightTokensFromNode(child,next,false)));return tokens}}
function richHighlightLines(ctx,maxWidth){{const editor=document.getElementById('highlightText');if(!editor||!editor.innerText.trim())return [];const tokens=highlightTokensFromNode(editor,{{bold:false,italic:false,color:'#fff'}},true);const lines=[[]];const fontFor=t=>`${{t.italic?'italic ':''}}${{t.bold?'700':'400'}} 28px Globotipo, Arial`;const lineWidth=line=>line.reduce((sum,t)=>{{ctx.font=fontFor(t);return sum+ctx.measureText(t.text).width}},0);const newLine=()=>{{if(lines[lines.length-1].some(t=>String(t.text).trim()))lines.push([])}};const pushChunk=(token,chunk)=>{{if(!chunk)return;ctx.font=fontFor(token);let current=lines[lines.length-1];if(lineWidth(current)+ctx.measureText(chunk).width<=maxWidth||!chunk.trim()){{current.push({{...token,text:chunk,color:token.color||'#fff'}});return}}if(current.length&&chunk.trim()){{lines.push([]);current=lines[lines.length-1]}};if(ctx.measureText(chunk).width<=maxWidth){{current.push({{...token,text:chunk,color:token.color||'#fff'}});return}};let part='';for(const char of Array.from(chunk)){{if(ctx.measureText(part+char).width>maxWidth&&part){{current.push({{...token,text:part,color:token.color||'#fff'}});lines.push([]);current=lines[lines.length-1];part=char}}else part+=char}}if(part)current.push({{...token,text:part,color:token.color||'#fff'}})}};tokens.forEach(token=>{{if(token.newline){{newLine();return}};String(token.text||'').split(/(\\n)/).forEach(chunk=>{{if(chunk==='\\n'){{newLine();return}};chunk.split(/(\\s+)/).forEach(part=>pushChunk(token,part))}})}});return lines.filter(line=>line.some(t=>String(t.text).trim())).slice(0,7)}}
function drawRichHighlight(ctx,lines,x,y,w){{if(!lines.length)return 0;const h=74+lines.length*38;ctx.fillStyle='rgba(0,30,90,.32)';ctx.beginPath();ctx.roundRect(x,y,w,h,14);ctx.fill();ctx.fillStyle='rgba(255,255,255,.92)';ctx.font='700 28px Globotipo, Arial';ctx.fillText('Destaques',x+24,y+42);let cy=y+82;lines.forEach(line=>{{let cx=x+28;line.forEach(t=>{{ctx.font=`${{t.italic?'italic ':''}}${{t.bold?'700':'400'}} 28px Globotipo, Arial`;ctx.fillStyle=t.color||'#fff';ctx.fillText(t.text,cx,cy);cx+=ctx.measureText(t.text).width}});cy+=38}});return h}}
async function generateHighlightsImage(){{if(document.fonts?.ready)await document.fonts.ready;const d=new Date(DATA.meta.date+'T00:00:00');const W=1080,H=1920,S=2;const metricKeys=['globo','nic','record','sbt'];const metricItems=metricKeys.map(key=>{{const c=channelMap[key]||{{key,label:key.toUpperCase(),color:'#005cef'}};const row=DATA.summaryBars.find(x=>x.key===key)||{{}};return {{...c,aud:row.aud,share:row.share}}}});const globoShares=(DATA.rankings.globo||[]).filter(r=>r.share!==null&&r.share!==undefined);const topShare=[...globoShares].sort((a,b)=>(b.share||0)-(a.share||0)).slice(0,3);const bottomShare=[...globoShares].sort((a,b)=>(a.share||0)-(b.share||0)).slice(0,3);const canvas=document.createElement('canvas');canvas.width=W*S;canvas.height=H*S;canvas.style.width=W+'px';canvas.style.height=H+'px';const ctx=canvas.getContext('2d');ctx.scale(S,S);ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';const grad=ctx.createLinearGradient(0,0,W,H);grad.addColorStop(0,'#003bb2');grad.addColorStop(.55,'#005cef');grad.addColorStop(1,'#65c7ff');ctx.fillStyle=grad;ctx.fillRect(0,0,W,H);ctx.fillStyle='rgba(255,255,255,.08)';ctx.beginPath();ctx.arc(850,120,420,0,Math.PI*2);ctx.fill();ctx.fillStyle='rgba(255,255,255,.06)';ctx.beginPath();ctx.arc(120,1780,390,0,Math.PI*2);ctx.fill();const loadImg=src=>new Promise(resolve=>{{if(!src)return resolve(null);const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=src;}});const intelLogo=await loadImg(DATA.assets.inteligencia);const metricLogos=Object.fromEntries(await Promise.all(metricItems.map(async item=>[item.key,await loadImg(item.logoData)])));function fitText(text,maxWidth,fontSize,weight='700'){{let size=fontSize;do{{ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;if(ctx.measureText(text).width<=maxWidth)return {{text,size}};size-=1}}while(size>=15);let clipped=text;while(clipped.length>3&&ctx.measureText(clipped+'...').width>maxWidth)clipped=clipped.slice(0,-1);return {{text:clipped+'...',size}}}}function centeredText(text,y,fontSize,weight='700',fill='#fff'){{ctx.font=`${{weight}} ${{fontSize}}px Globotipo, Arial`;ctx.fillStyle=fill;ctx.textAlign='center';ctx.fillText(text,W/2,y);ctx.textAlign='left'}}function wrapText(text,x,y,maxWidth,lineHeight,fontSize,weight='400',fill='rgba(255,255,255,.88)'){{ctx.font=`${{weight}} ${{fontSize}}px Globotipo, Arial`;ctx.fillStyle=fill;let line='',cy=y;String(text).split(/\\s+/).forEach(word=>{{const test=line?line+' '+word:word;if(ctx.measureText(test).width>maxWidth&&line){{ctx.fillText(line,x,cy);line=word;cy+=lineHeight}}else line=test}});if(line)ctx.fillText(line,x,cy);return cy}}function drawChannelMark(item,x,y,size){{ctx.fillStyle=item.color||'#005cef';ctx.beginPath();ctx.arc(x+size/2,y+size/2,size/2,0,Math.PI*2);ctx.fill();const logo=metricLogos[item.key];if(logo){{ctx.save();ctx.beginPath();ctx.arc(x+size/2,y+size/2,size*.42,0,Math.PI*2);ctx.clip();ctx.drawImage(logo,x+size*.18,y+size*.18,size*.64,size*.64);ctx.restore();}}else{{ctx.fillStyle='#fff';ctx.font='700 18px Globotipo, Arial';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(item.textLogo||item.label,x+size/2,y+size/2+1);ctx.textBaseline='alphabetic';ctx.textAlign='left'}}}}function metricCard(x,y,w,h,item){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,22);ctx.fill();ctx.fillStyle=item.color||'#005cef';ctx.fillRect(x,y,9,h);drawChannelMark(item,x+26,y+26,60);ctx.fillStyle='#607086';ctx.font='700 27px Globotipo, Arial';ctx.fillText(item.label,x+105,y+48);ctx.fillStyle='#101521';ctx.font='700 21px Globotipo, Arial';ctx.fillText('Aud.',x+105,y+94);ctx.fillText('Share',x+230,y+94);ctx.font='700 35px Globotipo, Arial';ctx.fillText(fmt(item.aud,2),x+105,y+130);ctx.fillStyle=item.color||'#005cef';ctx.fillText(fmt(item.share,2)+'%',x+230,y+130);ctx.fillStyle='#101521'}}function rankCard(x,y,w,h,title,items,accent){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,24);ctx.fill();ctx.fillStyle=accent;ctx.fillRect(x,y,9,h);ctx.fillStyle='#607086';ctx.font='700 28px Globotipo, Arial';ctx.textAlign='center';ctx.fillText(title,x+w/2,y+52);ctx.textAlign='left';items.forEach((item,i)=>{{const lineY=y+122+i*62;const fitted=fitText(`${{i+1}}. ${{item.program}}`,w-210,25,'700');ctx.fillStyle='#101521';ctx.font=`700 ${{fitted.size}}px Globotipo, Arial`;ctx.fillText(fitted.text,x+36,lineY);ctx.fillStyle=accent;ctx.font='700 25px Globotipo, Arial';ctx.textAlign='right';ctx.fillText(fmt(item.share,2)+'%',x+w-36,lineY);ctx.textAlign='left'}})}}if(intelLogo){{const logoW=560,logoH=logoW/(intelLogo.width/intelLogo.height);ctx.drawImage(intelLogo,(W-logoW)/2,70,logoW,logoH)}}centeredText('DESEMPENHO DI\u00c1RIO',218,58,'700');centeredText(`${{DATA.meta.weekday}} \u2022 ${{d.toLocaleDateString('pt-BR')}}`,286,34,'400');ctx.fillStyle='#fff';ctx.font='700 42px Globotipo, Arial';ctx.fillText('M\u00c9DIA | 07h \u00e0s 24h',80,395);const cardW=430,cardH=160;metricCard(80,470,cardW,cardH,metricItems[0]||{{}});metricCard(570,470,cardW,cardH,metricItems[1]||{{}});metricCard(80,670,cardW,cardH,metricItems[2]||{{}});metricCard(570,670,cardW,cardH,metricItems[3]||{{}});rankCard(80,900,920,320,'Melhores desempenhos (SHR%)',topShare,'#6b7280');rankCard(80,1270,920,320,'Piores desempenhos (SHR%)',bottomShare,'#374151');ctx.fillStyle='rgba(255,255,255,.94)';ctx.font='700 30px Globotipo, Arial';ctx.fillText('Relat\u00f3rio completo em anexo.',80,1708);wrapText(sourceText(),80,1760,920,28,20,'400','rgba(255,255,255,.88)');const a=document.createElement('a');a.href=canvas.toDataURL('image/png');a.download=`desempenho_diario_df_${{dateToken()}}.png`;a.click()}}
async function generateHighlightsImage(){{
  if(document.fonts?.ready)await document.fonts.ready;
  const d=new Date(DATA.meta.date+'T00:00:00');
  const W=1080,H=1920,S=2;
  const metricKeys=['globo','nic','record','sbt'];
  const metricItems=metricKeys.map(key=>{{const c=channelMap[key]||{{key,label:key.toUpperCase(),color:'#005cef'}};const row=DATA.summaryBars.find(x=>x.key===key)||{{}};return {{...c,aud:row.aud,share:row.share}}}});
  const globoShares=(DATA.rankings.globo||[]).filter(r=>r.share!==null&&r.share!==undefined);
  const performanceRanking=[...globoShares].sort((a,b)=>(b.share||0)-(a.share||0));
  const canvas=document.createElement('canvas');canvas.width=W*S;canvas.height=H*S;canvas.style.width=W+'px';canvas.style.height=H+'px';
  const ctx=canvas.getContext('2d');ctx.scale(S,S);ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality='high';
  const grad=ctx.createLinearGradient(0,0,W,H);grad.addColorStop(0,'#003bb2');grad.addColorStop(.55,'#005cef');grad.addColorStop(1,'#65c7ff');ctx.fillStyle=grad;ctx.fillRect(0,0,W,H);
  ctx.fillStyle='rgba(255,255,255,.08)';ctx.beginPath();ctx.arc(850,120,420,0,Math.PI*2);ctx.fill();
  ctx.fillStyle='rgba(255,255,255,.06)';ctx.beginPath();ctx.arc(120,1780,390,0,Math.PI*2);ctx.fill();
  const loadImg=src=>new Promise(resolve=>{{if(!src)return resolve(null);const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>resolve(null);img.src=src;}});
  const intelLogo=await loadImg(DATA.assets.inteligencia);
  const confLogo=await loadImg(DATA.assets.conf);
  const metricLogos=Object.fromEntries(await Promise.all(metricItems.map(async item=>[item.key,await loadImg(item.logoData)])));
  function fitText(text,maxWidth,fontSize,weight='700'){{let size=fontSize;do{{ctx.font=`${{weight}} ${{size}}px Globotipo, Arial`;if(ctx.measureText(text).width<=maxWidth)return {{text,size}};size-=1}}while(size>=15);let clipped=text;while(clipped.length>3&&ctx.measureText(clipped+'...').width>maxWidth)clipped=clipped.slice(0,-1);return {{text:clipped+'...',size}}}}
  function centeredText(text,y,fontSize,weight='700',fill='#fff'){{ctx.font=`${{weight}} ${{fontSize}}px Globotipo, Arial`;ctx.fillStyle=fill;ctx.textAlign='center';ctx.fillText(text,W/2,y);ctx.textAlign='left'}}
  function wrapText(text,x,y,maxWidth,lineHeight,fontSize,weight='400',fill='rgba(255,255,255,.88)'){{ctx.font=`${{weight}} ${{fontSize}}px Globotipo, Arial`;ctx.fillStyle=fill;let line='',cy=y;String(text).split(/\\s+/).forEach(word=>{{const test=line?line+' '+word:word;if(ctx.measureText(test).width>maxWidth&&line){{ctx.fillText(line,x,cy);line=word;cy+=lineHeight}}else line=test}});if(line)ctx.fillText(line,x,cy);return cy}}
  function drawChannelMark(item,x,y,size){{ctx.fillStyle=item.color||'#005cef';ctx.beginPath();ctx.arc(x+size/2,y+size/2,size/2,0,Math.PI*2);ctx.fill();const logo=metricLogos[item.key];if(logo){{ctx.save();ctx.beginPath();ctx.arc(x+size/2,y+size/2,size*.42,0,Math.PI*2);ctx.clip();ctx.drawImage(logo,x+size*.18,y+size*.18,size*.64,size*.64);ctx.restore();}}else{{ctx.fillStyle='#fff';ctx.font='700 18px Globotipo, Arial';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(item.textLogo||item.label,x+size/2,y+size/2+1);ctx.textBaseline='alphabetic';ctx.textAlign='left'}}}}
  function metricCard(x,y,w,h,item){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,22);ctx.fill();ctx.fillStyle=item.color||'#005cef';ctx.beginPath();ctx.roundRect(x+18,y+24,7,h-48,4);ctx.fill();drawChannelMark(item,x+52,y+37,56);ctx.fillStyle='#101521';ctx.font='700 19px Globotipo, Arial';ctx.fillText('Aud.',x+132,y+58);ctx.fillText('Share',x+265,y+58);ctx.font='700 33px Globotipo, Arial';ctx.fillText(fmt(item.aud,2),x+132,y+95);ctx.fillStyle=item.color||'#005cef';ctx.fillText(fmt(item.share,2)+'%',x+265,y+95);ctx.fillStyle='#101521'}}
  function rankCard(x,y,w,h,title,items,accent){{ctx.fillStyle='rgba(255,255,255,.94)';ctx.beginPath();ctx.roundRect(x,y,w,h,24);ctx.fill();ctx.fillStyle=accent;ctx.beginPath();ctx.roundRect(x+18,y+22,7,h-44,4);ctx.fill();ctx.fillStyle='#607086';ctx.font='700 28px Globotipo, Arial';ctx.textAlign='center';ctx.fillText(title,x+w/2,y+52);ctx.textAlign='left';items.forEach((item,i)=>{{const lineY=y+103+i*43;const st=item.shareStatus||statusFromVar(item.shareVar);const dotColor=st==='estavel'?'#9aa4b2':st==='cresceu'?'#27d66f':'#ef3340';const dg=ctx.createRadialGradient(x+52,lineY-7,3,x+52,lineY-7,11);dg.addColorStop(0,'#fff');dg.addColorStop(1,dotColor);ctx.fillStyle=dg;ctx.beginPath();ctx.arc(x+52,lineY-7,10,0,Math.PI*2);ctx.fill();const fitted=fitText(`${{i+1}}. ${{item.program}}`,w-270,22,'700');ctx.fillStyle='#101521';ctx.font=`700 ${{fitted.size}}px Globotipo, Arial`;ctx.fillText(fitted.text,x+74,lineY);ctx.fillStyle=accent;ctx.font='700 22px Globotipo, Arial';ctx.textAlign='right';ctx.fillText(fmt(item.share,2)+'%',x+w-36,lineY);ctx.textAlign='left'}});const legendY=y+h-42;const lg=ctx.createRadialGradient(x+58,legendY-7,3,x+58,legendY-7,10);lg.addColorStop(0,'#fff');lg.addColorStop(1,'#9aa4b2');ctx.fillStyle=lg;ctx.beginPath();ctx.arc(x+58,legendY-7,9,0,Math.PI*2);ctx.fill();ctx.fillStyle='#607086';ctx.font='700 22px Globotipo, Arial';ctx.fillText('Comparativo com 4 semanas anteriores.',x+78,legendY)}}
  if(intelLogo){{const logoW=560,logoH=logoW/(intelLogo.width/intelLogo.height);ctx.drawImage(intelLogo,(W-logoW)/2,70,logoW,logoH)}}
  if(confLogo){{const confW=86,confH=confW/(confLogo.width/confLogo.height);ctx.drawImage(confLogo,W-confW-72,62,confW,confH)}}
  centeredText('DESEMPENHO DIÁRIO',218,58,'700');
  centeredText(`${{DATA.meta.weekday}} • ${{d.toLocaleDateString('pt-BR')}}`,286,34,'400');
  const highlightLines=richHighlightLines(ctx,840);
  let mediaY=395;
  if(highlightLines.length)mediaY=348+drawRichHighlight(ctx,highlightLines,80,348,920)+74;
  ctx.fillStyle='#fff';ctx.font='700 42px Globotipo, Arial';ctx.fillText('MÉDIA | 07h às 24h',80,mediaY);
  const cardW=430,cardH=128,metricY1=mediaY+58,metricY2=metricY1+142;
  metricCard(80,metricY1,cardW,cardH,metricItems[0]||{{}});
  metricCard(570,metricY1,cardW,cardH,metricItems[1]||{{}});
  metricCard(80,metricY2,cardW,cardH,metricItems[2]||{{}});
  metricCard(570,metricY2,cardW,cardH,metricItems[3]||{{}});
  const rankH=154+performanceRanking.length*43,rankY1=metricY2+176;
  rankCard(80,rankY1,920,rankH,'Ranking de desempenho (SHR%)',performanceRanking,'#6b7280');
  const footerY=rankY1+rankH+62;
  ctx.fillStyle='rgba(255,255,255,.94)';ctx.font='700 30px Globotipo, Arial';ctx.fillText('Relatório completo em anexo.',80,footerY);
  const sourceBottom=wrapText(sourceText(),80,footerY+52,920,28,20,'400','rgba(255,255,255,.88)');
  const cropH=Math.min(H,Math.max(1280,Math.ceil(sourceBottom+56)));
  const out=document.createElement('canvas');out.width=W*S;out.height=cropH*S;const outCtx=out.getContext('2d');outCtx.drawImage(canvas,0,0,W*S,cropH*S,0,0,W*S,cropH*S);
  const a=document.createElement('a');a.href=out.toDataURL('image/png');a.download=`desempenho_diario_df_${{dateToken()}}.png`;a.click();
}}
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
function readTotalProfileJs(rows){{const out={{}};const total=(rows||[]).find(r=>r&&r[0]==='Total');if(!total)return out;const headerRows={{target:3,entity:4}};const maxCols=Math.max(...rows.slice(0,headerRows.entity+1).map(r=>r.length));for(let i=1;i<maxCols;i++){{const target=(rows[headerRows.target]||[])[i]||'';const entity=(rows[headerRows.entity]||[])[i]||'';const key=channelForRaw(entity);if(target&&key){{out[target]??={{}};out[target][key]={{adh:fnumJs(total[i])}}}}}}return out}}
function loadMinuteDataJs(rows){{const minutes=[];const targets=rows[1]||[], vars=rows[2]||[], entities=rows[3]||[];const maxCols=Math.max(targets.length,vars.length,entities.length);for(const row of rows.slice(5)){{const minute=parseTimebandJs(row[0]);if(minute===null)continue;const item={{time:hhmmFromMinJs(minute),minute,aud:{{}},share:{{}}}};for(let i=1;i<maxCols;i++){{if(!normalizeText(targets[i]).includes('total domic'))continue;const metric=metricFromVarJs(vars[i]);if(metric!=='aud'&&metric!=='share')continue;const key=channelForRaw(entities[i]);if(key)item[metric][key]=fnumJs(row[i])}}minutes.push(item)}}return minutes}}
function loadProgramsJs(programRows,crosstabRows){{const programs=[];const targets=programRows[1]||[], vars=programRows[2]||[];for(const row of programRows.slice(4)){{if(!row||!row[0])continue;const p={{name:row[0],start:fracToMinutesJs(row[1]),end:fracToMinutesJs(row[2]),targets:{{}}}};for(let i=3;i<targets.length;i++){{const target=targets[i], metric=metricFromVarJs(vars[i]);if(metric==='aud'||metric==='share'){{p.targets[target]??={{}};p.targets[target][metric]=fnumJs(row[i])}}}}programs.push(p)}}return [programs,readMatrixJs(crosstabRows,{{target:2,var:3,entity:4,start:6}},['aud','share'])]}}
function loadRankingsJs(currentBook,previousBook){{const map={{globo:'Globo',record:'Record',sbt:'SBT',band:'BAND'}}, rankings={{}};for(const [key,sheet] of Object.entries(map)){{const cur=currentBook.sheets[sheet]||[], prev=previousBook.sheets[sheet]||[], prevBy={{}};for(const row of prev.slice(3))if(row[1])prevBy[row[1]]={{aud:fnumJs(row[2]),share:fnumJs(row[3])}};rankings[key]=cur.slice(3,13).filter(r=>r[1]).map((row,idx)=>{{const aud=fnumJs(row[2]),share=fnumJs(row[3]),p=prevBy[row[1]]||{{}};const variation=(now,before)=>now===null||before===null||before===undefined||before===0?null:(now-before)/before*100;const audVar=variation(aud,p.aud), shareVar=variation(share,p.share);return {{rank:idx+1,program:row[1],aud,audPrev:p.aud,audVar,audStatus:statusFromVar(audVar),share,sharePrev:p.share,shareVar,shareStatus:statusFromVar(shareVar)}}}})}}return rankings}}
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
  const profileTotal=readTotalProfileJs(profileBook.sheets.Crosstab1||[]);
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
  const totalAdh=(target,key)=>{{const bucket=profileTotal[target]?.[key]||{{}};return bucket.adh??Object.values(bucket)[0]}};
  const profileData={{}};
  profileData['07h-24h']={{}};
  for(const ch of DATA.leadershipChannels){{
    const metrics=summaryBars.find(r=>r.key===ch.key)||{{}};
    profileData['07h-24h'][ch.key]={{aud:metrics.aud,share:metrics.share,gender:{{Homem:totalAdh('Masculino',ch.key),Mulher:totalAdh('Feminino',ch.key)}},classes:{{AB1:totalAdh('AB1',ch.key),B2:totalAdh('B2',ch.key),C1:totalAdh('C1',ch.key),C2:totalAdh('C2',ch.key),DE:totalAdh('DE',ch.key)}},ages:{{'4-11':totalAdh('4-11 anos',ch.key),'12-17':totalAdh('12-17 anos',ch.key),'18-24':totalAdh('18-24 anos',ch.key),'25-34':totalAdh('25-34 anos',ch.key),'35-49':totalAdh('35-49 anos',ch.key),'50+':totalAdh('50+',ch.key)}}}};
  }}
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
async function handleUploadedBases(fileList){{fileStatus.style.display='block';fileStatus.textContent='Lendo bases...';const books=[];for(const file of [...fileList]){{const lower=file.name.toLowerCase();if(lower.endsWith('.xlsx'))books.push(await readXlsxFile(file));else if(lower.endsWith('.zip')){{const pack=await JSZip.loadAsync(file);for(const entry of Object.values(pack.files)){{if(!entry.dir&&entry.name.toLowerCase().endsWith('.xlsx')){{const bytes=await entry.async('uint8array');books.push(await readXlsxFile(bytes,entry.name.split(/[\\/]/).pop()))}}}}}}}}DATA=buildDataFromBooks(books);fillSelect(programSelect,DATA.programs.map(p=>p.name));fillSelect(profileSelect,profileChoices());profileSelect.value='07h-24h';fillSelect(targetSelect,DATA.targets);fileStatus.textContent='Bases carregadas';revealDashboard()}}
function activateTab(btn){{if(!btn||!btn.dataset||!btn.dataset.tab)return;document.querySelectorAll('.tabs .tab').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));btn.classList.add('active');document.getElementById(btn.dataset.tab).classList.add('active');setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil()}},0)}}
function routeNavEvent(e){{const target=e.target instanceof Element?e.target:e.target?.parentElement;if(!target)return;const tab=target.closest('.tabs .tab');if(tab){{activateTab(tab);return}}const insight=target.closest('.insight-tab');if(insight){{renderDailyInsights(insight.dataset.insight)}}}}
document.addEventListener('click',routeNavEvent);
function routeControlEvent(e){{const id=e.target&&e.target.id;if(id==='programSelect'||id==='targetSelect')requestAnimationFrame(renderProgramas);if(id==='profileSelect')requestAnimationFrame(renderPerfil)}}
document.addEventListener('change',routeControlEvent,true);
document.addEventListener('input',routeControlEvent,true);
const programSelect=document.getElementById('programSelect'),targetSelect=document.getElementById('targetSelect'),profileSelect=document.getElementById('profileSelect'),programTitle=document.getElementById('programTitle'),profileTitle=document.getElementById('profileTitle'),minuteTable=document.getElementById('minuteTable'),profileGrid=document.getElementById('profileGrid'),baseUpload=document.getElementById('baseUpload'),fileStatus=document.getElementById('fileStatus');
fillSelect(programSelect,DATA.programs.map(p=>p.name));fillSelect(profileSelect,profileChoices());profileSelect.value='07h-24h';fillSelect(targetSelect,DATA.targets);
programSelect.addEventListener('change',renderProgramas);targetSelect.addEventListener('change',renderProgramas);profileSelect.addEventListener('change',renderPerfil);window.addEventListener('resize',()=>{{renderResumo();renderProgramas();renderPerfil()}});
function revealDashboard(){{document.body.className=document.body.className.replace(/\\bawaiting-bases\\b/g,'').trim();const actions=document.querySelector('header .actions');if(actions)actions.style.display='none';setTimeout(()=>{{renderResumo();renderProgramas();renderPerfil();window.scrollTo({{top:0,behavior:'smooth'}})}},0)}}
const uploadButton=document.getElementById('uploadBtn'),exportButton=document.getElementById('exportBtn'),imageButton=document.getElementById('imageBtn');
if(uploadButton&&baseUpload)uploadButton.addEventListener('click',()=>baseUpload.click());if(baseUpload)baseUpload.addEventListener('change',async()=>{{try{{const files=[...baseUpload.files].map(f=>f.name);fileStatus.textContent=files.length?`${{files.length}} arquivo(s): ${{files.join(', ')}}`:'';if(files.length)await handleUploadedBases(baseUpload.files)}}catch(err){{if(fileStatus){{fileStatus.style.display='block';fileStatus.textContent='Erro ao ler bases: '+err.message}}console.error(err)}}}});if(exportButton)exportButton.addEventListener('click',exportHtml);if(imageButton)imageButton.addEventListener('click',generateHighlightsImage);
document.querySelectorAll('[data-command]').forEach(btn=>btn.addEventListener('click',()=>{{document.getElementById('highlightText')?.focus();document.execCommand(btn.dataset.command,false,null)}}));
document.querySelectorAll('[data-insert]').forEach(btn=>btn.addEventListener('click',()=>{{document.getElementById('highlightText')?.focus();document.execCommand('insertText',false,btn.dataset.insert)}}));
renderResumo();renderProgramas();renderPerfil();
</script>
</body>
</html>"""


if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    data = build_data()
    OUT.write_text(render_html(data), encoding="utf-8")
    print(OUT)




