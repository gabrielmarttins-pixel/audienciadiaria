import html
import json
import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "exemplo_desempenho_diario_visual.pdf"
ASSETS = ROOT / "assets"
HTML = (ROOT / "outputs" / "index.html").read_text(encoding="utf-8")

match = re.search(r'<script id="data" type="application/json">(.*?)</script>', HTML, re.S)
if not match:
    raise SystemExit("DATA json not found")
DATA = json.loads(html.unescape(match.group(1)))

pdfmetrics.registerFont(TTFont("Globotipo", str(ASSETS / "GlobotipoCorporativa-Regular.ttf")))
pdfmetrics.registerFont(TTFont("GlobotipoBold", str(ASSETS / "GlobotipoCorporativa-Bold.ttf")))

PAGE_W, PAGE_H = A4
M = 16 * mm
BLUE = colors.HexColor("#005cef")
BG = colors.HexColor("#f4f7fb")
INK = colors.HexColor("#101521")
MUTED = colors.HexColor("#607086")
LINE = colors.HexColor("#dbe4f0")

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="TitleGlobo", fontName="GlobotipoBold", fontSize=24, leading=28, textColor=BLUE))
styles.add(ParagraphStyle(name="H1Globo", fontName="GlobotipoBold", fontSize=18, leading=22, textColor=INK, spaceBefore=8, spaceAfter=8))
styles.add(ParagraphStyle(name="H2Globo", fontName="GlobotipoBold", fontSize=13, leading=16, textColor=INK, spaceBefore=6, spaceAfter=6))
styles.add(ParagraphStyle(name="BodyGlobo", fontName="Globotipo", fontSize=9.5, leading=12.5, textColor=colors.HexColor("#40506a")))
styles.add(ParagraphStyle(name="SmallGlobo", fontName="Globotipo", fontSize=7.5, leading=9.5, textColor=MUTED))
styles.add(ParagraphStyle(name="Cell", fontName="Globotipo", fontSize=7.5, leading=9, textColor=INK))
styles.add(ParagraphStyle(name="CellBold", fontName="GlobotipoBold", fontSize=7.5, leading=9, textColor=INK))

channel_map = {c["key"]: c for c in DATA["channels"]}
for c in DATA["channels"]:
    c["pdfColor"] = colors.HexColor(c.get("color") or "#607086")

logo_files = {
    "globo": ASSETS / "GLOBO.png",
    "record": ASSETS / "RECORD.png",
    "sbt": ASSETS / "SBT.png",
    "band": ASSETS / "BAND.png",
    "plim": ASSETS / "PLIMPLIM_BRANCO.png",
    "intel": ASSETS / "inteligencia.png",
}


def fmt(value, digits=2):
    if value is None:
        return "-"
    try:
        return f"{float(value):.{digits}f}".replace(".", ",")
    except Exception:
        return "-"


def esc(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def p(text, style="BodyGlobo"):
    return Paragraph(esc(text), styles[style])


def header(canvas, _doc):
    canvas.saveState()
    canvas.setFillColor(BG)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.setFillColor(BLUE)
    canvas.rect(0, PAGE_H - 30 * mm, PAGE_W, 30 * mm, fill=1, stroke=0)
    if logo_files["plim"].exists():
        canvas.drawImage(ImageReader(str(logo_files["plim"])), M, PAGE_H - 23 * mm, 15 * mm, 15 * mm, mask="auto")
    canvas.setFont("GlobotipoBold", 19)
    canvas.setFillColor(colors.white)
    canvas.drawString(M + 20 * mm, PAGE_H - 15 * mm, "DESEMPENHO DIÁRIO")
    canvas.setFont("Globotipo", 9)
    canvas.drawString(M + 20 * mm, PAGE_H - 21 * mm, f"{DATA['meta']['weekday']} - {date_br()}")
    if logo_files["intel"].exists():
        canvas.drawImage(ImageReader(str(logo_files["intel"])), PAGE_W - 62 * mm, PAGE_H - 18 * mm, 48 * mm, 7 * mm, mask="auto")
    canvas.restoreState()


def date_br():
    date = DATA["meta"]["date"]
    return f"{date[8:10]}/{date[5:7]}/{date[:4]}"


def table_style(accent=BLUE):
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eaf2ff")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17335c")),
        ("FONTNAME", (0, 0), (-1, 0), "GlobotipoBold"),
        ("FONTNAME", (0, 1), (-1, -1), "Globotipo"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("LEADING", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBEFORE", (0, 0), (0, -1), 3, accent),
    ])


def metrics_table():
    rows = [[p("Emissora", "CellBold"), p("Audiência", "CellBold"), p("Share", "CellBold")]]
    for key in ["globo", "nic", "record", "sbt", "band", "paytv", "tle"]:
        row = next((r for r in DATA.get("summaryBars", []) if r["key"] == key), None)
        if not row:
            continue
        c = channel_map.get(key, row)
        rows.append([p(c["label"], "CellBold"), p(fmt(row.get("aud")), "CellBold"), p("-" if row.get("share") is None else fmt(row.get("share")) + "%", "CellBold")])
    table = Table(rows, colWidths=[70 * mm, 42 * mm, 42 * mm])
    table.setStyle(table_style(BLUE))
    return table


def leadership_table():
    rows = [[p("Emissora", "CellBold"), p("Minutos", "CellBold"), p("%", "CellBold"), p("Horas", "CellBold")]]
    for item in DATA.get("leadership", []):
        rows.append([p(item["label"], "CellBold"), p(item["minutes"], "Cell"), p(fmt(item.get("percent"), 1) + "%", "Cell"), p(item.get("hours", "-"), "Cell")])
    table = Table(rows, colWidths=[55 * mm, 32 * mm, 32 * mm, 35 * mm])
    table.setStyle(table_style(MUTED))
    return table


def ranking_tables():
    output = []
    for key in ["globo", "record", "sbt", "band"]:
        c = channel_map[key]
        rows = [[p("#", "CellBold"), p("Programa", "CellBold"), p("Aud.", "CellBold"), p("Méd. 4 sem.", "CellBold"), p("Var.", "CellBold"), p("Share", "CellBold"), p("Méd. 4 sem.", "CellBold"), p("Var.", "CellBold")]]
        for r in DATA.get("rankings", {}).get(key, [])[:10]:
            rows.append([
                p(r.get("rank", ""), "Cell"),
                p(r.get("program", ""), "CellBold"),
                p(fmt(r.get("aud")), "Cell"),
                p(fmt(r.get("audPrev")), "Cell"),
                p("-" if r.get("audVar") is None else fmt(r.get("audVar"), 1) + "%", "Cell"),
                p(fmt(r.get("share")) + "%", "Cell"),
                p(fmt(r.get("sharePrev")) + "%", "Cell"),
                p("-" if r.get("shareVar") is None else fmt(r.get("shareVar"), 1) + "%", "Cell"),
            ])
        table = Table(rows, colWidths=[10 * mm, 48 * mm, 18 * mm, 22 * mm, 18 * mm, 18 * mm, 22 * mm, 18 * mm])
        table.setStyle(table_style(c["pdfColor"]))
        output += [p(c["label"], "H2Globo"), table, Spacer(1, 5 * mm)]
    return output


def top_items(profile, key, count):
    vals = profile.get(key, {}) if isinstance(profile, dict) else {}
    items = sorted([(k, v) for k, v in vals.items() if v is not None], key=lambda x: x[1], reverse=True)[:count]
    return ", ".join(f"{k} ({fmt(v, 1)}%)" for k, v in items) or "-"


def profile_text(profile):
    return f"Gênero: {top_items(profile, 'gender', 1)} | Classes: {top_items(profile, 'classes', 2)} | Faixas: {top_items(profile, 'ages', 2)}"


def program_blocks():
    output = []
    channels = ["globo", "nic", "record", "sbt", "band", "paytv"]
    for program in DATA.get("programs", []):
        target = next((t for t in DATA.get("programCompetition", {}).get(program["name"], {}).keys() if "total domic" in t.lower()), DATA.get("targets", ["Total Domicílios"])[0])
        comp = DATA.get("programCompetition", {}).get(program["name"], {}).get(target, {})
        rows = [[p("Emissora", "CellBold"), p("Aud.", "CellBold"), p("Share", "CellBold"), p("Perfil", "CellBold")]]
        for key in channels:
            c = channel_map.get(key, {"label": key.upper()})
            metric = comp.get(key, {})
            profile = DATA.get("profile", {}).get(program["name"], {}).get(key, {})
            rows.append([p(c["label"], "CellBold"), p(fmt(metric.get("aud")), "Cell"), p(fmt(metric.get("share")) + "%", "Cell"), p(profile_text(profile), "Cell")])
        table = Table(rows, colWidths=[28 * mm, 18 * mm, 20 * mm, 104 * mm])
        table.setStyle(table_style(BLUE))
        output.append(KeepTogether([p(program["name"], "H2Globo"), p("Target: " + target, "SmallGlobo"), table, Spacer(1, 5 * mm)]))
    return output


story = [p("Resumo do Distrito Federal", "H1Globo"), metrics_table(), Spacer(1, 6 * mm)]
story += [p("Destaques do dia", "H1Globo")]
story += [p(f"A Globo fechou o período 07h-24h com média de {fmt(DATA['meta']['dailyAvg'])} pontos. O relatório em PDF consolida ranking, liderança e detalhamento por programa com audiência, share e perfil de público.", "BodyGlobo")]
story += [Spacer(1, 6 * mm), p("Minutos na liderança", "H1Globo"), leadership_table(), Spacer(1, 6 * mm), p("Ranking por emissora", "H1Globo")]
story += ranking_tables()
story += [PageBreak(), p("Detalhamento por programa", "H1Globo")]
story += program_blocks()
story += [Spacer(1, 6 * mm), p(f"Fonte: Ibope. Instar Analytics. DF. Total Domicílios. Aud%, Shr% Adh%. Atividades: Live. Data: {date_br()}.", "SmallGlobo")]

doc = SimpleDocTemplate(str(OUT), pagesize=A4, rightMargin=M, leftMargin=M, topMargin=38 * mm, bottomMargin=12 * mm)
doc.build(story, onFirstPage=header, onLaterPages=header)
print(OUT)
