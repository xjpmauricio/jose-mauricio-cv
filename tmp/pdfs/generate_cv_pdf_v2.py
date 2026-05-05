import html
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

DOCX_PATH = Path(r"C:\Users\bluem\Cvs\Jose-Mauricio-CV-ENG-2025.docx")
PHOTO_PATH = Path(r"C:\Users\bluem\Cvs\codex-cv\jose-mauricio-49.jpg")
OUT_MAIN = Path(r"C:\Users\bluem\Cvs\codex-cv\index-en.pdf")
OUT_ALT = Path(r"C:\Users\bluem\Cvs\codex-cv\output\pdf\jose-mauricio-cv-en.pdf")


def register_fonts():
    font_candidates = {
        "body": [
            ("Manrope", Path(r"C:\Windows\Fonts\segoeui.ttf")),
            ("Helvetica", None),
        ],
        "body_bold": [
            ("Manrope-Bold", Path(r"C:\Windows\Fonts\segoeuib.ttf")),
            ("Helvetica-Bold", None),
        ],
        "heading": [
            ("Fraunces-Alt", Path(r"C:\Windows\Fonts\georgiab.ttf")),
            ("Times-Bold", None),
        ],
    }

    resolved = {}
    for key, options in font_candidates.items():
        chosen = None
        for name, path in options:
            if path is None:
                chosen = name
                break
            if path.exists():
                if name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(name, str(path)))
                chosen = name
                break
        resolved[key] = chosen
    return resolved


def extract_docx_lines(path: Path):
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    xml = xml.replace("</w:p>", "</w:p>\n")
    xml = xml.replace("<w:tab/>", "\t").replace("<w:br/>", "\n").replace("<w:cr/>", "\n")
    text = re.sub(r"<[^>]+>", "", xml)
    text = html.unescape(text)
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    return [ln for ln in lines if ln]


def norm(s: str):
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.lower().strip())


def clean_text(text: str):
    if not text:
        return ""
    t = re.sub(r"\s+", " ", str(text)).strip()
    replacements = {
        "results..": "results.",
        "OpenAI Vision AI": "OpenAI Vision",
        "An hotel management platform": "A hotel management platform",
        "jQueryUI": "jQuery UI",
        "Sénior": "Senior",
        "Junho 2007 a Junho 2009": "June 2007 to June 2009",
        "ASP Classico": "Classic ASP",
        "–": "-",
        "—": "-",
        "‑": "-",
        "’": "'",
        "“": '"',
        "”": '"',
        "\u00a0": " ",
    }
    for old, new in replacements.items():
        t = t.replace(old, new)
    return t


def find_heading(lines, heading: str):
    target = norm(heading)
    for i, line in enumerate(lines):
        if norm(line) == target:
            return i
    raise ValueError(f"Heading not found: {heading}")


def split_csv_smart(value: str):
    parts = []
    buf = []
    depth = 0
    for ch in (value or ""):
        if ch == "(":
            depth += 1
        elif ch == ")" and depth > 0:
            depth -= 1
        if ch == "," and depth == 0:
            token = "".join(buf).strip()
            if token:
                parts.append(token)
            buf = []
        else:
            buf.append(ch)
    tail = "".join(buf).strip()
    if tail:
        parts.append(tail)
    return parts


def parse_content(lines):
    idx_personal = find_heading(lines, "Personal Details")
    idx_projects = find_heading(lines, "Most Relevant Projects")
    idx_education = find_heading(lines, "Education")
    idx_experience = find_heading(lines, "Summary of Professional Experience")
    idx_tech = find_heading(lines, "Technologies Summary")
    idx_languages = find_heading(lines, "Spoken Languages")
    idx_about = find_heading(lines, "About Me")

    data = {}
    data["name"] = lines[0]
    data["main_role"] = ""
    data["other_roles"] = ""

    for i, line in enumerate(lines[:idx_personal]):
        if norm(line) == "main role" and i + 1 < idx_personal:
            data["main_role"] = lines[i + 1]
        if norm(line) == "other roles" and i + 1 < idx_personal:
            data["other_roles"] = lines[i + 1]

    personal = {}
    key_map = {
        "name": "Name",
        "mobile phone": "Mobile Phone",
        "email": "Email",
        "linkedin": "LinkedIn",
        "date of birth": "Date of Birth",
        "hobby": "Hobby",
    }
    i = idx_personal + 1
    while i < idx_projects:
        key = norm(lines[i])
        if key in key_map and i + 1 < idx_projects:
            personal[key_map[key]] = lines[i + 1]
            i += 2
        else:
            i += 1
    data["personal"] = personal

    projects = []
    project_lines = lines[idx_projects + 1:idx_education]
    i = 0
    while i < len(project_lines):
        if norm(project_lines[i]) == "start:":
            p = {
                "start": project_lines[i + 1] if i + 1 < len(project_lines) else "",
                "title": project_lines[i + 2] if i + 2 < len(project_lines) else "",
                "end": "",
                "description": [],
                "client": "",
                "role": "",
                "team": "",
                "technologies": "",
            }
            i += 3
            if i < len(project_lines) and norm(project_lines[i]) == "end:":
                p["end"] = project_lines[i + 1] if i + 1 < len(project_lines) else ""
                i += 2
            while i < len(project_lines) and norm(project_lines[i]) != "client:":
                p["description"].append(project_lines[i])
                i += 1
            if i < len(project_lines) and norm(project_lines[i]) == "client:":
                p["client"] = project_lines[i + 1] if i + 1 < len(project_lines) else ""
                i += 2
            if i < len(project_lines) and norm(project_lines[i]) == "role:":
                p["role"] = project_lines[i + 1] if i + 1 < len(project_lines) else ""
                i += 2
            if i < len(project_lines) and norm(project_lines[i]) == "team:":
                p["team"] = project_lines[i + 1] if i + 1 < len(project_lines) else ""
                i += 2
            if i < len(project_lines) and norm(project_lines[i]) == "technologies":
                p["technologies"] = project_lines[i + 1] if i + 1 < len(project_lines) else ""
                i += 2
            projects.append(p)
        else:
            i += 1
    data["projects"] = projects

    period_re = re.compile(r"^\d{4}(\s*[^\d\w]+\s*\d{4})?$")
    edu = []
    edu_lines = lines[idx_education + 1:idx_experience]
    i = 0
    while i < len(edu_lines):
        if period_re.match(norm(edu_lines[i])):
            e = {
                "period": edu_lines[i],
                "course": edu_lines[i + 1] if i + 1 < len(edu_lines) else "",
                "description": "",
            }
            i += 2
            desc = []
            while i < len(edu_lines) and not period_re.match(norm(edu_lines[i])):
                desc.append(edu_lines[i])
                i += 1
            e["description"] = " ".join(desc)
            edu.append(e)
        else:
            i += 1
    data["education"] = edu

    exp_lines = lines[idx_experience + 1:idx_tech]

    def is_period(s: str):
        n = norm(s)
        return bool(re.search(r"\d{4}", n) and ((" to " in n) or (" a " in n) or (" ate " in n)))

    exp = []
    i = 0
    while i < len(exp_lines):
        if is_period(exp_lines[i]):
            e = {
                "period": exp_lines[i],
                "company": exp_lines[i + 1] if i + 1 < len(exp_lines) else "",
                "role": exp_lines[i + 2] if i + 2 < len(exp_lines) else "",
                "description": "",
            }
            i += 3
            desc = []
            while i < len(exp_lines) and not is_period(exp_lines[i]):
                desc.append(exp_lines[i])
                i += 1
            e["description"] = " ".join(desc)
            exp.append(e)
        else:
            i += 1
    data["experience"] = exp

    tech_lines = lines[idx_tech + 1:idx_languages]
    tech = []
    i = 0
    cat_map = {
        "Frameworks de Javascript": "JavaScript Frameworks",
        "Gestores de Conteúdos": "Content Management Systems",
    }
    while i < len(tech_lines) - 1:
        category = cat_map.get(tech_lines[i], tech_lines[i])
        tech.append({"category": category, "items": tech_lines[i + 1]})
        i += 2
    data["technology"] = tech

    lang_lines = lines[idx_languages + 1:idx_about]
    lang_map = {
        "Português": "Portuguese",
        "Inglês": "English",
        "Francês": "French",
        "Espanhol": "Spanish",
    }
    langs = []
    i = 0
    while i < len(lang_lines) - 1:
        langs.append({"level": lang_lines[i], "language": lang_map.get(lang_lines[i + 1], lang_lines[i + 1])})
        i += 2
    data["languages"] = langs

    about_lines = lines[idx_about + 1:]
    about = []
    signature = ""
    for line in about_lines:
        if line == ".":
            continue
        if norm(line).startswith("jose mauricio"):
            signature = line
        else:
            about.append(line)
    data["about"] = about
    data["signature"] = signature
    return data


def fmt_date(raw: str):
    months = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
        "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
    }
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", str(raw).strip())
    if not m:
        return clean_text(raw)
    year, month, _ = m.groups()
    return f"{months.get(month, month)} {year}"


def card(content, width, border="#E2E8F0", bg="#FFFFFF", pad=9):
    tbl = Table([[content]], colWidths=[width], hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(bg)),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(border)),
        ("LEFTPADDING", (0, 0), (-1, -1), pad),
        ("RIGHTPADDING", (0, 0), (-1, -1), pad),
        ("TOPPADDING", (0, 0), (-1, -1), pad - 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad - 1),
    ]))
    return tbl


def grid(cards, cols, col_width, hpad=6):
    rows = []
    for i in range(0, len(cards), cols):
        row = cards[i:i + cols]
        if len(row) < cols:
            row += [Spacer(1, 1)] * (cols - len(row))
        rows.append(row)
    tbl = Table(rows, colWidths=[col_width] * cols, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), hpad),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), hpad),
    ]))
    return tbl


def build_story(data, fonts):
    palette = {
        "ink": colors.HexColor("#0F172A"),
        "muted": colors.HexColor("#475569"),
        "line": colors.HexColor("#E2E8F0"),
        "primary": colors.HexColor("#0F766E"),
        "accent": colors.HexColor("#EA580C"),
        "paper": colors.HexColor("#FFFFFF"),
        "soft": colors.HexColor("#FFF8F2"),
        "chip": colors.HexColor("#F8FAFC"),
    }

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="HeroName", fontName=fonts["heading"], fontSize=24, leading=26, textColor=palette["ink"]))
    styles.add(ParagraphStyle(name="Role", fontName=fonts["body_bold"], fontSize=11.3, leading=13.5, textColor=palette["primary"]))
    styles.add(ParagraphStyle(name="Tiny", fontName=fonts["body"], fontSize=8.25, leading=10.2, textColor=palette["muted"]))
    styles.add(ParagraphStyle(name="Section", fontName=fonts["heading"], fontSize=15.2, leading=18.2, textColor=palette["ink"], spaceAfter=2))
    styles.add(ParagraphStyle(name="SectionSub", fontName=fonts["body"], fontSize=8.9, leading=10.8, textColor=palette["muted"], spaceAfter=5))
    styles.add(ParagraphStyle(name="CardTitle", fontName=fonts["body_bold"], fontSize=10.6, leading=12.6, textColor=palette["ink"]))
    styles.add(ParagraphStyle(name="Meta", fontName=fonts["body"], fontSize=8.2, leading=10, textColor=palette["muted"]))
    styles.add(ParagraphStyle(name="MetaStrong", fontName=fonts["body_bold"], fontSize=8.2, leading=10, textColor=palette["muted"]))
    styles.add(ParagraphStyle(name="Body", fontName=fonts["body"], fontSize=8.7, leading=10.8, textColor=palette["ink"]))
    styles.add(ParagraphStyle(name="BodyStrong", fontName=fonts["body_bold"], fontSize=8.5, leading=10.4, textColor=palette["primary"]))
    styles.add(ParagraphStyle(name="ChipLine", fontName=fonts["body"], fontSize=8.0, leading=9.8, textColor=palette["ink"]))

    story = []
    page_width = A4[0] - 26 * mm

    # Hero
    photo = Spacer(1, 1)
    if PHOTO_PATH.exists():
        photo = Image(str(PHOTO_PATH), width=31 * mm, height=31 * mm)

    personal = data["personal"]
    email = clean_text(personal.get("Email", ""))
    phone = clean_text(personal.get("Mobile Phone", ""))
    linkedin = clean_text(personal.get("LinkedIn", ""))
    dob = clean_text(personal.get("Date of Birth", ""))
    hobby = clean_text(personal.get("Hobby", ""))

    start_years = [int(m.group(1)) for p in data["projects"] for m in [re.match(r"^(\d{4})-", p.get("start", ""))] if m]
    first_year = min(start_years) if start_years else datetime.now().year
    years_exp = datetime.now().year - first_year + 1
    tech_count = sum(len(split_csv_smart(t["items"])) for t in data["technology"])

    role_items = [clean_text(x.strip()) for x in data["other_roles"].split(",") if x.strip()]
    role_line = " | ".join(role_items[:6])
    contact_line = " | ".join([x for x in [phone, email, linkedin] if x])
    stat_line = (
        "<font color='#EA580C'><b>{}</b></font> projects &nbsp;&nbsp; "
        "<font color='#EA580C'><b>{}+</b></font> years &nbsp;&nbsp; "
        "<font color='#EA580C'><b>{}+</b></font> technologies"
    ).format(len(data["projects"]), years_exp, tech_count)

    hero_right = [
        Paragraph(clean_text(data["name"]), styles["HeroName"]),
        Paragraph(clean_text(data["main_role"]), styles["Role"]),
        Spacer(1, 1.5),
        Paragraph("<b>Other roles:</b> " + role_line, styles["Tiny"]),
        Paragraph(contact_line, styles["Tiny"]),
    ]
    if dob:
        hero_right.append(Paragraph("Born: " + dob, styles["Tiny"]))
    if hobby:
        hero_right.append(Paragraph("Portfolio: " + hobby, styles["Tiny"]))
    hero_right.append(Spacer(1, 2))
    hero_right.append(Paragraph(stat_line, styles["Tiny"]))

    hero_tbl = Table(
        [[photo, hero_right]],
        colWidths=[38 * mm, page_width - 38 * mm],
        hAlign="LEFT",
    )
    hero_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), palette["soft"]),
        ("BOX", (0, 0), (-1, -1), 0.9, palette["line"]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 11),
        ("RIGHTPADDING", (0, 0), (-1, -1), 11),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))

    story.extend([hero_tbl, Spacer(1, 10)])

    # Experience
    story.append(Paragraph("Professional Experience Summary", styles["Section"]))
    story.append(Paragraph("Career path by company and period.", styles["SectionSub"]))
    exp_cards = []
    col_width = (page_width - 8) / 2
    for item in data["experience"]:
        content = [
            Paragraph(clean_text(item["period"]), styles["MetaStrong"]),
            Paragraph(clean_text(item["company"]), styles["CardTitle"]),
            Paragraph(clean_text(item["role"]), styles["BodyStrong"]),
        ]
        desc = clean_text(item.get("description", ""))
        if desc:
            content.append(Paragraph(desc, styles["Meta"]))
        exp_cards.append(card(content, col_width))
    story.extend([grid(exp_cards, 2, col_width), Spacer(1, 4)])

    # Technologies summary
    story.append(Paragraph("Technologies Summary", styles["Section"]))
    story.append(Paragraph("Skills grouped by category.", styles["SectionSub"]))
    tech_cards = []
    for item in data["technology"]:
        chips = [clean_text(x) for x in split_csv_smart(item["items"])]
        rows = []
        line = []
        for chip in chips:
            line.append(chip)
            if len(line) == 4:
                rows.append(" | ".join(line))
                line = []
        if line:
            rows.append(" | ".join(line))
        content = [Paragraph(clean_text(item["category"]), styles["CardTitle"])]
        for row in rows[:5]:
            content.append(Paragraph(row, styles["ChipLine"]))
        tech_cards.append(card(content, col_width, bg="#FFFFFF"))
    story.extend([grid(tech_cards, 2, col_width), Spacer(1, 4)])

    # Projects
    story.append(Paragraph("Most Relevant Projects", styles["Section"]))
    story.append(Paragraph("Content extracted directly from the English CV source.", styles["SectionSub"]))
    for item in data["projects"]:
        title_row = Table(
            [[Paragraph(clean_text(item["title"]), styles["CardTitle"]), Paragraph(fmt_date(item["start"]) + " - " + fmt_date(item["end"]), styles["MetaStrong"])]],
            colWidths=[page_width - 62 * mm, 62 * mm],
        )
        title_row.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        body = [
            title_row,
            Spacer(1, 2),
            Paragraph(
                f"<b>Client:</b> {clean_text(item['client'])} &nbsp;&nbsp; "
                f"<b>Role:</b> {clean_text(item['role'])} &nbsp;&nbsp; "
                f"<b>Team:</b> {clean_text(item['team'])}",
                styles["Meta"],
            ),
            Spacer(1, 3),
            Paragraph("Description and technologies", styles["BodyStrong"]),
        ]
        for paragraph in item["description"]:
            p = clean_text(paragraph)
            if p:
                body.append(Paragraph(p, styles["Body"]))
        body.append(Spacer(1, 2))
        tech_line = ", ".join(clean_text(x) for x in split_csv_smart(item["technologies"]))
        body.append(Paragraph("<b>Technologies:</b> " + tech_line, styles["ChipLine"]))

        story.append(card(body, page_width))
        story.append(Spacer(1, 5))

    # Education
    story.append(Spacer(1, 2))
    story.append(Paragraph("Education", styles["Section"]))
    edu_cards = []
    for item in data["education"]:
        content = [
            Paragraph(clean_text(item["period"]), styles["MetaStrong"]),
            Paragraph(clean_text(item["course"]), styles["CardTitle"]),
            Paragraph(clean_text(item["description"]), styles["Meta"]),
        ]
        edu_cards.append(card(content, col_width))
    story.extend([grid(edu_cards, 2, col_width), Spacer(1, 4)])

    # Languages
    story.append(Paragraph("Spoken Languages", styles["Section"]))
    lang_width = (page_width - 9) / 4
    lang_cards = []
    for item in data["languages"]:
        content = [
            Paragraph(clean_text(item["level"]), styles["MetaStrong"]),
            Paragraph(clean_text(item["language"]), styles["CardTitle"]),
        ]
        lang_cards.append(card(content, lang_width, bg="#F8FAFC"))
    story.extend([grid(lang_cards, 4, lang_width, hpad=3), Spacer(1, 6)])

    # About
    story.append(Paragraph("About Me", styles["Section"]))
    about_parts = []
    for paragraph in data["about"]:
        p = clean_text(paragraph)
        if p:
            about_parts.append(Paragraph(p, styles["Body"]))
            about_parts.append(Spacer(1, 1.5))
    if data["signature"]:
        about_parts.append(Spacer(1, 2))
        about_parts.append(Paragraph(clean_text(data["signature"]), styles["MetaStrong"]))
    story.append(card(about_parts, page_width))
    return story


def draw_page(canvas, doc):
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(colors.HexColor("#F0FDFA"))
    canvas.circle(22 * mm, height - 20 * mm, 16 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#FFF7ED"))
    canvas.circle(width - 18 * mm, height - 58 * mm, 18 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#64748B"))
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(width - 12 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf(data):
    OUT_MAIN.parent.mkdir(parents=True, exist_ok=True)
    OUT_ALT.parent.mkdir(parents=True, exist_ok=True)

    fonts = register_fonts()
    story = build_story(data, fonts)

    doc = SimpleDocTemplate(
        str(OUT_MAIN),
        pagesize=A4,
        leftMargin=13 * mm,
        rightMargin=13 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title=f"{clean_text(data['name'])} - CV",
        author=clean_text(data["name"]),
    )
    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    OUT_ALT.write_bytes(OUT_MAIN.read_bytes())


if __name__ == "__main__":
    lines = extract_docx_lines(DOCX_PATH)
    data = parse_content(lines)
    build_pdf(data)
    print(f"Generated: {OUT_MAIN}")
    print(f"Generated: {OUT_ALT}")
