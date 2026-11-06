import base64
import functools
import io
import json
import logging
import math
import os
import re
import sqlite3
import uuid
from datetime import datetime
from typing import Literal, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from openai import OpenAI
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("edu-assistant")

# النموذج الرئيسي (كتابة الجذاذات والامتحانات) ونموذج الصور (قراءة الأوراق والتصحيح).
# يمكن تغييرهما بدون لمس الكود: أضف OPENAI_MODEL و OPENAI_VISION_MODEL في Secrets على Streamlit.
DEFAULT_MODEL = "gpt-6.1-sol"
DEFAULT_VISION_MODEL = "gpt-6.1-terra"
FALLBACK_MODEL = "gpt-4o"  # يُستعمل تلقائياً إن لم يتوفر النموذج في حسابك
DB_PATH = "history.db"

# ---- المحاولات المجانية: أوقفناها للتجربة. ضعها True عند ربط الدفع ----
LIMITS_ENABLED = False
GUEST_FREE = 3
MEMBER_FREE = 15

IMG_TYPES = ["png", "jpg", "jpeg"]
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# =====================================================================
# نماذج البيانات (Structured Outputs)
# =====================================================================
class Block(BaseModel):
    heading: str
    content: str


class LessonModel(BaseModel):
    title: str
    field: str
    lesson_blocks: list[Block]


class LessonTurn(BaseModel):
    reply: str
    draft: Optional[LessonModel]


class Chart(BaseModel):
    kind: Literal["bar", "line", "pie"]
    title: str
    x_label: str
    y_label: str
    labels: list[str]
    values: list[float]


class Q(BaseModel):
    text: str
    points: float
    level: Literal["معرفة", "تطبيق", "تحليل"]
    objective: str
    answer: str
    lines: int
    chart: Optional[Chart]


class Part(BaseModel):
    title: str
    questions: list[Q]


class ExamModel(BaseModel):
    title: str
    parts: list[Part]


class ExamTurn(BaseModel):
    reply: str
    draft: Optional[ExamModel]


class Mistake(BaseModel):
    description: str
    correction: str
    category: Literal["لغوي", "مفهومي", "تطبيقي", "منهجي"]


class StudentResult(BaseModel):
    student_name: str
    transcription: str
    score: float
    max_score: float
    strengths: list[str]
    mistakes: list[Mistake]
    comment: str


# <<WORD-START>>
# =====================================================================
# التسميات حسب اللغة
# =====================================================================
LABELS = {
    "العربية": {
        "ministry": "الجمهورية التونسية\nوزارة التربية", "prof": "الأستاذ(ة)", "subj": "المادة", "level_": "المستوى",
        "dur": "المدة", "date": "التاريخ", "pts": "ن", "annex": "الوثائق المرفقة", "spec": "جدول المواصفات",
        "obj": "الهدف", "level": "المستوى", "qn": "السؤال", "points": "النقاط", "key": "سلّم الإصلاح",
        "total": "المجموع", "answer": "الإصلاح", "name": "الاسم واللقب", "cls": "القسم", "mark": "العدد",
        "partw": "الجزء", "grading": "جدول التنقيط", "luck": "بالتوفيق", "topic": "الموضوع",
        "lesson": "الدرس", "field": "مدار الاهتمام",
    },
    "الفرنسية": {
        "ministry": "République Tunisienne\nMinistère de l'Éducation", "prof": "Prof", "subj": "Matière", "level_": "Niveau",
        "dur": "Durée", "date": "Date", "pts": "pts", "annex": "Annexes", "spec": "Tableau de spécifications",
        "obj": "Objectif", "level": "Niveau", "qn": "Question", "points": "Points", "key": "Corrigé et barème",
        "total": "Total", "answer": "Corrigé", "name": "Nom et prénom", "cls": "Classe", "mark": "Note",
        "partw": "Partie", "grading": "Barème", "luck": "Bon travail", "topic": "Sujet",
        "lesson": "Leçon", "field": "Centre d'intérêt",
    },
    "الإنجليزية": {
        "ministry": "Republic of Tunisia\nMinistry of Education", "prof": "Teacher", "subj": "Subject", "level_": "Level",
        "dur": "Duration", "date": "Date", "pts": "pts", "annex": "Annexes", "spec": "Specification table",
        "obj": "Objective", "level": "Level", "qn": "Question", "points": "Points", "key": "Answer key and marking scheme",
        "total": "Total", "answer": "Answer", "name": "Full name", "cls": "Class", "mark": "Mark",
        "partw": "Part", "grading": "Marking grid", "luck": "Good luck", "topic": "Topic",
        "lesson": "Lesson", "field": "Theme",
    },
}
DUR_FR = {"ساعة واحدة": "1h", "ساعتان": "2h", "3 ساعات": "3h", "4 ساعات": "4h"}
DUR_EN = {"ساعة واحدة": "1 hour", "ساعتان": "2 hours", "3 ساعات": "3 hours", "4 ساعات": "4 hours"}

PAGE_W = 6.6  # العرض المتاح بالبوصة (A4 بهوامش 0.8)

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]


def _tag(el):
    return el.tag.split("}")[1]


def ppr_insert(pPr, el):
    """يُدخل عنصراً في pPr في موضعه الصحيح حسب مخطط Word."""
    idx = PPR_ORDER.index(_tag(el))
    for child in pPr:
        name = _tag(child)
        if name in PPR_ORDER and PPR_ORDER.index(name) > idx:
            child.addprevious(el)
            return
    pPr.append(el)


def style_run(run, lang, size=12, bold=False):
    rtl = lang == "العربية"
    font = "Arial" if rtl else "Times New Roman"
    code = {"العربية": "ar-TN", "الفرنسية": "fr-FR"}.get(lang, "en-US")
    bold = bold or bool(run.font.bold)
    run.font.name = font
    run.font.size = Pt(size)
    if bold:
        run.font.bold = True
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), font)
    if rtl:
        run.font.rtl = True
        if bold:
            run.font.cs_bold = True
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(size * 2)))
        rPr.find(qn("w:sz")).addnext(szCs)
    noproof = OxmlElement("w:noProof")
    rPr.find(qn("w:sz")).addprevious(noproof)
    le = OxmlElement("w:lang")
    le.set(qn("w:val"), code)
    if rtl:
        le.set(qn("w:bidi"), code)
    rPr.append(le)


def fmt(p, lang, align="start", bold=False, size=12):
    """align: start (بداية السطر: يمين للعربية ويسار لغيرها) | center"""
    rtl = lang == "العربية"
    if rtl:
        ppr_insert(p._p.get_or_add_pPr(), OxmlElement("w:bidi"))
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif not rtl:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        style_run(run, lang, size, bold)


def add_par(doc, text, lang, align="start", bold=False, size=12, after=None):
    p = doc.add_paragraph()
    p.add_run(text)
    fmt(p, lang, align, bold, size)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    return p


def spacer(doc, pts=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = Pt(pts)
    return p


def dotted_line(doc):
    """سطر منقّط بعرض الصفحة (tab بنقاط قائدة) يظهر مستقلاً مهما تكرر."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(22)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), str(int(PAGE_W * 1440) - 20))
    tabs.append(tab)
    ppr_insert(p._p.get_or_add_pPr(), tabs)
    p.add_run("\t")
    return p


def rtl_table(t):
    tblPr = t._tbl.tblPr
    el = OxmlElement("w:bidiVisual")
    anchor = tblPr.find(qn("w:tblW"))
    if anchor is not None:
        anchor.addprevious(el)
    else:
        tblPr.append(el)


def make_table(doc, rows, widths, lang, grid=True):
    t = doc.add_table(rows=rows, cols=len(widths))
    if grid:
        t.style = "Table Grid"
    t.autofit = False
    if lang == "العربية":
        rtl_table(t)
    for i, w in enumerate(widths):
        t.columns[i].width = Inches(w)
    for r in t.rows:
        for i, w in enumerate(widths):
            r.cells[i].width = Inches(w)
    return t


def shade(cell, fill="E7E6E6"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def cell_text(cell, text, lang, size=11, bold=False, align="center", valign=True):
    p = cell.paragraphs[0]
    p.add_run(text)
    fmt(p, lang, align, bold, size)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if valign:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


_SEP_RE = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$")


def _inline(p, text):
    """يدعم **عريض** فقط (لا يلمس علامة * في العمليات الحسابية)."""
    for part in re.split(r"(\*\*[^*]+\*\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            p.add_run(part[2:-2]).bold = True
        else:
            p.add_run(part)


def _md_table(container, rows, lang, width):
    cols = max(len(r) for r in rows)
    t = container.add_table(len(rows), cols)
    t.style = "Table Grid"
    if lang == "العربية":
        rtl_table(t)
    for j in range(cols):
        t.columns[j].width = Inches(width / cols)
    for i, row in enumerate(rows):
        for j in range(cols):
            cell = t.cell(i, j)
            cell.width = Inches(width / cols)
            if i == 0:
                shade(cell, "F2F2F2")
            p = cell.paragraphs[0]
            _inline(p, row[j] if j < len(row) else "")
            fmt(p, lang, "center", i == 0, 10.5)


def add_markdown(container, text, lang, size=12, width=PAGE_W, align="start"):
    """يكتب نص Markdown (جداول، **عريض**، عناوين #، قوائم -) داخل مستند أو خلية Word."""
    if not text or not str(text).strip():
        return
    is_cell = hasattr(container, "_tc")
    first = container.paragraphs[0] if is_cell and not container.paragraphs[0].runs else None

    def new_par():
        nonlocal first
        if first is not None:
            p, first = first, None
            return p
        return container.add_paragraph()

    lines = str(text).replace("\r", "").split("\n")
    i, last_table = 0, False
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and line.count("|") >= 2:
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            rows = [[c.strip() for c in b.strip("|").split("|")] for b in block if not _SEP_RE.match(b)]
            if rows:
                _md_table(container, rows, lang, width)
                last_table = True
            continue
        i += 1
        if not line:
            continue
        last_table = False
        bold, sz = False, size
        h = re.match(r"^#{1,6}\s+(.*)$", line)
        if h:
            line, bold, sz = h.group(1), True, size + 1
        else:
            b = re.match(r"^[-*•]\s+(.*)$", line)
            if b:
                line = "• " + b.group(1)
        p = new_par()
        _inline(p, line)
        fmt(p, lang, align, bold, sz)
        p.paragraph_format.space_after = Pt(3)
    if first is not None:
        first.paragraph_format.line_spacing = Pt(2)
    if is_cell and last_table:
        container.add_paragraph().paragraph_format.line_spacing = Pt(2)


def add_figure(container, png, width_in):
    p = container.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(io.BytesIO(png), width=Inches(width_in))


@functools.lru_cache(maxsize=64)
def _chart_cached(key, lang):
    from matplotlib.figure import Figure

    ch = json.loads(key)
    rtl = lang == "العربية"

    def sh(t):
        t = str(t or "")
        if rtl:
            try:
                import arabic_reshaper
                from bidi.algorithm import get_display

                return get_display(arabic_reshaper.reshape(t))
            except Exception:
                return t
        return t

    labels = [sh(x) for x in ch.get("labels", [])]
    vals = [float(v) for v in ch.get("values", [])]
    n = min(len(labels), len(vals))
    if n == 0:
        return None
    labels, vals = labels[:n], vals[:n]
    fig = Figure(figsize=(5.4, 3.1))
    ax = fig.subplots()
    kind = ch.get("kind")
    if kind == "pie":
        ax.pie([abs(v) for v in vals], labels=labels, autopct="%1.0f%%", startangle=90)
        ax.axis("equal")
    elif kind == "line":
        ax.plot(labels, vals, marker="o", color="#2F5597")
        ax.grid(True, alpha=0.3)
    else:
        ax.bar(labels, vals, color="#4472C4")
        ax.grid(axis="y", alpha=0.3)
    if kind != "pie":
        ax.set_xlabel(sh(ch.get("x_label")))
        ax.set_ylabel(sh(ch.get("y_label")))
    ax.set_title(sh(ch.get("title")), fontsize=11)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160)
    return buf.getvalue()


def chart_png(ch, lang):
    if not ch:
        return None
    try:
        return _chart_cached(json.dumps(ch, sort_keys=True, ensure_ascii=False), lang)
    except Exception as e:
        log.warning("تعذر رسم المبيان: %s", e)
        return None


def add_page_border(doc):
    for section in doc.sections:
        sectPr = section._sectPr
        for old in sectPr.findall(qn("w:pgBorders")):
            sectPr.remove(old)
        pg = OxmlElement("w:pgBorders")
        pg.set(qn("w:offsetFrom"), "page")
        for name in ("top", "left", "bottom", "right"):
            b = OxmlElement(f"w:{name}")
            b.set(qn("w:val"), "double")
            b.set(qn("w:sz"), "12")
            b.set(qn("w:space"), "24")
            b.set(qn("w:color"), "000000")
            pg.append(b)
        # مخطط Word يشترط أن يأتي pgBorders بعد pgMar مباشرة
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            pgMar.addnext(pg)
        else:
            sectPr.append(pg)


def new_doc():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin = s.right_margin = Inches(0.8)
    s.top_margin = s.bottom_margin = Inches(0.8)
    add_page_border(doc)
    return doc


def finish(doc):
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def exam_type_for(lang, title):
    if lang == "الفرنسية":
        return (title.replace("فرض مراقبة", "Devoir de Contrôle").replace("فرض تأليفي", "Devoir de Synthèse")
                .replace("عدد", "n°"))
    if lang == "الإنجليزية":
        return title.replace("فرض مراقبة", "Control Test").replace("فرض تأليفي", "Synthesis Test").replace("عدد", "No.")
    return title


def duration_for(lang, d):
    return {"الفرنسية": DUR_FR, "الإنجليزية": DUR_EN}.get(lang, {}).get(d, d)


def exam_header(doc, lang, m, suffix=""):
    L = LABELS[lang]
    t = make_table(doc, 3, [PAGE_W / 2, PAGE_W / 2], lang)
    cell_text(t.cell(0, 0), f"{L['ministry']}\n{m['school']}", lang, 10, True)
    cell_text(t.cell(0, 1), f"{L['prof']}: {m['teacher']}\n{L['subj']}: {m['subject']}\n{L['level_']}: {m['grade']}", lang, 10, True)
    t.rows[1].cells[0].merge(t.rows[1].cells[1])
    cell_text(t.cell(1, 0), exam_type_for(lang, m["exam_title"]) + suffix, lang, 15, True)
    shade(t.cell(1, 0))
    t.rows[2].cells[0].merge(t.rows[2].cells[1])
    cell_text(t.cell(2, 0), f"{L['dur']}: {duration_for(lang, m['duration'])}   |   {L['date']}: {m['date']}", lang, 10)
    spacer(doc, 6)


def student_line(doc, lang, m):
    L = LABELS[lang]
    t = make_table(doc, 1, [3.2, 1.5, 1.9], lang)
    cell_text(t.cell(0, 0), f"{L['name']}: ..............................", lang, 11, False, "start")
    cell_text(t.cell(0, 1), f"{L['cls']}: ........", lang, 11, False, "start")
    cell_text(t.cell(0, 2), f"{L['mark']}:      / {m['total']:g}", lang, 12, True, "center")
    spacer(doc, 8)


def all_questions(exam):
    return [(p["title"], q) for p in exam["parts"] for q in p["questions"]]


def exam_total(exam):
    return round(sum(q["points"] for _, q in all_questions(exam)), 2)


def rescale_points(exam, target):
    qs = [q for _, q in all_questions(exam)]
    tot = sum(q["points"] for q in qs)
    if not qs or tot <= 0:
        return exam
    f = target / tot
    for q in qs:
        q["points"] = max(0.25, round(q["points"] * f * 4) / 4)
    diff = round(target - sum(q["points"] for q in qs), 2)
    big = max(qs, key=lambda q: q["points"])
    big["points"] = round(big["points"] + diff, 2)
    return exam


def add_images_annex(doc, lang, images):
    if not images:
        return
    doc.add_page_break()
    add_par(doc, LABELS[lang]["annex"], lang, "center", True, 16)
    for img in images:
        try:
            doc.add_picture(io.BytesIO(img), width=Inches(5.5))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            log.warning("تعذر إدراج صورة في الملحق: %s", e)


def create_exam_doc(exam, lang, m, images=None, suffix=""):
    L = LABELS[lang]
    doc = new_doc()
    exam_header(doc, lang, m, suffix)
    student_line(doc, lang, m)
    if exam.get("title"):
        add_par(doc, f"{L['topic']}: {exam['title']}", lang, "center", True, 12, after=6)

    n = 0
    part_totals = []
    for i, part in enumerate(exam["parts"], 1):
        ptot = round(sum(q["points"] for q in part["questions"]), 2)
        part_totals.append((i, ptot))
        ht = make_table(doc, 1, [PAGE_W], lang)
        shade(ht.cell(0, 0))
        cell_text(ht.cell(0, 0), f"{part['title']}   ({ptot:g} {L['pts']})", lang, 13, True, "start")
        spacer(doc, 6)
        for q in part["questions"]:
            n += 1
            t = make_table(doc, 1, [0.45, PAGE_W - 1.25, 0.8], lang)
            cell_text(t.cell(0, 0), str(n), lang, 12, True)
            shade(t.cell(0, 0), "F2F2F2")
            qcell = t.cell(0, 1)
            add_markdown(qcell, q["text"], lang, 12, width=PAGE_W - 1.45)
            png = chart_png(q.get("chart"), lang)
            if png:
                add_figure(qcell, png, 4.4)
            else:
                qcell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell_text(t.cell(0, 2), f"{q['points']:g} {L['pts']}", lang, 11, True)
            lines = max(0, min(int(q.get("lines", 0)), 14))
            if lines:
                spacer(doc, 4)
                for _ in range(lines):
                    dotted_line(doc)
            spacer(doc, 10)

    # جدول التنقيط
    if part_totals:
        add_par(doc, L["grading"], lang, "center", True, 12, after=4)
        cols = len(part_totals) + 1
        w = PAGE_W / cols
        t = make_table(doc, 2, [w] * cols, lang)
        for j, (i, ptot) in enumerate(part_totals):
            cell_text(t.cell(0, j), f"{L['partw']} {i}", lang, 10, True)
            shade(t.cell(0, j), "F2F2F2")
            cell_text(t.cell(1, j), f"      / {ptot:g}", lang, 10)
        cell_text(t.cell(0, cols - 1), L["total"], lang, 10, True)
        shade(t.cell(0, cols - 1), "D9D9D9")
        cell_text(t.cell(1, cols - 1), f"      / {exam_total(exam):g}", lang, 10, True)
    spacer(doc, 8)
    add_par(doc, L["luck"], lang, "center", True, 12)
    add_images_annex(doc, lang, images)
    return finish(doc)


def create_answer_key_doc(exam, lang, m):
    L = LABELS[lang]
    doc = new_doc()
    exam_header(doc, lang, m, f" — {L['key']}")

    add_par(doc, L["spec"], lang, "center", True, 12, after=4)
    t = make_table(doc, 1, [0.9, PAGE_W - 3.3, 1.4, 1.0], lang)
    for i, name in enumerate([L["qn"], L["obj"], L["level"], L["points"]]):
        cell_text(t.rows[0].cells[i], name, lang, 10, True)
        shade(t.rows[0].cells[i], "F2F2F2")
    for n, (_, q) in enumerate(all_questions(exam), 1):
        cells = t.add_row().cells
        for i, w in enumerate([0.9, PAGE_W - 3.3, 1.4, 1.0]):
            cells[i].width = Inches(w)
        cell_text(cells[0], str(n), lang, 10)
        cell_text(cells[1], q["objective"], lang, 10, False, "start")
        cell_text(cells[2], q["level"], lang, 10)
        cell_text(cells[3], f"{q['points']:g}", lang, 10)
    spacer(doc, 10)

    n = 0
    for part in exam["parts"]:
        ptot = round(sum(q["points"] for q in part["questions"]), 2)
        ht = make_table(doc, 1, [PAGE_W], lang)
        shade(ht.cell(0, 0))
        cell_text(ht.cell(0, 0), f"{part['title']}   ({ptot:g} {L['pts']})", lang, 12, True, "start")
        spacer(doc, 6)
        for q in part["questions"]:
            n += 1
            add_par(doc, f"{n}-  ({q['points']:g} {L['pts']})", lang, "start", True, 12, after=2)
            add_markdown(doc, q["text"], lang, 12)
            png = chart_png(q.get("chart"), lang)
            if png:
                add_figure(doc, png, 4.0)
            add_markdown(doc, f"**{L['answer']}:** {q['answer']}", lang, 12)
            spacer(doc, 10)
    add_par(doc, f"{L['total']}: {exam_total(exam):g}", lang, "center", True, 13)
    return finish(doc)


def create_lesson_doc(data, lang, m, images=None):
    L = LABELS[lang]
    doc = new_doc()
    t = make_table(doc, 2, [PAGE_W / 2, PAGE_W / 2], lang)
    cell_text(t.cell(0, 0), f"{L['ministry']}\n{m['school']}", lang, 10, True)
    cell_text(t.cell(0, 1), f"{L['prof']}: {m['teacher']}\n{L['subj']}: {m['subject']}\n{L['level_']}: {m['grade']}", lang, 10, True)
    cell_text(t.cell(1, 0), f"{L['field']}: {data.get('field', '')}", lang, 11, True)
    cell_text(t.cell(1, 1), f"{L['lesson']}: {data.get('title', '')}", lang, 11, True)
    spacer(doc, 10)
    for block in data.get("lesson_blocks", []):
        ht = make_table(doc, 1, [PAGE_W], lang)
        shade(ht.cell(0, 0), "F2F2F2")
        cell_text(ht.cell(0, 0), block["heading"], lang, 13, True, "start")
        add_markdown(doc, block["content"], lang, 12)
        spacer(doc, 8)
    add_images_annex(doc, lang, images)
    return finish(doc)


# <<WORD-END>>

# =====================================================================
# قاعدة البيانات والحسابات
# =====================================================================
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def db():
    return sqlite3.connect(DB_PATH)


def init_db():
    conn = db()
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS chats (id TEXT PRIMARY KEY, teacher TEXT, title TEXT, timestamp TEXT, chat_type TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS messages (id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id TEXT, role TEXT, content TEXT, image BLOB)""")
    c.execute("""CREATE TABLE IF NOT EXISTS accounts (email TEXT PRIMARY KEY, display_name TEXT, used INTEGER DEFAULT 0, created TEXT, profile TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS drafts (chat_id TEXT PRIMARY KEY, kind TEXT, data TEXT)""")
    try:
        c.execute("ALTER TABLE accounts ADD COLUMN profile TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def get_or_create_account(email):
    conn = db()
    row = conn.execute("SELECT display_name FROM accounts WHERE email=?", (email,)).fetchone()
    if row:
        conn.close()
        return row[0]
    display = email.split("@")[0]
    conn.execute("INSERT INTO accounts (email, display_name, used, created) VALUES (?,?,0,?)",
                 (email, display, datetime.now().strftime("%d/%m/%Y")))
    conn.commit()
    conn.close()
    return display


def get_profile(email):
    conn = db()
    row = conn.execute("SELECT profile FROM accounts WHERE email=?", (email,)).fetchone()
    conn.close()
    try:
        return json.loads(row[0]) if row and row[0] else {}
    except ValueError:
        return {}


def save_profile(email, data):
    conn = db()
    conn.execute("UPDATE accounts SET profile=? WHERE email=?", (json.dumps(data, ensure_ascii=False), email))
    conn.commit()
    conn.close()


def get_used(email):
    conn = db()
    row = conn.execute("SELECT used FROM accounts WHERE email=?", (email,)).fetchone()
    conn.close()
    return row[0] if row else 0


def add_used(email, n):
    conn = db()
    conn.execute("UPDATE accounts SET used = used + ? WHERE email=?", (n, email))
    conn.commit()
    conn.close()


def ensure_chat(chat_id, owner, kind):
    conn = db()
    if not conn.execute("SELECT 1 FROM chats WHERE id=?", (chat_id,)).fetchone():
        default = "جذاذة جديدة" if kind == "lesson" else "امتحان جديد"
        conn.execute("INSERT INTO chats VALUES (?,?,?,?,?)", (chat_id, owner, default, datetime.now().strftime("%d/%m %H:%M"), kind))
        conn.commit()
    conn.close()


def maybe_set_title(chat_id, text):
    if not text:
        return
    conn = db()
    row = conn.execute("SELECT title FROM chats WHERE id=?", (chat_id,)).fetchone()
    if row and row[0] in ("جذاذة جديدة", "امتحان جديد", "درس جديد", "فرض جديد"):
        conn.execute("UPDATE chats SET title=? WHERE id=?", (text.strip().replace("\n", " ")[:35], chat_id))
        conn.commit()
    conn.close()


def save_message(chat_id, role, content, image=None):
    conn = db()
    conn.execute("INSERT INTO messages (chat_id, role, content, image) VALUES (?,?,?,?)", (chat_id, role, content, image))
    conn.commit()
    conn.close()


def load_chat_history(chat_id):
    conn = db()
    rows = conn.execute("SELECT role, content, image FROM messages WHERE chat_id=? ORDER BY id ASC", (chat_id,)).fetchall()
    conn.close()
    out = []
    for role, content, image in rows:
        # صور إضافية بلا نص تُدمج مع الرسالة السابقة
        if role == "user" and not content and image and out and out[-1]["role"] == "user":
            out[-1]["images"].append(image)
            continue
        out.append({"role": role, "content": content or "", "images": [image] if image else []})
    return out


def save_draft(chat_id, kind, data):
    conn = db()
    conn.execute("INSERT OR REPLACE INTO drafts VALUES (?,?,?)", (chat_id, kind, json.dumps(data, ensure_ascii=False)))
    conn.commit()
    conn.close()


def load_draft(chat_id):
    conn = db()
    row = conn.execute("SELECT data FROM drafts WHERE chat_id=?", (chat_id,)).fetchone()
    conn.close()
    try:
        return json.loads(row[0]) if row else None
    except ValueError:
        return None


def get_all_chats(owner):
    conn = db()
    rows = conn.execute("SELECT id, title, timestamp, chat_type FROM chats WHERE teacher=? ORDER BY rowid DESC", (owner,)).fetchall()
    conn.close()
    return rows


def delete_chat(chat_id):
    conn = db()
    conn.execute("DELETE FROM messages WHERE chat_id=?", (chat_id,))
    conn.execute("DELETE FROM drafts WHERE chat_id=?", (chat_id,))
    conn.execute("DELETE FROM chats WHERE id=?", (chat_id,))
    conn.commit()
    conn.close()


init_db()

# =====================================================================
# المفتاح وعميل OpenAI
# =====================================================================
def configured_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("OPENAI_API_KEY", "")


def get_client():
    key = configured_key()
    if not key:
        log.error("OPENAI_API_KEY غير مضبوط في إعدادات الخادم")
        raise RuntimeError("الخدمة غير متاحة مؤقتاً. يرجى المحاولة لاحقاً أو التواصل مع المشرف.")
    return OpenAI(api_key=key)


def _setting(name, default):
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name, default)


def model_for(kind):
    if kind == "vision":
        return _setting("OPENAI_VISION_MODEL", DEFAULT_VISION_MODEL)
    return _setting("OPENAI_MODEL", DEFAULT_MODEL)


def _is_legacy(model):
    return model.startswith(("gpt-4", "gpt-3"))


def _params(model, temperature, effort, max_out):
    """نماذج الاستدلال الحديثة لا تقبل temperature وتقبل reasoning_effort؛ والقديمة العكس."""
    kw = {}
    if _is_legacy(model):
        if temperature is not None:
            kw["temperature"] = temperature
        if max_out:
            kw["max_completion_tokens"] = max_out
    elif effort:
        kw["reasoning_effort"] = effort
    return kw


def _model_unavailable(e):
    msg = str(e).lower()
    return "model" in msg and any(k in msg for k in ("does not exist", "not found", "access", "not supported", "unsupported", "invalid", "unknown"))


def llm(kind, messages, temperature=None, effort="medium", max_out=None, response_format=None):
    """نداء موحّد لـ OpenAI مع رجوع تلقائي إلى gpt-4o إن لم يتوفر النموذج الجديد."""
    client = get_client()
    models = [model_for(kind)]
    if models[0] != FALLBACK_MODEL:
        models.append(FALLBACK_MODEL)
    for m in models:
        kw = _params(m, temperature, effort, max_out)
        try:
            if response_format is not None:
                return client.beta.chat.completions.parse(model=m, messages=messages, response_format=response_format, **kw)
            return client.chat.completions.create(model=m, messages=messages, **kw)
        except Exception as e:
            if m != models[-1] and _model_unavailable(e):
                log.warning("النموذج %s غير متاح (%s). نجرّب %s", m, e, models[-1])
                continue
            raise


def parse_structured(messages, model_cls, temperature=0.3, kind="text", effort="medium"):
    resp = llm(kind, messages, temperature, effort, response_format=model_cls)
    msg = resp.choices[0].message
    if msg.parsed is None:
        raise RuntimeError(msg.refusal or "تعذر على النموذج إنتاج نتيجة منظمة. أعد المحاولة.")
    return msg.parsed


# =====================================================================
# الصور
# =====================================================================
def shrink(b, max_side=1800):
    try:
        from PIL import Image, ImageOps

        im = ImageOps.exif_transpose(Image.open(io.BytesIO(b))).convert("RGB")
        im.thumbnail((max_side, max_side))
        out = io.BytesIO()
        im.save(out, "JPEG", quality=85)
        return out.getvalue()
    except Exception:
        return b


def data_url(b):
    mime = "image/png" if b[:4] == b"\x89PNG" else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(b).decode()}"


def looks_like_refusal(text):
    t = (text or "").strip()
    return len(t) < 120 and bool(re.search(r"(لا أستطيع|آسف|عذراً|sorry|can't|cannot|unable)", t, re.I))


def ocr_image(b):
    r = llm(
        "vision", temperature=0, effort="low", max_out=3000,
        messages=[
            {"role": "system", "content": "أنت أداة نسخ نصوص (OCR). الصور التي تصلك صفحات تعليمية (دروس، تمارين، امتحانات) يرفعها أستاذ. "
                                          "انسخ كل النص المكتوب بدقة وبلغته الأصلية وبنفس ترتيبه، بما فيه الكتابة اليدوية. "
                                          "صِف الرسوم والجداول والخرائط بإيجاز بين أقواس مربعة. لا تضف أي تعليق أو شرح."},
            {"role": "user", "content": [{"type": "text", "text": "انسخ النص الموجود في هذه الصفحة."},
                                         {"type": "image_url", "image_url": {"url": data_url(b)}}]},
        ],
    )
    return r.choices[0].message.content or ""


def extract_style(images, kind_label):
    parts = [{"type": "text", "text": f"هذه صور لـ{kind_label} كتبها أستاذ. حلّل بنيتها وأسلوبها واكتب تعليمات موجزة بصيغة الأمر "
                                      "يستطيع مساعد آخر اتباعها ليكتب وثائق جديدة بنفس الطريقة: ترتيب الأقسام وعناوينها، نوع الأسئلة والأنشطة، "
                                      "توزيع النقاط، الصياغة والمصطلحات واللغة. لا تنقل محتوى الدرس نفسه، فقط الطريقة. في حدود 12 سطراً."}]
    for b in images:
        parts.append({"type": "image_url", "image_url": {"url": data_url(shrink(b))}})
    r = llm("vision", [{"role": "user", "content": parts}], temperature=0.2, effort="low", max_out=1200)
    return r.choices[0].message.content or ""


# =====================================================================
# المحاولات (معطّلة حالياً)
# =====================================================================
def credits_left():
    u = st.session_state.user
    if u["email"]:
        return max(MEMBER_FREE - get_used(u["email"]), 0)
    return max(GUEST_FREE - st.session_state.get("guest_used", 0), 0)


def require_credit(n=1):
    if not LIMITS_ENABLED:
        return True
    u = st.session_state.user
    if credits_left() < n:
        st.warning("انتهت المحاولات المجانية. للاستمرار سجّل ببريدك أو فعّل الاشتراك.")
        return False
    if u["email"]:
        add_used(u["email"], n)
    else:
        st.session_state.guest_used = st.session_state.get("guest_used", 0) + n
    return True


# =====================================================================
# ملف الأستاذ
# =====================================================================
GRADES_SECONDARY = ["السابعة أساسي", "الثامنة أساسي", "التاسعة أساسي", "الأولى ثانوي", "الثانية ثانوي", "الثالثة ثانوي", "البكالوريا"]
NO_SECTION_GRADES = ["السابعة أساسي", "الثامنة أساسي", "التاسعة أساسي"]
SECTION_NONE = "— بدون —"
SECTIONS = [SECTION_NONE, "جذع مشترك", "آداب", "رياضيات", "علوم تجريبية", "علوم الإعلامية", "تقنية", "اقتصاد وتصرف", "رياضة"]

PRESET_NONE = "— بدون —"
PRESETS = {
    "الرياضيات والعلوم والفيزياء": {
        "lesson": "أبدأ بنشاط اكتشاف أو مسألة انطلاق، ثم التعريفات والخاصيات (ألخّصها في جدول)، ثم البراهين خطوة بخطوة، ثم أمثلة محلولة، ثم تمارين تطبيقية متدرجة الصعوبة. أكتب الرموز بوضوح (x²، √، π).",
        "exam": "تمارين مستقلة متدرجة الصعوبة، في كل تمرين أسئلة فرعية بنقاطها، وتمرين أو مسألة إدماجية في الأخير. أكتب المعطيات بوضوح، وأدرج جدولاً أو مبياناً عندما يحتاجه السؤال.",
    },
    "التاريخ والجغرافيا": {
        "lesson": "أبدأ بإشكالية، ثم سندات ووثائق (نص تاريخي، خريطة، جدول إحصائي)، ثم تحليل، ثم استنتاج. أذكر التواريخ والأسماء بدقة، وألخّص المقارنات في جداول.",
        "exam": "جزء للمعرفة (تواريخ ومفاهيم)، وجزء لتحليل وثيقة (سند ثم تعليمات)، وجزء للإنتاج الكتابي (فقرة أو تركيب).",
    },
    "التربية المدنية والإسلامية": {
        "lesson": "مدار الاهتمام، ثم وضعية انطلاق، ثم سند، ثم تعليمة وأنشطة، ثم خلاصة، ثم تقييم.",
        "exam": "جزء للمعرفة، وجزء للتطبيق على سند، ووضعية إدماجية في الأخير.",
    },
    "الفلسفة": {
        "lesson": "أبدأ بالإشكالية وطرح المشكل، ثم محاولة الحل: موقف أول ونقده، موقف ثانٍ ونقده، ثم التركيب والخاتمة. أستشهد بأقوال الفلاسفة بين قوسين.",
        "exam": "تحليل نص فلسفي أو معالجة إشكالية: فهم، وتحليل، ومناقشة، وموقف شخصي معلَّل.",
    },
    "اللغات (عربية، فرنسية، إنجليزية)": {
        "lesson": "نص اعتماد، ثم ملاحظة، ثم شرح وتحليل (مفردات وصور بيانية)، ثم قواعد أو بلاغة (ألخّصها في جدول)، ثم تعبير وإنتاج.",
        "exam": "جزء لفهم النص، وجزء للقواعد والصرف والبلاغة، وجزء للإنتاج الكتابي.",
    },
    "الاقتصاد والتصرف": {
        "lesson": "المفاهيم، ثم القوانين والعلاقات، ثم جداول تحليلية، ثم دراسة حالة رقمية، ثم تطبيق عددي.",
        "exam": "جزء للمفاهيم، وجزء للتطبيق العددي (بجداول ومعطيات)، ودراسة حالة في الأخير.",
    },
    "الإعلامية والتكنولوجيا": {
        "lesson": "أبدأ بوضعية مشكل، ثم المفاهيم والتعريفات، ثم أمثلة وخوارزميات أو مخططات، ثم تمارين تطبيقية.",
        "exam": "جزء للمعرفة، وتمارين تطبيقية متدرجة (برمجة أو مخططات أو جداول)، ووضعية إدماجية.",
    },
}

PROFILE_DEFAULTS = {
    "name": "", "school": "المعهد الثانوي", "cycle": "إعدادي وثانوي", "grade": "السابعة أساسي", "section": SECTION_NONE,
    "subject": "", "language": "العربية", "creativity": 0.7,
    "method_lesson": "", "method_exam": "", "notes": "",
}


SUBJECT_CHOICES = ["— اختر —", "التربية المدنية", "التربية الإسلامية", "اللغة العربية", "اللغة الفرنسية", "اللغة الإنجليزية",
                   "الرياضيات", "علوم الحياة والأرض", "الفيزياء", "التاريخ", "الجغرافيا", "الفلسفة", "الإعلامية",
                   "التكنولوجيا", "الاقتصاد", "التصرف", "التربية التشكيلية", "التربية الموسيقية", "التربية البدنية"]


def pick_subject():
    v = st.session_state.get("subj_pick")
    if v and v != SUBJECT_CHOICES[0]:
        st.session_state["pf_subject"] = v


def prof(k):
    return st.session_state.get(f"pf_{k}", PROFILE_DEFAULTS[k])


def grade_label():
    g, sec = prof("grade"), prof("section")
    if g in NO_SECTION_GRADES or sec in ("", SECTION_NONE):
        return g
    return f"{g} {sec}"


def init_profile_state():
    if st.session_state.get("pf_init"):
        return
    u = st.session_state.user
    saved = get_profile(u["email"]) if u["email"] else {}
    for k, default in PROFILE_DEFAULTS.items():
        v = saved.get(k, default)
        if k == "name" and not v:
            v = u["display"] if u["email"] else ""
        st.session_state.setdefault(f"pf_{k}", v)
    st.session_state.pf_init = True


def doc_meta():
    return {
        "school": prof("school"), "teacher": prof("name") or st.session_state.user["display"],
        "subject": prof("subject") or "........", "grade": grade_label(),
    }


def safe_name(s):
    return re.sub(r"[^\w\-]+", "_", s, flags=re.UNICODE).strip("_") or "file"


# =====================================================================
# الـ prompts (تُبنى من ملف الأستاذ)
# =====================================================================
def method_text(kind):
    method = prof(f"method_{kind}").strip()
    notes = prof("notes").strip()
    out = method or "لم يحدد الأستاذ طريقة بعد: اقترح هيكلاً بسيطاً مناسباً لمادته وأخبره به ليعدّله."
    if notes:
        out += f"\nملاحظات خاصة: {notes}"
    return out


def lesson_system():
    draft = st.session_state.get("draft_lesson")
    return f"""أنت مساعد تربوي يعمل مع أستاذ(ة) {prof('name') or 'الأستاذ'} في مادة {prof('subject') or '(غير محددة)'} لمستوى {grade_label()} (المنهج التونسي). لغة المحتوى: {prof('language')}.

طريقة الأستاذ في التحضير:
{method_text('lesson')}

مهمتك: تحاور الأستاذ ثم تبني «مسودة جذاذة» منظمة: عنوان الدرس، مدار الاهتمام (أو المحور)، وكتل (عنوان فرعي + محتوى فعلي).

القواعد:
1. اكتب المحتوى الفعلي مباشرة (نصوص، أسئلة، أنشطة، أمثلة)، ولا تكتب أبداً «يقوم المعلم بـ» أو «يطرح الأستاذ».
2. ابنِ فقط على ما قدمه الأستاذ (عنوان، نص، صفحة مستخرجة من صورة). إن نقصك مصدر أو معلومة جوهرية فاسأله في reply وأعد draft = null، ولا تخترع.
3. عند أي طلب تعديل (أضف/احذف/غيّر/اختصر) عدّل المسودة الحالية وأعدها كاملة في draft.
4. اجعل المستوى مناسباً لتلاميذ {grade_label()}.
5. تحقق داخلياً من دقة المعلومات قبل الإجابة، ولا تكتب هذه المراجعة.
6. في reply جملة أو جملتان فقط: ما فعلته أو ما تحتاجه. لا تكرر المسودة في reply.
7. التنسيق داخل content: استعمل **العريض** والقوائم (- عنصر) والعناوين (# عنوان)، وجداول Markdown (| عمود | عمود |) عند الحاجة لتلخيص خاصيات أو مقارنات. كل ذلك يظهر منسقاً في Word.

المسودة الحالية:
{json.dumps(draft, ensure_ascii=False) if draft else 'لا توجد بعد.'}"""


def exam_system(exam_title, total):
    draft = st.session_state.get("draft_exam")
    return f"""أنت مساعد تربوي يعمل مع أستاذ(ة) {prof('name') or 'الأستاذ'} (مادة {prof('subject') or '(غير محددة)'}، مستوى {grade_label()}، المنهج التونسي) لبناء «{exam_title}» بلغة {prof('language')}.

طريقة الأستاذ في الامتحانات:
{method_text('exam')}

مهمتك: تحاور الأستاذ ثم تبني مسودة امتحان قابلة للتعديل. أجزاء الامتحان (عددها وعناوينها وتوزيع نقاطها) يحددها الأستاذ بنفسه، ولا هيكل مفروضاً عليه.

القواعد:
1. إن لم يحدد الأستاذ بعد مواضيع الامتحان أو مصدر الأسئلة (درس، صورة، موضوع)، فاسأله في reply وأعد draft = null. لا تخترع أسئلة عامة لا علاقة بدروسه.
2. إن حدد المواضيع ولم يحدد الأجزاء، فابنِ مسودة أولى بهيكل مناسب لمادته، وأخبره في reply بما اخترت ليعدّله.
3. مجموع نقاط كل الأسئلة = {total:g} بالضبط (من إعدادات الفرض)، والنقاط من مضاعفات 0.25.
4. لكل سؤال: text (نص السؤال فقط، بلا ترقيم ولا نقاط)، points، level (معرفة/تطبيق/تحليل)، objective (الهدف التعلمي)، answer (الإصلاح النموذجي المختصر مع توزيع النقاط عند الحاجة)، lines (عدد أسطر مساحة الإجابة من 0 إلى 12، متناسب مع النقاط وطول الجواب المتوقع، و0 للاختيار من متعدد). title في المسودة هو عنوان موضوع الفرض.
5. عند أي طلب تعديل (أضف سؤالاً/احذف/غيّر النقاط/بدّل الجزء) عدّل المسودة الحالية وأعدها كاملة في draft مع الحفاظ على المجموع {total:g}. إن أخلّ الطلب بالمجموع فعدّل نقاطاً أخرى وأخبره بذلك.
6. الأسئلة مناسبة لمستوى {grade_label()} وتغطي مستويات معرفية متنوعة.
7. تحقق داخلياً من صحة الأسئلة والإصلاح، ولا تكتب هذه المراجعة.
8. في reply جملتان على الأكثر: ما فعلته أو سؤالك. لا تكرر نص الامتحان.
9. الجداول: عندما يحتاج السؤال إلى جدول (معطيات، جدول قيم، جدول إحصائي...) فاكتبه داخل text بصيغة جدول Markdown (| عمود | عمود |) وسطر الفاصل (|---|---|)، فيظهر في ورقة Word جدولاً حقيقياً. والنقاط والترقيم لا تدخل في text.
10. المبيانات: عندما يحتاج السؤال إلى سند بياني فاستعمل الحقل chart (kind: bar/line/pie، title، x_label، y_label، labels، values) بأرقام قدّمها الأستاذ أو أرقام معروفة وصحيحة. لا تخترع إحصاءات وأرقاماً واقعية. وإن لم يحتج السؤال إلى مبيان فاجعل chart = null.
11. الرموز الرياضية والعلمية: اكتبها نصاً واضحاً يظهر في Word (x²، √، π، ≤، →...).

المسودة الحالية:
{json.dumps(draft, ensure_ascii=False) if draft else 'لا توجد بعد.'}"""


# =====================================================================
# ورشة المحادثة
# =====================================================================
GREETINGS = {
    "lesson": "أهلاً بك 👋 اكتب عنوان الدرس (وارفع صورة صفحته إن شئت)، وسنبني الجذاذة معاً، ثم تعدّلها بالكلام: أضف، احذف، غيّر.",
    "exam": "أهلاً بك 👋 أخبرني بمواضيع الفرض (أو ارفع صفحات الدروس) وكيف تريد الأجزاء والنقاط، وسأبني المسودة وتعدّلها معي.",
}


def reset_chat(kind):
    st.session_state[f"chat_id_{kind}"] = str(uuid.uuid4())
    st.session_state[f"messages_{kind}"] = [{"role": "assistant", "content": GREETINGS[kind], "images": []}]
    st.session_state[f"draft_{kind}"] = None
    st.session_state[f"ver_{kind}"] = st.session_state.get(f"ver_{kind}", 0) + 1
    st.session_state.pop("exam_b", None)


def load_chat(chat_id, kind):
    st.session_state[f"chat_id_{kind}"] = chat_id
    st.session_state[f"messages_{kind}"] = [{"role": "assistant", "content": GREETINGS[kind], "images": []}] + load_chat_history(chat_id)
    st.session_state[f"draft_{kind}"] = load_draft(chat_id)
    st.session_state[f"ver_{kind}"] = st.session_state.get(f"ver_{kind}", 0) + 1
    st.session_state.pop("exam_b", None)


def remove_chat(chat_id, kind):
    delete_chat(chat_id)
    if st.session_state.get(f"chat_id_{kind}") == chat_id:
        reset_chat(kind)


def build_api_messages(messages, system):
    out = [{"role": "system", "content": system}]
    for m in messages:
        if m["role"] == "assistant":
            if m.get("content"):
                out.append({"role": "assistant", "content": m["content"]})
        elif m["role"] == "source":
            out.append({"role": "user", "content": "[نص مستخرج من صورة رفعها الأستاذ]\n" + m["content"]})
        elif m.get("content"):
            out.append({"role": "user", "content": m["content"]})
    return out


def render_message(m):
    if m["role"] == "source":
        with st.chat_message("assistant", avatar="📎"):
            with st.expander("📎 النص المستخرج من الصورة (تأكد من صحته)"):
                st.text(m["content"])
        return
    with st.chat_message(m["role"]):
        if m.get("content"):
            st.markdown(m["content"])
        for b in m.get("images", []):
            st.image(b, width=260)


def workshop(kind, placeholder, system_fn, turn_cls, temperature):
    mk, ik, dk = f"messages_{kind}", f"chat_id_{kind}", f"draft_{kind}"
    box = st.container(height=520, border=True)
    with box:
        for m in st.session_state[mk]:
            render_message(m)
    prompt = st.chat_input(placeholder, accept_file="multiple", file_type=IMG_TYPES, key=f"input_{kind}")
    if not prompt:
        return
    if not require_credit():
        return

    text = getattr(prompt, "text", None)
    if text is None:
        text = str(prompt)
    text = text.strip()
    files = [shrink(f.getvalue()) for f in (getattr(prompt, "files", None) or [])]

    chat_id = st.session_state[ik]
    owner = st.session_state.user["owner"]
    ensure_chat(chat_id, owner, kind)
    maybe_set_title(chat_id, text or "صورة مرفقة")

    user_msg = {"role": "user", "content": text, "images": files}
    st.session_state[mk].append(user_msg)
    save_message(chat_id, "user", text, files[0] if files else None)
    for extra in files[1:]:
        save_message(chat_id, "user", "", extra)

    with box:
        render_message(user_msg)
        with st.chat_message("assistant"):
            with st.spinner("جاري العمل..."):
                try:
                    for img in files:
                        src = ocr_image(img)
                        if looks_like_refusal(src) or not src.strip():
                            st.warning("تعذرت قراءة إحدى الصور. جرّب صورة أوضح (إضاءة جيدة، الصفحة كاملة، دون تصوير شاشة).")
                            continue
                        smsg = {"role": "source", "content": src, "images": []}
                        st.session_state[mk].append(smsg)
                        save_message(chat_id, "source", src)
                    turn = parse_structured(build_api_messages(st.session_state[mk], system_fn()), turn_cls, temperature)
                except Exception as e:
                    log.exception("turn failed")
                    st.error(f"حدث خطأ: {e}")
                    return
    reply = turn.reply.strip() or "تم."
    st.session_state[mk].append({"role": "assistant", "content": reply, "images": []})
    save_message(chat_id, "assistant", reply)
    if turn.draft is not None:
        data = turn.draft.model_dump()
        st.session_state[dk] = data
        st.session_state[f"ver_{kind}"] = st.session_state.get(f"ver_{kind}", 0) + 1
        st.session_state.pop("exam_b", None)
        save_draft(chat_id, kind, data)
    st.rerun()


# =====================================================================
# المعاينات
# =====================================================================
def chat_images(kind):
    return [b for m in st.session_state[f"messages_{kind}"] for b in m.get("images", [])]


def preview_lesson(lang, attach):
    d = st.session_state.get("draft_lesson")
    if not d:
        st.info("📄 ستظهر هنا مسودة الجذاذة بعد أن تحدثني عن الدرس.")
        return
    st.markdown(f"### {d['title']}")
    st.caption(f"{LABELS[lang]['field']}: {d['field']}")
    for b in d["lesson_blocks"]:
        with st.expander(b["heading"], expanded=True):
            st.markdown(b["content"])
    data = create_lesson_doc(d, lang, doc_meta(), chat_images("lesson") if attach else None).getvalue()
    st.download_button("⬇️ تحميل الجذاذة (Word)", data, f"Jadhatha_{safe_name(prof('subject'))}_{safe_name(grade_label())}.docx",
                       DOCX_MIME, type="primary", use_container_width=True)


def preview_exam(lang, meta, target, attach):
    d = st.session_state.get("draft_exam")
    if not d:
        st.info("📝 ستظهر هنا مسودة الامتحان بعد أن تحدد مواضيعه.")
        return
    L = LABELS[lang]
    chat_id = st.session_state["chat_id_exam"]
    ver = st.session_state.get("ver_exam", 0)

    tot = exam_total(d)
    if abs(tot - target) > 0.01:
        st.warning(f"⚠️ مجموع النقاط الحالي {tot:g} وليس {target:g}.")
        if st.button(f"⚖️ وزّع النقاط تلقائياً ليصبح المجموع {target:g}", key="rescale"):
            rescale_points(d, target)
            save_draft(chat_id, "exam", d)
            st.session_state.ver_exam = ver + 1
            st.session_state.pop("exam_b", None)
            st.rerun()
    else:
        st.success(f"✅ المجموع {tot:g}")

    t_prev, t_edit, t_spec = st.tabs(["👁️ المعاينة", "✏️ تعديل مباشر", "📊 جدول المواصفات"])
    with t_prev:
        if d.get("title"):
            st.markdown(f"### {d['title']}")
        n = 0
        for part in d["parts"]:
            ptot = round(sum(q["points"] for q in part["questions"]), 2)
            st.markdown(f"#### {part['title']} — {ptot:g} {L['pts']}")
            for q in part["questions"]:
                n += 1
                st.markdown(f"**{n}.** {q['text']}  \n`{q['points']:g} {L['pts']}` · {q['level']}")
                png = chart_png(q.get("chart"), lang)
                if png:
                    st.image(png, width=380)
    with t_edit:
        st.caption("عدّل نص السؤال أو نقاطه أو مستواه مباشرة. ولتغيير بنية الامتحان (إضافة، حذف، أجزاء) اكتب لي في المحادثة.")
        rows = [{"الجزء": pt, "السؤال": q["text"], "النقاط": float(q["points"]), "المستوى": q["level"]} for pt, q in all_questions(d)]
        edited = st.data_editor(
            pd.DataFrame(rows), key=f"ed_exam_{ver}", hide_index=True, use_container_width=True, disabled=["الجزء"],
            column_config={
                "السؤال": st.column_config.TextColumn(width="large"),
                "النقاط": st.column_config.NumberColumn(min_value=0.0, step=0.25),
                "المستوى": st.column_config.SelectboxColumn(options=["معرفة", "تطبيق", "تحليل"]),
            },
        )
        changed, i = False, 0
        for part in d["parts"]:
            for q in part["questions"]:
                row = edited.iloc[i]
                i += 1
                txt, pts, lvl = row["السؤال"], row["النقاط"], row["المستوى"]
                if isinstance(txt, str) and txt.strip() and txt != q["text"]:
                    q["text"], changed = txt, True
                if pts is not None and not pd.isna(pts) and float(pts) != float(q["points"]):
                    q["points"], changed = float(pts), True
                if lvl in ("معرفة", "تطبيق", "تحليل") and lvl != q["level"]:
                    q["level"], changed = lvl, True
        if changed:
            save_draft(chat_id, "exam", d)
            st.session_state.pop("exam_b", None)
            st.rerun()
    with t_spec:
        qs = all_questions(d)
        df = pd.DataFrame([{"السؤال": i, "الجزء": pt, "الهدف": q["objective"], "المستوى": q["level"], "النقاط": q["points"]}
                           for i, (pt, q) in enumerate(qs, 1)])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.bar_chart(df.groupby("المستوى")["النقاط"].sum())

    imgs = chat_images("exam") if attach else None
    base = f"{safe_name(prof('subject'))}_{safe_name(grade_label())}"
    c1, c2 = st.columns(2)
    c1.download_button("⬇️ ورقة الامتحان", create_exam_doc(d, lang, meta, imgs).getvalue(), f"Examen_{base}.docx",
                       DOCX_MIME, type="primary", use_container_width=True)
    c2.download_button("⬇️ ملف الإصلاح (للأستاذ)", create_answer_key_doc(d, lang, meta).getvalue(), f"Corrige_{base}.docx",
                       DOCX_MIME, use_container_width=True)

    if st.button("🔀 توليد نسخة B (لمنع الغش)", use_container_width=True):
        if require_credit():
            with st.spinner("جاري توليد النسخة B..."):
                try:
                    b = parse_structured(
                        [{"role": "system", "content": f"أعد كتابة هذا الامتحان كنسخة B مكافئة في الصعوبة والأهداف والمستويات وعدد الأسئلة ومجموع النقاط ({target:g})، "
                                                      f"مع تغيير الأرقام والأمثلة والصياغة، وحافظ على نفس الأجزاء. اللغة: {lang}. أعد أيضاً الإصلاح لكل سؤال."},
                         {"role": "user", "content": json.dumps(d, ensure_ascii=False)}],
                        ExamModel, 0.6).model_dump()
                    if abs(exam_total(b) - target) > 0.01:
                        rescale_points(b, target)
                    st.session_state.exam_b = b
                except Exception as e:
                    log.exception("version B")
                    st.error(f"حدث خطأ: {e}")
    b = st.session_state.get("exam_b")
    if b:
        d1, d2 = st.columns(2)
        d1.download_button("⬇️ الامتحان (نسخة B)", create_exam_doc(b, lang, meta, imgs, " (B)").getvalue(), f"Examen_B_{base}.docx",
                           DOCX_MIME, use_container_width=True)
        d2.download_button("⬇️ إصلاح النسخة B", create_answer_key_doc(b, lang, meta).getvalue(), f"Corrige_B_{base}.docx",
                           DOCX_MIME, use_container_width=True)


# =====================================================================
# الدخول
# =====================================================================
def email_box(key):
    email = st.text_input("📧 بريدك الإلكتروني", key=f"email_{key}", placeholder="name@example.com")
    if st.button("متابعة", type="primary", key=f"go_{key}"):
        e = email.strip().lower()
        if not EMAIL_RE.match(e):
            st.error("اكتب بريداً إلكترونياً صحيحاً.")
        else:
            name = get_or_create_account(e)
            st.session_state.clear()
            st.session_state.user = {"email": e, "owner": e, "display": name}
            st.rerun()


st.set_page_config(page_title="المساعد التربوي الذكي", page_icon="🎓", layout="wide")

if "user" not in st.session_state:
    st.title("🎓 المساعد التربوي الذكي")
    st.markdown("### حضّر جذاذاتك، ابنِ امتحاناتك، وصحّح أوراق تلاميذك بطريقتك أنت")
    c1, c2, c3 = st.columns(3)
    c1.info("📄 **جذاذات**\n\nتحاور المساعد ثم تعدّل المسودة بالكلام")
    c2.info("📝 **امتحانات**\n\nأجزاء ونقاط حسب طلبك، وورقة مؤطرة جاهزة للطباعة")
    c3.info("🔍 **تصحيح**\n\nارفع صور الأوراق واحصل على العلامات وجدول Excel")
    st.divider()
    if st.button("🚀 جرّب الآن بدون تسجيل", type="primary", use_container_width=True):
        st.session_state.user = {"email": None, "owner": "guest:" + str(uuid.uuid4()), "display": "الأستاذ(ة)"}
        st.rerun()
    st.markdown("**أو ادخل ببريدك الإلكتروني** (بدون كلمة سر) ليُحفظ ملفك وسجل محادثاتك:")
    email_box("landing")
    st.stop()

user = st.session_state.user
init_profile_state()
for _k in ("lesson", "exam"):
    if f"chat_id_{_k}" not in st.session_state:
        reset_chat(_k)

# =====================================================================
# القائمة الجانبية: الحساب + السجل
# =====================================================================
with st.sidebar:
    if user["email"]:
        st.caption(f"👤 {user['email']}")
        if st.button("🚪 خروج", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    else:
        st.caption("👤 زائر — لن يُحفظ ملفك. ادخل ببريدك لحفظه.")
        if st.button("📧 دخول ببريدي", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    if LIMITS_ENABLED:
        st.metric("🎟️ المحاولات المتبقية", credits_left())

    st.divider()
    st.subheader("📋 بيانات المادة")
    st.selectbox("اختر المادة", SUBJECT_CHOICES, key="subj_pick", on_change=pick_subject)
    st.text_input("أو اكتب مادتك", key="pf_subject", placeholder="مثال: التربية المدنية")
    st.session_state.pf_cycle = "إعدادي وثانوي"
    st.selectbox("المرحلة", ["إعدادي وثانوي"], key="pf_cycle")
    if st.session_state.pf_grade not in GRADES_SECONDARY:
        st.session_state.pf_grade = GRADES_SECONDARY[0]
    st.selectbox("المستوى", GRADES_SECONDARY, key="pf_grade")
    if st.session_state.pf_grade not in NO_SECTION_GRADES:
        if st.session_state.pf_section not in SECTIONS:
            st.session_state.pf_section = SECTION_NONE
        st.selectbox("الشعبة / الاختصاص", SECTIONS, key="pf_section")
    st.selectbox("لغة المحتوى", ["العربية", "الفرنسية", "الإنجليزية"], key="pf_language")

    st.divider()
    st.subheader("📂 سجل المحادثات")
    n1, n2 = st.columns(2)
    n1.button("➕ جذاذة", on_click=reset_chat, args=("lesson",), use_container_width=True)
    n2.button("➕ امتحان", on_click=reset_chat, args=("exam",), use_container_width=True)
    for c_id, c_title, c_time, c_type in get_all_chats(user["owner"]):
        kind = "lesson" if c_type in ("lesson",) else "exam"
        row = st.columns([6, 1])
        row[0].button(f"{'📄' if kind == 'lesson' else '📝'} {c_title}", key=f"open_{c_id}", help=c_time,
                      on_click=load_chat, args=(c_id, kind), use_container_width=True)
        row[1].button("🗑", key=f"del_{c_id}", help="حذف", on_click=remove_chat, args=(c_id, kind))

# =====================================================================
# التبويبات
# =====================================================================
tab_profile, tab_lesson, tab_exam, tab_grader = st.tabs(["👤 ملفي", "📄 جذاذة", "📝 امتحان", "🔍 تصحيح الأوراق"])


# ---------- ملفي (يُنفَّذ أولاً لتكون القيم محدّثة) ----------
def profile_ui():
    st.subheader("👤 ملفي البيداغوجي")
    st.caption("املأه مرة واحدة. كل ما تكتبه هنا يُستعمل تلقائياً في الجذاذات والامتحانات.")
    c1, c2 = st.columns(2)
    c1.text_input("الاسم واللقب (يظهر في الوثائق)", key="pf_name")
    c2.text_input("المؤسسة", key="pf_school")
    st.caption("المادة والمستوى واللغة تختارها من القائمة الجانبية على اليسار 👈")
    st.slider("درجة الإبداع", 0.1, 1.0, step=0.1, key="pf_creativity", help="0.1 دقيق جداً (رياضيات)، 1.0 إبداعي (تاريخ)")

    st.markdown("#### 🧭 طريقتي في العمل")
    st.caption("اكتب بكلامك كيف تحضّر درسك وكيف تبني امتحانك، أو دع التطبيق يتعلمها من وثائقك السابقة.")
    pc1, pc2 = st.columns([3, 1])
    with pc1:
        preset = st.selectbox("قالب جاهز حسب المجال (اختياري، تعدّله كما تشاء)", [PRESET_NONE] + list(PRESETS), key="preset_pick")
    with pc2:
        st.write("")
        preset_go = st.button("تطبيق القالب", key="preset_go")
    if preset_go and preset != PRESET_NONE:
        st.session_state["replace_method_lesson"] = PRESETS[preset]["lesson"]
        st.session_state["replace_method_exam"] = PRESETS[preset]["exam"]
        st.rerun()
    for kind, label, ph in (
        ("lesson", "طريقتي في تحضير الدرس", "مثال: أبدأ بوضعية انطلاق، ثم سند ووثيقة، ثم أنشطة، وأختم بتقييم من 3 أسئلة."),
        ("exam", "طريقتي في بناء الامتحان", "مثال: جزء للمعرفة (8 ن)، جزء للتطبيق (8 ن)، ووضعية إدماجية (4 ن)."),
    ):
        replace = st.session_state.pop(f"replace_method_{kind}", None)
        if replace:
            st.session_state[f"pf_method_{kind}"] = replace
        pending = st.session_state.pop(f"pending_method_{kind}", None)
        if pending:
            cur = st.session_state.get(f"pf_method_{kind}", "").strip()
            st.session_state[f"pf_method_{kind}"] = (cur + "\n\n" if cur else "") + pending
        st.text_area(label, key=f"pf_method_{kind}", height=140, placeholder=ph)
    st.text_area("ملاحظات خاصة (مؤسستي، منهاجي، تفضيلاتي...)", key="pf_notes", height=90)

    with st.expander("✨ قلّدني: ارفع وثيقة سابقة لك ليتعلّم التطبيق أسلوبك"):
        kind_label = st.radio("نوع الوثيقة", ["جذاذة", "امتحان"], horizontal=True, key="mimic_kind")
        imgs = st.file_uploader("صور الوثيقة", type=IMG_TYPES, accept_multiple_files=True, key="mimic_files")
        if st.button("🔍 حلّل أسلوبي", key="mimic_go"):
            if not imgs:
                st.warning("ارفع صورة واحدة على الأقل.")
            else:
                with st.spinner("جاري تحليل أسلوبك..."):
                    try:
                        res = extract_style([f.getvalue() for f in imgs], kind_label)
                        st.session_state["pending_method_" + ("lesson" if kind_label == "جذاذة" else "exam")] = res
                        st.rerun()
                    except Exception as e:
                        log.exception("mimic")
                        st.error(f"حدث خطأ: {e}")
        st.caption("يُضاف الوصف إلى خانة «طريقتي» أعلاه لتراجعه وتعدّله قبل الحفظ.")

    if st.button("💾 حفظ ملفي", type="primary", key="save_profile"):
        if user["email"]:
            save_profile(user["email"], {k: st.session_state[f"pf_{k}"] for k in PROFILE_DEFAULTS})
            st.success("تم حفظ ملفك ✅")
        else:
            st.info("أنت زائر: التعديلات مطبّقة في هذه الجلسة فقط. ادخل ببريدك لحفظ ملفك.")


with tab_profile:
    profile_ui()

LANG = prof("language")
if not prof("subject").strip():
    st.warning("👈 اختر المادة والمستوى من القائمة الجانبية لنتائج أدق.")

# ---------- الجذاذة ----------
with tab_lesson:
    c_set, c_chat, c_prev = st.columns([1, 2, 2])
    with c_set:
        st.markdown("#### ⚙️ إعدادات الجذاذة")
        st.caption(f"المادة: {prof('subject') or '—'}  \nالمستوى: {grade_label()}  \nاللغة: {LANG}")
        st.caption("تُغيَّر من القائمة الجانبية 👈")
        lesson_attach = st.checkbox("إرفاق صور المحادثة كملحق", key="lesson_attach")
        st.download_button(
            "📄 تصدير الحوار (.txt)",
            "\n\n".join(f"[{m['role']}]: {m['content']}" for m in st.session_state.messages_lesson if m.get("content")),
            "chat_lesson.txt", "text/plain", use_container_width=True)
    with c_chat:
        st.markdown("#### 💬 المحادثة")
        workshop("lesson", "اكتب عنوان الدرس، أو اطلب تعديلاً...", lesson_system, LessonTurn, prof("creativity"))
    with c_prev:
        st.markdown("#### 📄 المسودة")
        preview_lesson(LANG, lesson_attach)

# ---------- الامتحان ----------
with tab_exam:
    c_set, c_chat, c_prev = st.columns([1, 2, 2])
    with c_set:
        st.markdown("#### ⚙️ إعدادات الفرض")
        ex_type = st.selectbox("نوع الفرض", ["فرض مراقبة عدد 1", "فرض مراقبة عدد 2", "فرض تأليفي عدد 1", "فرض تأليفي عدد 2",
                                              "فرض تأليفي عدد 3", "فرض منزلي", "اختبار"], key="ex_type")
        ex_custom = st.text_input("عنوان مخصص (اختياري)", key="ex_custom")
        ex_duration = st.selectbox("المدة", ["ساعة واحدة", "ساعتان", "3 ساعات", "4 ساعات"], index=1, key="ex_duration")
        ex_date = st.date_input("التاريخ", datetime.now(), key="ex_date").strftime("%d/%m/%Y")
        ex_total = st.number_input("المجموع", min_value=1.0, max_value=100.0, value=20.0, step=1.0, key="ex_total")
        exam_attach = st.checkbox("إرفاق صور المحادثة كملحق", key="exam_attach")
    exam_title = ex_custom.strip() or ex_type
    meta = doc_meta() | {"exam_title": exam_title, "duration": ex_duration, "date": ex_date, "total": float(ex_total)}
    with c_chat:
        st.markdown("#### 💬 المحادثة")
        workshop("exam", "اكتب مواضيع الامتحان، أو اطلب تعديلاً...",
                 lambda: exam_system(exam_title, float(ex_total)), ExamTurn, min(prof("creativity"), 0.5))
    with c_prev:
        st.markdown("#### 📝 المسودة")
        preview_exam(LANG, meta, float(ex_total), exam_attach)

# ---------- المصحح ----------
with tab_grader:
    st.subheader("🔍 المصحح البيداغوجي الآلي (جماعي)")
    st.info("📸 ارفع صور أوراق عدة تلاميذ دفعة واحدة. ستحصل على جدول نتائج (Excel) وتحليل لأخطاء القسم.")
    draft = st.session_state.get("draft_exam")
    use_key = st.checkbox("استعمل سلّم إصلاح الامتحان الحالي", value=bool(draft), disabled=not draft, key="use_key")
    student_imgs = st.file_uploader("أوراق التلاميذ", type=IMG_TYPES, accept_multiple_files=True, key="student_files")
    grading_instructions = st.text_area("شروط التصحيح وسلّم التنقيط", "قم بتصحيح الأخطاء، واعط علامة مع ملاحظة مشجعة.", key="grading_text")

    if st.button("🚀 تصحيح الأوراق", type="primary", key="grade_go"):
        if not student_imgs:
            st.warning("⚠️ الرجاء رفع ورقة واحدة على الأقل.")
        elif require_credit(len(student_imgs)):
            instr, max_note = grading_instructions, ""
            if use_key and draft:
                key_text = "\n".join(f"{i}. {q['text']} ({q['points']:g} ن) ← الإصلاح: {q['answer']}"
                                     for i, (_, q) in enumerate(all_questions(draft), 1))
                instr += f"\n\nسلّم الإصلاح الرسمي:\n{key_text}"
                max_note = f" العلامة القصوى {exam_total(draft):g}."
            results, bar = [], st.progress(0.0, text="بدء التصحيح...")
            for i, f in enumerate(student_imgs, 1):
                bar.progress((i - 1) / len(student_imgs), text=f"تصحيح الورقة {i}/{len(student_imgs)}: {f.name}")
                system = f"""أنت أستاذ ومفتش تربوي تونسي خبير في المادة ({prof('subject')}) والمستوى ({grade_label()}).
حلّل ورقة التلميذ المرفقة بدقة:
1. اقرأ اسم التلميذ (إن لم يظهر فاكتب اسم الملف: {f.name}) وانسخ إجابته.
2. حدّد النقاط القوية.
3. رصد الأخطاء وصنّفها (لغوي/مفهومي/تطبيقي/منهجي) مع التصويب.
4. أعطِ علامة حسب التعليمات التالية: ({instr}).{max_note}
5. اكتب ملاحظة بيداغوجية مشجعة."""
                try:
                    r = parse_structured(
                        [{"role": "system", "content": system},
                         {"role": "user", "content": [{"type": "image_url", "image_url": {"url": data_url(shrink(f.getvalue()))}}]}],
                        StudentResult, 0.2, kind="vision")
                    results.append(r.model_dump())
                except Exception as e:
                    log.exception("grading %s", f.name)
                    st.error(f"تعذر تصحيح {f.name}: {e}")
            bar.progress(1.0, text="اكتمل التصحيح")
            st.session_state.grades = results
            st.session_state.pop("class_analysis", None)

    results = st.session_state.get("grades")
    if results:
        df = pd.DataFrame([{"التلميذ": r["student_name"], "العلامة": r["score"], "من": r["max_score"], "الملاحظة": r["comment"]} for r in results])
        st.dataframe(df, use_container_width=True, hide_index=True)
        scores20 = [r["score"] / r["max_score"] * 20 for r in results if r["max_score"]]
        if scores20:
            m1, m2, m3 = st.columns(3)
            m1.metric("معدل القسم /20", f"{sum(scores20) / len(scores20):.2f}")
            m2.metric("أعلى علامة /20", f"{max(scores20):.2f}")
            m3.metric("أدنى علامة /20", f"{min(scores20):.2f}")

        xbio = io.BytesIO()
        with pd.ExcelWriter(xbio, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="النتائج")
            pd.DataFrame([{"التلميذ": r["student_name"], "نوع الخطأ": mk["category"], "الخطأ": mk["description"], "التصويب": mk["correction"]}
                          for r in results for mk in r["mistakes"]]).to_excel(w, index=False, sheet_name="الأخطاء")
        st.download_button("⬇️ تحميل جدول النتائج (Excel)", xbio.getvalue(), f"Resultats_{safe_name(prof('subject'))}.xlsx", XLSX_MIME)

        for i, r in enumerate(results):
            with st.expander(f"📄 {r['student_name']} — {r['score']:g}/{r['max_score']:g}"):
                if r["strengths"]:
                    st.markdown("**✅ نقاط القوة:** " + "؛ ".join(r["strengths"]))
                for mk in r["mistakes"]:
                    st.markdown(f"- ❌ ({mk['category']}) {mk['description']}  \n  ✔️ {mk['correction']}")
                st.info(r["comment"])

        if st.button("🧠 تحليل أخطاء القسم واقتراح حصة علاج", key="analyze"):
            if require_credit():
                with st.spinner("جاري التحليل..."):
                    try:
                        summary = "\n".join(f"- ({mk['category']}) {mk['description']}" for r in results for mk in r["mistakes"])
                        resp = llm(
                            "text", temperature=0.4,
                            messages=[{"role": "user", "content": f"""أنت مفتش تربوي. هذه أخطاء تلاميذ قسم في {prof('subject')} ({grade_label()}):
{summary}

1) لخّص أكثر الأخطاء تكراراً مرتبة حسب النسبة التقريبية.
2) حدّد الأسباب المحتملة.
3) اقترح حصة علاج مفصلة (الأهداف، الأنشطة، التمارين، التقييم) في حدود 45 دقيقة."""}])
                        st.session_state.class_analysis = resp.choices[0].message.content
                    except Exception as e:
                        log.exception("class analysis")
                        st.error(f"حدث خطأ: {e}")
        if st.session_state.get("class_analysis"):
            st.markdown(st.session_state.class_analysis)