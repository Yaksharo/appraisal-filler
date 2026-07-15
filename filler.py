"""Fill the Appraisal Sheet and Report of Rating docx templates."""
import copy
import io
import os
import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docxcompose.composer import Composer

def resource_path(rel):
    """Resolve a bundled resource, both in dev and inside a PyInstaller exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


APPRAISAL_TPL = resource_path(os.path.join("templates_docx", "Appraisal_Sheet.docx"))
REPORT_TPL = resource_path(os.path.join("templates_docx", "Report_of_Rating.docx"))

FONT = "Calibri"


def _set_cell_valign_top(cell):
    """Anchor cell content to the top (Word sometimes inherits bottom
    alignment from the table style, pushing values to the lower line)."""
    from docx.oxml.ns import qn
    tc_pr = cell._tc.get_or_add_tcPr()
    v = tc_pr.find(qn('w:vAlign'))
    if v is None:
        v = tc_pr.makeelement(qn('w:vAlign'), {})
        tc_pr.append(v)
    v.set(qn('w:val'), 'top')


def _set_cell(cell, text, size=10, center=False, bold=False):
    # keep exactly one paragraph in the cell
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run(str(text))
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_cell_valign_top(cell)


def _total_units(subjects):
    total = 0.0
    for s in subjects:
        try:
            total += float(s["units"])
        except ValueError:
            pass
    return f"{total:.1f}"


def _add_data_row(table, template_row):
    """Clone a data row so tables can grow past the template's blank rows."""
    new = copy.deepcopy(template_row._tr)
    template_row._tr.addnext(new)


def _delete_row(table, row):
    row._tr.getparent().remove(row._tr)


W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def _strip_trailing_empty_paragraphs(doc):
    """Remove empty paragraphs at the end of the body (they can create a
    blank trailing page, like in the Report of Rating template)."""
    body = doc.element.body
    for el in reversed(list(body.iterchildren())):
        tag = el.tag.split('}')[1]
        if tag == 'sectPr':
            continue
        is_empty_p = (tag == 'p'
                      and not ''.join(el.itertext()).strip()
                      and not el.findall(f'.//{W_NS}drawing'))
        if is_empty_p:
            body.remove(el)
        else:
            break


def term_sy_label(term, sy):
    short = term.replace(" Term", "").strip()          # "1st"
    yy = sy.split("-")                                  # 2025-2026 -> 25-26
    return f"{short}/{yy[0][2:]}-{yy[1][2:]}" if len(yy) == 2 else f"{short}/{sy}"


# ---------------------------------------------------------------- appraisal

def build_appraisal(student, faculty_map=None):
    """Return a filled Document for one student (all terms)."""
    faculty_map = faculty_map or {}
    doc = Document(APPRAISAL_TPL)

    # Name / ID line
    for p in doc.paragraphs:
        if "Name" in p.text and "ID Number" in p.text:
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            for text, bold in [("Name: ", True), (student["name"], False),
                               ("\t\t", False), ("ID Number: ", True),
                               (student["id"], False)]:
                run = p.add_run(text)
                run.bold = bold
                run.font.name = FONT
                run.font.size = Pt(11)
            break

    # sort the student's terms: earlier SY first, 1st before 2nd
    order = {"1st Term": 0, "2nd Term": 1, "3rd Term": 2, "4th Term": 3}
    keys = sorted(student["terms"].keys(), key=lambda k: (k[1], order.get(k[0], 9)))

    for slot, key in enumerate(keys[:4]):
        term, sy = key
        subjects = student["terms"][key]
        table = doc.tables[slot]
        label = term_sy_label(term, sy)
        n_data_rows = len(table.rows) - 2  # minus header and totals rows
        while n_data_rows < len(subjects):
            _add_data_row(table, table.rows[1])
            n_data_rows += 1
        for i, subj in enumerate(subjects):
            row = table.rows[1 + i]
            _set_cell(row.cells[0], subj["code"], center=True)
            _set_cell(row.cells[1], subj["title"])
            _set_cell(row.cells[2], subj["units"], center=True)
            _set_cell(row.cells[3], subj["grade"], center=True)
            _set_cell(row.cells[4], faculty_map.get(subj["code"].replace(" ", ""), ""))
            _set_cell(row.cells[5], label, center=True)
        totals = table.rows[-1]
        _set_cell(totals.cells[2], _total_units(subjects), center=True, bold=True)

    _strip_trailing_empty_paragraphs(doc)
    return doc


def fill_appraisal(student, faculty_map=None):
    return _to_buffer(build_appraisal(student, faculty_map))


# ----------------------------------------------------------------- report

def build_report(student, term, sy, subjects, faculty_map=None, adviser="",
                 trim_rows=True):
    """Return a filled Document for one student and one term."""
    faculty_map = faculty_map or {}
    doc = Document(REPORT_TPL)

    info = doc.tables[0]
    _set_cell(info.rows[0].cells[1], student["name"], size=11, bold=True)
    section = subjects[0]["section"] if subjects else student.get("course", "")
    _set_cell(info.rows[0].cells[3], section, size=11, bold=True)
    _set_cell(info.rows[1].cells[1], term, size=11, bold=True)
    _set_cell(info.rows[1].cells[3], sy, size=11, bold=True)

    grades = doc.tables[1]
    first_data = 2                      # rows 0-1 are the header
    last_data = len(grades.rows) - 2    # last row is TOTAL
    n_slots = last_data - first_data + 1

    # grow if a term ever has more subjects than blank rows
    while n_slots < len(subjects):
        _add_data_row(grades, grades.rows[last_data])
        last_data += 1
        n_slots += 1

    for i, subj in enumerate(subjects):
        row = grades.rows[first_data + i]
        _set_cell(row.cells[0], subj["code"], center=True)
        _set_cell(row.cells[1], subj["title"])
        _set_cell(row.cells[2], subj["units"], center=True)
        _set_cell(row.cells[3], subj["rating"], center=True)
        _set_cell(row.cells[4], subj["grade"], center=True)
        _set_cell(row.cells[5], faculty_map.get(subj["code"].replace(" ", ""), ""))

    # remove the unused blank rows between the last subject and TOTAL
    if trim_rows:
        for idx in range(last_data, first_data + len(subjects) - 1, -1):
            _delete_row(grades, grades.rows[idx])

    totals = grades.rows[-1]
    _set_cell(totals.cells[2], _total_units(subjects), center=True, bold=True)

    if adviser:
        sig = doc.tables[2]
        _set_cell(sig.rows[0].cells[0], adviser, size=11, bold=True)

    _strip_trailing_empty_paragraphs(doc)
    return doc


def fill_report(student, term, sy, subjects, faculty_map=None, adviser="",
                trim_rows=True):
    return _to_buffer(build_report(student, term, sy, subjects,
                                   faculty_map, adviser, trim_rows))


# ----------------------------------------------------------------- combine

def _page_break_before(doc):
    """Mark the first paragraph of a document to start on a new page."""
    from docx.oxml.ns import qn
    body = doc.element.body
    first_p = None
    for el in body.iterchildren():
        if el.tag == f'{W_NS}p':
            first_p = el
            break
        if el.tag == f'{W_NS}tbl':
            # body starts with a table: insert a tiny paragraph before it
            p = body.makeelement(qn('w:p'), {})
            el.addprevious(p)
            first_p = p
            break
    if first_p is None:
        return
    pPr = first_p.find(qn('w:pPr'))
    if pPr is None:
        pPr = first_p.makeelement(qn('w:pPr'), {})
        first_p.insert(0, pPr)
    if pPr.find(qn('w:pageBreakBefore')) is None:
        pPr.append(pPr.makeelement(qn('w:pageBreakBefore'), {}))


def combine_documents(docs):
    """Merge several filled Documents into one, each starting on a new page."""
    if not docs:
        return None
    composer = Composer(docs[0])
    for doc in docs[1:]:
        _page_break_before(doc)
        composer.append(doc)
    return _to_buffer(composer.doc)


def _to_buffer(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
