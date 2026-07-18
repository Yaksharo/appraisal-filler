"""Create a PDF summary of notable student statuses in grade listings.

The Periodic Grades Listing has subject-level grades, not an official
student-enrollment status.  This module therefore keeps the raw grade codes
visible and labels any roster comparison as an inference for adviser review.
"""
import io
import re
from datetime import datetime
from html import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)


# Categories intentionally overlap: a student can have more than one
# subject-level status, and may also be absent from the reference-term roster.
CATEGORY_DEFINITIONS = (
    ("incomplete", "Incomplete (INC)",
     "Subject records marked INC in the uploaded listing."),
    ("possible_not_enrolled", "Possible not enrolled / no attendance (NA)",
     "Subject records marked NA. Verify the student's enrollment with the registrar."),
    ("unofficially_dropped", "Unofficially dropped (UD)",
     "Subject records marked UD in the uploaded listing."),
    ("officially_dropped", "Officially dropped (OD)",
     "Subject records marked OD in the uploaded listing."),
    ("dropped", "Dropped (DRP)",
     "Subject records marked DRP in the uploaded listing."),
    ("not_found_in_reference", "Not found in reference term",
     "Present in an earlier uploaded term but absent from the latest term. "
     "This is an inference only, not an official enrollment decision."),
    ("failed", "Failed subject (5.00)",
     "Subject records with a final grade of 5.00."),
    ("other_status", "Other registrar status (LR, NFE, NGA, W)",
     "Other non-numeric status codes copied exactly from the uploaded listing."),
    ("no_final_grade", "No final grade recorded",
     "A subject was parsed without a final grade. This can be normal for an "
     "in-progress or incomplete listing."),
    ("no_subjects", "No subjects parsed / review needed",
     "A student entry was found for this term, but no subject rows were parsed."),
)

_DIRECT_GRADE_CATEGORIES = {
    "INC": "incomplete",
    "NA": "possible_not_enrolled",
    "UD": "unofficially_dropped",
    "OD": "officially_dropped",
    "DRP": "dropped",
}
_OTHER_STATUS_CODES = {"LR", "NFE", "NGA", "W"}
_TERM_ORDER = {
    "1st term": 1, "first term": 1,
    "2nd term": 2, "second term": 2,
    "3rd term": 3, "third term": 3,
    "4th term": 4, "fourth term": 4,
    "mid year": 5, "midyear": 5, "summer": 5,
}


def _term_sort_key(term_key):
    """Return a stable chronological key for ``(term, school_year)``."""
    term, school_year = term_key
    match = re.search(r"(\d{4})\s*-\s*(\d{4})", school_year or "")
    if match:
        year_key = (int(match.group(1)), int(match.group(2)))
    else:
        # Put an unrecognised school year after recognised ones while keeping
        # its order deterministic for the report.
        year_key = (9999, 9999)
    normal_term = " ".join((term or "").lower().split())
    return (*year_key, _TERM_ORDER.get(normal_term, 99), normal_term,
            school_year or "")


def _term_label(term_key):
    term, school_year = term_key
    return " ".join(part for part in (term, school_year) if part) or "Unspecified term"


def _student_sort_key(student):
    return ((student.get("name") or "").casefold(), student.get("id") or "")


def _subject_label(subject):
    code = (subject.get("code") or "").strip()
    title = (subject.get("title") or "").strip()
    if code and title:
        return f"{code} — {title}"
    return code or title or "Subject not identified"


def _record(student, term_key, subject=None, note="", grade=""):
    """Normalise a detail row used by every category."""
    return {
        "name": student.get("name") or "Unnamed student",
        "id": student.get("id") or "—",
        "course": student.get("course") or "—",
        "term": _term_label(term_key),
        "subject": note or _subject_label(subject or {}),
        "grade": grade or "—",
    }


def _category_for_grade(grade):
    """Classify one final-grade value without assigning hidden meanings."""
    grade = (grade or "").strip().upper()
    if grade in _DIRECT_GRADE_CATEGORIES:
        return _DIRECT_GRADE_CATEGORIES[grade]
    if grade in _OTHER_STATUS_CODES:
        return "other_status"
    if not grade:
        return "no_final_grade"
    try:
        if float(grade) == 5.0:
            return "failed"
    except ValueError:
        pass
    return None


def analyze_students(students):
    """Return status categories derived from merged parser student records.

    ``students`` may be an iterable of records or the dictionary returned by
    :func:`parser.merge_terms`.  The last chronological uploaded term is used
    only as a *reference roster*: earlier students missing from it are shown
    for review, not treated as officially dropped or unenrolled.
    """
    if isinstance(students, dict):
        students = students.values()
    students = sorted(list(students), key=_student_sort_key)
    term_keys = sorted({key for student in students
                        for key in student.get("terms", {})},
                       key=_term_sort_key)
    records = {key: [] for key, _title, _description in CATEGORY_DEFINITIONS}

    for student in students:
        terms = student.get("terms", {})
        for term_key in sorted(terms, key=_term_sort_key):
            subjects = terms[term_key]
            if not subjects:
                records["no_subjects"].append(_record(
                    student, term_key, note="No subject rows were parsed"))
                continue
            for subject in subjects:
                raw_grade = (subject.get("grade") or "").strip().upper()
                category = _category_for_grade(raw_grade)
                if category:
                    records[category].append(_record(
                        student, term_key, subject, grade=raw_grade))

    # A term-comparison needs at least two unique terms.  A gap can mean a
    # number of things (missing PDF, transfer, graduation, etc.), so its
    # report label deliberately avoids claiming a definite enrollment status.
    reference_term = term_keys[-1] if len(term_keys) > 1 else None
    if reference_term:
        for student in students:
            student_terms = student.get("terms", {})
            earlier_terms = [key for key in student_terms if key != reference_term]
            if earlier_terms and reference_term not in student_terms:
                last_seen = max(earlier_terms, key=_term_sort_key)
                records["not_found_in_reference"].append(_record(
                    student, reference_term,
                    note=f"Last seen: {_term_label(last_seen)}",
                    grade="Not listed"))

    categories = []
    for key, title, description in CATEGORY_DEFINITIONS:
        category_records = records[key]
        categories.append({
            "key": key,
            "title": title,
            "description": description,
            "records": category_records,
            "student_count": len({row["id"] for row in category_records}),
            "record_count": len(category_records),
        })
    return {
        "students": students,
        "student_count": len(students),
        "term_keys": term_keys,
        "reference_term": reference_term,
        "categories": categories,
    }


def _paragraph(text, style):
    return Paragraph(escape(str(text)).replace("\n", "<br/>"), style)


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
    canvas.line(doc.leftMargin, 10 * mm, doc.pagesize[0] - doc.rightMargin,
                10 * mm)
    canvas.setFillColor(colors.HexColor("#4B5563"))
    canvas.setFont("Helvetica", 7)
    canvas.drawString(doc.leftMargin, 6 * mm, "Student Status Summary")
    canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, 6 * mm,
                           f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def build_status_report_pdf(students):
    """Build and return a ready-to-read PDF status report buffer."""
    analysis = analyze_students(students)
    buffer = io.BytesIO()
    page_size = landscape(A4)
    doc = SimpleDocTemplate(
        buffer, pagesize=page_size,
        leftMargin=14 * mm, rightMargin=14 * mm,
        topMargin=14 * mm, bottomMargin=16 * mm,
        title="Student Status Summary",
        author="Advisee Document Filler",
    )

    base = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "StatusTitle", parent=base["Title"], fontName="Helvetica-Bold",
        fontSize=17, leading=21, textColor=colors.HexColor("#0F3D66"),
        spaceAfter=2 * mm,
    )
    subtitle_style = ParagraphStyle(
        "StatusSubtitle", parent=base["Normal"], fontName="Helvetica",
        fontSize=8.5, leading=11, textColor=colors.HexColor("#4B5563"),
        spaceAfter=2 * mm,
    )
    section_style = ParagraphStyle(
        "StatusSection", parent=base["Heading2"], fontName="Helvetica-Bold",
        fontSize=11, leading=14, textColor=colors.HexColor("#0F3D66"),
        spaceBefore=5 * mm, spaceAfter=1 * mm,
    )
    detail_style = ParagraphStyle(
        "StatusDetail", parent=base["Normal"], fontName="Helvetica",
        fontSize=7.2, leading=9, textColor=colors.HexColor("#111827"),
    )
    table_header_style = ParagraphStyle(
        "StatusTableHeader", parent=detail_style, fontName="Helvetica-Bold",
        fontSize=7.2, leading=8.5, textColor=colors.white,
    )
    summary_header_style = ParagraphStyle(
        "StatusSummaryHeader", parent=detail_style, fontName="Helvetica-Bold",
        fontSize=8, leading=9.5, textColor=colors.white, alignment=TA_CENTER,
    )
    summary_cell_style = ParagraphStyle(
        "StatusSummaryCell", parent=detail_style, fontSize=8, leading=10,
    )
    summary_number_style = ParagraphStyle(
        "StatusSummaryNumber", parent=summary_cell_style, alignment=TA_CENTER,
    )

    coverage = ", ".join(_term_label(key) for key in analysis["term_keys"])
    coverage = coverage or "No terms found"
    generated = datetime.now().astimezone().strftime("%d %B %Y, %I:%M %p")
    story = [
        _paragraph("Student Status Summary", title_style),
        _paragraph(f"Generated {generated} · Students found: "
                   f"{analysis['student_count']}", subtitle_style),
        _paragraph(f"Uploaded terms: {coverage}", subtitle_style),
    ]
    if analysis["reference_term"]:
        story.append(_paragraph(
            "Reference-term comparison: "
            f"{_term_label(analysis['reference_term'])}. "
            "Students absent from this listing are flagged only for review.",
            subtitle_style))
    else:
        story.append(_paragraph(
            "A reference-term comparison needs at least two uploaded terms.",
            subtitle_style))

    story.append(Spacer(1, 2 * mm))
    summary_data = [[
        _paragraph("Category", summary_header_style),
        _paragraph("Students", summary_header_style),
        _paragraph("Records", summary_header_style),
    ]]
    for category in analysis["categories"]:
        summary_data.append([
            _paragraph(category["title"], summary_cell_style),
            _paragraph(category["student_count"], summary_number_style),
            _paragraph(category["record_count"], summary_number_style),
        ])
    summary = Table(summary_data, colWidths=[178 * mm, 28 * mm, 28 * mm],
                    repeatRows=1, hAlign="LEFT")
    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F3D66")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6 * mm),
    ]))
    story.append(summary)
    story.append(_paragraph(
        "Important: Grade codes are subject-level values copied from the "
        "uploaded PDFs. Categories can overlap. “Not found in reference term” "
        "only identifies a roster gap; confirm enrollment, transfer, graduation, "
        "or drop status with the registrar.",
        subtitle_style))

    detail_headers = ["Student", "ID", "Course", "Term",
                      "Subject / note", "Grade / status"]
    detail_widths = [46 * mm, 27 * mm, 23 * mm, 34 * mm, 111 * mm, 24 * mm]
    for category in analysis["categories"]:
        if not category["records"]:
            continue
        story.append(_paragraph(
            f"{category['title']} — {category['student_count']} student(s), "
            f"{category['record_count']} record(s)", section_style))
        story.append(_paragraph(category["description"], subtitle_style))
        detail_data = [[_paragraph(label, table_header_style)
                        for label in detail_headers]]
        for row in category["records"]:
            detail_data.append([
                _paragraph(row["name"], detail_style),
                _paragraph(row["id"], detail_style),
                _paragraph(row["course"], detail_style),
                _paragraph(row["term"], detail_style),
                _paragraph(row["subject"], detail_style),
                _paragraph(row["grade"], detail_style),
            ])
        detail = Table(detail_data, colWidths=detail_widths, repeatRows=1,
                       hAlign="LEFT", splitByRow=1)
        detail.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4E89")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F8FAFC")]),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 1.7 * mm),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1.7 * mm),
            ("TOPPADDING", (0, 0), (-1, -1), 1.1 * mm),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.1 * mm),
        ]))
        story.append(detail)

    if not any(category["records"] for category in analysis["categories"]):
        story.append(_paragraph("No notable status records were found in the "
                                "uploaded listings.", section_style))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buffer.seek(0)
    return buffer
