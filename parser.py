"""Parse UNP 'Periodic Grades Listing' PDFs into structured student records."""
import re
import pdfplumber

STUDENT_RE = re.compile(r"^([A-ZÑ][A-ZÑa-z ,.\-'’]+?)\s*\((\d{2}-\d{4,6})\)\s*$")
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
UNITS_RE = re.compile(r"^\d\.\d$")            # 3.0, 2.0
GRADE_NUM_RE = re.compile(r"^\d\.\d{2}$")     # 2.50, 5.00
GRADE_TXT_RE = re.compile(r"^(INC|UD|NA|DRP|LR|NFE|NGA|W)$", re.I)
RATING_RE = re.compile(r"^\d(\.\d{2})?$")     # periodic rating (1..5)
PERIOD_RE = re.compile(r"Period:\s*(.+?)\s+(\d{4}-\d{4})")
CODE_NUM_RE = re.compile(r"^\d{2,3}[A-Z]?$")


def _parse_subject_line(line):
    toks = line.split()
    if len(toks) < 8 or not toks[0].isdigit() or not DATE_RE.match(toks[-1]):
        return None
    date_posted = toks.pop()
    # walk from the right: units, optional grade, optional periodic rating
    if not UNITS_RE.match(toks[-1]):
        return None
    units = toks.pop()
    grade, rating = "", ""
    if toks and (GRADE_NUM_RE.match(toks[-1]) or GRADE_TXT_RE.match(toks[-1])):
        grade = toks.pop()
    if toks and RATING_RE.match(toks[-1]) and len(toks[-1]) == 1:
        rating = toks.pop()
    # lab units, lec units
    if len(toks) < 4 or not UNITS_RE.match(toks[-1]) or not UNITS_RE.match(toks[-2]):
        return None
    lab = toks.pop()
    lec = toks.pop()
    # section = last two tokens (e.g. "DCT 1C")
    sec2 = toks.pop()
    sec1 = toks.pop()
    section = f"{sec1} {sec2}"
    # row no, subject code (may be split: "CSS 102" or joined: "CSS101")
    toks.pop(0)  # row number
    code = toks.pop(0)
    if toks and CODE_NUM_RE.match(toks[0]):
        code = f"{code} {toks.pop(0)}"
    title = " ".join(toks)
    return {
        "code": code, "title": title, "section": section,
        "lec": lec, "lab": lab, "rating": rating, "grade": grade,
        "units": units, "date_posted": date_posted,
    }


def parse_pdf(path_or_file):
    """Return {'term': str, 'sy': str, 'students': [ {...} ]}."""
    term, sy = "", ""
    students, current = [], None
    with pdfplumber.open(path_or_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                if not term:
                    m = PERIOD_RE.search(line)
                    if m:
                        term, sy = m.group(1).strip(), m.group(2)
                m = STUDENT_RE.match(line)
                if m:
                    current = {
                        "name": re.sub(r"\s+", " ", m.group(1)).strip(),
                        "id": m.group(2),
                        "course": "",
                        "subjects": [],
                    }
                    students.append(current)
                    continue
                if current is not None:
                    if not current["course"] and re.fullmatch(r"[A-Z]{2,6}", line):
                        current["course"] = line
                        continue
                    subj = _parse_subject_line(line)
                    if subj:
                        current["subjects"].append(subj)
    return {"term": term, "sy": sy, "students": students}


def merge_terms(parsed_list):
    """Merge several parsed PDFs into one dict keyed by student id."""
    merged = {}
    for parsed in parsed_list:
        key = (parsed["term"], parsed["sy"])
        for s in parsed["students"]:
            rec = merged.setdefault(s["id"], {
                "name": s["name"], "id": s["id"],
                "course": s["course"], "terms": {},
            })
            if s["course"] and not rec["course"]:
                rec["course"] = s["course"]
            rec["terms"][key] = s["subjects"]
    return merged
