"""Web app that fills Appraisal Sheets and Reports of Rating from grade listing PDFs."""
import io
import re
import zipfile
from flask import Flask, render_template, request, send_file

import parser as grade_parser
import filler

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024

DOCX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".wordprocessingml.document")


def _safe(name):
    return re.sub(r"[^A-Za-z0-9 _.\-]", "", name).strip().replace(" ", "_")


def _parse_faculty_map(text):
    """Lines like 'CT103 = Juan M. Dela Cruz' -> {'CT103': 'Juan M. Dela Cruz'}"""
    out = {}
    for line in (text or "").splitlines():
        if "=" in line:
            code, name = line.split("=", 1)
            if name.strip():
                out[code.strip().replace(" ", "").upper()] = name.strip()
    return out


def _term_key_str(term, sy):
    return f"{term}|{sy}"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    files = request.files.getlist("pdfs")
    parsed = [grade_parser.parse_pdf(f.stream) for f in files if f.filename]
    merged = grade_parser.merge_terms(parsed)
    students = sorted(merged.values(), key=lambda s: s["name"])
    terms = sorted({k for s in students for k in s["terms"]})
    term_opts = [{"key": _term_key_str(t, sy), "label": f"{t} {sy}"}
                 for t, sy in terms]
    codes = sorted({subj["code"] for s in students
                    for subs in s["terms"].values() for subj in subs})
    faculty_prefill = "\n".join(f"{c} = " for c in codes)
    return render_template("index.html", students=students,
                           term_opts=term_opts, n_files=len(parsed),
                           faculty_prefill=faculty_prefill)


@app.route("/generate", methods=["POST"])
def generate():
    files = request.files.getlist("pdfs")
    parsed = [grade_parser.parse_pdf(f.stream) for f in files if f.filename]
    merged = grade_parser.merge_terms(parsed)

    make_appraisal = "appraisal" in request.form.getlist("docs")
    make_report = "report" in request.form.getlist("docs")
    adviser = request.form.get("adviser", "").strip()
    faculty_map = _parse_faculty_map(request.form.get("faculty_map", ""))
    only_ids = set(request.form.getlist("student"))
    only_terms = set(request.form.getlist("terms"))      # e.g. "1st Term|2025-2026"
    output_mode = request.form.get("output_mode", "individual")
    trim_rows = request.form.get("trim_rows", "yes") == "yes"

    students = [s for sid, s in sorted(merged.items(),
                                       key=lambda kv: kv[1]["name"])
                if not only_ids or sid in only_ids]

    def term_selected(term, sy):
        return not only_terms or _term_key_str(term, sy) in only_terms

    outputs = []  # list of (zip_path, filename, bytes)

    if output_mode == "combined":
        if make_appraisal:
            docs = [filler.build_appraisal(s, faculty_map) for s in students]
            buf = filler.combine_documents(docs)
            if buf:
                outputs.append(("", "Appraisal_Sheets_ALL.docx", buf.read()))
        if make_report:
            all_terms = sorted({k for s in students for k in s["terms"]})
            for term, sy in all_terms:
                if not term_selected(term, sy):
                    continue
                docs = [filler.build_report(s, term, sy, s["terms"][(term, sy)],
                                            faculty_map, adviser, trim_rows)
                        for s in students if (term, sy) in s["terms"]]
                buf = filler.combine_documents(docs)
                if buf:
                    tslug = term.replace(" ", "")
                    outputs.append(
                        ("", f"Report_of_Rating_{tslug}_{sy}_ALL.docx",
                         buf.read()))
    else:
        for s in students:
            base = _safe(s["name"])
            if make_appraisal:
                buf = filler.fill_appraisal(s, faculty_map)
                outputs.append(("Appraisal_Sheets/",
                                f"{base}_{s['id']}.docx", buf.read()))
            if make_report:
                for (term, sy), subjects in sorted(s["terms"].items()):
                    if not term_selected(term, sy):
                        continue
                    buf = filler.fill_report(s, term, sy, subjects,
                                             faculty_map, adviser, trim_rows)
                    tslug = term.replace(" ", "")
                    outputs.append((f"Reports_of_Rating/{tslug}_{sy}/",
                                    f"{base}_{s['id']}.docx", buf.read()))

    if not outputs:
        return "Nothing to generate. Check your selections.", 400

    # single combined file: send the docx directly, no zip needed
    if len(outputs) == 1 and output_mode == "combined":
        _, fname, data = outputs[0]
        return send_file(io.BytesIO(data), as_attachment=True,
                         download_name=fname, mimetype=DOCX_MIME)

    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, fname, data in outputs:
            zf.writestr(folder + fname, data)
    zbuf.seek(0)
    return send_file(zbuf, as_attachment=True,
                     download_name="advisee_documents.zip",
                     mimetype="application/zip")


def _free_private_port():
    """Pick a free port in the dynamic/private range (49152-65535)."""
    import random
    import socket
    for _ in range(50):
        port = random.randint(49152, 65535)
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return 0  # let the OS choose


if __name__ == "__main__":
    port = _free_private_port()
    print(f"\n  Advisee Document Filler running at http://127.0.0.1:{port}\n")
    app.run(debug=False, host="127.0.0.1", port=port)
