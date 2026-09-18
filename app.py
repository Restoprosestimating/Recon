#!/usr/bin/env python3
"""Restopros Recon Tracker — Internal TAM → cost / royalty / profit workbook."""
from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for

from excel_builder import build_workbook
from parser import PRINT_SELECTION_NAMES, parse_internal_tam

ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
OUTPUTS = ROOT / "outputs"
UPLOADS.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("RECON_SECRET", "restopros-recon-tracker")
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024


def _safe_name(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s or "job")
    return s[:80] or "job"


@app.get("/")
def index():
    return render_template("index.html", print_selections=PRINT_SELECTION_NAMES)


@app.post("/process")
def process():
    f = request.files.get("estimate")
    if not f or not f.filename:
        flash("Upload an Xactimate Internal TAM PDF.")
        return redirect(url_for("index"))
    if not f.filename.lower().endswith(".pdf"):
        flash("File must be a PDF.")
        return redirect(url_for("index"))

    tmp = UPLOADS / _safe_name(f.filename)
    f.save(tmp)

    parsed = parse_internal_tam(str(tmp))
    settings = {
        "rate_sup": request.form.get("rate_sup", 30),
        "rate_wkr": request.form.get("rate_wkr", 22),
        "burden": request.form.get("burden", 0.18),
        "royalty": request.form.get("royalty", 0.08),
        "contingency": request.form.get("contingency", 0.05),
        "tax": request.form.get("tax", 0.0825),
        "negotiated": request.form.get("negotiated") or None,
    }
    # empty negotiated → None so builder uses RCV
    if settings["negotiated"] in ("", None):
        settings["negotiated"] = None
    else:
        settings["negotiated"] = settings["negotiated"].replace("$", "").replace(",", "")

    if not parsed.ok:
        return render_template(
            "index.html",
            print_selections=PRINT_SELECTION_NAMES,
            errors=parsed.errors,
            warnings=parsed.warnings,
            found=parsed.found_sections,
            missing=parsed.missing_sections,
            preview=parsed.raw_text_preview,
        )

    out_name = f"Restopros_Recon_{_safe_name(parsed.estimate_id or tmp.stem)}.xlsx"
    out_path = OUTPUTS / out_name
    build_workbook(parsed, settings, str(out_path))

    hours = sum(r.hours for r in parsed.labor_rows)
    return render_template(
        "result.html",
        parsed=parsed,
        hours=hours,
        filename=out_name,
        settings=settings,
        print_selections=PRINT_SELECTION_NAMES,
    )


@app.get("/download/<path:name>")
def download(name):
    path = OUTPUTS / Path(name).name
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=path.name)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5055)), debug=True)
