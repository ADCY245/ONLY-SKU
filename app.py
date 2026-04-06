import os
import uuid
from io import BytesIO

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
import pandas as pd

from analyzer import analyze_excel


app = Flask(__name__)
app.config["SECRET_KEY"] = "sku-analyzer-secret"
RESULT_CACHE = {}


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        preview_rows=None,
        preview_columns=None,
        download_token=None,
        page=None,
        page_size=None,
        total_rows=None,
        total_pages=None,
        page_numbers=None,
    )


@app.route("/analyze", methods=["POST"])
def analyze():
    uploaded_file = request.files.get("file")

    if not uploaded_file or uploaded_file.filename == "":
        flash("Please choose an Excel file first.")
        return redirect(url_for("index"))

    if not uploaded_file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Only Excel files (.xlsx or .xls) are supported.")
        return redirect(url_for("index"))

    try:
        output_df = analyze_excel(uploaded_file)
    except ValueError as exc:
        flash(str(exc))
        return redirect(url_for("index"))
    except Exception as exc:
        flash(f"Could not process the file: {exc}")
        return redirect(url_for("index"))

    output_stream = BytesIO()
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Analyzed Output")

    output_stream.seek(0)
    original_name = uploaded_file.filename.rsplit(".", 1)[0]
    download_name = f"{original_name}_analyzed.xlsx"

    download_token = str(uuid.uuid4())
    columns = list(output_df.columns)
    rows = output_df.fillna("").to_dict(orient="records")
    RESULT_CACHE[download_token] = {
        "filename": download_name,
        "content": output_stream.getvalue(),
        "columns": columns,
        "rows": rows,
    }

    return redirect(url_for("preview", token=download_token, page=1))


@app.route("/preview/<token>", methods=["GET"])
def preview(token):
    result = RESULT_CACHE.get(token)
    if not result:
        flash("That preview is no longer available. Please analyze the file again.")
        return redirect(url_for("index"))

    try:
        page = int(request.args.get("page", 1))
    except (TypeError, ValueError):
        page = 1

    page_size = 100
    rows = result.get("rows") or []
    columns = result.get("columns") or []
    total_rows = len(rows)
    total_pages = max(1, (total_rows + page_size - 1) // page_size)
    page = max(1, min(page, total_pages))
    page_numbers = list(range(1, total_pages + 1))

    start = (page - 1) * page_size
    end = start + page_size
    preview_rows = rows[start:end]

    return render_template(
        "index.html",
        preview_rows=preview_rows,
        preview_columns=columns,
        download_token=token,
        page=page,
        page_size=page_size,
        total_rows=total_rows,
        total_pages=total_pages,
        page_numbers=page_numbers,
    )


@app.route("/download/<token>", methods=["GET"])
def download(token):
    result = RESULT_CACHE.get(token)
    if not result:
        flash("That preview is no longer available. Please analyze the file again.")
        return redirect(url_for("index"))

    return send_file(
        BytesIO(result["content"]),
        as_attachment=True,
        download_name=result["filename"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
