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
    return render_template("index.html", preview_rows=None, preview_columns=None, download_token=None)


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
    RESULT_CACHE[download_token] = {
        "filename": download_name,
        "content": output_stream.getvalue(),
    }

    preview_df = output_df.head(100).fillna("")
    return render_template(
        "index.html",
        preview_rows=preview_df.to_dict(orient="records"),
        preview_columns=list(preview_df.columns),
        download_token=download_token,
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
