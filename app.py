from io import BytesIO

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
import pandas as pd

from analyzer import analyze_excel


app = Flask(__name__)
app.config["SECRET_KEY"] = "sku-analyzer-secret"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


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

    return send_file(
        output_stream,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
