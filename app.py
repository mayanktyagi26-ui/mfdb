from flask import Flask, render_template, request, send_file
import os
from analysis import run_analysis

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():

    summary_table = None
    download_link = None

    if request.method == "POST":

        file = request.files["file"]
        start_date = request.form["start_date"]
        end_date = request.form["end_date"]

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        summary_df, output_file = run_analysis(file_path, start_date, end_date)

        if summary_df is not None:
            summary_table = summary_df.to_html(classes="table table-bordered")
            download_link = output_file

    return render_template("index.html",
                           table=summary_table,
                           download_link=download_link)


@app.route("/download")
def download():
    return send_file("outputs/MF_Combined_NAV_YOY_Screening.xlsx",
                     as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)

