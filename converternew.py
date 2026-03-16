from flask import Flask, request, send_file, render_template
import os
import subprocess
from pdf2docx import Converter
from PIL import Image

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "converted"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return render_template("indexnew.html")


@app.route('/convert', methods=['POST'])
def convert_file():

    file = request.files['file']
    conversion_type = request.form['type']

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    filename = os.path.splitext(file.filename)[0]

    # PDF → DOCX
    if conversion_type == "pdf-docx":

        output_path = os.path.join(OUTPUT_FOLDER, filename + ".docx")

        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()

        return send_file(output_path, as_attachment=True)


    # DOCX → PDF
    elif conversion_type == "docx-pdf":

        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            input_path,
            "--outdir",
            OUTPUT_FOLDER
        ])

        output_path = os.path.join(OUTPUT_FOLDER, filename + ".pdf")

        return send_file(output_path, as_attachment=True)


    # TXT → PDF
    elif conversion_type == "txt-pdf":

        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            input_path,
            "--outdir",
            OUTPUT_FOLDER
        ])

        output_path = os.path.join(OUTPUT_FOLDER, filename + ".pdf")

        return send_file(output_path, as_attachment=True)


    # IMAGE → PDF
    elif conversion_type == "image-pdf":

        output_path = os.path.join(OUTPUT_FOLDER, filename + ".pdf")

        image = Image.open(input_path)
        image = image.convert("RGB")
        image.save(output_path)

        return send_file(output_path, as_attachment=True)


    return "Invalid conversion type"


if __name__ == "__main__":
    app.run(debug=True)