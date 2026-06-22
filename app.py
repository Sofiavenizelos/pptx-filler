import os
import io
import copy
from flask import Flask, request, send_file, jsonify
from pptx import Presentation

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

# Groups of tag prefixes that form one "ordering table" row.
# If ALL fields in a row are empty, that row gets deleted from the table.
ORDERING_ROW_FIELDS = [
    ["part_number_{i}", "part_description_{i}", "compatibility_{i}", "branded_equiv_{i}"]
    for i in range(1, 5)
]


def fill_text_frame(text_frame, data):
    for para in text_frame.paragraphs:
        for run in para.runs:
            for key, val in data.items():
                tag = f"{{{{{key}}}}}"
                if tag in run.text:
                    run.text = run.text.replace(tag, "" if val is None else str(val))


def fill_table(table, data):
    for row in table.rows:
        for cell in row.cells:
            fill_text_frame(cell.text_frame, data)


def remove_empty_ordering_rows(prs, data):
    """Delete rows in any table where all 4 ordering fields for that row index are empty."""
    for row_index, field_templates in enumerate(ORDERING_ROW_FIELDS, start=1):
        field_names = [f.format(i=row_index) for f in field_templates]
        values = [data.get(f, "") for f in field_names]
        row_is_empty = all(v == "" or v is None for v in values)
        if not row_is_empty:
            continue

        # Find a table row that originally contained these tags (now filled with "")
        # and remove it. We match by checking if the row's cells are now blank
        # AND the row is part of a 4-row ordering table (heuristic: table has 4+ rows).
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_table:
                    continue
                table = shape.table
                if len(table.rows) < 4:
                    continue  # not the ordering table
                # Walk rows from bottom to top so removal doesn't shift indices we still need
                for tr in list(table.rows)[::-1]:
                    cell_texts = [c.text.strip() for c in tr.cells]
                    if all(t == "" for t in cell_texts):
                        tr._tr.getparent().remove(tr._tr)


def fill_template(data):
    prs = Presentation(TEMPLATE_PATH)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                fill_text_frame(shape.text_frame, data)
            if shape.has_table:
                fill_table(shape.table, data)

    remove_empty_ordering_rows(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    try:
        file_buf = fill_template(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    filename = f"{data.get('part_number_1', 'datasheet')}.pptx"
    return send_file(
        file_buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))