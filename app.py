from flask import Flask, request, send_file, render_template_string
from dd1750_core import parse_bom_and_fill_dd1750
import tempfile

app = Flask(__name__)

HTML_TEMPLATE = """
<!doctype html>
<title>BOM ➜ DD1750 (Flask v4 patched)</title>
<h2>BOM ➜ DD1750 (Flask v4 patched)</h2>
<form method=post enctype=multipart/form-data>
  <p><b>BOM (PDF)</b><br><input type=file name=bom_file></p>
  <p><b>DD1750 template (flat PDF)</b><br><input type=file name=template_file></p>
  <p><b>Start parsing at page (0-based)</b><br>
  <input type=text name=start_page value="0"> Use 0 for normal BOMs.</p>
  <p><input type=submit value="Generate DD1750"></p>
</form>
{% if error %}<p style="color:red;"><b>{{ error }}</b></p>{% endif %}
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        bom = request.files.get("bom_file")
        template = request.files.get("template_file")
        start_page = int(request.form.get("start_page", 0))

        if not bom or not template:
            return render_template_string(HTML_TEMPLATE, error="Both BOM and DD1750 template files are required.")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as bom_temp,              tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as template_temp,              tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as output_temp:

            bom.save(bom_temp.name)
            template.save(template_temp.name)

            try:
                parse_bom_and_fill_dd1750(bom_temp.name, template_temp.name, output_temp.name, start_page)
                return send_file(output_temp.name, as_attachment=True, download_name="DD1750_filled.pdf")
            except Exception as e:
                return render_template_string(HTML_TEMPLATE, error=f"Error: {e}")

    return render_template_string(HTML_TEMPLATE)