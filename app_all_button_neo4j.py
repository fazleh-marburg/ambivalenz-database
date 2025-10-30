from flask import Flask, render_template, request, redirect, url_for
from openpyxl import Workbook, load_workbook
from werkzeug.utils import secure_filename
import os
import spacy
from neo4j import GraphDatabase  # <-- added Neo4j driver

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

EXCEL_FILE = "objekt_data.xlsx"

# Load spaCy German model
nlp = spacy.load("de_core_news_sm")

# Neo4j connection
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "neo4j"  # replace with your Neo4j password
driver = GraphDatabase.driver(NEO4J_URI, auth=(NEO4J_USER, NEO4J_PASSWORD))

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.title = "Objekt Informationen"
    ws.append(["Property", "Value", "Status", "Name Entity"])
    wb.save(EXCEL_FILE)


def get_field_list():
    """Read 'Property' column from Excel file, or fallback default list."""
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        properties = []
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            val = str(row[0]).strip() if row[0] else ""
            if val:
                properties.append(val)
        if properties:
            return properties
    return [
        "Titel", "Künstler*in/Autor*in", "Sichtbare oder genannte Personen",
        "Entstehungsjahr", "Ort der Entstehung / Nutzung", "Gattung / Genre", "Technik",
        "Dimensionen", "Kurzbeschreibung", "Motive / Topoi", "Narrative / Diskurse",
        "Historischer Kontext", "Antiziganistische / Stigmatisierende Elemente",
        "Agency", "Verknüpfung", "Narrativwandel bei Medienwechsel", "Rezeptionsweg",
        "Sammlung / Archiv", "Provenienz", "Literatur",
        "Ausstellungen / Aufführungen / Veröffentlichungen", "Objekt- oder Werkteil",
        "Rechte / Lizenzen", "Digitalisat-Link/Pfad", "Metadaten-Status",
        "Erfasst von", "Erfassungsdatum", "Kommentar / Anmerkung", "Versionsgeschichte"
    ]


def extract_name_entities(text):
    """Return list of PER, ORG, LOC entities"""
    if not text:
        return []
    doc = nlp(text)
    entities = [ent.text.strip() for ent in doc.ents if ent.label_ in ("PER", "ORG", "LOC")]
    seen = []
    for e in entities:
        if e and e not in seen:
            seen.append(e)
    return seen


def save_to_neo4j(fields_dict):
    """Save each property-value pair as an Object node"""
    with driver.session() as session:
        # Create a single Object node with all properties
        session.run(
            "CREATE (o:Object $props)",
            props=fields_dict
        )


@app.route('/')
def form():
    fields = get_field_list()
    prefill = request.args.get('prefill', '').strip()
    data = {}
    if prefill:
        data[fields[0]] = prefill
    colors = {}
    name_entities = {}
    return render_template(
        "objekt_form_buttons.html",
        fields=fields, data=data, colors=colors, name_entities=name_entities
    )


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files or request.files['file'].filename == '':
        return "❌ Keine Datei ausgewählt", 400

    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    wb = load_workbook(filepath)
    ws = wb.active

    data, colors, name_entities, properties = {}, {}, {}, []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        key = str(row[0]).strip()
        value = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        data[key] = value
        colors[key] = "yellow" if value else "red"
        name_entities[key] = extract_name_entities(value)
        properties.append(key)

    fields = properties
    return render_template(
        "objekt_form_buttons.html",
        fields=fields, data=data, colors=colors, name_entities=name_entities
    )


@app.route('/extract_entities', methods=['POST'])
def extract_entities():
    fields = get_field_list()
    data, colors, name_entities = {}, {}, {}

    for key in fields:
        value = request.form.get(key, "").strip()
        data[key] = value
        ne = extract_name_entities(value)
        name_entities[key] = ne
        user_color = request.form.get(f"traffic_{key}")
        if user_color:
            colors[key] = user_color
        else:
            colors[key] = "green" if ne else "yellow" if value else "red"

    return render_template(
        "objekt_form_buttons.html",
        fields=fields, data=data, colors=colors, name_entities=name_entities
    )


@app.route('/submit', methods=['POST'])
def submit():
    fields = get_field_list()
    posted = request.form
    data_to_save = {}

    # Collect standard fields
    for field in fields:
        val = posted.get(field, "").strip()
        data_to_save[field] = val

    # Collect custom fields
    for key, value in posted.items():
        if key.startswith("custom_property_"):
            index = key.split("_")[-1]
            prop_name = value.strip()
            prop_val = posted.get(f"custom_value_{index}", "").strip()
            if prop_name:
                data_to_save[prop_name] = prop_val

    # Save to Neo4j
    save_to_neo4j(data_to_save)

    # Optional: also save to Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Objekt Informationen"
    ws.append(["Property", "Value"])
    for k, v in data_to_save.items():
        ws.append([k, v])
    wb.save(EXCEL_FILE)

    return redirect(url_for('success'))


@app.route('/success')
def success():
    return "<h3>✅ Daten erfolgreich gespeichert!</h3>"


if __name__ == '__main__':
    app.run(debug=True)
