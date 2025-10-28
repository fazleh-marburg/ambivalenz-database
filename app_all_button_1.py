from flask import Flask, render_template, request, redirect, url_for
from openpyxl import Workbook, load_workbook
from werkzeug.utils import secure_filename
import os
import spacy

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

EXCEL_FILE = "objekt_data.xlsx"

# Load spaCy German model
nlp = spacy.load("de_core_news_sm")

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.title = "Objekt Informationen"
    ws.append(["Property", "Value", "Status", "Name Entity"])
    wb.save(EXCEL_FILE)

def get_field_list():
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
    return seen  # always return a list

@app.route('/')
def form():
    fields = get_field_list()
    prefill = request.args.get('prefill', '').strip()
    data = {}
    if prefill:
        # Prefill the first field (Titel) with clicked entity
        data[fields[0]] = prefill
    colors = {}
    name_entities = {}
    return render_template("objekt_form_buttons.html",
                           fields=fields, data=data, colors=colors, name_entities=name_entities)

import pandas as pd

@app.route('/upload', methods=['POST'])
def upload():
    """Upload Excel or CSV and pre-fill form"""
    if 'file' not in request.files or request.files['file'].filename == '':
        return "❌ Keine Datei ausgewählt", 400

    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    # Try to read with pandas (auto-detects .xlsx, .xls, .csv, etc.)
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_csv(filepath)
        except Exception as e:
            return f"❌ Fehler beim Lesen der Datei: {str(e)}", 500

    data = {}
    colors = {}
    name_entities = {}

    # Expecting columns: Property | Value
    for _, row in df.iterrows():
        if len(row) < 2:
            continue
        key = str(row.iloc[0]).strip()
        value = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ""
        if not key:
            continue
        data[key] = value
        colors[key] = "yellow" if value else "red"
        name_entities[key] = extract_name_entities(value)

    fields = get_field_list()
    return render_template(
        "objekt_form_buttons.html",
        fields=fields,
        data=data,
        colors=colors,
        name_entities=name_entities
    )


@app.route('/extract_entities', methods=['POST'])
def extract_entities():
    """Extract named entities from all fields"""
    fields = get_field_list()
    data = {}
    colors = {}
    name_entities = {}

    for key in fields:
        value = request.form.get(key, "").strip()
        data[key] = value
        ne = extract_name_entities(value)
        name_entities[key] = ne
        colors[key] = "green" if ne else "yellow" if value else "red"

    return render_template("objekt_form_buttons.html",
                           fields=fields, data=data, colors=colors, name_entities=name_entities)

@app.route('/submit', methods=['POST'])
def submit():
    """Save form data to Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Objekt Informationen"
    ws.append(["Property", "Value", "Status", "Name Entity"])

    fields = get_field_list()
    posted = request.form

    for field in fields:
        val = posted.get(field, "").strip()
        status = posted.get(f"traffic_{field}", "red")
        ne = extract_name_entities(val)
        ws.append([field, val, status, ", ".join(ne)])

    # Handle custom fields if any
    for key, value in posted.items():
        if key.startswith("custom_property_"):
            index = key.split("_")[-1]
            prop_name = value.strip()
            prop_val = posted.get(f"custom_value_{index}", "").strip()
            prop_status = posted.get(f"traffic_custom_{index}", "red")
            ne = extract_name_entities(prop_val)
            if prop_name:
                ws.append([prop_name, prop_val, prop_status, ", ".join(ne)])

    wb.save(EXCEL_FILE)
    return redirect(url_for('success'))

@app.route('/success')
def success():
    return "<h3>✅ Daten erfolgreich gespeichert!</h3>"

if __name__ == '__main__':
    app.run(debug=True)
