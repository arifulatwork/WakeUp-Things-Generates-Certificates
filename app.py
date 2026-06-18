from flask import Flask, request, send_file, render_template_string, jsonify
import pandas as pd
import openpyxl
from docxtpl import DocxTemplate
import os
import zipfile
import io
import re
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), "generated")
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "certificate_template.docx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Certificate Generator</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; background: #f0f4f8; min-height: 100vh;
         display: flex; align-items: center; justify-content: center; padding: 24px; }
  .card { background: #fff; border-radius: 12px; box-shadow: 0 4px 24px rgba(0,0,0,.1);
          padding: 40px; width: 100%; max-width: 620px; }
  .logo { text-align: center; margin-bottom: 28px; }
  .logo svg { width: 56px; height: 56px; }
  h1 { font-size: 22px; color: #1F4E79; text-align: center; margin-bottom: 6px; }
  p.sub { font-size: 13px; color: #666; text-align: center; margin-bottom: 28px; }

  .drop-zone { border: 2px dashed #BDD7EE; border-radius: 8px; padding: 32px 16px;
               text-align: center; cursor: pointer; transition: all .2s; background: #f7fbff;
               position: relative; }
  .drop-zone.drag-over { border-color: #1F4E79; background: #EBF5FB; }
  .drop-zone input[type=file] { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .drop-zone .icon { font-size: 36px; margin-bottom: 8px; }
  .drop-zone p { font-size: 14px; color: #555; }
  .drop-zone .hint { font-size: 12px; color: #999; margin-top: 4px; }
  #file-name { margin-top: 10px; font-size: 13px; color: #1F4E79; font-weight: bold;
               min-height: 18px; text-align: center; }

  .info-box { background: #EBF5FB; border-left: 4px solid #1F4E79; border-radius: 4px;
              padding: 12px 14px; margin: 20px 0; font-size: 12.5px; color: #444; line-height: 1.6; }
  .info-box strong { color: #1F4E79; }

  .checkbox-group { margin: 16px 0; padding: 12px 14px; background: #f7fbff; border-radius: 6px; border: 1px solid #d0e0f0; }
  .checkbox-group label { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #333; cursor: pointer; }
  .checkbox-group input[type=checkbox] { width: 16px; height: 16px; accent-color: #1F4E79; }

  button[type=submit] { width: 100%; padding: 14px; background: #1F4E79; color: #fff;
                        border: none; border-radius: 8px; font-size: 15px; font-weight: bold;
                        cursor: pointer; transition: background .2s; margin-top: 8px; }
  button[type=submit]:hover { background: #16365a; }
  button[type=submit]:disabled { background: #aaa; cursor: not-allowed; }

  #progress { display: none; margin-top: 16px; }
  .bar-wrap { background: #e0e0e0; border-radius: 99px; height: 8px; overflow: hidden; }
  .bar { height: 100%; background: #1F4E79; border-radius: 99px;
         animation: indeterminate 1.5s infinite ease-in-out; }
  @keyframes indeterminate {
    0%   { transform: translateX(-100%); width: 60%; }
    100% { transform: translateX(200%); width: 60%; }
  }
  #status { font-size: 13px; color: #555; margin-top: 8px; text-align: center; }

  #result { display: none; margin-top: 20px; text-align: center; }
  #result .success { color: #1a7a1a; font-size: 15px; font-weight: bold; margin-bottom: 12px; }
  #result .dl-btn { display: inline-block; padding: 12px 28px; background: #217a21;
                    color: #fff; border-radius: 8px; text-decoration: none; font-weight: bold;
                    font-size: 14px; }
  #result .dl-btn:hover { background: #175c17; }

  #error-box { display: none; margin-top: 14px; padding: 12px 14px; background: #fff0f0;
               border-left: 4px solid #c00; border-radius: 4px; font-size: 13px; color: #c00; }

  .step-indicator { display: flex; justify-content: space-between; margin: 16px 0 24px; font-size: 12px; color: #888; }
  .step-indicator .active { color: #1F4E79; font-weight: bold; }
</style>
</head>
<body>
<div class="card">
  <div class="logo">
    <svg viewBox="0 0 56 56" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="56" height="56" rx="12" fill="#1F4E79"/>
      <path d="M14 38V20a2 2 0 0 1 2-2h24a2 2 0 0 1 2 2v18" stroke="#BDD7EE" stroke-width="2" stroke-linecap="round"/>
      <rect x="20" y="26" width="16" height="2" rx="1" fill="#fff"/>
      <rect x="20" y="30" width="10" height="2" rx="1" fill="#BDD7EE"/>
      <circle cx="28" cy="42" r="4" fill="#BDD7EE"/>
      <path d="M25 42h6M28 39v6" stroke="#1F4E79" stroke-width="1.5" stroke-linecap="round"/>
    </svg>
  </div>
  <h1>Certificate Generator</h1>
  <p class="sub">Upload your Excel file to generate internship certificates</p>

  <div class="step-indicator">
    <span class="active">1. Upload Excel</span>
    <span>2. Generate</span>
    <span>3. Download ZIP</span>
  </div>

  <form id="upload-form" enctype="multipart/form-data">
    <div class="drop-zone" id="drop-zone">
      <input type="file" name="file" id="file-input" accept=".xlsx,.xls">
      <div class="icon">📊</div>
      <p>Drag &amp; drop your Excel file here</p>
      <p class="hint">or click to browse &nbsp;·&nbsp; .xlsx / .xls</p>
    </div>
    <div id="file-name"></div>

    <div class="checkbox-group" id="sheet-picker" style="display:none;">
      <label for="sheet-select" style="display:block; margin-bottom:8px; cursor:default;">
        Generate certificates for which group / sheet?
      </label>
      <select name="sheet_name" id="sheet-select" disabled
              style="width:100%; padding:8px; border-radius:6px; border:1px solid #BDD7EE; font-size:13px;">
        <option value="">Loading sheets…</option>
      </select>
    </div>

    <div class="info-box">
      <strong>Expected Excel format:</strong><br>
      Pick the file and choose which group/cohort sheet to generate from the dropdown above.
      Each student sheet can have its header row anywhere near the top (the app auto-detects it),
      and should include columns like:<br>
      <code>Name · Surname · Age · Sector · Sector Database · Experience · Languages · Alergies ·
      COMPANY · Address · Tutor · Telephone · Email · VAT · PIC · OID · Website · Follow Up ·
      Comentarios · Dress Code · Pickup · Apresentação</code><br><br>
      <strong>Optional:</strong> A <code>TASK DATABASE</code> sheet with columns:<br>
      <code>Sector · Detailed tasks · Job-related knowledge, skills &amp; competences</code>
    </div>

    <button type="submit" id="submit-btn" disabled>Generate Certificates</button>
  </form>

  <div id="progress">
    <div class="bar-wrap"><div class="bar"></div></div>
    <div id="status">Processing…</div>
  </div>

  <div id="result">
    <div class="success">✅ Certificates generated successfully!</div>
    <a id="dl-link" class="dl-btn" href="#">⬇ Download ZIP</a>
  </div>

  <div id="error-box"></div>
</div>

<script>
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');
const fileName   = document.getElementById('file-name');
const sheetPicker= document.getElementById('sheet-picker');
const sheetSelect= document.getElementById('sheet-select');
const submitBtn  = document.getElementById('submit-btn');
const form       = document.getElementById('upload-form');
const progress   = document.getElementById('progress');
const statusEl   = document.getElementById('status');
const resultEl   = document.getElementById('result');
const errorBox   = document.getElementById('error-box');
const dlLink     = document.getElementById('dl-link');

async function loadSheets(file) {
  sheetPicker.style.display = 'block';
  sheetSelect.disabled = true;
  sheetSelect.innerHTML = '<option value="">Loading sheets…</option>';
  submitBtn.disabled = true;
  errorBox.style.display = 'none';

  const fd = new FormData();
  fd.append('file', file);
  try {
    const res = await fetch('/list-sheets', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Could not read sheets');
    if (!data.sheets || data.sheets.length === 0) {
      throw new Error('No group sheets found in this file.');
    }
    sheetSelect.innerHTML = '<option value="">— Select a group —</option>' +
      data.sheets.map(s => `<option value="${s.replace(/"/g, '&quot;')}">${s}</option>`).join('');
    sheetSelect.disabled = false;
  } catch (err) {
    sheetSelect.innerHTML = '<option value="">— could not load —</option>';
    errorBox.textContent = '❌ ' + err.message;
    errorBox.style.display = 'block';
  }
}

sheetSelect.addEventListener('change', () => {
  submitBtn.disabled = !(fileInput.files[0] && sheetSelect.value);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) {
    fileName.textContent = '📄 ' + fileInput.files[0].name;
    loadSheets(fileInput.files[0]);
  }
});

['dragenter','dragover'].forEach(e => dropZone.addEventListener(e, ev => {
  ev.preventDefault(); dropZone.classList.add('drag-over');
}));
['dragleave','drop'].forEach(e => dropZone.addEventListener(e, ev => {
  ev.preventDefault(); dropZone.classList.remove('drag-over');
}));
dropZone.addEventListener('drop', ev => {
  const f = ev.dataTransfer.files[0];
  if (f) {
    fileInput.files = ev.dataTransfer.files;
    fileName.textContent = '📄 ' + f.name;
    loadSheets(f);
  }
});

form.addEventListener('submit', async e => {
  e.preventDefault();
  errorBox.style.display = 'none';
  resultEl.style.display = 'none';
  progress.style.display = 'block';
  submitBtn.disabled = true;
  statusEl.textContent = 'Generating certificates…';

  const fd = new FormData(form);
  try {
    const res = await fetch('/generate-certificates', { method: 'POST', body: fd });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Unknown error' }));
      throw new Error(err.error || 'Server error');
    }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    dlLink.href = url;
    dlLink.download = 'certificates.zip';
    progress.style.display = 'none';
    resultEl.style.display = 'block';
    submitBtn.disabled = false;
  } catch (err) {
    progress.style.display = 'none';
    errorBox.textContent = '❌ ' + err.message;
    errorBox.style.display = 'block';
    submitBtn.disabled = false;
  }
});
</script>
</body>
</html>
"""

# Column name candidates, ordered so more specific names are tried first
# where two columns could both match (e.g. "Sector" vs "Sector Database").
COL_MAP = {
    'name': ['name', 'student', 'participant', 'first name', 'nome'],
    'surname': ['surname', 'last name', 'apellido', 'sobrenome'],
    'birth_date': ['date of birth', 'birth date', 'birth_date', 'dob', 'birthday'],
    'start_date': ['start date', 'start_date', 'start', 'begin', 'begin date'],
    'end_date': ['end date', 'end_date', 'end', 'finish', 'finish date'],
    # "sector database" is the normalized sector used to match TASK DATABASE -
    # it must be tried BEFORE the generic "sector" so it wins when both exist.
    'sector': ['sector database', 'sector', 'area', 'field', 'department', 'study course'],
    'company': ['company', 'host org', 'host_org', 'host organisation', 'organization', 'empresa'],
    'company_address': ['address', 'company address', 'host address', 'host_address', 'endereco', 'direccion'],
    'tutor': ['tutor', 'supervisor', 'mentor', 'contact'],
    'telephone': ['telephone', 'phone', 'contact number', 'tel', 'telefone'],
    'email': ['email', 'e-mail', 'mail', 'correo'],
    'project_code': ['project code', 'project id', 'project number', 'projeto'],
    'notes': ['notes', 'comments', 'observations', 'comentarios', 'feedback'],
    'dress_code': ['dress code', 'uniform', 'attire'],
    'pickup': ['pickup', 'pick up', 'collection'],
    'presentation': ['presentation', 'apresentação', 'start time', 'induction'],
    'age': ['age', 'idade', 'edad'],
    'languages': ['languages', 'idiomas', 'language'],
    'allergies': ['allergies', 'alergies', 'allergens', 'alergias'],
    'experience': ['experience', 'experiencia', 'work experience'],
    'vat': ['vat', 'nif', 'tax id', 'vat number'],
    'pic': ['pic', 'participant identification code'],
    'oid': ['oid', 'organisation id'],
    'website': ['website', 'web', 'site', 'pagina web'],
}

# Required header markers used to auto-detect which row is the real header row.
HEADER_MARKERS = ['name', 'surname', 'sector']

# Matches a date range like "17/06/2026 - 02/07/2026" anywhere in a cell.
DATE_RANGE_RE = re.compile(
    r'(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–—]\s*(\d{1,2}/\d{1,2}/\d{2,4})'
)
# Matches an Erasmus+/KA1xx style project code, e.g. "2025-1-ES01-KA121-VET-000339264".
PROJECT_CODE_RE = re.compile(r'\b\d{4}-\d+-[A-Z0-9-]*KA\d[A-Z0-9-]+\b')


def normalize_sector(sector):
    """Clean and normalize sector names for matching."""
    if not sector or pd.isna(sector):
        return ""
    sector = str(sector).strip()
    sector = re.sub(r'\s*[\(\[].*?[\)\]]', '', sector)
    sector = re.sub(r'\s*[-–—]\s*.*$', '', sector)
    return sector.strip().lower()


def find_header_row(sheet_path, sheet_name, max_scan_rows=15, markers=None, min_matches=2):
    """
    Scan the top of a sheet to find which row actually contains the column
    headers (e.g. 'Name', 'Surname', 'Sector'), instead of assuming it's
    always row 0 or row 1. Sheets in this workbook put a title (and
    sometimes a project-number row) above the real header row, and that
    offset varies sheet to sheet.

    Returns the 0-based row index to pass to pandas as `header=`.
    """
    if markers is None:
        markers = HEADER_MARKERS
    wb = openpyxl.load_workbook(sheet_path, data_only=True, read_only=True)
    try:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=max_scan_rows):
            values = [str(c.value).strip().lower() if c.value is not None else '' for c in row]
            matches = sum(1 for marker in markers if any(marker in v for v in values))
            if matches >= min_matches:
                return row[0].row - 1  # convert to 0-based for pandas header=
    finally:
        wb.close()
    return None  # couldn't find it; caller falls back to a default


def extract_sheet_metadata(sheet_path, sheet_name, header_row_idx, max_scan_rows=15):
    """
    Pull a shared date range and/or project code from the title rows sitting
    above the detected header row (these sheets often print one combined
    date range / project code for the whole group instead of per-student
    columns).
    """
    metadata = {'start_date': '', 'end_date': '', 'project_code': ''}
    if header_row_idx is None:
        scan_limit = max_scan_rows
    else:
        scan_limit = header_row_idx + 1  # only look above the header row

    wb = openpyxl.load_workbook(sheet_path, data_only=True, read_only=True)
    try:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=scan_limit):
            for cell in row:
                if cell.value is None:
                    continue
                text = str(cell.value)
                date_match = DATE_RANGE_RE.search(text)
                if date_match and not metadata['start_date']:
                    metadata['start_date'] = date_match.group(1)
                    metadata['end_date'] = date_match.group(2)
                code_match = PROJECT_CODE_RE.search(text)
                if code_match and not metadata['project_code']:
                    metadata['project_code'] = code_match.group(0)
    finally:
        wb.close()
    return metadata


def match_column(df_columns, possible):
    """
    Find the best-matching column name for a standard field.

    Two sheets in this workbook can have BOTH a generic column (e.g. 'Sector')
    and a more specific one (e.g. 'Sector Database') that both technically
    match. Picking "whichever comes first in the sheet" is unreliable -
    instead, score every column against the candidate list and pick the best
    score: an exact match always beats a substring match, and earlier entries
    in `possible` (more specific terms) always beat later ones.
    """
    best_col, best_score = None, None
    for col in df_columns:
        col_lower = col.lower().strip()
        col_score = None
        for priority, term in enumerate(possible):
            if col_lower == term:
                score = (0, priority)
            elif term in col_lower:
                score = (1, priority)
            else:
                continue
            if col_score is None or score < col_score:
                col_score = score
        if col_score is not None and (best_score is None or col_score < best_score):
            best_score = col_score
            best_col = col
    return best_col


def extract_student_data(df):
    """Extract student data from a dataframe with various column name variations."""
    students = []

    actual_cols = {}
    for standard, possible in COL_MAP.items():
        col = match_column(df.columns, possible)
        if col is not None:
            actual_cols[standard] = col

    for idx, row in df.iterrows():
        name_col = actual_cols.get('name', 'name')
        surname_col = actual_cols.get('surname', 'surname')

        if name_col not in df.columns and surname_col not in df.columns:
            continue

        first_name = str(row.get(name_col, '')).strip() if name_col in df.columns else ''
        surname = str(row.get(surname_col, '')).strip() if surname_col in df.columns else ''

        if (not first_name or first_name.lower() == 'nan') and (not surname or surname.lower() == 'nan'):
            continue

        student = {}

        if first_name and first_name.lower() != 'nan':
            if surname and surname.lower() != 'nan':
                student['name'] = f"{first_name} {surname}"
            else:
                student['name'] = first_name
        elif surname and surname.lower() != 'nan':
            student['name'] = surname
        else:
            continue

        for standard, col in actual_cols.items():
            if standard not in ['name', 'surname']:
                if col in df.columns:
                    val = row.get(col, '')
                    if pd.isna(val):
                        val = ''
                    elif isinstance(val, (pd.Timestamp, datetime)):
                        val = val.strftime('%d/%m/%Y')
                    student[standard] = str(val).strip()
                else:
                    student[standard] = ''

        for key in ['birth_date', 'start_date', 'end_date', 'sector', 'company', 'company_address',
                    'tutor', 'telephone', 'email', 'project_code', 'notes', 'dress_code', 'pickup',
                    'presentation', 'age', 'languages', 'allergies', 'experience', 'vat', 'pic',
                    'oid', 'website']:
            student.setdefault(key, '')

        students.append(student)

    return students


# Sheet names that are never cohorts of students - reused by both the
# /list-sheets route (so the dropdown doesn't show them) and
# /generate-certificates (so they're never scanned for students).
NON_COHORT_SHEET_NAMES = {
    'task database', 'company database', 'task database - olivetrain',
}


def is_cohort_sheet(sheet_name):
    """True if a sheet looks like a student/group sheet rather than an
    admin sheet (Task Database, Company Database, Template, etc.)."""
    name_lower = sheet_name.strip().lower()
    if name_lower in NON_COHORT_SHEET_NAMES:
        return False
    if name_lower.startswith('template'):
        return False
    return True


@app.route("/")
def home():
    return render_template_string(HTML)


@app.route("/list-sheets", methods=["POST"])
def list_sheets():
    """Return the cohort/group sheet names in an uploaded workbook so the
    frontend can offer them in a dropdown, instead of hardcoding one
    specific group name like 'O-LIVE TRAIN'."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    excel_file = request.files["file"]
    if excel_file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_path)

    try:
        xl = pd.ExcelFile(excel_path)
        sheets = [s for s in xl.sheet_names if is_cohort_sheet(s)]
    except Exception as e:
        return jsonify({"error": f"Could not read Excel file: {e}"}), 400
    finally:
        try:
            os.remove(excel_path)
        except Exception:
            pass

    return jsonify({"sheets": sheets})


@app.route("/generate-certificates", methods=["POST"])
def generate_certificates():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    excel_file = request.files["file"]
    if excel_file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    sheet_name_filter = request.form.get('sheet_name', '').strip()
    if not sheet_name_filter:
        return jsonify({"error": "Please choose a group/sheet to generate certificates for."}), 400

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_path)

    try:
        xl = pd.ExcelFile(excel_path)
        all_students = []

        for sheet_name in xl.sheet_names:
            if sheet_name.lower() != sheet_name_filter.lower():
                continue
            if not is_cohort_sheet(sheet_name):
                continue

            try:
                header_row_idx = find_header_row(excel_path, sheet_name)
                if header_row_idx is None:
                    # Couldn't auto-detect a header row on this sheet - skip it
                    # rather than silently reading garbage column names.
                    continue

                df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row_idx)
                df.columns = [str(col).strip() for col in df.columns]

                if len(df) == 0:
                    continue

                has_name = any('name' in str(col).lower() or 'surname' in str(col).lower() for col in df.columns)
                has_sector = any('sector' in str(col).lower() or 'study' in str(col).lower() for col in df.columns)

                if not (has_name or has_sector):
                    continue

                sheet_meta = extract_sheet_metadata(excel_path, sheet_name, header_row_idx)

                students = extract_student_data(df)
                for s in students:
                    s['sheet_name'] = sheet_name
                    # Fill in shared start/end date & project code from the
                    # sheet title row if the student row didn't have its own.
                    if not s.get('start_date') and sheet_meta['start_date']:
                        s['start_date'] = sheet_meta['start_date']
                    if not s.get('end_date') and sheet_meta['end_date']:
                        s['end_date'] = sheet_meta['end_date']
                    if not s.get('project_code') and sheet_meta['project_code']:
                        s['project_code'] = sheet_meta['project_code']
                all_students.extend(students)
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {e}")
                continue

        if not all_students:
            return jsonify({"error": f"No student data found in sheet '{sheet_name_filter}'. Make sure it has columns like 'Name' or 'Surname'."}), 400

        sector_map = {}
        try:
            task_sheet_names = ['Task Database', 'TASK DATABASE', 'Task database', 'tasks', 'Sectors']
            task_df = None
            for name in task_sheet_names:
                if name not in xl.sheet_names:
                    continue
                try:
                    task_header_idx = find_header_row(
                        excel_path, name,
                        markers=['sector', 'detailed', 'task', 'competence', 'knowledge'],
                        min_matches=2,
                    )
                    if task_header_idx is None:
                        task_header_idx = 0
                    task_df = pd.read_excel(excel_path, sheet_name=name, header=task_header_idx)
                    break
                except Exception:
                    continue

            if task_df is not None:
                task_df.columns = [str(col).strip() for col in task_df.columns]

                sector_col = None
                tasks_col = None
                competences_col = None

                for col in task_df.columns:
                    col_lower = col.lower()
                    if 'sector' in col_lower:
                        sector_col = col
                    elif 'detailed' in col_lower or 'task' in col_lower or 'Detailed tasks' in col:
                        tasks_col = col
                    elif 'job' in col_lower or 'competence' in col_lower or 'knowledge' in col_lower or 'Job-related' in col:
                        competences_col = col

                if sector_col and tasks_col:
                    for _, row in task_df.iterrows():
                        sector = str(row.get(sector_col, "")).strip()
                        if sector and sector.lower() != "nan":
                            sector_key = normalize_sector(sector)
                            sector_map[sector_key] = {
                                "detailed_tasks": str(row.get(tasks_col, "") or ""),
                                "job_related_competences": str(row.get(competences_col, "") or "") if competences_col else "",
                            }
        except Exception:
            pass

    except Exception as e:
        return jsonify({"error": f"Could not read Excel file: {e}"}), 400

    generated_files = []
    errors = []
    processed_count = 0

    for student in all_students:
        try:
            notes = student.get('notes', '').lower()
            experience = student.get('experience', '').lower()
            if 'will not come' in notes or "didn't show up" in notes or 'will not come' in experience or "didn't show up" in experience:
                continue

            ctx = {
                "NAME": student.get('name', ''),
                "BIRTH_DATE": student.get('birth_date', ''),
                "START": student.get('start_date', ''),
                "END": student.get('end_date', ''),
                "SECTOR": student.get('sector', ''),
                "HOST_ORG": student.get('company', ''),
                "HOST_ORG_ADDRESS": student.get('company_address', ''),
                "PROJECT_CODE": student.get('project_code', ''),
                "TUTOR": student.get('tutor', ''),
                "TELEPHONE": student.get('telephone', ''),
                "EMAIL": student.get('email', ''),
                "NOTES": student.get('notes', ''),
                "DRESS_CODE": student.get('dress_code', ''),
                "PICKUP": student.get('pickup', ''),
                "PRESENTATION": student.get('presentation', ''),
                "AGE": student.get('age', ''),
                "LANGUAGES": student.get('languages', ''),
                "ALLERGIES": student.get('allergies', ''),
                "EXPERIENCE": student.get('experience', ''),
                "VAT": student.get('vat', ''),
                "PIC": student.get('pic', ''),
                "OID": student.get('oid', ''),
                "WEBSITE": student.get('website', ''),
                "SHEET_NAME": student.get('sheet_name', ''),
                "detailed_tasks": "",
                "job_related_competences": "",
            }

            sector_key = normalize_sector(ctx["SECTOR"])
            if sector_key in sector_map:
                ctx["detailed_tasks"] = sector_map[sector_key].get("detailed_tasks", "")
                ctx["job_related_competences"] = sector_map[sector_key].get("job_related_competences", "")
            else:
                for key in sector_map:
                    if key in sector_key or sector_key in key:
                        ctx["detailed_tasks"] = sector_map[key].get("detailed_tasks", "")
                        ctx["job_related_competences"] = sector_map[key].get("job_related_competences", "")
                        break

            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(ctx)

            safe_name = re.sub(r'[^\w\s-]', '', student.get('name', 'student'))
            safe_name = re.sub(r'[-\s]+', '_', safe_name)
            out_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}.docx")
            doc.save(out_path)
            generated_files.append(out_path)
            processed_count += 1

        except Exception as e:
            errors.append(f"Error processing {student.get('name', 'Unknown')}: {str(e)}")

    if not generated_files:
        error_msg = "No certificates generated."
        if errors:
            error_msg += " Errors: " + "; ".join(errors[:3])
        error_msg += f" No students with valid data found in '{sheet_name_filter}'."
        return jsonify({"error": error_msg}), 400

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in generated_files:
            zf.write(f, os.path.basename(f))
    zip_buffer.seek(0)

    for f in generated_files:
        try:
            os.remove(f)
        except Exception:
            pass

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="certificates.zip",
        mimetype="application/zip"
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)