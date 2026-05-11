from datetime import date
from io import BytesIO
from pathlib import Path
import re

import pandas as pd
from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import load_workbook
from openpyxl import Workbook


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates_excel"

REPORT_TYPES = {
    "monthly": {
        "label": "Monthly",
        "template": TEMPLATE_DIR / "monthly_template.xlsx",
        "filename_prefix": "Monthly_Alert_Report",
    },
    "weekly": {
        "label": "Weekly",
        "template": TEMPLATE_DIR / "weekly_template.xlsx",
        "filename_prefix": "Weekly_Alert_Report",
    },
    "daily": {
        "label": "Daily",
        "template": TEMPLATE_DIR / "daily_template.xlsx",
        "filename_prefix": "Daily_Alert_Report",
    },
}

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
OUTPUT_FIELDS = [
    "Date",
    "Alert Id",
    "Type of Alert",
    "Alert Level",
    "Severity",
    "Origin",
    "Entity",
    "Entity Type",
    "Status",
    "Message",
    "SLA (Y/N)",
]
EXCEL_COLUMNS = {
    "Date": 1,
    "Alert Id": 2,
    "Type of Alert": 5,
    "Alert Level": 8,
    "Severity": 11,
    "Origin": 12,
    "Entity": 14,
    "Entity Type": 16,
    "Status": 17,
    "Message": 18,
    "SLA (Y/N)": 19,
}
REQUIRED_RAW_COLUMNS = [
    "Time",
    "Alert Id",
    "Type",
    "Severity",
    "Confidence",
    "Entity Type",
    "Status",
    "Message",
    "Assignee Email",
]


def today_string():
    return date.today().isoformat()


def clean_text(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime().replace(tzinfo=None)
    if hasattr(value, "isoformat") and not isinstance(value, str):
        value = value.isoformat(sep=" ") if hasattr(value, "hour") else value.isoformat()
    text = str(value)
    text = re.sub(r"[\ud800-\udfff]", "", text)
    text = "".join(ch for ch in text if ch == "\n" or ch == "\t" or ch == "\r" or ord(ch) >= 32)
    return text.strip()


def validate_report_type(report_type):
    key = clean_text(report_type).lower()
    if key not in REPORT_TYPES:
        raise ValueError("Select a valid report type: Monthly, Weekly, or Daily.")
    return key


def validate_upload(file_storage):
    if not file_storage or not file_storage.filename:
        raise ValueError("Please upload a raw alert Excel export.")
    suffix = Path(file_storage.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ValueError("Invalid file type. Upload a .xlsx or .xls raw alert export.")
    return suffix


def choose_column(columns, preferred, fallback=None):
    column_set = set(columns)
    if preferred in column_set:
        return preferred
    if fallback and fallback in column_set:
        return fallback
    return None


def parse_alert_file(file_storage):
    validate_upload(file_storage)
    data = BytesIO(file_storage.read())
    try:
        df = pd.read_excel(data)
    except Exception as exc:
        raise ValueError(f"Unable to read the Excel file. Check that it is a valid .xlsx or .xls export. Details: {exc}") from exc

    columns = [clean_text(col) for col in df.columns]
    df.columns = columns

    missing = [col for col in REQUIRED_RAW_COLUMNS if col not in df.columns]
    origin_col = choose_column(df.columns, "Origin.1", "Origin")
    entity_col = choose_column(df.columns, "Entity.1", "Entity")
    if origin_col is None:
        missing.append("Origin.1 or Origin")
    if entity_col is None:
        missing.append("Entity.1 or Entity")
    if missing:
        raise ValueError("Missing required raw alert columns: " + ", ".join(missing))

    rows = []
    for _, record in df.iterrows():
        assignee = clean_text(record.get("Assignee Email"))
        rows.append(
            {
                "Date": clean_text(record.get("Time")),
                "Alert Id": clean_text(record.get("Alert Id")),
                "Type of Alert": clean_text(record.get("Type")),
                "Alert Level": clean_text(record.get("Severity")),
                "Severity": clean_text(record.get("Confidence")),
                "Origin": clean_text(record.get(origin_col)),
                "Entity": clean_text(record.get(entity_col)),
                "Entity Type": clean_text(record.get("Entity Type")),
                "Status": clean_text(record.get("Status")),
                "Message": clean_text(record.get("Message")),
                "SLA (Y/N)": "Y" if "@" in assignee else "N",
            }
        )
    return rows


def compute_stats(rows):
    def norm(value):
        return clean_text(value).upper()

    return {
        "total": len(rows),
        "critical": sum(1 for row in rows if norm(row.get("Alert Level")) == "CRITICAL"),
        "major": sum(1 for row in rows if norm(row.get("Alert Level")) == "MAJOR"),
        "sla_y": sum(1 for row in rows if norm(row.get("SLA (Y/N)")) == "Y"),
    }


def sanitize_rows(rows):
    if not isinstance(rows, list):
        raise ValueError("Processed rows must be provided as a JSON array.")
    sanitized = []
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError("Each processed row must be an object.")
        sanitized.append({field: clean_text(row.get(field, "")) for field in OUTPUT_FIELDS})
    return sanitized


def preview_workbook(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Alert Preview"
    worksheet.append(OUTPUT_FIELDS)
    for row in rows:
        worksheet.append([row.get(field, "") for field in OUTPUT_FIELDS])

    for column_cells in worksheet.columns:
        max_length = max(len(clean_text(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 60)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def filled_report_workbook(report_type, rows):
    config = REPORT_TYPES[report_type]
    template_path = config["template"]
    if not template_path.exists():
        raise ValueError(f"Missing local template file: {template_path.relative_to(BASE_DIR)}")

    try:
        workbook = load_workbook(template_path)
    except Exception as exc:
        raise ValueError(f"Unable to load the local template. Details: {exc}") from exc

    if "Report" not in workbook.sheetnames:
        raise ValueError('The selected template must contain a worksheet named "Report".')

    worksheet = workbook["Report"]
    start_row = 15
    available_rows = max(worksheet.max_row - start_row + 1, 0)
    if len(rows) > available_rows:
        raise ValueError(
            f"The selected template has {available_rows} alert body rows available, "
            f"but {len(rows)} records were processed."
        )

    for row_idx in range(start_row, worksheet.max_row + 1):
        for col_idx in EXCEL_COLUMNS.values():
            worksheet.cell(row=row_idx, column=col_idx).value = None

    for offset, row in enumerate(rows):
        excel_row = start_row + offset
        for field, col_idx in EXCEL_COLUMNS.items():
            worksheet.cell(row=excel_row, column=col_idx).value = row.get(field, "")

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def excel_response(buffer, filename):
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/")
def index():
    template_status = {
        key: {
            "label": value["label"],
            "path": str(value["template"].relative_to(BASE_DIR)),
            "exists": value["template"].exists(),
        }
        for key, value in REPORT_TYPES.items()
    }
    return render_template_string(INDEX_HTML, template_status=template_status)


@app.post("/process")
def process():
    try:
        report_type = validate_report_type(request.form.get("report_type", ""))
        rows = parse_alert_file(request.files.get("raw_file"))
        return jsonify(
            {
                "ok": True,
                "report_type": report_type,
                "report_label": REPORT_TYPES[report_type]["label"],
                "rows": rows,
                "count": len(rows),
                "stats": compute_stats(rows),
                "template": {
                    "path": str(REPORT_TYPES[report_type]["template"].relative_to(BASE_DIR)),
                    "exists": REPORT_TYPES[report_type]["template"].exists(),
                },
            }
        )
    except ValueError as exc:
        return jsonify({"ok": False, "error": clean_text(exc)}), 400


@app.post("/download_preview")
def download_preview():
    try:
        payload = request.get_json(silent=True) or {}
        rows = sanitize_rows(payload.get("rows", []))
        return excel_response(preview_workbook(rows), f"Alert_Preview_{today_string()}.xlsx")
    except ValueError as exc:
        return jsonify({"ok": False, "error": clean_text(exc)}), 400


@app.post("/download")
def download():
    try:
        payload = request.get_json(silent=True) or {}
        report_type = validate_report_type(payload.get("report_type", ""))
        rows = sanitize_rows(payload.get("rows", []))
        filename = f"{REPORT_TYPES[report_type]['filename_prefix']}_{today_string()}.xlsx"
        return excel_response(filled_report_workbook(report_type, rows), filename)
    except ValueError as exc:
        return jsonify({"ok": False, "error": clean_text(exc)}), 400


@app.errorhandler(413)
def file_too_large(_error):
    return jsonify({"ok": False, "error": "The uploaded file is too large. Maximum size is 50 MB."}), 413


INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Monthly, Weekly, Daily Alert Report Generator</title>
  <style>
    :root {
      color-scheme: dark;
      --bg: #0f1720;
      --panel: #151f2b;
      --panel-2: #1b2634;
      --line: #2c3b4f;
      --text: #e7edf5;
      --muted: #9aacbd;
      --accent: #28a7c7;
      --accent-2: #8bd450;
      --danger: #ff5d66;
      --warn: #ffb24a;
      --minor: #57d68d;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      min-height: 100vh;
    }
    .app {
      display: grid;
      grid-template-columns: minmax(300px, 360px) minmax(0, 1fr);
      min-height: 100vh;
    }
    aside {
      border-right: 1px solid var(--line);
      background: #121c27;
      padding: 24px;
      overflow: auto;
    }
    main {
      padding: 24px;
      overflow: auto;
    }
    h1 {
      font-size: 22px;
      line-height: 1.2;
      margin: 0 0 8px;
      letter-spacing: 0;
    }
    h2 {
      font-size: 15px;
      margin: 22px 0 10px;
      color: #cdd9e6;
      letter-spacing: 0;
    }
    p, li, label, small { color: var(--muted); }
    .subtitle { margin: 0 0 18px; }
    .group {
      border-top: 1px solid var(--line);
      padding-top: 18px;
      margin-top: 18px;
    }
    .segmented {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 8px;
    }
    .segmented input { position: absolute; opacity: 0; pointer-events: none; }
    .segmented span {
      display: block;
      text-align: center;
      border: 1px solid var(--line);
      background: var(--panel);
      padding: 10px 8px;
      border-radius: 6px;
      color: var(--text);
      cursor: pointer;
      font-weight: 700;
    }
    .segmented input:checked + span {
      background: #103244;
      border-color: var(--accent);
      color: #d9f8ff;
    }
    .dropzone {
      border: 1px dashed #52708d;
      background: var(--panel);
      border-radius: 8px;
      padding: 18px;
      text-align: center;
      transition: border-color .15s ease, background .15s ease;
    }
    .dropzone.dragover {
      border-color: var(--accent);
      background: #132b38;
    }
    input[type="file"] { display: none; }
    .file-button, button {
      appearance: none;
      border: 1px solid transparent;
      border-radius: 6px;
      padding: 10px 12px;
      min-height: 42px;
      font-weight: 800;
      color: #061019;
      background: var(--accent);
      cursor: pointer;
      width: 100%;
      margin-top: 10px;
    }
    button.secondary {
      background: #263649;
      color: var(--text);
      border-color: #3b4c62;
    }
    button.success {
      background: var(--accent-2);
    }
    button:disabled {
      cursor: not-allowed;
      opacity: .45;
    }
    .mapping {
      display: grid;
      gap: 6px;
      margin: 0;
      padding: 0;
      list-style: none;
      font-size: 13px;
    }
    .mapping li {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      border-bottom: 1px solid rgba(255,255,255,.06);
      padding-bottom: 6px;
    }
    .mapping strong { color: var(--text); text-align: right; }
    .status {
      margin-top: 12px;
      padding: 10px 12px;
      border-radius: 6px;
      background: var(--panel);
      color: var(--muted);
      border: 1px solid var(--line);
      min-height: 42px;
      line-height: 1.35;
    }
    .status.ok { color: #c8f7ce; border-color: #326b44; }
    .status.error { color: #ffd4d7; border-color: #8a3037; }
    .template-pill {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 12px;
      font-size: 13px;
    }
    .dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      flex: 0 0 auto;
      background: var(--danger);
    }
    .dot.ok { background: var(--accent-2); }
    .stats {
      display: grid;
      grid-template-columns: repeat(4, minmax(120px, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .stat {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
    }
    .stat b { display: block; font-size: 28px; color: var(--text); }
    .stat span { color: var(--muted); font-size: 13px; }
    .table-wrap {
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: auto;
      max-height: calc(100vh - 190px);
      background: var(--panel);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 1200px;
    }
    th, td {
      padding: 10px 12px;
      border-bottom: 1px solid rgba(255,255,255,.07);
      text-align: left;
      vertical-align: top;
      font-size: 13px;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #203044;
      color: #f3f7fb;
    }
    tr:hover td { background: rgba(255,255,255,.03); }
    .badge {
      display: inline-block;
      border-radius: 999px;
      padding: 3px 8px;
      font-weight: 800;
      font-size: 12px;
    }
    .critical { background: rgba(255,93,102,.16); color: #ff9aa0; }
    .major { background: rgba(255,178,74,.16); color: #ffc979; }
    .minor { background: rgba(87,214,141,.16); color: #8ff0b5; }
    .sla-y { color: #8ff0b5; font-weight: 900; }
    .sla-n { color: #ff9aa0; font-weight: 900; }
    .empty {
      color: var(--muted);
      padding: 34px;
      text-align: center;
    }
    @media (max-width: 980px) {
      .app { grid-template-columns: 1fr; }
      aside { border-right: 0; border-bottom: 1px solid var(--line); }
      .stats { grid-template-columns: repeat(2, 1fr); }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside>
      <h1>Monthly, Weekly, Daily Alert Report Generator</h1>
      <p class="subtitle">Local Excel conversion using project templates only.</p>

      <form id="processForm">
        <div class="group">
          <h2>Report Type</h2>
          <div class="segmented" role="radiogroup" aria-label="Report type">
            <label><input type="radio" name="report_type" value="monthly" checked><span>Monthly</span></label>
            <label><input type="radio" name="report_type" value="weekly"><span>Weekly</span></label>
            <label><input type="radio" name="report_type" value="daily"><span>Daily</span></label>
          </div>
        </div>

        <div class="group">
          <h2>Raw Alert Upload</h2>
          <div id="dropzone" class="dropzone">
            <p id="fileName">Drop a .xlsx or .xls raw alert export here</p>
            <label class="file-button" for="rawFile">Choose Excel File</label>
            <input id="rawFile" name="raw_file" type="file" accept=".xlsx,.xls">
          </div>
        </div>

        <div class="group">
          <h2>Local Template</h2>
          <div class="template-pill">
            <span id="templatePath"></span>
            <span id="templateDot" class="dot" aria-hidden="true"></span>
          </div>
        </div>

        <div class="group">
          <h2>Save Location</h2>
          <div class="template-pill">
            <span id="saveCapability"></span>
          </div>
        </div>

        <div class="group">
          <h2>Field Mapping</h2>
          <ul class="mapping">
            <li><span>Time</span><strong>Date</strong></li>
            <li><span>Alert Id</span><strong>Alert Id</strong></li>
            <li><span>Type</span><strong>Type of Alert</strong></li>
            <li><span>Severity</span><strong>Alert Level</strong></li>
            <li><span>Confidence</span><strong>Severity</strong></li>
            <li><span>Origin.1 / Origin</span><strong>Origin</strong></li>
            <li><span>Entity.1 / Entity</span><strong>Entity</strong></li>
            <li><span>Assignee Email</span><strong>SLA (Y/N)</strong></li>
          </ul>
        </div>

        <div class="group">
          <button id="processBtn" type="submit">Process & Preview</button>
          <button id="saveBtn" class="success" type="button" disabled>Save Report to Folder</button>
          <button id="previewBtn" class="secondary" type="button" disabled>Download Excel Preview</button>
          <button id="reportBtn" class="secondary" type="button" disabled>Download Generated Report</button>
        </div>
        <div id="status" class="status">Ready.</div>
      </form>
    </aside>

    <main>
      <div class="stats">
        <div class="stat"><b id="totalStat">0</b><span>Total records</span></div>
        <div class="stat"><b id="criticalStat">0</b><span>Critical count</span></div>
        <div class="stat"><b id="majorStat">0</b><span>Major count</span></div>
        <div class="stat"><b id="slaStat">0</b><span>SLA Y count</span></div>
      </div>
      <div class="table-wrap">
        <table id="previewTable">
          <thead></thead>
          <tbody><tr><td class="empty">Process a raw alert export to preview records.</td></tr></tbody>
        </table>
      </div>
    </main>
  </div>

  <script>
    const templateStatus = {{ template_status | tojson }};
    const fields = ["Date", "Alert Id", "Type of Alert", "Alert Level", "Severity", "Origin", "Entity", "Entity Type", "Status", "Message", "SLA (Y/N)"];
    let processedRows = [];
    let currentReportType = "monthly";

    const form = document.getElementById("processForm");
    const rawFile = document.getElementById("rawFile");
    const dropzone = document.getElementById("dropzone");
    const fileName = document.getElementById("fileName");
    const statusBox = document.getElementById("status");
    const templatePath = document.getElementById("templatePath");
    const templateDot = document.getElementById("templateDot");
    const saveCapability = document.getElementById("saveCapability");
    const processBtn = document.getElementById("processBtn");
    const saveBtn = document.getElementById("saveBtn");
    const previewBtn = document.getElementById("previewBtn");
    const reportBtn = document.getElementById("reportBtn");

    function selectedReportType() {
      return new FormData(form).get("report_type");
    }

    function setStatus(message, type = "") {
      statusBox.textContent = message;
      statusBox.className = "status" + (type ? " " + type : "");
    }

    function updateTemplateStatus() {
      currentReportType = selectedReportType();
      const info = templateStatus[currentReportType];
      templatePath.textContent = info.path + (info.exists ? " found" : " missing");
      templateDot.className = "dot" + (info.exists ? " ok" : "");
    }

    function updateButtons(enabled) {
      saveBtn.disabled = !enabled;
      previewBtn.disabled = !enabled;
      reportBtn.disabled = !enabled;
    }

    function validExcelFile(file) {
      return file && /\\.(xlsx|xls)$/i.test(file.name);
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, char => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
      }[char]));
    }

    function badgeForLevel(value) {
      const normalized = String(value || "").toUpperCase();
      if (normalized === "CRITICAL") return `<span class="badge critical">${escapeHtml(value)}</span>`;
      if (normalized === "MAJOR") return `<span class="badge major">${escapeHtml(value)}</span>`;
      if (normalized === "MINOR") return `<span class="badge minor">${escapeHtml(value)}</span>`;
      return escapeHtml(value);
    }

    function slaCell(value) {
      const normalized = String(value || "").toUpperCase();
      const cls = normalized === "Y" ? "sla-y" : "sla-n";
      return `<span class="${cls}">${escapeHtml(value)}</span>`;
    }

    function renderPreview(rows) {
      const table = document.getElementById("previewTable");
      table.querySelector("thead").innerHTML = `<tr>${fields.map(field => `<th>${escapeHtml(field)}</th>`).join("")}</tr>`;
      if (!rows.length) {
        table.querySelector("tbody").innerHTML = `<tr><td class="empty" colspan="${fields.length}">No records found.</td></tr>`;
        return;
      }
      table.querySelector("tbody").innerHTML = rows.map(row => {
        return `<tr>${fields.map(field => {
          if (field === "Alert Level") return `<td>${badgeForLevel(row[field])}</td>`;
          if (field === "SLA (Y/N)") return `<td>${slaCell(row[field])}</td>`;
          return `<td>${escapeHtml(row[field])}</td>`;
        }).join("")}</tr>`;
      }).join("");
    }

    function updateStats(stats) {
      document.getElementById("totalStat").textContent = stats.total || 0;
      document.getElementById("criticalStat").textContent = stats.critical || 0;
      document.getElementById("majorStat").textContent = stats.major || 0;
      document.getElementById("slaStat").textContent = stats.sla_y || 0;
    }

    async function downloadBlob(url, body, preferredName, saveToFolder = false) {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
      if (!response.ok) {
        const errorPayload = await response.json().catch(() => ({}));
        throw new Error(errorPayload.error || "The download request failed.");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename="?([^"]+)"?/i);
      const filename = match ? match[1] : preferredName;

      if (saveToFolder && window.showSaveFilePicker) {
        const handle = await window.showSaveFilePicker({
          suggestedName: filename,
          types: [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return filename;
      }

      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(link.href);
      return filename;
    }

    form.addEventListener("change", event => {
      if (event.target.name === "report_type") {
        updateTemplateStatus();
        currentReportType = selectedReportType();
      }
    });

    rawFile.addEventListener("change", () => {
      const file = rawFile.files[0];
      fileName.textContent = file ? file.name : "Drop a .xlsx or .xls raw alert export here";
      if (file && !validExcelFile(file)) setStatus("Invalid file type. Upload .xlsx or .xls.", "error");
    });

    ["dragenter", "dragover"].forEach(eventName => {
      dropzone.addEventListener(eventName, event => {
        event.preventDefault();
        dropzone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach(eventName => {
      dropzone.addEventListener(eventName, event => {
        event.preventDefault();
        dropzone.classList.remove("dragover");
      });
    });
    dropzone.addEventListener("drop", event => {
      const file = event.dataTransfer.files[0];
      if (!validExcelFile(file)) {
        setStatus("Invalid file type. Upload .xlsx or .xls.", "error");
        return;
      }
      rawFile.files = event.dataTransfer.files;
      fileName.textContent = file.name;
    });

    form.addEventListener("submit", async event => {
      event.preventDefault();
      const file = rawFile.files[0];
      if (!validExcelFile(file)) {
        setStatus("Choose a valid .xlsx or .xls raw alert export.", "error");
        return;
      }
      processBtn.disabled = true;
      updateButtons(false);
      setStatus("Processing raw alert export...");
      const formData = new FormData();
      formData.append("report_type", selectedReportType());
      formData.append("raw_file", file);
      try {
        const response = await fetch("/process", { method: "POST", body: formData });
        const payload = await response.json();
        if (!response.ok || !payload.ok) throw new Error(payload.error || "Processing failed.");
        processedRows = payload.rows;
        currentReportType = payload.report_type;
        updateStats(payload.stats);
        renderPreview(processedRows);
        updateButtons(true);
        setStatus(`${payload.report_label} preview ready with ${payload.count} records.`, "ok");
      } catch (error) {
        processedRows = [];
        updateStats({});
        renderPreview([]);
        setStatus(error.message, "error");
      } finally {
        processBtn.disabled = false;
      }
    });

    previewBtn.addEventListener("click", async () => {
      try {
        setStatus("Preparing preview workbook...");
        const filename = await downloadBlob("/download_preview", { rows: processedRows }, "Alert_Preview.xlsx");
        setStatus(`Preview ready: ${filename}`, "ok");
      } catch (error) {
        setStatus(error.message, "error");
      }
    });

    reportBtn.addEventListener("click", async () => {
      try {
        setStatus("Generating report from the selected local template...");
        const filename = await downloadBlob("/download", { report_type: currentReportType, rows: processedRows }, "Alert_Report.xlsx");
        setStatus(`Report ready: ${filename}`, "ok");
      } catch (error) {
        setStatus(error.message, "error");
      }
    });

    saveBtn.addEventListener("click", async () => {
      try {
        setStatus("Preparing report save...");
        const filename = await downloadBlob(
          "/download",
          { report_type: currentReportType, rows: processedRows },
          "Alert_Report.xlsx",
          Boolean(window.showSaveFilePicker)
        );
        setStatus(window.showSaveFilePicker ? `Saved report: ${filename}` : `Downloaded report: ${filename}`, "ok");
      } catch (error) {
        if (error.name === "AbortError") {
          setStatus("Save cancelled.");
        } else {
          setStatus(error.message, "error");
        }
      }
    });

    saveCapability.textContent = window.showSaveFilePicker
      ? "Folder save supported by this browser"
      : "Browser download fallback will be used";
    updateTemplateStatus();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
