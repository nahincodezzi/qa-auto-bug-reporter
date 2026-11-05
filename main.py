import os
import sys
import re
import io
import mimetypes
from datetime import datetime
from pathlib import Path
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageGrab

# ==============================================================================
# 📋 15 EXACT EXCEL COLUMNS (AS PER QA TEMPLATE)
# ==============================================================================
COLUMNS = [
    "TID",
    "Module",
    "Test Case Title",
    "Test Scenario",
    "Test Data",
    "Test Step",
    "Expected Result",
    "Actual Result",
    "Status",
    "Screenshort",
    "Dev Comment",
    "QA Comment",
    "Status",
    "Dev",
    "Dev Comennt"
]

EXCEL_FILE = "bug_report.xlsx"

# ==============================================================================
# ☁️ CLOUD SCREENSHOT UPLOADER (LIGHTSHOT STYLE)
# ==============================================================================
def upload_image_bytes(image_bytes: bytes, filename: str = "screenshot.png") -> str:
    """Uploads raw image bytes to cloud and returns permanent direct public URL."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    mime_type = mimetypes.guess_type(filename)[0] or "image/png"

    # Primary Host: Freeimage.host
    try:
        data = {
            "key": "6d207e02198a847aa98d0a2a901485a5",
            "action": "upload",
            "format": "json"
        }
        res = requests.post(
            "https://freeimage.host/api/1/upload",
            data=data,
            files={"source": (filename, image_bytes, mime_type)},
            headers=headers,
            timeout=25
        )
        if res.status_code == 200:
            url = res.json().get("image", {}).get("url")
            if url:
                return url
    except Exception:
        pass

    # Secondary Host: Catbox.moe fallback
    try:
        res = requests.post(
            "https://catbox.moe/user/api.php",
            data={"reqtype": "fileupload"},
            files={"fileToUpload": (filename, image_bytes, mime_type)},
            headers=headers,
            timeout=25
        )
        if res.status_code == 200 and res.text.strip().startswith("http"):
            return res.text.strip()
    except Exception as e:
        print(f"❌ Upload Error: {e}")

    return "Upload failed"

def get_screenshot_cloud_url(ss_input: str) -> str:
    """
    Handles screenshot from:
    1. File path (drag & drop)
    2. System clipboard ('clip' or empty if clipboard has image)
    3. Direct URL (if already a link)
    4. Multiple links separated by comma or newline
    """
    if not ss_input or not str(ss_input).strip():
        return "-"

    text = ss_input.strip("\"' ")

    # Case A: Already a cloud link
    if text.startswith("http://") or text.startswith("https://"):
        return text

    # Case B: Clipboard grab
    if text.lower() == "clip":
        print("📋 Checking system clipboard for screenshot...")
        clip_data = ImageGrab.grabclipboard()
        if isinstance(clip_data, Image.Image):
            buf = io.BytesIO()
            clip_data.save(buf, format="PNG")
            print("☁️ Uploading clipboard image to cloud...")
            url = upload_image_bytes(buf.getvalue(), "clipboard_screenshot.png")
            print(f"✅ Cloud link created: {url}")
            return url
        else:
            print("⚠️ No image found in clipboard.")
            return "-"

    # Case C: Multiple files separated by comma or semicolon
    if "," in text or ";" in text:
        parts = [p.strip("\"' ") for p in re.split(r"[,;]", text) if p.strip()]
        urls = []
        for i, p in enumerate(parts, 1):
            url = get_screenshot_cloud_url(p)
            if url and url != "-":
                urls.append(f"SS {i}: {url}")
        return "\n".join(urls) if urls else "-"

    # Case D: Local file path
    path = Path(text)
    if not path.exists():
        print(f"⚠️ Image file not found: {text}")
        return "-"

    print(f"☁️ Uploading '{path.name}' to cloud...")
    with open(path, "rb") as f:
        file_bytes = f.read()

    url = upload_image_bytes(file_bytes, path.name)
    print(f"✅ Cloud link created: {url}")
    return url

# ==============================================================================
# 🧠 SMART QA CONVERTER (RAW TEXT / BANGLA -> PROFESSIONAL QA ENGLISH)
# ==============================================================================
def parse_markdown_table_or_structured(text: str) -> dict:
    """Detects if user pasted an existing table row with pipes '|' and extracts it."""
    if "|" in text:
        parts = [p.strip() for p in text.split("|") if p.strip()]
        # Check if this is a header row or data row
        if len(parts) >= 8 and not any(h in parts[0].upper() for h in ["TID", "MODULE", "STATUS"]):
            return {
                "tid": parts[0] if len(parts) > 0 else "",
                "module": parts[1] if len(parts) > 1 else "",
                "title": parts[2] if len(parts) > 2 else "",
                "scenario": parts[3] if len(parts) > 3 else "",
                "test_data": parts[4] if len(parts) > 4 else "N/A",
                "steps": parts[5] if len(parts) > 5 else "",
                "expected": parts[6] if len(parts) > 6 else "",
                "actual": parts[7] if len(parts) > 7 else "",
                "status": parts[8] if len(parts) > 8 else "Fail",
                "screenshort": parts[9] if len(parts) > 9 else "",
                "dev_comment": parts[10] if len(parts) > 10 else "",
                "qa_comment": parts[11] if len(parts) > 11 else "",
                "defect_status": "New",
                "dev": "Unassigned",
                "dev_comennt": ""
            }
    return None

def convert_problem_to_qa_report(problem_text: str, image_bytes: bytes = None, mime_type: str = "image/png") -> dict:
    """
    Converts any raw problem description (Bangla / Banglish / English) and optional screenshot
    into standard, 100% professional simple English QA Test Case and Bug Report fields.
    Uses Gemini Vision AI if API key is present, with built-in smart QA engine fallback.
    NEVER leaks raw Banglish or Bengali words into the output fields.
    """
    # 1. Check if user passed already structured table format
    structured = parse_markdown_table_or_structured(problem_text)
    if structured:
        return structured

    # 2. Try Gemini AI if API key is set in environment or .env
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key and Path(".env").exists():
        from dotenv import load_dotenv
        load_dotenv()
        api_key = os.getenv("GEMINI_API_KEY", "").strip()

    if api_key:
        try:
            from google import genai
            from google.genai import types
            import json
            client = genai.Client(api_key=api_key)

            contents = []
            if image_bytes:
                contents.append(types.Part.from_bytes(data=image_bytes, mime_type=mime_type or "image/png"))

            prompt = f"""You are a Senior QA Test Engineer. Analyze this bug report (written in Banglish, Bengali, or informal English) and the screenshot (if provided).
Generate a complete, 100% SIMPLE PROFESSIONAL CORPORATE ENGLISH QA bug report.
Do NOT output any Bengali or Banglish words. Every single word in the output must be clear, simple English.

Input description: "{problem_text}"

Respond with ONLY valid JSON:
{{
  "module": "Application Module (e.g. Applicant Portal > Dashboard, Super Admin > Life Cycle)",
  "title": "Descriptive title starting with 'Verify...'",
  "scenario": "Clear QA test scenario statement",
  "test_data": "Applicable test data or parameters (e.g. Applicant account with 1 submitted application)",
  "steps": "1. First step\\n2. Second step\\n3. Third step",
  "expected": "Exact expected behavior in clear English",
  "actual": "Exact defect / failure in clean professional English",
  "qa_comment": "Technical QA observation or root cause clue in English"
}}"""
            contents.append(prompt)

            res = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2)
            )
            data = json.loads(res.text)
            if data and data.get("title") and data.get("actual"):
                return {
                    "module": data.get("module", "Application Functionality"),
                    "title": data.get("title", ""),
                    "scenario": data.get("scenario", ""),
                    "test_data": data.get("test_data", "N/A"),
                    "steps": data.get("steps", ""),
                    "expected": data.get("expected", ""),
                    "actual": data.get("actual", ""),
                    "status": "Fail",
                    "screenshort": "-",
                    "dev_comment": "",
                    "qa_comment": data.get("qa_comment", "Logged by QA"),
                    "defect_status": "New",
                    "dev": "Unassigned",
                    "dev_comennt": ""
                }
        except Exception as e:
            print(f"[Notice] Gemini AI skipped ({e}), using smart QA engine...")

    # 3. Smart Rule-Based Engine (Produces 100% pure professional English)
    text_lower = problem_text.lower()

    # Smart Module Detection
    if "dashboard" in text_lower or "dash" in text_lower:
        if any(w in text_lower for w in ["applicant", "client", "user"]):
            module = "Applicant Portal > Dashboard"
        elif "admin" in text_lower:
            module = "Super Admin > Dashboard"
        else:
            module = "Dashboard & Analytics"
    elif any(w in text_lower for w in ["grant application", "form builder", "application from"]):
        module = "Super Admin > Grant Application Form"
    elif any(w in text_lower for w in ["grant event", "event"]):
        module = "Super Admin > Grant Event"
    elif any(w in text_lower for w in ["life cycle", "stage", "sequence"]):
        module = "Super Admin > Life Cycle"
    elif "my application" in text_lower:
        module = "Applicant Portal > My Applications"
    elif any(w in text_lower for w in ["login", "auth", "password"]):
        module = "Authentication > Login"
    elif any(w in text_lower for w in ["user", "profile"]):
        module = "User Management"
    else:
        module = "Application Functionality"

    # Pattern A: Prolonged inactivity / navigation freeze / refresh required
    if any(w in text_lower for w in ["onk somoy", "onek somoy", "dore", "thaka hoy"]) and any(w in text_lower for w in ["user page", "users", "page ta ase na", "load hoy na", "refresh"]):
        module = "Super Admin > User Management / Navigation"
        title = "Verify navigation to Users page after prolonged activity in Grant Application Form"
        scenario = "Verify client-side route navigation to the Users page without requiring a browser refresh."
        test_data = "Super Admin active session"
        steps = "1. Log in as Super Admin.\n2. Navigate to 'Grant Application Form' and remain on the page for an extended duration.\n3. Click on the 'Users' menu from the left navigation sidebar.\n4. Observe whether the Users page loads properly."
        expected = "The Users page should load seamlessly and display the user list without requiring a manual page refresh."
        actual = "The Users page fails to load after remaining on the Grant Application Form for an extended period, while other pages function normally. A manual browser refresh is required to load the page."
        qa_comment = "Client-side route navigation freeze or component state lock in Grant Application Form preventing transition to Users page."

    # Pattern B: Dashboard count zero / missing
    elif ("dash" in text_lower or "dashboard" in text_lower) and any(w in text_lower for w in ["zero", "0", "missing", "not showing", "dekha jacce na", "deka jacce na", "ashche na"]):
        module = "Applicant Portal > Dashboard"
        title = "Verify that Applicant Dashboard reflects total submitted application count"
        scenario = "Verify real-time data synchronization between submitted applications and Dashboard metrics counter."
        test_data = "Applicant account with active submitted application(s)"
        steps = "1. Log in as an Applicant.\n2. Navigate to 'My Applications' and verify that an active application exists.\n3. Return to the 'Dashboard'.\n4. Observe the 'Total Applications' metric counter widget."
        expected = "The Dashboard counter should accurately display the count of submitted applications (e.g., displaying 1 instead of 0)."
        actual = "The Dashboard counter displays 0 applications despite an active submitted application existing under the applicant account."
        qa_comment = "Data aggregation failure. The dashboard metrics endpoint returns 0 or fails to query the applicant's submission records."

    # Pattern C: Form Builder / Dynamic Field synchronization defect
    elif any(w in text_lower for w in ["form", "builder", "field", "schema", "tab", "project information"]):
        module = "Super Admin > Grant Application Form"
        title = f"Verify dynamic form fields created in {module} render on client portal"
        scenario = f"Verify end-to-end synchronization between Form Builder schema configuration and user-facing submission forms."
        test_data = "Configured form schema and custom fields under Project Information"
        steps = "1. Log in as Super Admin.\n2. Open Form Builder and configure custom fields under the 'Project Information' tab.\n3. Save and update form schema.\n4. Log in to the applicant portal and open the application submission form.\n5. Verify if custom fields render properly."
        expected = "All custom fields and tabs configured in Form Builder should dynamically render on the applicant submission form."
        actual = "Configured fields from the Form Builder do not render on the applicant portal. The form continues to display static default fields."
        qa_comment = "Dynamic form schema is not synchronized or fetched by the client-side application create endpoint."

    # Pattern D: Delete operation failure / 404 / 500 error
    elif "delete" in text_lower:
        title = f"Verify Delete functionality in {module}"
        scenario = f"Verify that an authorized user can successfully delete an existing record in {module} without server errors."
        test_data = "Target record selected for deletion"
        steps = f"1. Navigate to {module}.\n2. Locate an existing record.\n3. Click on the Delete (trash) icon.\n4. Confirm deletion."
        expected = "The record should be deleted successfully, removed from the list, and a success confirmation notification should appear."
        actual = "Clicking the Delete button fails and triggers an unexpected server error or HTTP 404 Not Found response."
        qa_comment = "Backend API error during DELETE request. Verify endpoint route registration and parameter mapping."

    # Pattern E: Validation error / Duplicate values
    elif any(w in text_lower for w in ["duplicate", "sequence", "unique", "validation"]):
        title = f"Verify uniqueness and field validation in {module}"
        scenario = f"Verify that the system enforces uniqueness and prevents duplicate records in {module}."
        test_data = "Duplicate input values (e.g. existing Sequence number or Name)"
        steps = f"1. Navigate to {module}.\n2. Edit an existing record or create a new entry.\n3. Enter a value already in use by another record.\n4. Save changes."
        expected = "The system should display a validation error message and prevent saving duplicate values."
        actual = "The system allows duplicate values to be saved without triggering unique constraint validation."
        qa_comment = "Missing unique validation rule on both frontend form validation and backend database constraints."

    # Pattern F: Button unresponsive / Nothing happens
    elif any(w in text_lower for w in ["kichu hoy na", "kaj kore na", "response nai", "unresponsive"]):
        title = f"Verify user action triggers expected response in {module}"
        scenario = f"Verify that user interaction on designated controls in {module} successfully triggers the corresponding operation."
        test_data = "Standard interaction event"
        steps = f"1. Navigate to {module}.\n2. Trigger the primary action or submit button.\n3. Observe system behavior."
        expected = "The system should execute the action promptly and display proper feedback or status updates."
        actual = "The action is completely unresponsive; clicking the button fails to trigger any network request or UI confirmation."
        qa_comment = "Unbound event listener, unhandled promise rejection, or disabled button state in frontend code."

    # Pattern G: General clean fallback (STRICTLY NO BANGLISH!)
    else:
        title = f"Verify functional workflow in {module}"
        scenario = f"Verify functional stability and user interactions in {module}"
        test_data = "Standard application user data"
        steps = f"1. Log in to the application.\n2. Navigate to {module}.\n3. Perform the designated user action.\n4. Observe system response."
        expected = "The system should perform the requested action smoothly according to business requirements."
        actual = f"The expected functional operation in {module} failed during execution. An unexpected defect occurred."
        qa_comment = "Functional defect observed during manual exploratory QA testing. Developer investigation required."

    return {
        "module": module,
        "title": title,
        "scenario": scenario,
        "test_data": test_data,
        "steps": steps,
        "expected": expected,
        "actual": actual,
        "status": "Fail",
        "screenshort": "-",
        "dev_comment": "",
        "qa_comment": qa_comment,
        "defect_status": "New",
        "dev": "Unassigned",
        "dev_comennt": ""
    }

# ==============================================================================
# 📊 EXCEL MANAGER (ONE SINGLE bug_report.xlsx FILE, FOREVER)
# ==============================================================================
def get_next_tid(ws=None, module: str = "") -> str:
    """
    Auto-generates TID in the format: TID<month><day><sequence:02d>
    Example on November 5th:
        TID11501, TID11502, TID11503...
    Where:
        - 11 = Month (November)
        - 5  = Day / Date (5th)
        - 01, 02... = 2-digit sequential defect counter for that day
    """
    now = datetime.now()
    prefix = f"TID{now.month}{now.day}"

    if ws is None:
        return f"{prefix}01"

    existing_seqs = []
    for r in range(2, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "").strip()
        if val.startswith(prefix):
            suffix = val[len(prefix):]
            if suffix.isdigit():
                existing_seqs.append(int(suffix))
            else:
                m = re.search(r"(\d+)$", val)
                if m:
                    existing_seqs.append(int(m.group(1)))

    next_seq = (max(existing_seqs) + 1) if existing_seqs else 1
    return f"{prefix}{next_seq:02d}"

def compact_worksheet(ws) -> None:
    """
    Removes trailing blank rows, fixes desynchronized cell hyperlinks,
    and removes any empty ghost rows in the middle so that data rows are always consecutive and contiguous.
    """
    # 1. Update all cell hyperlink coordinates so openpyxl doesn't resurrect phantom rows
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink:
                cell.hyperlink.ref = cell.coordinate

    # 2. Remove trailing empty rows from the bottom
    for r in range(ws.max_row, 1, -1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        has_data = any(v is not None and str(v).strip() not in ("", "-") for v in vals)
        if not has_data:
            ws.delete_rows(r)
        else:
            break

    # 3. Remove any accidental empty rows in the middle
    r = 2
    while r <= ws.max_row:
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        has_data = any(v is not None and str(v).strip() not in ("", "-") for v in vals)
        if not has_data:
            ws.delete_rows(r)
        else:
            r += 1

    # 4. Re-align hyperlink references again after row deletion
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink:
                cell.hyperlink.ref = cell.coordinate

def delete_columns_from_excel(col_indices: list, file_path: str = EXCEL_FILE) -> list:
    """
    Deletes specified columns from the Excel file (1-based index).
    Deletes in descending order to avoid index shifting.
    """
    p = Path(file_path)
    if not p.exists():
        return []

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Sort in descending order
    sorted_indices = sorted([int(i) for i in col_indices], reverse=True)
    for idx in sorted_indices:
        if 1 <= idx <= ws.max_column:
            ws.delete_cols(idx)

    compact_worksheet(ws)

    # Save
    try:
        wb.save(file_path)
    except PermissionError:
        raise PermissionError(f"'{file_path}' is currently open in Microsoft Excel. Please close Excel and try again!")

    # Return updated headers
    return [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1) if ws.cell(1, c).value]

def delete_row_from_excel(row_num: int, file_path: str = EXCEL_FILE) -> bool:
    """Deletes a specific data row from the Excel file (1-based index) and compacts the sheet."""
    p = Path(file_path)
    if not p.exists():
        return False

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    if 2 <= row_num <= ws.max_row:
        ws.delete_rows(row_num)
        compact_worksheet(ws)
        try:
            wb.save(file_path)
            return True
        except PermissionError:
            raise PermissionError(f"'{file_path}' is currently open in Microsoft Excel. Please close Excel and try again!")
    return False

def append_to_bug_report(bug_dict: dict, file_path: str = EXCEL_FILE) -> str:
    """
    Appends a new bug report directly into the single Excel spreadsheet.
    Handles Excel file locks automatically with friendly user retry!
    """
    p = Path(file_path)
    if p.exists():
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bug Reports"
        ws.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        ws.row_dimensions[1].height = 28

        col_widths = {
            "A": 15, "B": 26, "C": 35, "D": 32, "E": 20,
            "F": 38, "G": 34, "H": 36, "I": 12, "J": 35,
            "K": 18, "L": 32, "M": 14, "N": 16, "O": 18
        }
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = col_widths.get(col_letter, 22)

    # Clean existing ghost rows first
    compact_worksheet(ws)

    # Determine auto TID (format: TID<month><day><sequence:02d>)
    tid = bug_dict.get("tid")
    if not tid or not str(tid).strip() or str(tid).startswith("TC-"):
        tid = get_next_tid(ws, bug_dict.get("module", ""))

    next_row = ws.max_row + 1

    # Read active headers from row 1 of the sheet
    headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
    while headers and not headers[-1]:
        headers.pop()
    if not headers:
        headers = COLUMNS

    # Build values dynamically according to active headers
    status_count = 0
    values = []
    header_names = []
    for c_idx, h in enumerate(headers, 1):
        hn = h.strip()
        header_names.append(hn)
        if hn == "TID":
            values.append(tid)
        elif hn == "Module":
            values.append(bug_dict.get("module", "General"))
        elif hn == "Test Case Title":
            values.append(bug_dict.get("title", ""))
        elif hn == "Test Scenario":
            values.append(bug_dict.get("scenario", ""))
        elif hn == "Test Data":
            values.append(bug_dict.get("test_data", "N/A"))
        elif hn == "Test Step":
            values.append(bug_dict.get("steps", ""))
        elif hn == "Expected Result":
            values.append(bug_dict.get("expected", ""))
        elif hn == "Actual Result":
            values.append(bug_dict.get("actual", ""))
        elif hn == "Screenshort":
            values.append(bug_dict.get("screenshort", "-"))
        elif hn == "Dev Comment":
            values.append(bug_dict.get("dev_comment", ""))
        elif hn == "QA Comment":
            values.append(bug_dict.get("qa_comment", "Logged by QA"))
        elif hn == "Status":
            status_count += 1
            if status_count == 1:
                values.append(bug_dict.get("status", "Fail"))
            else:
                values.append(bug_dict.get("defect_status", "New"))
        elif hn == "Dev":
            values.append(bug_dict.get("dev", "Unassigned"))
        elif "Dev Com" in hn:
            values.append(bug_dict.get("dev_comennt", ""))
        else:
            values.append(bug_dict.get(hn, ""))

    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if next_row % 2 == 0 else PatternFill(fill_type=None)

    for col_idx, (hn, val) in enumerate(zip(header_names, values), start=1):
        cell = ws.cell(row=next_row, column=col_idx)
        cell.value = val
        cell.border = thin_border
        cell.fill = zebra_fill
        cell.font = Font(name="Segoe UI", size=10)
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # Highlight Status columns
        if hn == "Status":
            if str(val).lower() in ["fail", "bug", "blocked"]:
                cell.fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="9B1C1C")
                cell.alignment = Alignment(horizontal="center", vertical="top")
            elif str(val).lower() in ["new", "open", "in progress"]:
                cell.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="713F12")
                cell.alignment = Alignment(horizontal="center", vertical="top")
            elif str(val).lower() in ["pass", "passed", "resolved", "closed"]:
                cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="065F46")
                cell.alignment = Alignment(horizontal="center", vertical="top")

        elif hn == "Screenshort":
            if str(val).startswith("http"):
                cell.hyperlink = val
                cell.font = Font(name="Segoe UI", size=10, color="0B69A3", underline="single")

        elif hn in ("TID", "Dev"):
            cell.alignment = Alignment(horizontal="center", vertical="top")

    # Dynamic row height based on content lines (especially multiple screenshots & steps)
    lines_count = max([len(str(v).split("\n")) for v in values] + [1])
    ws.row_dimensions[next_row].height = max(40, min(160, lines_count * 20 + 15))

    # Compact worksheet before saving
    compact_worksheet(ws)

    # Save with Excel lock handling
    try:
        wb.save(file_path)
    except PermissionError:
        raise PermissionError(f"'{file_path}' is currently open in Microsoft Excel. Please close Excel and try again!")

    return tid

# ==============================================================================
# 🚀 MAIN RUNNER
# ==============================================================================
def main():
    print("=" * 68)
    print("   🐞 Permanent QA Bug Reporter (Screenshot -> Cloud -> Excel)")
    print("=" * 68)

    # 1. Screenshot input
    if len(sys.argv) > 1:
        ss_input = sys.argv[1].strip("\"'")
    else:
        print("\n📸 Screenshot options:")
        print("   - Drag & drop image file here")
        print("   - Type 'clip' to grab screenshot from Windows Clipboard")
        print("   - Press [Enter] to skip screenshot")
        ss_input = input("👉 Screenshot input: ").strip("\"'")

    # 2. Problem description input
    if len(sys.argv) > 2:
        problem_desc = " ".join(sys.argv[2:]).strip()
    else:
        print("\n📝 Problem description:")
        print("   (Enter bug in Bengali, Banglish, English, or paste table)")
        problem_desc = input("👉 Problem: ").strip()

    if not problem_desc:
        print("❌ Error: Problem description cannot be empty.")
        return

    # 3. Upload screenshot to cloud
    cloud_url = "-"
    if ss_input and ss_input != "-":
        cloud_url = get_screenshot_cloud_url(ss_input)

    # 4. Convert problem to Professional QA Report
    print("🧠 Structuring defect into professional QA format...")
    qa_data = convert_problem_to_qa_report(problem_desc)
    qa_data["screenshort"] = cloud_url

    # 5. Append to Excel
    tid = append_to_bug_report(qa_data, EXCEL_FILE)

    print("\n" + "=" * 68)
    print(f"🎉 Successfully saved [{tid}] into '{EXCEL_FILE}'!")
    print(f"📌 TID:         {tid}")
    print(f"📦 Module:      {qa_data['module']}")
    print(f"🎯 Title:       {qa_data['title']}")
    print(f"📝 Expected:    {qa_data['expected']}")
    print(f"💥 Actual:      {qa_data['actual']}")
    print(f"🔗 Cloud Link:  {cloud_url}")
    print("=" * 68 + "\n")

if __name__ == "__main__":
    main()
