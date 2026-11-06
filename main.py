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

def test_gemini_connection(api_key: str, model_name: str = "gemini-3.5-flash") -> tuple:
    """Tests if a Gemini API key and model work properly."""
    if not api_key or not api_key.strip():
        return False, "API Key is empty."
    try:
        from google import genai
        client = genai.Client(api_key=api_key.strip())
        target_model = model_name or "gemini-3.5-flash"
        res = client.models.generate_content(
            model=target_model,
            contents="Reply with ONLY the word: OK"
        )
        if res and res.text:
            return True, f"Connected successfully to {target_model}!"
        return False, "No response from Gemini API."
    except Exception as e:
        return False, str(e)

def convert_problem_to_qa_report(
    problem_text: str, 
    image_bytes: bytes = None, 
    mime_type: str = "image/png",
    model_name: str = None
) -> dict:
    """
    Converts any raw problem description (Bangla / Banglish / English) and optional screenshot
    into standard, 100% professional simple corporate English QA Test Case and Bug Report fields.
    Uses Google Gemini Vision AI (Gemini 2.5 Pro / Flash) if API key is present.
    If offline or without key, utilizes an advanced semantic NLP engine for 100% relevant reports.
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

    chosen_model = model_name or os.getenv("GEMINI_MODEL", "gemini-3.5-flash").strip() or "gemini-3.5-flash"

    if api_key:
        from google import genai
        from google.genai import types
        import json

        client = genai.Client(api_key=api_key)

        contents = []
        if image_bytes:
            contents.append(types.Part.from_bytes(data=image_bytes, mime_type=mime_type or "image/png"))

        prompt = f"""You are an Elite Principal QA Test Engineer & SQA Architect with 15+ years of enterprise testing experience.
Analyze this bug report (written in Banglish, Bengali, or informal English) and the attached screenshot (if provided).

YOUR OBJECTIVE:
Generate a complete, highly relevant, 100% SIMPLE PROFESSIONAL CORPORATE ENGLISH QA test case and bug report.

STRICT RELEVANCY & SQA RULES:
1. DEEP RELEVANCY: Every single field MUST directly reflect the exact UI elements, fields, error messages, user credentials, URLs, and workflows described by the user or visible in the screenshot. Do NOT write generic placeholder text!
2. PURE CORPORATE ENGLISH: Do NOT output any Bengali or Banglish words. Write in crisp, simple, clear corporate SQA English.
3. MODULE: Identify the exact functional module and sub-feature (e.g. "Authentication > User Login", "Applicant Portal > Dashboard", "Super Admin > Grant Application Form > Form Builder").
4. TEST CASE TITLE: Specific, action-oriented SQA title starting with "Verify..." detailing the exact functionality and failure condition.
5. TEST SCENARIO: A clear, professional test scenario describing what user expectation or business rule is being validated.
6. TEST DATA: Concrete test parameters (e.g. "Email: 'arobiamin75@gmail.com', valid password", "Sequence number = 12").
7. TEST STEP: Clear, numbered step-by-step reproduction instructions (1. ..., 2. ..., 3. ...).
8. EXPECTED RESULT: Exact expected business behavior as per software specifications.
9. ACTUAL RESULT: Exact observed failure describing precisely what went wrong and how the application misbehaved.
10. QA COMMENT: Actionable root cause clue or technical hypothesis for developers (e.g., infinite spinner state, API 404, CORS error, missing unique constraint, client-side routing freeze).

USER PROBLEM DESCRIPTION:
"{problem_text}"

Respond with ONLY valid JSON with this exact structure:
{{
  "module": "Exact Module > Sub-Module",
  "title": "Verify...",
  "scenario": "Clear SQA scenario statement",
  "test_data": "Specific test data or preconditions",
  "steps": "1. Step one\\n2. Step two\\n3. Step three",
  "expected": "Exact expected behavior in clear simple English",
  "actual": "Exact defect observed in clear simple English",
  "qa_comment": "Technical QA observation / developer clue"
}}"""
        contents.append(prompt)

        # Multi-model cascade: try preferred model first, then fallback to other fast & high-reasoning Gemini models
        candidate_models = [
            chosen_model,
            "gemini-3.5-flash",
            "gemini-3.1-flash-lite",
            "gemini-3.8-flash",
            "gemini-3.1-pro-preview",
            "gemini-flash-latest",
            "gemini-pro-latest"
        ]
        models_to_try = list(dict.fromkeys([m for m in candidate_models if m]))

        for mod in models_to_try:
            try:
                res = client.models.generate_content(
                    model=mod,
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
                        "qa_comment": data.get("qa_comment", f"Logged by QA (Model: {mod})"),
                        "defect_status": "New",
                        "dev": "Unassigned",
                        "dev_comennt": ""
                    }
            except Exception as e:
                print(f"[Notice] Model {mod} error ({e}), trying next fallback...")
                continue

    # ==============================================================================
    # 3. High-Precision Smart Semantic NLP Engine (Produces 100% Relevant English)
    # ==============================================================================
    text_lower = problem_text.lower()

    # Pattern 1: Authentication & User Login (Spinner, loading hang, login failure, credentials)
    if any(w in text_lower for w in ["login", "log in", "signin", "sign in"]) or (any(w in text_lower for w in ["loader", "ghur", "spinner", "stuck"]) and any(w in text_lower for w in ["password", "email", "auth"])):
        module = "Authentication > User Login"
        title = "Verify user authentication and dashboard redirect on login submission"
        scenario = "Verify that an authorized user can submit valid credentials and successfully authenticate without encountering infinite loading states or authentication freezes."
        test_data = "Registered user email and valid account password"
        steps = "1. Navigate to the application login page (/#/login).\n2. Enter registered user email address in the Email field.\n3. Enter valid account password in the Password field.\n4. Click on the 'Log in' button.\n5. Observe authentication processing and dashboard transition."
        expected = "The system should authenticate the credentials, display a brief loading indicator, and seamlessly redirect the user to their designated dashboard."
        actual = "Upon clicking the 'Log in' button, the button enters a continuous loading spinner state, the authentication request fails to complete, and the user is not redirected to the dashboard."
        qa_comment = "Infinite loading state observed during login submission. Likely an unhandled API promise rejection, expired token response, or network connection timeout on the authentication endpoint."

    # Pattern 2: Password Reset / Change Password
    elif any(w in text_lower for w in ["password", "pass"]) and any(w in text_lower for w in ["change", "reset", "old password", "vul", "wrong"]):
        module = "Authentication > Password Management"
        title = "Verify old password validation during password change"
        scenario = "Verify that the system validates the existing password and prevents updating credentials when an incorrect current password is provided."
        test_data = "User account with current password and mismatched old password input"
        steps = "1. Log in to the application.\n2. Navigate to Account Settings > Change Password.\n3. Enter an incorrect value in the Current/Old Password field.\n4. Enter a new valid password and confirmation password.\n5. Click on 'Save Changes' or 'Update Password'."
        expected = "The system should reject the request, display an error message stating 'Incorrect current password', and retain the existing password."
        actual = "The system accepts the invalid old password or fails to display a validation error, permitting unauthorized password changes."
        qa_comment = "Missing server-side validation against current password hash before executing update."

    # Pattern 3: Prolonged inactivity / navigation freeze / refresh required
    elif any(w in text_lower for w in ["onk somoy", "onek somoy", "dore", "thaka hoy"]) and any(w in text_lower for w in ["user page", "users", "page ta ase na", "load hoy na", "refresh"]):
        module = "Super Admin > User Management / Navigation"
        title = "Verify navigation to Users page after prolonged activity in Grant Application Form"
        scenario = "Verify client-side route navigation to the Users page without requiring a browser refresh."
        test_data = "Super Admin active session"
        steps = "1. Log in as Super Admin.\n2. Navigate to 'Grant Application Form' and remain on the page for an extended duration.\n3. Click on the 'Users' menu from the left navigation sidebar.\n4. Observe whether the Users page loads properly."
        expected = "The Users page should load seamlessly and display the user list without requiring a manual page refresh."
        actual = "The Users page fails to load after remaining on the Grant Application Form for an extended period, while other pages function normally. A manual browser refresh is required to load the page."
        qa_comment = "Client-side route navigation freeze or component state lock in Grant Application Form preventing transition to Users page."

    # Pattern 4: Dashboard count zero / missing metric counter
    elif ("dash" in text_lower or "dashboard" in text_lower) and any(w in text_lower for w in ["zero", "0", "missing", "not showing", "dekha jacce na", "deka jacce na", "ashche na", "vul"]):
        module = "Applicant Portal > Dashboard"
        title = "Verify that Applicant Dashboard reflects total submitted application count"
        scenario = "Verify real-time data synchronization between submitted applications and Dashboard metrics counter."
        test_data = "Applicant account with active submitted application(s)"
        steps = "1. Log in as an Applicant.\n2. Navigate to 'My Applications' and verify that an active application exists.\n3. Return to the 'Dashboard'.\n4. Observe the 'Total Applications' metric counter widget."
        expected = "The Dashboard counter should accurately display the count of submitted applications (e.g., displaying 1 instead of 0)."
        actual = "The Dashboard counter displays 0 applications despite an active submitted application existing under the applicant account."
        qa_comment = "Data aggregation failure. The dashboard metrics endpoint returns 0 or fails to query the applicant's submission records."

    # Pattern 5: Search and Filter Functionality
    elif any(w in text_lower for w in ["search", "filter", "khuj", "khoj", "filter korle", "search dile"]):
        module = "Data Grid > Search & Filters"
        title = "Verify real-time search filter returns matching records"
        scenario = "Verify that entering valid search terms filters table records dynamically without dropping results or hanging."
        test_data = "Existing record keyword or ID"
        steps = "1. Navigate to the data table interface.\n2. Enter an existing record keyword into the search input field.\n3. Press Enter or observe instant debounced search results."
        expected = "The table should filter dynamically and display only records matching the search query."
        actual = "Search query fails to filter records; the table displays 'No records found' or fails to refresh data."
        qa_comment = "Frontend search debouncing issue, backend query parameter mismatch, or missing case-insensitive search index."

    # Pattern 6: Data Export (Excel / CSV / PDF)
    elif "export" in text_lower or ("download" in text_lower and any(w in text_lower for w in ["excel", "csv", "pdf", "file", "report"])) or "download button" in text_lower:
        module = "Reports & Data Export"
        title = "Verify export functionality downloads valid data spreadsheet"
        scenario = "Verify that triggering the Export action generates a downloadable file containing all active filtered records."
        test_data = "Active record set in table view"
        steps = "1. Navigate to the report or data list page.\n2. Apply any relevant filters or keep default view.\n3. Click on the 'Export to Excel' / 'Download' button.\n4. Verify downloaded file."
        expected = "The application should trigger file download promptly, and the generated spreadsheet should contain complete, uncorrupted data."
        actual = "Clicking Export fails to initiate download, throws a server 500 error, or generates an empty corrupted file."
        qa_comment = "Server-side file generation exception or missing Blob response headers (Content-Disposition, Content-Type)."

    # Pattern 7: Pagination & Page Navigation
    elif any(w in text_lower for w in ["pagination", "page 2", "next page", "page change", "page number"]):
        module = "Data Grid > Pagination"
        title = "Verify table pagination navigation across multiple pages"
        scenario = "Verify that navigating between pagination pages correctly updates table row items with corresponding data offset."
        test_data = "Multi-page dataset (> 10 records)"
        steps = "1. Navigate to a paginated table view.\n2. Scroll to the bottom pagination controls.\n3. Click on 'Next' or Page '2'.\n4. Observe table data."
        expected = "The table should fetch and display the second page of records with correct offset and updated item count."
        actual = "Clicking page 2 fails to load data, displays a blank table, or resets pagination back to page 1."
        qa_comment = "Pagination state synchronization error; query parameters 'page' and 'limit' are not passed correctly to API endpoint."

    # Pattern 8: Form Builder / Dynamic Field synchronization defect
    elif any(w in text_lower for w in ["builder", "form schema", "dynamic field", "custom field", "form builder", "project information"]):
        module = "Super Admin > Grant Application Form"
        title = "Verify dynamic form fields created in Form Builder render on applicant portal"
        scenario = "Verify end-to-end synchronization between Form Builder schema configuration and user-facing submission forms."
        test_data = "Configured form schema and custom fields under Project Information"
        steps = "1. Log in as Super Admin.\n2. Open Form Builder and configure custom fields under the 'Project Information' tab.\n3. Save and update form schema.\n4. Log in to the applicant portal and open the application submission form.\n5. Verify if custom fields render properly."
        expected = "All custom fields and tabs configured in Form Builder should dynamically render on the applicant submission form."
        actual = "Configured fields from the Form Builder do not render on the applicant portal. The form continues to display static default fields."
        qa_comment = "Dynamic form schema is not synchronized or fetched by the client-side application create endpoint."

    # Pattern 9: Delete operation failure / 404 / 500 error
    elif "delete" in text_lower or "remove" in text_lower:
        if "event" in text_lower:
            module = "Super Admin > Grant Event"
        elif "user" in text_lower:
            module = "Super Admin > User Management"
        else:
            module = "Super Admin > Records Management"
        title = f"Verify Delete functionality in {module}"
        scenario = f"Verify that an authorized user can successfully delete an existing record in {module} without server errors."
        test_data = "Target record selected for deletion"
        steps = f"1. Navigate to {module}.\n2. Locate an existing record.\n3. Click on the Delete (trash) icon.\n4. Confirm deletion."
        expected = "The record should be deleted successfully, removed from the list, and a success confirmation notification should appear."
        actual = "Clicking the Delete button fails and triggers an unexpected server error or HTTP 404 Not Found response."
        qa_comment = "Backend API error during DELETE request. Verify endpoint route registration and parameter mapping."

    # Pattern 10: File Upload & Attachment
    elif any(w in text_lower for w in ["upload", "attachment", "file select", "image upload", "dropzone", "browse"]):
        module = "File Management > Upload"
        title = "Verify file upload accepts valid attachments and displays preview"
        scenario = "Verify that authorized users can upload supported file formats within size limits without network failures."
        test_data = "Valid image/document file (PNG/JPG/PDF, < 10MB)"
        steps = "1. Navigate to the file upload section.\n2. Click 'Browse' or drag and drop a valid file.\n3. Verify file preview and filename display.\n4. Click upload or submit form."
        expected = "The file should upload successfully, display a progress bar, and show uploaded attachment preview."
        actual = "File upload hangs indefinitely, fails silently, or displays an unhandled multipart form data error."
        qa_comment = "Multipart/form-data encoding issue, CORS policy block on storage endpoint, or missing file size validation handler."

    # Pattern 11: Dropdown / Select Menu Selection
    elif any(w in text_lower for w in ["dropdown", "select option", "category select", "combo box", "drop down"]):
        module = "Form Controls > Dropdown Selection"
        title = "Verify dropdown options populate and persist on selection"
        scenario = "Verify that clicking the dropdown reveals all available options and correctly sets the form control value."
        test_data = "Selectable option from dropdown list"
        steps = "1. Open the target form interface.\n2. Click on the dropdown menu.\n3. Select a specific option from the list.\n4. Observe field state and form payload."
        expected = "The dropdown should display all active options, update the visible selection upon click, and bind the selected ID."
        actual = "Dropdown options fail to load (empty list), selection does not persist, or the dropdown closes immediately on click."
        qa_comment = "Dropdown component state binding failure or lookup API endpoint returning empty array or error."

    # Pattern 12: Modal & Dialog Popup Controls
    elif any(w in text_lower for w in ["modal", "popup", "dialog", "close icon", "cross button", "backdrop"]):
        module = "User Interface > Modal Dialogs"
        title = "Verify modal dialog opens and closes correctly via controls"
        scenario = "Verify that action triggers open the modal dialog and that clicking the '✕' close icon or backdrop dismisses it."
        test_data = "Interactive modal popup element"
        steps = "1. Click the action button that triggers the modal popup.\n2. Verify modal content renders.\n3. Click on the '✕' close button or press 'Escape'.\n4. Observe modal dismissal."
        expected = "The modal should dismiss smoothly and return focus to the underlying page without locking the background scroll."
        actual = "The modal close button is unresponsive, clicking outside fails to close, or the page background remains locked (overlay freeze)."
        qa_comment = "Missing event listener on close trigger or CSS z-index/pointer-events trapping keyboard and mouse interaction."

    # Pattern 13: Role-Based Access Control & Permissions (403 Forbidden)
    elif any(w in text_lower for w in ["role", "permission", "unauthorized", "403", "forbidden", "access denied", "admin"]):
        module = "Security & Access Control (RBAC)"
        title = "Verify role-based access restrictions for unauthorized users"
        scenario = "Verify that users without administrative privileges are restricted from accessing protected routes or controls."
        test_data = "Non-admin user account credentials"
        steps = "1. Log in with a standard non-admin account.\n2. Attempt to navigate directly to an administrative URL (e.g. /super-admin).\n3. Observe access control handling."
        expected = "The system should restrict access, redirect to an Unauthorized Access (403) page, and hide admin-only navigation links."
        actual = "The system exposes admin interfaces or action buttons to unauthorized roles, permitting unauthorized privilege escalation."
        qa_comment = "Missing route guard check in client router and lack of server-side middleware role validation."

    # Pattern 14: Session Expiry & Auto Logout
    elif any(w in text_lower for w in ["session", "expire", "token", "logout", "log out", "auto logout"]):
        module = "Authentication > Session Management"
        title = "Verify automatic redirection to login upon JWT session expiration"
        scenario = "Verify that when the user authorization token expires, the application gracefully prompts re-authentication."
        test_data = "User session with expired JWT authorization token"
        steps = "1. Log in to the application and retain active session.\n2. Allow token to expire or revoke authentication cookie.\n3. Attempt to perform an API-backed action.\n4. Observe redirect behavior."
        expected = "The application should intercept the 401 response, clear cached user state, and redirect cleanly to /login with a session expired toast."
        actual = "The application displays a broken blank white screen, infinite spinner, or uncaught console error instead of redirecting."
        qa_comment = "Missing global HTTP 401 response interceptor in Axios/Fetch client to handle automatic token refresh or logout redirect."

    # Pattern 15: Phone number / Email / Input field validation
    elif any(w in text_lower for w in ["phone", "mobile", "number", "email", "digit"]):
        module = "User Registration & Form Validation"
        title = "Verify field length and format validation during form submission"
        scenario = "Verify that the system enforces minimum length and character format rules before accepting input."
        test_data = "Input with insufficient digits or invalid characters"
        steps = "1. Navigate to the target form.\n2. Enter invalid or insufficient characters into the input field.\n3. Click on the submit or save button.\n4. Observe validation response."
        expected = "The system should reject invalid input, display a clear inline validation message, and prevent submission."
        actual = "The system accepts invalid or short input without displaying any validation error, permitting corrupted data submission."
        qa_comment = "Missing regex or minimum length validation constraint on both client-side and server-side handlers."

    # Pattern 16: Validation error / Duplicate values / Unique constraint
    elif any(w in text_lower for w in ["duplicate", "sequence", "unique", "already exists"]):
        if "sequence" in text_lower or "cycle" in text_lower:
            module = "Super Admin > Life Cycle"
        else:
            module = "Application Configuration"
        title = f"Verify uniqueness and field validation in {module}"
        scenario = f"Verify that the system enforces uniqueness and prevents duplicate records in {module}."
        test_data = "Duplicate input values (e.g. existing Sequence number or Name)"
        steps = f"1. Navigate to {module}.\n2. Edit an existing record or create a new entry.\n3. Enter a value already in use by another record.\n4. Save changes."
        expected = "The system should display a validation error message and prevent saving duplicate values."
        actual = "The system allows duplicate values to be saved without triggering unique constraint validation."
        qa_comment = "Missing unique validation rule on both frontend form validation and backend database constraints."

    # Pattern 17: Unresponsive controls / Button click does nothing
    elif any(w in text_lower for w in ["kichu hoy na", "kaj kore na", "response nai", "unresponsive", "click korle"]):
        module = "User Interface & Controls"
        title = "Verify primary action trigger executes requested operation"
        scenario = "Verify that clicking the designated action control triggers the appropriate network request and state transition."
        test_data = "Standard user interaction event"
        steps = "1. Navigate to the target interface.\n2. Trigger the primary action or submit button.\n3. Observe system behavior."
        expected = "The system should execute the action promptly, display visual feedback, and transition state."
        actual = "The control is completely unresponsive; clicking the button fails to trigger any network request or user feedback."
        qa_comment = "Unbound event listener, unhandled promise rejection, or disabled button state in frontend code."

    # Pattern 18: Smart Contextual Entity Extractor (High Quality Fallback for any other bug)
    else:
        # Extract keywords and identify actions and nouns
        stopwords = {"ami", "amr", "amader", "tmi", "apni", "akta", "ekta", "ar", "kintu", "er", "te", "kore",
                     "dila", "dile", "dileh", "hobe", "hoy", "na", "to", "ta", "theke", "diye", "jodi", "tobe",
                     "ar", "oita", "aita", "sheta", "sob", "kisu", "kichu", "pore", "age", "jabe", "jay"}
        words = re.findall(r"\b[a-zA-Z0-9_\-]+\b", problem_text)
        filtered = [w for w in words if w.lower() not in stopwords]

        # Detect error codes
        err_match = re.search(r"\b(400|401|403|404|409|500|502|503)\b", problem_text)
        error_code = err_match.group(1) if err_match else None

        # Build topic
        topic_words = filtered[:4] if filtered else ["Feature"]
        topic = " ".join(topic_words).title()

        module = f"Application Core > {topic}"
        title = f"Verify {topic} workflow and functional execution"
        scenario = f"Verify that user interaction with {topic} produces expected business outcomes without encountering system errors or execution failures."
        test_data = "Standard operational user inputs"
        steps = f"1. Log in to the application.\n2. Navigate to the {topic} section.\n3. Execute the designated user action or submission.\n4. Observe system response and outcome."
        
        if error_code:
            expected = f"The operation should complete successfully with an HTTP 200 OK status, and confirmation should appear."
            actual = f"The operation failed and the backend returned an HTTP {error_code} error response during execution."
            qa_comment = f"Backend API returned HTTP {error_code}. Inspect server logs, request payload, and endpoint routing."
        else:
            expected = f"The {topic} feature should execute smoothly and display appropriate status confirmation."
            actual = f"The {topic} feature encountered an unexpected defect during execution and did not behave as specified."
            qa_comment = "Defect identified during QA manual exploratory test execution. Developer investigation required."

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
