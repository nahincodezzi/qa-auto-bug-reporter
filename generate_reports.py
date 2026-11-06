import os
import sys
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
    HRFlowable
)
from reportlab.pdfgen import canvas

# Create reports directory
REPORTS_DIR = Path("reports")
REPORTS_DIR.mkdir(exist_ok=True)

TECH_PDF_PATH = REPORTS_DIR / "Technical_Report.pdf"
USER_PDF_PATH = REPORTS_DIR / "User_Guideline.pdf"

# Numbered Canvas for "Page X of Y" and Running Headers/Footers
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#64748B"))
        
        # Don't draw header/footer on cover page if page 1
        if self._pageNumber > 1:
            # Header
            self.drawString(54, 750, getattr(self, "doc_title", "QA Auto Bug Reporter Pro"))
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(54, 742, 558, 742)
            
            # Footer
            self.line(54, 50, 558, 50)
            self.drawString(54, 38, "Confidential - SQA Enterprise Automation Infrastructure")
            page_str = f"Page {self._pageNumber} of {page_count}"
            self.drawRightString(558, 38, page_str)
        else:
            # Cover page footer
            self.line(54, 50, 558, 50)
            self.drawString(54, 38, "QA Auto Bug Reporter Pro - Technical Documentation Suite")
            self.drawRightString(558, 38, f"Total Pages: {page_count}")
        self.restoreState()

def create_tech_canvas(*args, **kwargs):
    c = NumberedCanvas(*args, **kwargs)
    c.doc_title = "Technical Architecture & System Design Report | QA Auto Bug Reporter Pro"
    return c

def create_user_canvas(*args, **kwargs):
    c = NumberedCanvas(*args, **kwargs)
    c.doc_title = "User Operations & Workflow Manual | QA Auto Bug Reporter Pro"
    return c

def build_styles():
    styles = getSampleStyleSheet()
    
    primary = colors.HexColor("#1B365D")
    secondary = colors.HexColor("#2563EB")
    dark = colors.HexColor("#0F172A")
    slate = colors.HexColor("#334155")
    
    styles.add(ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        textColor=primary,
        spaceAfter=8
    ))
    styles.add(ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=16,
        textColor=secondary,
        spaceAfter=20
    ))
    styles.add(ParagraphStyle(
        'MetaBox',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=14,
        textColor=slate
    ))
    styles.add(ParagraphStyle(
        'SectionHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=primary,
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    ))
    styles.add(ParagraphStyle(
        'SubSectionHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=secondary,
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True
    ))
    styles.add(ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        textColor=dark,
        spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        'CustomBullet',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=dark,
        leftIndent=14,
        firstLineIndent=-10,
        spaceAfter=4
    ))
    styles.add(ParagraphStyle(
        'CalloutText',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#1E3A8A")
    ))
    styles.add(ParagraphStyle(
        'CodeStyle',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#0F172A")
    ))
    styles.add(ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.white,
        alignment=0
    ))
    styles.add(ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=dark
    ))
    styles.add(ParagraphStyle(
        'TableCellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=primary
    ))
    return styles

# ==============================================================================
# 📄 1. TECHNICAL ARCHITECTURE & SYSTEM DESIGN REPORT
# ==============================================================================
def generate_technical_report():
    doc = SimpleDocTemplate(
        str(TECH_PDF_PATH),
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    styles = build_styles()
    story = []
    
    # Header Banner
    story.append(Paragraph("QA Auto Bug Reporter Pro", styles['DocTitle']))
    story.append(Paragraph("Technical Architecture & System Design Specification", styles['DocSubTitle']))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1B365D"), spaceAfter=14))
    
    # Metadata Table
    meta_data = [
        [
            Paragraph("<b>System Name:</b> QA Auto Bug Reporter Pro", styles['MetaBox']),
            Paragraph("<b>Target Framework:</b> Python 3.10+ / Flask Microservice", styles['MetaBox'])
        ],
        [
            Paragraph("<b>Primary AI Model:</b> Google Gemini 2.5 Pro Vision", styles['MetaBox']),
            Paragraph("<b>Persistence Store:</b> Microsoft Excel (OpenPyXL Engine)", styles['MetaBox'])
        ],
        [
            Paragraph("<b>Repository:</b> git@github.com:nahincodezzi/qa-auto-bug-reporter.git", styles['MetaBox']),
            Paragraph("<b>Deployment Target:</b> Windows OS Standalone Executable / Desktop Web", styles['MetaBox'])
        ]
    ]
    t_meta = Table(meta_data, colWidths=[250, 254])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F1F5F9")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 14))
    
    # Section 1
    story.append(Paragraph("1. Executive Summary & Architectural Mission", styles['SectionHeading']))
    story.append(Paragraph(
        "Modern enterprise software testing requires high velocity and meticulous documentation. However, Quality Assurance (QA) "
        "engineers frequently encounter operational friction when manually taking screenshots, uploading assets to cloud servers, copying URLs, "
        "and manually translating localized informal bug descriptions (Banglish, Bengali, informal notes) into standard English test cases. "
        "<b>QA Auto Bug Reporter Pro</b> is an automated end-to-end testing infrastructure designed to eliminate this friction entirely.",
        styles['CustomBody']
    ))
    story.append(Paragraph(
        "By integrating an asynchronous clipboard listener (<kbd>Ctrl+V</kbd>), automated binary cloud image streaming, Google's flagship "
        "<b>Gemini 2.5 Pro Multimodal Vision AI</b>, and an industrial-grade OpenPyXL spreadsheet automation engine, the system achieves "
        "zero-touch bug reporting directly into a persistent 15-column master corporate defect ledger (<code>bug_report.xlsx</code>).",
        styles['CustomBody']
    ))
    story.append(Spacer(1, 10))
    
    # Section 2: 5-Tier Architecture
    story.append(Paragraph("2. High-Level Multi-Tier System Architecture", styles['SectionHeading']))
    
    arch_data = [
        [Paragraph("Tier / Layer", styles['TableHeader']), Paragraph("Core Technology", styles['TableHeader']), Paragraph("Architectural Responsibility", styles['TableHeader'])],
        [
            Paragraph("<b>1. Presentation Layer</b>", styles['TableCellBold']),
            Paragraph("HTML5, CSS3, Vanilla ES6 JavaScript", styles['TableCell']),
            Paragraph("Direct clipboard interception (<kbd>Ctrl+V</kbd>), multi-file thumbnail gallery, dynamic column deletion UI, real-time AI status badge.", styles['TableCell'])
        ],
        [
            Paragraph("<b>2. Microservice Layer</b>", styles['TableCellBold']),
            Paragraph("Python 3.10+, Flask, Threading", styles['TableCell']),
            Paragraph("RESTful endpoints (/api/add_bug, /api/delete_row, /api/delete_columns), automatic browser launch via daemon thread, file-lock handling.", styles['TableCell'])
        ],
        [
            Paragraph("<b>3. AI & Vision Engine</b>", styles['TableCellBold']),
            Paragraph("Google GenAI SDK (gemini-2.5-pro)", styles['TableCell']),
            Paragraph("Multimodal screenshot inspection, SQA prompt engineering, multi-model fallback cascade, zero-Banglish leakage enforcement.", styles['TableCell'])
        ],
        [
            Paragraph("<b>4. Semantic NLP Fallback</b>", styles['TableCellBold']),
            Paragraph("Regex Pattern Classifier & Entity Engine", styles['TableCell']),
            Paragraph("Zero-key offline operation; parses authentication, validation, forms, delete operations, and constructs 100% relevant English SQA reports.", styles['TableCell'])
        ],
        [
            Paragraph("<b>5. Persistence Layer</b>", styles['TableCellBold']),
            Paragraph("OpenPyXL Spreadsheet Engine", styles['TableCell']),
            Paragraph("Manages master <code>bug_report.xlsx</code>, automated date-based TID generation, hyperlink re-alignment, and middle/trailing sheet compaction.", styles['TableCell'])
        ]
    ]
    t_arch = Table(arch_data, colWidths=[110, 130, 264])
    t_arch.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1B365D")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_arch)
    story.append(Spacer(1, 14))
    
    # Section 3: Deep Technical Breakdown
    story.append(Paragraph("3. In-Depth Component Specifications", styles['SectionHeading']))
    
    story.append(Paragraph("3.1 Asynchronous Clipboard & Memory Streaming", styles['SubSectionHeading']))
    story.append(Paragraph(
        "Traditional web file upload forms require saving screenshots to the physical filesystem before browsing and uploading. "
        "The system overrides the global browser clipboard handler (<code>window.addEventListener('paste')</code>). "
        "When an image is intercepted from the OS clipboard, it is converted into an in-memory <code>File</code> Blob and rendered instantly "
        "into a responsive thumbnail gallery with an active counter badge. Uploading is executed asynchronously via <code>multipart/form-data</code>.",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("3.2 Dual-Provider Cloud Image Streaming", styles['SubSectionHeading']))
    story.append(Paragraph(
        "Screenshots are transmitted as raw byte buffers (<code>io.BytesIO</code>) to freeimage.host API with an automated Catbox.moe "
        "fallback. The service returns a permanent, unexpiring, publicly resolvable HTTPS direct URL (e.g. <code>https://iili.io/ntf415X.png</code>). "
        "When multiple screenshots are attached, the system parses and numbers each link sequentially (<code>SS 1: ...\\nSS 2: ...</code>) "
        "and binds them into clickable hyperlink chips within both the Excel spreadsheet and the dashboard table.",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("3.3 Google Gemini 2.5 Pro Vision AI & Multi-Model Cascade", styles['SubSectionHeading']))
    story.append(Paragraph(
        "The application integrates Google's newly released <b>Gemini 2.5 Pro</b> flagship model. "
        "Unlike standard text-only LLMs, screenshot bytes are encoded as <code>types.Part.from_bytes()</code> and delivered directly to "
        "the multimodal vision pipeline. The AI visually reads the active browser URL bar, button states (e.g. loading spinners), input values, "
        "and HTTP error dialogs. A multi-model cascade guarantees zero system downtime:",
        styles['CustomBody']
    ))
    story.append(Paragraph("• <b>Primary:</b> <code>gemini-2.5-pro</code> — Deepest SQA reasoning and visual precision.", styles['CustomBullet']))
    story.append(Paragraph("• <b>Fallback 1:</b> <code>gemini-2.5-flash</code> — High-speed multimodal inference.", styles['CustomBullet']))
    story.append(Paragraph("• <b>Fallback 2:</b> <code>gemini-2.0-flash</code> — Next-generation ultra-fast inference.", styles['CustomBullet']))
    story.append(Paragraph("• <b>Fallback 3:</b> <code>gemini-1.5-pro</code> — Extended context window backup.", styles['CustomBullet']))
    
    story.append(Spacer(1, 10))
    story.append(PageBreak())
    
    # Page 2: Excel & Data Integrity
    story.append(Paragraph("4. Master Excel Schema & Data Integrity Engineering", styles['SectionHeading']))
    story.append(Paragraph(
        "The application enforces a rigorous, standardized 15-column schema aligned with corporate enterprise SQA benchmarks:",
        styles['CustomBody']
    ))
    
    schema_cols = [
        [Paragraph("Col #", styles['TableHeader']), Paragraph("Column Name", styles['TableHeader']), Paragraph("Type & Formatting", styles['TableHeader']), Paragraph("SQA Description", styles['TableHeader'])],
        [Paragraph("1", styles['TableCellBold']), Paragraph("TID", styles['TableCellBold']), Paragraph("String (Center / Navy Badge)", styles['TableCell']), Paragraph("Date-based unique defect identifier (e.g. TID11601).", styles['TableCell'])],
        [Paragraph("2", styles['TableCellBold']), Paragraph("Module", styles['TableCell']), Paragraph("String (Hierarchy)", styles['TableCell']), Paragraph("Hierarchical module path (e.g. Super Admin > Grant Event).", styles['TableCell'])],
        [Paragraph("3", styles['TableCellBold']), Paragraph("Test Case Title", styles['TableCell']), Paragraph("String (Bold Title)", styles['TableCell']), Paragraph("Action-oriented verification statement starting with 'Verify...'.", styles['TableCell'])],
        [Paragraph("4", styles['TableCellBold']), Paragraph("Test Scenario", styles['TableCell']), Paragraph("String (Pre-wrap)", styles['TableCell']), Paragraph("Business condition or specification rule under verification.", styles['TableCell'])],
        [Paragraph("5", styles['TableCellBold']), Paragraph("Test Data", styles['TableCell']), Paragraph("String (Pre-wrap)", styles['TableCell']), Paragraph("Concrete test input values, parameters, or session credentials.", styles['TableCell'])],
        [Paragraph("6", styles['TableCellBold']), Paragraph("Test Step", styles['TableCell']), Paragraph("Multi-line String (Numbered)", styles['TableCell']), Paragraph("Sequential reproduction steps (1. ..., 2. ..., 3. ...).", styles['TableCell'])],
        [Paragraph("7", styles['TableCellBold']), Paragraph("Expected Result", styles['TableCell']), Paragraph("String (Pre-wrap)", styles['TableCell']), Paragraph("Exact expected business outcome according to specification.", styles['TableCell'])],
        [Paragraph("8", styles['TableCellBold']), Paragraph("Actual Result", styles['TableCell']), Paragraph("String (Red / Pre-wrap)", styles['TableCell']), Paragraph("Observed defect behavior detailing exact deviation from expected.", styles['TableCell'])],
        [Paragraph("9", styles['TableCellBold']), Paragraph("Status", styles['TableCellBold']), Paragraph("Enum (Red Fill: Fail)", styles['TableCell']), Paragraph("Test execution outcome (Fail / Pass / Blocked).", styles['TableCell'])],
        [Paragraph("10", styles['TableCellBold']), Paragraph("Screenshort", styles['TableCell']), Paragraph("Hyperlink / Blue Underline", styles['TableCell']), Paragraph("Public cloud URL(s) to permanent screenshot evidence.", styles['TableCell'])],
        [Paragraph("11", styles['TableCellBold']), Paragraph("Dev Comment", styles['TableCell']), Paragraph("String (Nullable)", styles['TableCell']), Paragraph("Developer feedback, triage notes, or fix PR references.", styles['TableCell'])],
        [Paragraph("12", styles['TableCellBold']), Paragraph("QA Comment", styles['TableCell']), Paragraph("String (Pre-wrap)", styles['TableCell']), Paragraph("Technical QA observations, root cause clues, or API error codes.", styles['TableCell'])],
        [Paragraph("13", styles['TableCellBold']), Paragraph("Status", styles['TableCellBold']), Paragraph("Enum (Yellow Fill: New)", styles['TableCell']), Paragraph("Defect lifecycle state (New / Open / In Progress / Resolved).", styles['TableCell'])],
        [Paragraph("14", styles['TableCellBold']), Paragraph("Dev", styles['TableCell']), Paragraph("String (Center)", styles['TableCell']), Paragraph("Assigned developer or engineering team.", styles['TableCell'])],
        [Paragraph("15", styles['TableCellBold']), Paragraph("Dev Comennt", styles['TableCell']), Paragraph("String (Nullable)", styles['TableCell']), Paragraph("Additional secondary engineering commentary.", styles['TableCell'])]
    ]
    t_schema = Table(schema_cols, colWidths=[36, 120, 134, 214])
    t_schema.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1B365D")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_schema)
    story.append(Spacer(1, 14))
    
    # Section 5: Date-Based TID Algorithm
    story.append(Paragraph("5. Date-Based Dynamic TID Generation Algorithm", styles['SectionHeading']))
    story.append(Paragraph(
        "Test Case IDs are critical for cross-referencing between Jira, GitHub Issues, and automated test runs. "
        "The system generates TIDs using the deterministic formula: <code>TID&lt;Month&gt;&lt;Day&gt;&lt;Sequence:02d&gt;</code>.",
        styles['CustomBody']
    ))
    story.append(Paragraph("• <b>Month & Day:</b> Evaluated in real-time via <code>datetime.now()</code> (e.g. November 6th &rarr; <code>TID116</code>).", styles['CustomBullet']))
    story.append(Paragraph("• <b>Daily Sequential Counter:</b> Evaluated by inspecting existing row identifiers in Column 1. The regex pattern <code>r'(\\d+)$'</code> extracts historical sequence numbers for today's prefix, selecting <code>max(sequences) + 1</code>. This guarantees sequential continuity regardless of row deletions.", styles['CustomBullet']))
    
    story.append(Spacer(1, 10))
    
    # Section 6: OpenPyXL Sheet Compaction & Hyperlink Desync Resolution
    story.append(Paragraph("6. OpenPyXL Sheet Compaction & Ghost Row Resolution", styles['SectionHeading']))
    story.append(Paragraph(
        "A critical technical challenge discovered in openpyxl was the <b>Hyperlink Coordinate Desync Bug</b>. "
        "When <code>ws.delete_rows()</code> deletes an intermediary data row, openpyxl shifts cell data upward but fails to adjust "
        "the internal <code>ws.cell.hyperlink.ref</code> XML attributes. Upon saving, openpyxl resurrected phantom cells at obsolete row coordinates, "
        "creating empty ghost rows displaying <code>- | - | -</code>. This was permanently resolved via the <code>compact_worksheet()</code> algorithm:",
        styles['CustomBody']
    ))
    
    callout_data = [[
        Paragraph(
            "<b>Engineering Resolution (compact_worksheet):</b><br/>"
            "1. <b>Iterative Coordinate Re-alignment:</b> Iterates all active cells and re-binds <code>cell.hyperlink.ref = cell.coordinate</code>.<br/>"
            "2. <b>Bottom-up Trailing Purge:</b> Scans from <code>ws.max_row</code> backwards to Row 2, deleting fully empty trailing rows.<br/>"
            "3. <b>Middle Ghost Row Deletion:</b> Iterates row-by-row and verifies meaningful content across columns; deletes phantom rows.<br/>"
            "4. <b>Dynamic Header Mapping:</b> Reads active headers from Row 1, allowing safe dynamic column deletion without data corruption.",
            styles['CalloutText']
        )
    ]]
    t_callout = Table(callout_data, colWidths=[504])
    t_callout.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EFF6FF")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#BFDBFE")),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
    ]))
    story.append(t_callout)
    story.append(Spacer(1, 14))
    
    # Section 7: File Lock Protection
    story.append(Paragraph("7. Windows Concurrency & Excel File Lock Protection", styles['SectionHeading']))
    story.append(Paragraph(
        "When an end-user opens <code>bug_report.xlsx</code> in Microsoft Excel on Windows, the OS places an exclusive lock on the file descriptor. "
        "Attempts to write via openpyxl raise a <code>PermissionError: [Errno 13]</code>. Instead of crashing the server or hanging the terminal, "
        "the Flask controller catches <code>PermissionError</code> and returns a structured JSON HTTP 400 error. The UI displays an alert toast "
        "instructing the user to close Excel before retrying, ensuring flawless data safety.",
        styles['CustomBody']
    ))
    
    doc.build(story, canvasmaker=create_tech_canvas)
    print(f"✅ Generated: {TECH_PDF_PATH}")

# ==============================================================================
# 📄 2. USER GUIDELINE & OPERATIONS MANUAL
# ==============================================================================
def generate_user_guideline():
    doc = SimpleDocTemplate(
        str(USER_PDF_PATH),
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    styles = build_styles()
    story = []
    
    # Header Banner
    story.append(Paragraph("QA Auto Bug Reporter Pro", styles['DocTitle']))
    story.append(Paragraph("User Operations & Workflow Manual", styles['DocSubTitle']))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1B365D"), spaceAfter=14))
    
    # Introduction
    story.append(Paragraph("1. Welcome & Introduction", styles['SectionHeading']))
    story.append(Paragraph(
        "Welcome to <b>QA Auto Bug Reporter Pro</b>! This user manual provides step-by-step instructions on how to effortlessly "
        "report software defects, upload screenshots to the cloud, translate informal notes into 100% professional corporate English, "
        "and maintain a master Excel bug report spreadsheet.",
        styles['CustomBody']
    ))
    story.append(Spacer(1, 8))
    
    # Quick Launch
    story.append(Paragraph("2. Quick Launch (1-Click Start)", styles['SectionHeading']))
    story.append(Paragraph("Starting the application requires only a single click without terminal commands:", styles['CustomBody']))
    story.append(Paragraph("1. Open the project root folder: <code>d:\\make test excel test case auto</code>", styles['CustomBullet']))
    story.append(Paragraph("2. Double-click the file named <b><code>run.bat</code></b>.", styles['CustomBullet']))
    story.append(Paragraph("3. The black server terminal will initialize, and your default web browser will automatically open the dashboard at:", styles['CustomBullet']))
    story.append(Paragraph("<b>&nbsp;&nbsp;&nbsp;&nbsp;&rarr; URL:</b> <code>http://127.0.0.1:5000</code>", styles['CustomBullet']))
    story.append(Spacer(1, 10))
    
    # Step-by-step Workflow
    story.append(Paragraph("3. Core Defect Logging Workflow", styles['SectionHeading']))
    
    step_data = [
        [Paragraph("Step", styles['TableHeader']), Paragraph("Action", styles['TableHeader']), Paragraph("Best Practice & Guidelines", styles['TableHeader'])],
        [
            Paragraph("<b>Step 1</b>", styles['TableCellBold']),
            Paragraph("Capture Screenshot", styles['TableCellBold']),
            Paragraph("Take a screenshot using Windows Snipping Tool (<kbd>Win + Shift + S</kbd>). Click and drag to capture the buggy screen, URL bar, or error popup.", styles['TableCell'])
        ],
        [
            Paragraph("<b>Step 2</b>", styles['TableCellBold']),
            Paragraph("Direct Paste (<kbd>Ctrl + V</kbd>)", styles['TableCellBold']),
            Paragraph("Switch to your browser and press <kbd>Ctrl + V</kbd>. The image appears instantly in the gallery. You can paste multiple screenshots sequentially or drag & drop files.", styles['TableCell'])
        ],
        [
            Paragraph("<b>Step 3</b>", styles['TableCellBold']),
            Paragraph("Describe in Banglish / Bengali", styles['TableCellBold']),
            Paragraph("Write what went wrong in your own natural words (e.g. <i>'login button a click korle loader ghurteche kintu login hocche na'</i> or <i>'phone number 10 digit er kom dileo save hoy'</i>).", styles['TableCell'])
        ],
        [
            Paragraph("<b>Step 4</b>", styles['TableCellBold']),
            Paragraph("1-Click AI Translation", styles['TableCellBold']),
            Paragraph("Click the blue button: <b>🚀 Convert to Professional QA English & Add to Excel</b>. In 2-3 seconds, screenshots upload to cloud, text is converted, and the row is saved to Excel.", styles['TableCell'])
        ],
        [
            Paragraph("<b>Step 5</b>", styles['TableCellBold']),
            Paragraph("Inspect & Manage", styles['TableCellBold']),
            Paragraph("View the new defect in the live table below. Click on <b>📂 Open bug_report.xlsx</b> to view and edit directly in Microsoft Excel.", styles['TableCell'])
        ]
    ]
    t_step = Table(step_data, colWidths=[50, 130, 324])
    t_step.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1B365D")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_step)
    story.append(Spacer(1, 14))
    
    # Section 4: AI Settings
    story.append(Paragraph("4. AI Model & API Key Configuration (Recommended)", styles['SectionHeading']))
    story.append(Paragraph(
        "While the system includes a built-in Smart Semantic NLP engine that functions 100% offline, connecting a free "
        "Google Gemini API key enables <b>Computer Vision</b> to inspect your screenshots directly.",
        styles['CustomBody']
    ))
    story.append(Paragraph("<b>How to enable Gemini 2.5 Pro Vision AI:</b>", styles['CustomBody']))
    story.append(Paragraph("1. Click the top navbar button: <b><code>AI Model & API Key</code></b>.", styles['CustomBullet']))
    story.append(Paragraph("2. Click the link to get a 100% Free Gemini API key from Google AI Studio (takes 10 seconds).", styles['CustomBullet']))
    story.append(Paragraph("3. Paste your key (starts with <code>AIzaSy...</code>) into the input box.", styles['CustomBullet']))
    story.append(Paragraph("4. Select <b>🏆 Gemini 2.5 Pro (Recommended - Deepest SQA Reasoning)</b> from the model dropdown.", styles['CustomBullet']))
    story.append(Paragraph("5. Click <b>🔌 Test Connection</b> to verify with Google servers.", styles['CustomBullet']))
    story.append(Paragraph("6. Click <b>💾 Save Settings</b>. The status badge will glow green (<code>🟢 AI Active: Gemini 2.5 Pro</code>).", styles['CustomBullet']))
    story.append(Spacer(1, 10))
    
    # PageBreak for Table & Operations
    story.append(PageBreak())
    
    story.append(Paragraph("5. Spreadsheet & Table Management Tools", styles['SectionHeading']))
    story.append(Paragraph(
        "The live spreadsheet dashboard below the entry card mirrors your master <code>bug_report.xlsx</code> file in real time:",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("5.1 Opening Microsoft Excel Directly", styles['SubSectionHeading']))
    story.append(Paragraph(
        "Click the green <b>📂 Open bug_report.xlsx</b> button in the top navbar. Windows will immediately launch Microsoft Excel "
        "displaying your master file with full formatting, freeze panes, borders, and color coding.",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("5.2 Viewing Cloud Screenshots", styles['SubSectionHeading']))
    story.append(Paragraph(
        "In the <b>Screenshort</b> column, click on any <b>🔗 View SS</b> link. It opens the high-resolution screenshot in a new browser tab. "
        "Anyone in your team can view the evidence without needing local access to your computer.",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("5.3 Deleting Columns (Customizing Your Report Format)", styles['SubSectionHeading']))
    story.append(Paragraph(
        "If your team does not need certain columns (e.g. <i>Dev Comment</i> or duplicate <i>Status</i>):",
        styles['CustomBody']
    ))
    story.append(Paragraph("• <b>Single Column:</b> Click the small <b>✕</b> icon on any table column header.", styles['CustomBullet']))
    story.append(Paragraph("• <b>Bulk Deletion:</b> Click the <b>🗑️ Select & Delete Columns</b> button above the table. Select checkboxes for all columns you want to remove and confirm. The system permanently deletes them from <code>bug_report.xlsx</code> without shifting remaining indices.", styles['CustomBullet']))
    
    story.append(Paragraph("5.4 Deleting Rows", styles['SubSectionHeading']))
    story.append(Paragraph(
        "Click the red trash icon (<b>🗑️</b>) on the right side of any row. The row is removed from Excel, and the automated "
        "<b>Sheet Compaction Engine</b> eliminates blank lines, ensuring remaining rows stay tight and sequential.",
        styles['CustomBody']
    ))
    
    story.append(Paragraph("5.5 Real-Time Search Filter", styles['SubSectionHeading']))
    story.append(Paragraph(
        "Type any keyword into the <b>🔍 Search bugs...</b> input box above the table. The table instantly filters matching defects "
        "by TID, Module, Title, or QA observation without reloading.",
        styles['CustomBody']
    ))
    story.append(Spacer(1, 12))
    
    # Section 6: Real-World Transformation Examples
    story.append(Paragraph("6. Before & After Transformation Examples", styles['SectionHeading']))
    story.append(Paragraph(
        "See how the system transforms informal Banglish inputs into enterprise-grade corporate English SQA reports:",
        styles['CustomBody']
    ))
    
    example_data = [
        [Paragraph("Informal Input (Banglish)", styles['TableHeader']), Paragraph("Generated SQA Output (Corporate English)", styles['TableHeader'])],
        [
            Paragraph("<i>'login button a click korle loader ghurteche kintu login hocche na'</i>", styles['TableCell']),
            Paragraph(
                "<b>Module:</b> Authentication > User Login<br/>"
                "<b>Title:</b> Verify user authentication and dashboard redirect on login submission<br/>"
                "<b>Actual Result:</b> Upon clicking 'Log in', the button enters a continuous loading spinner state, authentication fails to complete, and the user is not redirected to dashboard.<br/>"
                "<b>QA Comment:</b> Infinite loading state on login submission. Likely unhandled promise rejection or CORS timeout.",
                styles['TableCell']
            )
        ],
        [
            Paragraph("<i>'applicant portal a submitted application 1 ta thakleo dash a zero dekay'</i>", styles['TableCell']),
            Paragraph(
                "<b>Module:</b> Applicant Portal > Dashboard<br/>"
                "<b>Title:</b> Verify that Applicant Dashboard reflects total submitted application count<br/>"
                "<b>Actual Result:</b> The Dashboard counter displays 0 applications despite an active submitted application existing under the applicant account.<br/>"
                "<b>QA Comment:</b> Data aggregation failure. Dashboard metrics endpoint returns 0 or fails to query applicant submission records.",
                styles['TableCell']
            )
        ],
        [
            Paragraph("<i>'phone number 10 digit er kom dileo submit hoye jacche kono validation error nai'</i>", styles['TableCell']),
            Paragraph(
                "<b>Module:</b> User Registration & Form Validation<br/>"
                "<b>Title:</b> Verify field length and format validation during form submission<br/>"
                "<b>Actual Result:</b> The system accepts invalid or short input without displaying any validation error, permitting corrupted data submission.<br/>"
                "<b>QA Comment:</b> Missing regex or minimum length validation constraint on client-side and server-side handlers.",
                styles['TableCell']
            )
        ]
    ]
    t_ex = Table(example_data, colWidths=[180, 324])
    t_ex.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1B365D")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_ex)
    story.append(Spacer(1, 14))
    
    # Section 7: Troubleshooting
    story.append(Paragraph("7. Frequently Asked Questions & Troubleshooting", styles['SectionHeading']))
    story.append(Paragraph("<b>Q: I got an error saying 'bug_report.xlsx is currently open in Microsoft Excel'. What should I do?</b>", styles['CustomBody']))
    story.append(Paragraph("<b>A:</b> Simply close the Excel window on your computer and click the button again. Windows locks open spreadsheets, but our system catches this safely so no data is lost.", styles['CustomBody']))
    story.append(Paragraph("<b>Q: Can I use this application without an internet connection?</b>", styles['CustomBody']))
    story.append(Paragraph("<b>A:</b> Yes! The built-in Smart Semantic NLP engine converts Banglish into professional English 100% locally. Screenshot cloud uploading will save locally when offline.", styles['CustomBody']))
    story.append(Paragraph("<b>Q: Where are my screenshots stored?</b>", styles['CustomBody']))
    story.append(Paragraph("<b>A:</b> Screenshots are hosted on high-availability cloud servers with direct HTTPS links accessible by developers, project managers, and clients anytime.", styles['CustomBody']))
    
    doc.build(story, canvasmaker=create_user_canvas)
    print(f"✅ Generated: {USER_PDF_PATH}")

if __name__ == "__main__":
    generate_technical_report()
    generate_user_guideline()

