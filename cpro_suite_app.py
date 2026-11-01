import os
import re
import sys
import time
import threading
import subprocess
import ctypes
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd

# Enable High-DPI crisp rendering on Windows
if sys.platform == "win32":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    sync_playwright = None
    PlaywrightTimeoutError = Exception


# ==============================================================================
# APPLE HUMAN INTERFACE DESIGN PALETTE & CONSTANTS
# ==============================================================================
THEME = {
    "bg": "#F5F5F7",
    "card": "#FFFFFF",
    "border": "#E5E5EA",
    "text": "#1D1D1F",
    "subtext": "#86868B",
    "blue": "#007AFF",
    "blue_hover": "#0062CC",
    "green": "#34C759",
    "green_hover": "#28A745",
    "red": "#FF3B30",
    "neutral_btn": "#E8E8ED",
    "neutral_hover": "#D2D2D7",
    "input_bg": "#F2F2F7",
    "row_found": "#E8F8EE",
    "row_missing": "#FFF2F2",
    "row_active": "#EBF5FF",
    "row_error": "#FFF4D6",
}

CASE_LISTING_URL = "https://hlbcclo.hlb.com.my/lawfirm/case-listing"
MAX_RETRIES_ON_KICKOUT = 5
GRID_SETTLE_TIMEOUT_MS = 30000


class GridLoadTimeoutError(RuntimeError):
    pass


class GridSearchError(RuntimeError):
    pass


class InvalidCredentialsError(RuntimeError):
    pass


def _find_writable_output_dirs():
    """Find the best place to save files, prioritizing the script's folder, then Desktop/Documents."""
    candidates = []
    
    # 1. Prioritize the folder where the script (or EXE) is located
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(script_dir)

    # 2. Desktop & Documents
    user_home = os.path.expanduser("~")
    if user_home:
        candidates.append(os.path.join(user_home, "Desktop", "CPro Case Suite Output"))
        candidates.append(os.path.join(user_home, "Documents", "CPro Case Suite Output"))

    # 3. Local AppData
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(os.path.join(local_app_data, "CProCaseSuite", "Output"))

    # 4. Temp folder
    candidates.append(os.path.join(tempfile.gettempdir(), "CProCaseSuite", "Output"))
    
    # 5. Current working directory
    candidates.append(os.getcwd())

    writable = []
    seen = set()
    for folder in candidates:
        folder = os.path.abspath(folder)
        normalized = os.path.normcase(folder)
        if normalized in seen:
            continue
        seen.add(normalized)
        try:
            os.makedirs(folder, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                prefix=".cpro_write_test_", suffix=".tmp", dir=folder, delete=True
            ) as probe:
                probe.write(b"ok")
                probe.flush()
            writable.append(folder)
        except (OSError, PermissionError):
            continue
    return writable


WRITABLE_OUTPUT_DIRS = _find_writable_output_dirs()
if not WRITABLE_OUTPUT_DIRS:
    # Fallback if nothing else works
    OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "CProCaseSuite", "Output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    WRITABLE_OUTPUT_DIRS = [OUTPUT_DIR]
else:
    OUTPUT_DIR = WRITABLE_OUTPUT_DIRS[0]

EXTRACTED_EXCEL = os.path.join(OUTPUT_DIR, "extracted_cases.xlsx")
REPORT_EXCEL = os.path.join(OUTPUT_DIR, "cpro_availability_report.xlsx")


# ==============================================================================
# OPTIMIZED PDF EXTRACTION ENGINE (MULTIPLE CASES PER PDF)
# ==============================================================================
class PDFExtractorEngine:
    def __init__(self):
        self.ocr = None

    def _ensure_ocr(self):
        if pymupdf is None:
            raise RuntimeError("Missing PDF packages. Please run: python -m pip install pymupdf")

    def extract_pdf(self, pdf_path):
        self._ensure_ocr()
        doc = pymupdf.open(pdf_path)
        full_text = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text("text").strip()
            # Fallback to OCR only if the page is a scanned image (very little text)
            if len(text) < 30 and RapidOCR is not None:
                if self.ocr is None:
                    self.ocr = RapidOCR()
                pix = page.get_pixmap(dpi=200)  # Reduced DPI for speed
                img_bytes = pix.tobytes("png")
                ocr_result, _ = self.ocr(img_bytes)
                if ocr_result:
                    text = "\n".join([item[1] for item in ocr_result])
            full_text.append(text)

        doc.close()
        combined_text = "\n".join(full_text)
        return self.parse_cases_from_text(combined_text, os.path.basename(pdf_path))

    @staticmethod
    def parse_cases_from_text(text, filename):
        cases = []
        seen_acc = set()
        
        # Matches Account No formats like: UPL-2611339-39596078602 or MG-2611300-26791032594
        acc_pattern = re.compile(r'\b([A-Z]{2,5})\s*-\s*(\d{7,8})\s*-\s*(\d{10,14})\b', re.IGNORECASE)
        
        for match in acc_pattern.finditer(text):
            prefix = match.group(1).upper()
            mid = match.group(2)
            suffix = match.group(3)
            acc_no = f"{prefix}-{mid}-{suffix}"
            
            # Deduplicate: If we already found this account, skip to next
            if acc_no in seen_acc:
                continue
                
            # Grab up to 150 characters immediately following the Account No
            start_idx = match.end()
            snippet = text[start_idx:start_idx+150]
            
            # Clean snippet
            snippet = re.sub(r'\s+', ' ', snippet)
            # Remove leading numbers, currency, punctuation (e.g. table cells "0.00")
            snippet = re.sub(r'^[\d\.\,\s\-\|]+', '', snippet)
            
            # Extract the Name (heuristic: uppercase letters, spaces, @, /, etc.)
            # Stops before numbers or known C-PRO/Invoice keywords
            name_match = re.search(r'^([A-Z][A-Z\s@\/\.\']{2,80})', snippet, re.IGNORECASE)
            
            name = ""
            if name_match:
                raw_name = name_match.group(1).strip()
                # Remove trailing noise
                raw_name = re.sub(r'\s+(SV-8|Legal fee|Notification|Letter|Unsealed|Writ|Summons|Billing|History|All|Pending|Approval|Approved|Rejected|Description).*$', '', raw_name, flags=re.IGNORECASE).strip()
                if len(raw_name) > 2 and not any(k in raw_name.upper() for k in ["NO.", "DATE", "PAGE"]):
                    name = raw_name
            
            if name:
                # Clean short name for C-PRO search (cuts off before BIN, BINTI, etc.)
                short_name = re.split(r'\s+(?:BIN|BINTI|B\.|BT\.|A/L|A/P|@)\s+', name, maxsplit=1, flags=re.IGNORECASE)[0].strip()
                
                cases.append({
                    "Source File": filename,
                    "Account No": acc_no,
                    "Search Name": short_name,
                    "Full Borrower Name": name,
                    "Availability Status": "Pending",
                    "Found In Account(s)": "-",
                    "Found In Section(s)": "-",
                    "C-PRO Row Details": "-"
                })
                seen_acc.add(acc_no)
                
        return cases


# ==============================================================================
# UNIFIED APPLE-STYLE DESKTOP APPLICATION
# ==============================================================================
class CProSuiteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("C-PRO Case Suite")
        self.root.geometry("1280x820")
        self.root.minsize(1080, 700)
        self.root.configure(bg=THEME["bg"])

        self.pdf_engine = PDFExtractorEngine()
        self.cases_data = []
        self.is_running = False
        self.stop_requested = False
        self.last_extracted_path = EXTRACTED_EXCEL
        self.last_report_path = REPORT_EXCEL
        self._last_searched_column = None

        self.pwd_tol1 = tk.StringVar()
        self.pwd_tol2 = tk.StringVar()
        self.pwd_melaka = tk.StringVar()
        self.pwd_melaka01 = tk.StringVar()
        self.pwd_putrajaya = tk.StringVar()
        self.show_pwd = tk.BooleanVar(value=False)

        self.enable_tol1 = tk.BooleanVar(value=True)
        self.enable_tol2 = tk.BooleanVar(value=False)
        self.enable_melaka = tk.BooleanVar(value=True)
        self.enable_melaka01 = tk.BooleanVar(value=False)
        self.enable_putrajaya = tk.BooleanVar(value=True)

        self.sec_in_progress = tk.BooleanVar(value=True)
        self.sec_withhold = tk.BooleanVar(value=True)
        self.sec_withdrawn = tk.BooleanVar(value=True)

        self.search_mode = tk.StringVar(value="ACCOUNT_THEN_NAME")
        self.show_browser = tk.BooleanVar(value=True)

        self._configure_styles()
        self._build_layout()
        self.log(f"Output folder: {OUTPUT_DIR}")
        self._auto_load_existing_excel()

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Apple.Treeview", background=THEME["card"], foreground=THEME["text"], fieldbackground=THEME["card"], rowheight=34, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Apple.Treeview.Heading", background="#F9F9FB", foreground=THEME["subtext"], relief="flat", borderwidth=0, padding=(10, 8), font=("Segoe UI Semibold", 9))
        style.map("Apple.Treeview", background=[("selected", "#D8EAFE")], foreground=[("selected", THEME["text"])])
        style.configure("Apple.Horizontal.TProgressbar", troughcolor="#E5E5EA", background=THEME["blue"], bordercolor="#E5E5EA", lightcolor=THEME["blue"], darkcolor=THEME["blue"], thickness=6)

    def _create_card(self, parent, padx=18, pady=16):
        outer = tk.Frame(parent, bg=THEME["border"], padx=1, pady=1)
        inner = tk.Frame(outer, bg=THEME["card"], padx=padx, pady=pady)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def _create_button(self, parent, text, command, bg, fg="#FFFFFF", hover_bg=None, width=None):
        btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg, activebackground=hover_bg or bg, activeforeground=fg, relief="flat", bd=0, padx=14, pady=7, cursor="hand2", font=("Segoe UI Semibold", 9))
        if width: btn.config(width=width)
        if hover_bg:
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg) if btn["state"] != "disabled" else None)
            btn.bind("<Leave>", lambda e: btn.config(bg=bg) if btn["state"] != "disabled" else None)
        return btn

    def _build_layout(self):
        header = tk.Frame(self.root, bg=THEME["bg"], padx=24, pady=16)
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=THEME["bg"])
        title_box.pack(side="left")
        tk.Label(title_box, text="C-PRO Case Suite", font=("Segoe UI", 18, "bold"), bg=THEME["bg"], fg=THEME["text"]).pack(anchor="w")
        tk.Label(title_box, text="Extracts multiple Account Numbers and Names from PDF files automatically.", font=("Segoe UI", 9), bg=THEME["bg"], fg=THEME["subtext"]).pack(anchor="w", pady=(2, 0))

        status_frame = tk.Frame(header, bg=THEME["bg"])
        status_frame.pack(side="right", anchor="center")
        self.status_badge = tk.Label(status_frame, text="●  Ready", font=("Segoe UI Semibold", 9), bg="#E8E8ED", fg=THEME["text"], padx=12, pady=5)
        self.status_badge.pack(side="right")

        workspace = tk.Frame(self.root, bg=THEME["bg"], padx=24, pady=4)
        workspace.pack(fill="both", expand=True)
        sidebar = tk.Frame(workspace, bg=THEME["bg"], width=330)
        sidebar.pack(side="left", fill="y", padx=(0, 16))
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        content = tk.Frame(workspace, bg=THEME["bg"])
        content.pack(side="left", fill="both", expand=True)
        self._build_content_area(content)

    def _build_sidebar(self, sidebar):
        card1_outer, card1 = self._create_card(sidebar)
        card1_outer.pack(fill="x", pady=(0, 14))
        tk.Label(card1, text="SESSION PASSWORDS", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w")
        tk.Label(card1, text="Passwords stay in memory until the app closes.", font=("Segoe UI", 8), bg=THEME["card"], fg=THEME["subtext"], wraplength=280, justify="left").pack(anchor="w", pady=(2, 10))

        self.pwd_entries = []
        accounts_ui = [
            ("TOL1 (shared group with TOL2)", self.pwd_tol1, self.enable_tol1),
            ("TOL2 (alternative to TOL1)", self.pwd_tol2, self.enable_tol2),
            ("TOLMELAKA (shared group with 01)", self.pwd_melaka, self.enable_melaka),
            ("TOLMELAKA01 (alternative to Melaka)", self.pwd_melaka01, self.enable_melaka01),
            ("TOLPUTRAJAYA", self.pwd_putrajaya, self.enable_putrajaya),
        ]
        for label_text, pwd_var, enable_var in accounts_ui:
            row_header = tk.Frame(card1, bg=THEME["card"])
            row_header.pack(fill="x", pady=(4, 2))
            tk.Checkbutton(row_header, text=label_text, variable=enable_var, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI Semibold", 9)).pack(side="left")
            entry_border = tk.Frame(card1, bg=THEME["border"], padx=1, pady=1)
            entry_border.pack(fill="x", pady=(0, 6))
            ent = tk.Entry(entry_border, textvariable=pwd_var, show="•", bg=THEME["input_bg"], fg=THEME["text"], relief="flat", bd=5, font=("Segoe UI", 9))
            ent.pack(fill="x")
            self.pwd_entries.append(ent)

        def toggle_pwd_visibility():
            char = "" if self.show_pwd.get() else "•"
            for e in self.pwd_entries: e.config(show=char)

        tk.Checkbutton(card1, text="Show Passwords", variable=self.show_pwd, command=toggle_pwd_visibility, bg=THEME["card"], fg=THEME["subtext"], activebackground=THEME["card"], font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        card2_outer, card2 = self._create_card(sidebar)
        card2_outer.pack(fill="x", pady=(0, 14))
        tk.Label(card2, text="SECTIONS TO CHECK", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 6))
        for sec_label, sec_var in [("Case In Progress", self.sec_in_progress), ("Case Withhold", self.sec_withhold), ("Case Withdrawn", self.sec_withdrawn)]:
            tk.Checkbutton(card2, text=sec_label, variable=sec_var, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9)).pack(anchor="w", pady=1)

        tk.Frame(card2, bg=THEME["border"], height=1).pack(fill="x", pady=10)
        tk.Label(card2, text="SEARCH STRATEGY", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 6))
        modes = [("Account No → Fallback to Name", "ACCOUNT_THEN_NAME"), ("Account No Only", "ACCOUNT_ONLY"), ("Borrower Name Only", "NAME_ONLY")]
        for text, val in modes:
            tk.Radiobutton(card2, text=text, variable=self.search_mode, value=val, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9)).pack(anchor="w", pady=1)

        tk.Frame(card2, bg=THEME["border"], height=1).pack(fill="x", pady=10)
        tk.Checkbutton(card2, text="Show Live Chrome Browser", variable=self.show_browser, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9)).pack(anchor="w")

        card3_outer, card3 = self._create_card(sidebar)
        card3_outer.pack(fill="x")
        tk.Label(card3, text="EXCEL FILES", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 8))
        
        self._create_button(card3, "Open Extracted Cases Excel", lambda: self._open_file(getattr(self, "last_extracted_path", EXTRACTED_EXCEL)), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x", pady=(0, 6))
        self._create_button(card3, "Open Final C-PRO Report Excel", lambda: self._open_file(getattr(self, "last_report_path", REPORT_EXCEL)), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x", pady=(0, 6))
        
        # NEW BUTTON TO OPEN OUTPUT FOLDER
        self._create_button(card3, "📂 Open Output Folder", lambda: self._open_folder(OUTPUT_DIR), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x")

    def _build_content_area(self, content):
        bar_outer, bar = self._create_card(content, padx=14, pady=12)
        bar_outer.pack(fill="x", pady=(0, 12))

        self.btn_upload_pdf = self._create_button(bar, "＋ Upload & Extract PDFs", self.on_upload_pdfs, bg=THEME["blue"], hover_bg=THEME["blue_hover"])
        self.btn_upload_pdf.pack(side="left", padx=(0, 8))
        self.btn_add_manual = self._create_button(bar, "＋ Add Manual Row", self.on_add_manual_row, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"])
        self.btn_add_manual.pack(side="left", padx=(0, 8))
        self.btn_load_excel = self._create_button(bar, "Import Excel", self.on_import_excel, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"])
        self.btn_load_excel.pack(side="left", padx=(0, 8))
        self.btn_clear = self._create_button(bar, "Clear List", self.on_clear_list, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"])
        self.btn_clear.pack(side="left")

        self.btn_stop = self._create_button(bar, "■ Stop", self.on_stop_automation, bg=THEME["red"], hover_bg="#D70015")
        self.btn_stop.pack(side="right")
        self.btn_stop.config(state="disabled")
        self.btn_run_cpro = self._create_button(bar, "▶ Start C-PRO Check", self.on_start_cpro_check, bg=THEME["green"], hover_bg=THEME["green_hover"])
        self.btn_run_cpro.pack(side="right", padx=(0, 8))

        table_outer, table_card = self._create_card(content, padx=0, pady=0)
        table_outer.pack(fill="both", expand=True, pady=(0, 12))
        table_header = tk.Frame(table_card, bg=THEME["card"], padx=16, pady=10)
        table_header.pack(fill="x")
        self.lbl_case_count = tk.Label(table_header, text="Cases Queue (0 records)", font=("Segoe UI Semibold", 10), bg=THEME["card"], fg=THEME["text"])
        self.lbl_case_count.pack(side="left")
        tk.Label(table_header, text="Tip: Double-click any row to edit before checking", font=("Segoe UI", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(side="right")
        tk.Frame(table_card, bg=THEME["border"], height=1).pack(fill="x")

        tree_frame = tk.Frame(table_card, bg=THEME["card"])
        tree_frame.pack(fill="both", expand=True)

        cols = ("no", "acc_no", "search_name", "status", "found_acc", "found_sec")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", style="Apple.Treeview", selectmode="extended")
        headings = [
            ("no", "#", 50, "center"),
            ("acc_no", "Account No", 200, "w"),
            ("search_name", "Borrower Name", 250, "w"),
            ("status", "Status", 110, "center"),
            ("found_acc", "Found In Account(s)", 150, "w"),
            ("found_sec", "Found In Section(s)", 180, "w"),
        ]
        for col_id, title, width, anchor in headings:
            self.tree.heading(col_id, text=title, anchor=anchor)
            self.tree.column(col_id, width=width, anchor=anchor)

        self.tree.tag_configure("AVAILABLE", background=THEME["row_found"])
        self.tree.tag_configure("NOT FOUND", background=THEME["row_missing"])
        self.tree.tag_configure("CHECKING", background=THEME["row_active"])
        self.tree.tag_configure("CHECK ERROR", background=THEME["row_error"])

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.on_double_click_row)
        self.tree.bind("<Delete>", lambda e: self.on_delete_selected())

        log_outer, log_card = self._create_card(content, padx=14, pady=10)
        log_outer.pack(fill="x")
        prog_row = tk.Frame(log_card, bg=THEME["card"])
        prog_row.pack(fill="x", pady=(0, 6))
        tk.Label(prog_row, text="LIVE ACTIVITY LOG", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(side="left")
        self.lbl_progress = tk.Label(prog_row, text="0%", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["blue"])
        self.lbl_progress.pack(side="right")

        self.progress = ttk.Progressbar(log_card, orient="horizontal", mode="determinate", style="Apple.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 8))
        self.log_text = tk.Text(log_card, height=6, bg="#F9F9FB", fg=THEME["text"], relief="flat", bd=6, font=("Consolas", 9), state="disabled")
        self.log_text.pack(fill="x")

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", line)
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _append)

    def set_status(self, text, bg="#E8E8ED", fg=THEME["text"]):
        self.root.after(0, lambda: self.status_badge.config(text=f"●  {text}", bg=bg, fg=fg))

    def set_progress(self, current, total):
        pct = int((current / total) * 100) if total > 0 else 0
        def _upd():
            self.progress["value"] = pct
            self.lbl_progress.config(text=f"{pct}% ({current}/{total})")
        self.root.after(0, _upd)

    def refresh_table(self, active_idx=None):
        def _rebuild():
            self.tree.delete(*self.tree.get_children())
            for i, row in enumerate(self.cases_data):
                status = row.get("Availability Status", "Pending")
                tag = ""
                if i == active_idx: tag = "CHECKING"
                elif status == "AVAILABLE": tag = "AVAILABLE"
                elif status == "NOT FOUND": tag = "NOT FOUND"
                elif status == "CHECK ERROR": tag = "CHECK ERROR"

                self.tree.insert("", "end", iid=str(i), values=(
                    i + 1,
                    row.get("Account No", ""),
                    row.get("Search Name", ""),
                    status,
                    row.get("Found In Account(s)", "-"),
                    row.get("Found In Section(s)", "-"),
                ), tags=(tag,) if tag else ())
            self.lbl_case_count.config(text=f"Cases Queue ({len(self.cases_data)} records)")
        self.root.after(0, _rebuild)

    def _auto_load_existing_excel(self):
        candidates = [EXTRACTED_EXCEL, os.path.join(os.getcwd(), "extracted_cases.xlsx"), os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_cases.xlsx")]
        source_path = next((path for path in candidates if os.path.isfile(path)), None)
        if not source_path: return
        try:
            df = pd.read_excel(source_path, dtype=str).fillna("")
            for rec in df.to_dict("records"):
                rec.setdefault("Availability Status", "Pending")
                rec.setdefault("Found In Account(s)", "-")
                rec.setdefault("Found In Section(s)", "-")
                rec.setdefault("C-PRO Row Details", "-")
                self.cases_data.append(rec)
            self.refresh_table()
            self.last_extracted_path = source_path
            self.log(f"Loaded {len(self.cases_data)} existing case(s) from '{source_path}'.")
        except Exception as e:
            self.log(f"Could not auto-load '{source_path}': {e}")

    @staticmethod
    def _directory_is_writable(folder):
        try:
            os.makedirs(folder, exist_ok=True)
            with tempfile.NamedTemporaryFile(prefix=".cpro_write_test_", suffix=".tmp", dir=folder, delete=True) as probe:
                probe.write(b"ok"); probe.flush()
            return True
        except OSError:
            return False

    def _write_with_fallback(self, preferred_path, write_callback, description):
        try:
            write_callback(preferred_path)
            return preferred_path
        except PermissionError as exc:
            first_error_text = str(exc)
            self.log(f"[!] Cannot write '{preferred_path}' (locked). Trying fallback folder.")
        base_name = os.path.basename(preferred_path)
        base, ext = os.path.splitext(base_name)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        candidate_dirs = [os.path.dirname(os.path.abspath(preferred_path))] + list(WRITABLE_OUTPUT_DIRS)
        checked = set()
        last_error = None
        for folder in candidate_dirs:
            folder = os.path.abspath(folder or os.getcwd())
            normalized = os.path.normcase(folder)
            if normalized in checked: continue
            checked.add(normalized)
            if not self._directory_is_writable(folder): continue
            for suffix in range(10):
                extra = f"_{suffix + 1}" if suffix else ""
                candidate = os.path.join(folder, f"{base}_{timestamp}{extra}{ext}")
                if os.path.exists(candidate): continue
                try:
                    write_callback(candidate)
                    self.log(f"[!] {description} saved to fallback: '{candidate}'.")
                    return candidate
                except Exception as e:
                    last_error = e
                    break
        raise PermissionError(f"Could not save {description}. Last error: {last_error or first_error_text}")

    def _save_extracted_excel(self):
        if not self.cases_data: return None
        df = pd.DataFrame(self.cases_data)
        export_cols = [c for c in ["Source File", "Account No", "Search Name", "Full Borrower Name"] if c in df.columns]

        def _write_workbook(path):
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df[export_cols].to_excel(writer, index=False, sheet_name="To_Check")
                ws = writer.sheets["To_Check"]
                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col) + 4
                    ws.column_dimensions[col[0].column_letter].width = max(max_len, 15)
        try:
            saved_path = self._write_with_fallback(EXTRACTED_EXCEL, _write_workbook, "extracted cases")
            self.last_extracted_path = saved_path
            return saved_path
        except Exception as e:
            self.log(f"[ERROR] Could not save extracted cases: {e}")
            return None

    def _open_file(self, filepath):
        if not os.path.exists(filepath):
            messagebox.showinfo("File Not Found", f"'{filepath}' has not been generated yet.")
            return
        try:
            if sys.platform == "win32": os.startfile(os.path.abspath(filepath))
            elif sys.platform == "darwin": subprocess.Popen(["open", os.path.abspath(filepath)])
            else: subprocess.Popen(["xdg-open", os.path.abspath(filepath)])
        except Exception as e:
            messagebox.showerror("Error Opening File", str(e))

    def _open_folder(self, folder):
        if not os.path.exists(folder):
            messagebox.showinfo("Folder Not Found", f"The folder '{folder}' does not exist yet.")
            return
        try:
            if sys.platform == "win32": os.startfile(os.path.abspath(folder))
            elif sys.platform == "darwin": subprocess.Popen(["open", os.path.abspath(folder)])
            else: subprocess.Popen(["xdg-open", os.path.abspath(folder)])
        except Exception as e:
            messagebox.showerror("Error Opening Folder", str(e))

    def on_upload_pdfs(self):
        if self.is_running: return
        pdf_files = filedialog.askopenfilenames(parent=self.root, title="Select Invoice / Case PDF Files", filetypes=[("PDF Files", "*.pdf")])
        if not pdf_files: return

        self.is_running = True
        self.btn_upload_pdf.config(state="disabled")
        self.btn_run_cpro.config(state="disabled")
        self.set_status("Extracting PDFs...", bg="#EBF5FF", fg=THEME["blue"])

        def _worker():
            total = len(pdf_files)
            for idx, path in enumerate(pdf_files, start=1):
                fname = os.path.basename(path)
                self.log(f"Scanning ({idx}/{total}): {fname} ...")
                try:
                    extracted_cases = self.pdf_engine.extract_pdf(path)
                    if not extracted_cases:
                        self.log(f"  [!] No Account Numbers found in {fname}")
                    for case in extracted_cases:
                        self.cases_data.append(case)
                        self.log(f"  -> {case['Account No']} | {case['Search Name']}")
                    self.refresh_table()
                except Exception as e:
                    self.log(f"  [ERROR] Failed to extract {fname}: {e}")
                self.set_progress(idx, total)

            saved_path = self._save_extracted_excel()
            if saved_path:
                self.log(f"Saved {len(self.cases_data)} total record(s) to '{saved_path}'.")
            self.set_status("Ready", bg="#E8E8ED", fg=THEME["text"])
            self.is_running = False
            self.root.after(0, lambda: self.btn_upload_pdf.config(state="normal"))
            self.root.after(0, lambda: self.btn_run_cpro.config(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def on_import_excel(self):
        if self.is_running: return
        path = filedialog.askopenfilename(parent=self.root, title="Select Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path: return
        try:
            df = pd.read_excel(path, dtype=str).fillna("")
            for rec in df.to_dict("records"):
                self.cases_data.append({
                    "Source File": rec.get("Source File", os.path.basename(path)),
                    "Account No": rec.get("Account No", ""),
                    "Search Name": rec.get("Search Name", rec.get("Borrower Name", "")),
                    "Full Borrower Name": rec.get("Full Borrower Name", rec.get("Search Name", "")),
                    "Availability Status": "Pending",
                    "Found In Account(s)": "-",
                    "Found In Section(s)": "-",
                    "C-PRO Row Details": "-"
                })
            self.refresh_table()
            self._save_extracted_excel()
            self.log(f"Imported {len(df)} row(s) from {os.path.basename(path)}.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def on_add_manual_row(self):
        if self.is_running: return
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Case Manually")
        dialog.geometry("380x230")
        dialog.configure(bg=THEME["card"])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = tk.Frame(dialog, bg=THEME["card"], padx=20, pady=16)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Account No:", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI Semibold", 9)).pack(anchor="w")
        ent_acc = tk.Entry(frame, bg=THEME["input_bg"], relief="flat", bd=5, font=("Segoe UI", 9))
        ent_acc.pack(fill="x", pady=(2, 10))

        tk.Label(frame, text="Borrower Name:", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI Semibold", 9)).pack(anchor="w")
        ent_name = tk.Entry(frame, bg=THEME["input_bg"], relief="flat", bd=5, font=("Segoe UI", 9))
        ent_name.pack(fill="x", pady=(2, 14))

        def _save():
            acc = ent_acc.get().strip()
            name = ent_name.get().strip()
            if not acc and not name:
                messagebox.showwarning("Input Required", "Please enter an Account No or Search Name.", parent=dialog)
                return
            self.cases_data.append({
                "Source File": "Manual Entry",
                "Account No": acc,
                "Search Name": name,
                "Full Borrower Name": name,
                "Availability Status": "Pending",
                "Found In Account(s)": "-",
                "Found In Section(s)": "-",
                "C-PRO Row Details": "-"
            })
            self.refresh_table()
            self._save_extracted_excel()
            dialog.destroy()

        self._create_button(frame, "Add to Queue", _save, bg=THEME["blue"], hover_bg=THEME["blue_hover"]).pack(fill="x")

    def on_double_click_row(self, event):
        if self.is_running: return
        sel = self.tree.selection()
        if not sel: return
        idx = int(sel[0])
        row_data = self.cases_data[idx]

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Row #{idx + 1}")
        dialog.geometry("400x250")
        dialog.configure(bg=THEME["card"])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = tk.Frame(dialog, bg=THEME["card"], padx=20, pady=16)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Account No:", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI Semibold", 9)).pack(anchor="w")
        ent_acc = tk.Entry(frame, bg=THEME["input_bg"], relief="flat", bd=5, font=("Segoe UI", 9))
        ent_acc.insert(0, row_data.get("Account No", ""))
        ent_acc.pack(fill="x", pady=(2, 10))

        tk.Label(frame, text="Search Name:", bg=THEME["card"], fg=THEME["text"], font=("Segoe UI Semibold", 9)).pack(anchor="w")
        ent_name = tk.Entry(frame, bg=THEME["input_bg"], relief="flat", bd=5, font=("Segoe UI", 9))
        ent_name.insert(0, row_data.get("Search Name", ""))
        ent_name.pack(fill="x", pady=(2, 14))

        def _update():
            row_data["Account No"] = ent_acc.get().strip()
            row_data["Search Name"] = ent_name.get().strip()
            self.refresh_table()
            self._save_extracted_excel()
            dialog.destroy()

        self._create_button(frame, "Save Changes", _update, bg=THEME["blue"], hover_bg=THEME["blue_hover"]).pack(fill="x")

    def on_delete_selected(self):
        if self.is_running: return
        sel = self.tree.selection()
        if not sel: return
        indices = sorted([int(i) for i in sel], reverse=True)
        for i in indices: del self.cases_data[i]
        self.refresh_table()
        self._save_extracted_excel()

    def on_clear_list(self):
        if self.is_running or not self.cases_data: return
        if messagebox.askyesno("Clear Queue", "Remove all cases from the current list?"):
            self.cases_data.clear()
            self.refresh_table()

    def on_stop_automation(self):
        if self.is_running:
            self.stop_requested = True
            self.log("[!] Stop requested by user. Finishing current action...")
            self.set_status("Stopping...", bg="#FFF2F2", fg=THEME["red"])

    def on_start_cpro_check(self):
        if self.is_running: return
        if not self.cases_data:
            messagebox.showwarning("Empty Queue", "Please upload PDFs or add cases before starting the C-PRO check.")
            return
        if sync_playwright is None:
            messagebox.showerror("Playwright Missing", "Playwright is not installed. Run:\npython -m pip install playwright\npython -m playwright install chromium")
            return

        active_accounts = []
        def add_one_from_group(primary_enabled, primary_name, primary_password, alias_enabled, alias_name, alias_password, group_label):
            chosen = None
            if primary_enabled.get(): chosen = (primary_name, primary_password)
            elif alias_enabled.get(): chosen = (alias_name, alias_password)
            if chosen:
                username, password_var = chosen
                active_accounts.append({"group_label": group_label, "username": username, "password": password_var.get().strip(), "password_var": password_var})

        add_one_from_group(self.enable_tol1, "TOL1", self.pwd_tol1, self.enable_tol2, "TOL2", self.pwd_tol2, "TOL1 / TOL2")
        add_one_from_group(self.enable_melaka, "TOLMELAKA", self.pwd_melaka, self.enable_melaka01, "TOLMELAKA01", self.pwd_melaka01, "TOLMELAKA / TOLMELAKA01")
        if self.enable_putrajaya.get():
            active_accounts.append({"group_label": "TOLPUTRAJAYA", "username": "TOLPUTRAJAYA", "password": self.pwd_putrajaya.get().strip(), "password_var": self.pwd_putrajaya})

        if not active_accounts:
            messagebox.showwarning("No Accounts Selected", "Please select at least one C-PRO account in the left sidebar.")
            return

        for acc in active_accounts:
            if not acc["password"]:
                pwd = simpledialog.askstring("Password Required", f"Enter C-PRO password for {acc['username']}:", show="•", parent=self.root)
                if not pwd: return
                acc["password"] = pwd.strip()
                acc["password_var"].set(acc["password"])

        active_sections = []
        if self.sec_in_progress.get(): active_sections.append("Case In Progress")
        if self.sec_withhold.get(): active_sections.append("Case Withhold")
        if self.sec_withdrawn.get(): active_sections.append("Case Withdrawn")

        if not active_sections:
            messagebox.showwarning("No Sections Selected", "Please select at least one section to check.")
            return

        self.is_running = True
        self.stop_requested = False
        self.btn_run_cpro.config(state="disabled")
        self.btn_upload_pdf.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.set_status("Checking C-PRO...", bg="#E8F8EE", fg=THEME["green"])

        threading.Thread(target=self._cpro_worker, args=(active_accounts, active_sections), daemon=True).start()

    @staticmethod
    def _is_kicked_or_unauthorized(page):
        try:
            if "login" in page.url.lower() or "auth" in page.url.lower(): return True
            if page.locator("button:has-text('Sign In')").is_visible(): return True
            toast_error = page.locator("text=/Error! Unauthenticated|Unauthenticated|Session expired|Logged in from another/i")
            if toast_error.count() > 0 and toast_error.first.is_visible(): return True
            unauth = page.locator("text=/unauthorized|session expired|logged in from another|invalid session|please login again/i")
            if unauth.count() > 0 and unauth.first.is_visible(): return True
        except Exception: return True
        return False

    @staticmethod
    def _wait_for_network_quiet(page, timeout=5000):
        try: page.wait_for_load_state("networkidle", timeout=timeout)
        except Exception: pass

    @staticmethod
    def _section_tab_candidates(page, section):
        exact = re.compile(rf"^\s*{re.escape(section)}(?:\s+\d+)?\s*$", re.IGNORECASE)
        selectors = ("[role='tab']", ".nav-tabs a", ".nav-tabs button", ".nav-link", ".dx-tab", "button", "a", "li")
        seen = set()
        for selector in selectors:
            locator = page.locator(selector)
            try: count = min(locator.count(), 200)
            except Exception: continue
            for index in range(count):
                item = locator.nth(index)
                try:
                    if not item.is_visible(): continue
                    text = re.sub(r"\s+", " ", item.inner_text(timeout=500)).strip()
                    if not exact.match(text): continue
                    key = item.evaluate("el => el.tagName + '|' + el.className + '|' + el.textContent")
                    if key in seen: continue
                    seen.add(key)
                    yield item
                except Exception: continue
        try:
            labels = page.get_by_text(section, exact=True)
            for index in range(min(labels.count(), 50)):
                item = labels.nth(index)
                try:
                    if item.is_visible(): yield item
                except Exception: continue
        except Exception: return

    def _wait_for_section_table(self, page, section, timeout_ms=30000, action=None):
        started = time.monotonic()
        deadline = started + timeout_ms / 1000.0
        last_signature = None
        stable_since = None
        last_network = started
        pending = set()
        saw_xhr = False

        def on_request(request):
            nonlocal saw_xhr, last_network
            try:
                if request.resource_type in ("xhr", "fetch"):
                    saw_xhr = True
                    pending.add(id(getattr(request, "_impl_obj", request)))
                    last_network = time.monotonic()
            except Exception: pass

        def on_done(request):
            nonlocal last_network
            try:
                if request.resource_type in ("xhr", "fetch"):
                    pending.discard(id(getattr(request, "_impl_obj", request)))
                    last_network = time.monotonic()
            except Exception: pass

        page.on("request", on_request)
        page.on("requestfinished", on_done)
        page.on("requestfailed", on_done)
        try:
            if action is not None: action()
            while time.monotonic() < deadline:
                if self.stop_requested: raise GridLoadTimeoutError(f"Stopped while waiting for {section}.")
                if self._is_kicked_or_unauthorized(page): raise InvalidCredentialsError(f"Session not authenticated while loading {section}.")

                signature, loading = self._grid_snapshot(page)
                try:
                    headers_ready = page.locator(".dx-datagrid-headers:visible, tr.dx-header-row:visible, [role='columnheader']:visible").count() > 0
                    filter_ready = page.locator("tr.dx-datagrid-filter-row:visible").count() > 0
                    rows_ready = page.locator("tr.dx-data-row:visible").count() > 0
                    empty_ready = page.locator(".dx-datagrid-nodata:visible, .dx-datagrid-rowsview .dx-empty-message:visible").count() > 0
                except Exception: headers_ready = filter_ready = rows_ready = empty_ready = False

                table_ready = headers_ready and filter_ready and (rows_ready or empty_ready)
                if signature != last_signature:
                    last_signature = signature
                    stable_since = time.monotonic()
                elif stable_since is None:
                    stable_since = time.monotonic()

                elapsed = time.monotonic() - started
                stable_for = time.monotonic() - (stable_since or time.monotonic())
                network_quiet = not pending and (time.monotonic() - last_network) >= 1.0
                min_wait = 1.5 if saw_xhr else 4.0

                if (table_ready and not loading and elapsed >= min_wait and stable_for >= 1.25 and (network_quiet if saw_xhr else True)):
                    self.log(f"[WAIT] {section} table ready after {elapsed:.1f}s")
                    return
                page.wait_for_timeout(250)

            raise GridLoadTimeoutError(f"Timed out waiting for {section} table.")
        finally:
            for event, handler in (("request", on_request), ("requestfinished", on_done), ("requestfailed", on_done)):
                try: page.remove_listener(event, handler)
                except Exception: pass

    def _switch_section_and_wait(self, page, section):
        tab = next(self._section_tab_candidates(page, section), None)
        if tab is None: raise GridLoadTimeoutError(f"Could not find a visible '{section}' tab.")
        self.log(f"[WAIT] Clicking section tab '{section}'...")
        self._last_searched_column = None
        self._wait_for_section_table(page, section, action=lambda: tab.click(timeout=8000))

    def _prompt_for_correct_password(self, acc, reason):
        result = {"password": None}
        ready = threading.Event()
        def ask():
            try:
                result["password"] = simpledialog.askstring("C-PRO Password Rejected", f"C-PRO rejected the password for {acc['username']}.\n{reason}\n\nEnter correct password:", show="•", parent=self.root)
            finally: ready.set()
        self.root.after(0, ask)
        while not ready.wait(0.2):
            if self.stop_requested: return None
        password = (result["password"] or "").strip()
        if password:
            acc["password"] = password
            if acc.get("password_var") is not None: self.root.after(0, lambda v=acc["password_var"], p=password: v.set(p))
        return password or None

    @staticmethod
    def _login_rejection_text(page):
        try: body = page.locator("body").inner_text(timeout=1500).casefold()
        except Exception: body = ""
        keywords = ("invalid password", "incorrect password", "invalid username", "invalid credentials", "login failed", "authentication failed", "password is incorrect", "wrong password")
        for keyword in keywords:
            if keyword in body: return keyword
        return ""

    def _login_and_restore_section(self, page, acc, section=None):
        username = acc["username"]
        password = acc["password"]
        self.log(f"[AUTH] Logging into {username}...")
        self.set_status(f"Authenticating {username}...", bg="#EBF5FF", fg=THEME["blue"])

        try:
            close_btns = page.locator(".toast-close-button, .close, button[aria-label='Close'], .swal2-close")
            for i in range(min(close_btns.count(), 20)):
                if close_btns.nth(i).is_visible(): close_btns.nth(i).click(timeout=500)
        except Exception: pass

        for login_attempt in range(2):
            page.goto(CASE_LISTING_URL, wait_until="domcontentloaded", timeout=20000)
            self._wait_for_network_quiet(page, timeout=3500)
            page.wait_for_timeout(300)

            sign_in = page.locator("button:has-text('Sign In')")
            user_input = page.locator("input[placeholder*='Username' i], input[name*='user' i], input[type='text']").first
            if sign_in.is_visible():
                if not user_input.is_visible(): user_input = page.locator("input:not([type]), input[type='email']").first
                pwd_input = page.locator("input[type='password']").first
                user_input.wait_for(state="visible", timeout=5000)
                pwd_input.wait_for(state="visible", timeout=5000)
                user_input.fill(username)
                pwd_input.fill(password)
                sign_in.click()

                deadline = time.monotonic() + 7.0
                rejection = ""
                while time.monotonic() < deadline:
                    rejection = self._login_rejection_text(page)
                    if rejection: break
                    if "case-listing" in page.url.lower() and not page.locator("input[type='password']").is_visible(): break
                    page.wait_for_timeout(250)

                rejection = rejection or self._login_rejection_text(page)
                if not rejection:
                    try:
                        if page.locator("button:has-text('Sign In')").is_visible() and page.locator("input[type='password']").is_visible():
                            rejection = "C-PRO remained on the sign-in form"
                    except Exception: pass

                if rejection:
                    self.log(f"[AUTH ERROR] C-PRO rejected login for {username}: {rejection}.")
                    if login_attempt == 1: raise InvalidCredentialsError(f"Corrected password was also rejected for {username}.")
                    corrected = self._prompt_for_correct_password(acc, rejection)
                    if not corrected: raise InvalidCredentialsError(f"Login cancelled for {username}.")
                    password = corrected
                    continue

                try: page.wait_for_url("**/case-listing**", timeout=8000)
                except PlaywrightTimeoutError: pass
                self._wait_for_network_quiet(page, timeout=4000)
                if self._is_kicked_or_unauthorized(page): raise InvalidCredentialsError(f"Could not confirm authenticated session for {username}.")
                break
            else:
                if self._is_kicked_or_unauthorized(page): raise InvalidCredentialsError(f"Sign-in form not available for {username}.")
                break
        else:
            raise InvalidCredentialsError(f"Login failed for {username}.")

        if section:
            self.log(f"[AUTH] Restoring section tab: {section}")
            self._switch_section_and_wait(page, section)

        self.set_status(f"Checking {username}...", bg="#E8F8EE", fg=THEME["green"])

    @staticmethod
    def _normalize_grid_text(value):
        return re.sub(r"\s+", " ", str(value or "")).strip().casefold()

    def _find_grid_column_index(self, page, column_name):
        wanted = re.sub(r"[^a-z0-9]+", "", column_name.casefold())
        aliases = {"accountno": {"accountno", "accountnumber", "account"}, "primaryborrowername": {"primaryborrowername", "borrowername", "primaryname"}}
        wanted_set = aliases.get(wanted, {wanted})
        grids = page.locator(".dx-datagrid:visible")
        for grid_idx in range(grids.count()):
            grid = grids.nth(grid_idx)
            header_rows = grid.locator(".dx-datagrid-headers tr.dx-header-row:visible")
            for row_idx in range(header_rows.count()):
                header_row = header_rows.nth(row_idx)
                cells = header_row.locator("td, th")
                for col_idx in range(cells.count()):
                    cell = cells.nth(col_idx)
                    try:
                        text = re.sub(r"[^a-z0-9]+", "", cell.inner_text(timeout=700).casefold())
                        aria = re.sub(r"[^a-z0-9]+", "", (cell.get_attribute("aria-label") or "").casefold())
                        candidates = {x for x in (text, aria) if x}
                        if any(candidate in wanted_set or any(candidate.endswith(alias) for alias in wanted_set) for candidate in candidates):
                            return grid, col_idx
                    except Exception: continue
            header_cells = grid.locator(".dx-datagrid-headers [role='columnheader']")
            for col_idx in range(header_cells.count()):
                cell = header_cells.nth(col_idx)
                try:
                    text = re.sub(r"[^a-z0-9]+", "", cell.inner_text(timeout=700).casefold())
                    aria = re.sub(r"[^a-z0-9]+", "", (cell.get_attribute("aria-label") or "").casefold())
                    candidates = {x for x in (text, aria) if x}
                    if any(candidate in wanted_set or any(candidate.endswith(alias) for alias in wanted_set) for candidate in candidates):
                        return grid, col_idx
                except Exception: continue
        return None, None

    def _get_grid_filter_input(self, page, column_index, column_name, grid=None):
        filter_rows = (grid if grid is not None else page).locator("tr.dx-datagrid-filter-row:visible")
        if filter_rows.count() == 0: raise RuntimeError("The current C-PRO grid has no filter row")

        filter_row = None
        for row_idx in range(filter_rows.count()):
            candidate = filter_rows.nth(row_idx)
            if candidate.locator("td, th").count() > column_index:
                filter_row = candidate
                break
        if filter_row is None: raise RuntimeError(f"Filter row does not contain column '{column_name}'")

        cell = filter_row.locator("td, th").nth(column_index)
        def visible_input_in_cell():
            inputs = cell.locator("input.dx-texteditor-input, input[type='text'], input:not([type])")
            for input_idx in range(inputs.count()):
                candidate = inputs.nth(input_idx)
                try:
                    if candidate.is_visible() and candidate.is_enabled(): return candidate
                except Exception: continue
            return None

        target = visible_input_in_cell()
        if target is not None: return target

        for selector in (".dx-icon-search", "button", "[role='button']"):
            controls = cell.locator(selector)
            for control_idx in range(controls.count()):
                control = controls.nth(control_idx)
                try:
                    if control.is_visible() and control.is_enabled():
                        control.click(timeout=1200)
                        page.wait_for_timeout(150)
                        target = visible_input_in_cell()
                        if target is not None: return target
                        overlay_inputs = page.locator(".dx-overlay-content input.dx-texteditor-input:visible, .dx-popup-content input.dx-texteditor-input:visible")
                        if overlay_inputs.count() > 0: return overlay_inputs.last
                except Exception: continue

        try:
            cell.click(timeout=1200)
            page.wait_for_timeout(150)
        except Exception: pass
        target = visible_input_in_cell()
        if target is not None: return target

        overlay_inputs = page.locator(".dx-overlay-content input.dx-texteditor-input:visible, .dx-popup-content input.dx-texteditor-input:visible")
        if overlay_inputs.count() > 0: return overlay_inputs.last
        raise RuntimeError(f"Could not activate the filter for '{column_name}'")

    def _clear_all_filters(self, page):
        try:
            filter_rows = page.locator("tr.dx-datagrid-filter-row:visible")
            cleared = False
            for row_idx in range(filter_rows.count()):
                row = filter_rows.nth(row_idx)
                inputs = row.locator("input.dx-texteditor-input:visible")
                for i in range(inputs.count()):
                    inp = inputs.nth(i)
                    try:
                        if inp.input_value():
                            inp.fill("")
                            cleared = True
                    except Exception: pass
            if cleared:
                first_input = page.locator("tr.dx-datagrid-filter-row:visible input.dx-texteditor-input:visible").first
                if first_input.count() > 0:
                    first_input.press("Enter")
                    page.wait_for_timeout(400)
        except Exception: pass

    @staticmethod
    def _grid_cell_matches(column_name, cell_text, search_value):
        cell = re.sub(r"\s+", " ", str(cell_text or "")).strip()
        query = re.sub(r"\s+", " ", str(search_value or "")).strip()
        if not cell or not query: return False

        if column_name.casefold() == "account no":
            cell_digits = re.sub(r"\D", "", cell)
            query_digits = re.sub(r"\D", "", query)
            if cell_digits and query_digits: return cell_digits == query_digits
            return cell.casefold() == query.casefold()
        return query.casefold() in cell.casefold()

    def _grid_snapshot(self, page):
        signature_parts = []
        try:
            rows = page.locator("tr.dx-data-row:visible")
            for row_idx in range(min(rows.count(), 100)):
                try: signature_parts.append(rows.nth(row_idx).inner_text(timeout=1000).strip())
                except Exception: signature_parts.append("<row-refreshing>")
        except Exception: signature_parts.append("<grid-refreshing>")

        for selector in (".dx-datagrid-nodata:visible", ".dx-datagrid-rowsview .dx-empty-message:visible"):
            try:
                labels = page.locator(selector)
                for i in range(min(labels.count(), 3)):
                    text = labels.nth(i).inner_text(timeout=500).strip()
                    if text: signature_parts.append("EMPTY:" + text)
            except Exception: pass

        loading = False
        for selector in (".dx-loadpanel:visible", ".dx-loadindicator:visible", ".dx-datagrid-loading:visible", ".dx-datagrid[aria-busy='true']"):
            try:
                loc = page.locator(selector)
                for i in range(min(loc.count(), 5)):
                    try:
                        if loc.nth(i).is_visible():
                            loading = True
                            break
                    except Exception: continue
                if loading: break
            except Exception: continue
        return tuple(signature_parts), loading

    def _trigger_grid_action_and_wait(self, page, action, description, timeout_ms=GRID_SETTLE_TIMEOUT_MS, require_activity=False):
        before_signature, _ = self._grid_snapshot(page)
        started = time.monotonic()
        last_network_activity = [started]
        seen_xhr = []
        pending_xhr = set()

        def on_request(request):
            try:
                if request.resource_type in ("xhr", "fetch"):
                    token = id(getattr(request, "_impl_obj", request))
                    seen_xhr.append(token)
                    pending_xhr.add(token)
                    last_network_activity[0] = time.monotonic()
            except Exception: pass

        def on_request_done(request):
            try:
                if request.resource_type in ("xhr", "fetch"):
                    pending_xhr.discard(id(getattr(request, "_impl_obj", request)))
                    last_network_activity[0] = time.monotonic()
            except Exception: pass

        page.on("request", on_request)
        page.on("requestfinished", on_request_done)
        page.on("requestfailed", on_request_done)
        last_signature = before_signature
        stable_samples = 0
        grid_change_seen = False
        loading_seen = False

        try:
            action()
            while (time.monotonic() - started) * 1000 < timeout_ms:
                if self.stop_requested: raise RuntimeError("Stopped by user while waiting for the C-PRO grid")
                if self._is_kicked_or_unauthorized(page): raise InvalidCredentialsError("Session expired or unauthenticated during grid refresh.")

                signature, loading = self._grid_snapshot(page)
                now = time.monotonic()
                elapsed = now - started

                if signature != last_signature:
                    last_signature = signature
                    stable_samples = 0
                else: stable_samples += 1
                if signature != before_signature: grid_change_seen = True
                if loading:
                    stable_samples = 0
                    loading_seen = True

                xhr_seen = bool(seen_xhr)
                network_quiet = (not pending_xhr and (now - last_network_activity[0]) >= 0.9)
                changed = signature != before_signature

                if xhr_seen:
                    min_elapsed = 1.0
                    network_ready = network_quiet
                    required_stable_samples = 4
                else:
                    min_elapsed = 1.5 if changed else 3.5
                    network_ready = True
                    required_stable_samples = 5

                if (elapsed >= min_elapsed and stable_samples >= required_stable_samples and not loading and network_ready):
                    if require_activity and not (xhr_seen or grid_change_seen or loading_seen):
                        # SMART FALLBACK: If the grid didn't change, check if it's already empty.
                        # If it is, the search probably worked and returned 0 results.
                        current_sig, _ = self._grid_snapshot(page)
                        if any("EMPTY:" in s or "No data" in s for s in current_sig):
                            return elapsed
                        raise GridSearchError(f"No XHR/loading/row change observed after submitting {description}.")
                    return elapsed
                page.wait_for_timeout(250)

            raise GridLoadTimeoutError(f"C-PRO grid did not finish '{description}' within {timeout_ms // 1000}s.")
        finally:
            for event, handler in (("request", on_request), ("requestfinished", on_request_done), ("requestfailed", on_request_done)):
                try: page.remove_listener(event, handler)
                except Exception: pass

    def _search_grid_with_relogin(self, page, acc, section, column_name, val):
        if not val: return []
        for attempt in range(1, MAX_RETRIES_ON_KICKOUT + 1):
            if self.stop_requested: return []
            try:
                if self._is_kicked_or_unauthorized(page):
                    self.log(f"[!] Session unauthenticated before search for '{val}'. Re-logging...")
                    self._login_and_restore_section(page, acc, section)

                if getattr(self, "_last_searched_column", None) != column_name:
                    self._clear_all_filters(page)
                    self._last_searched_column = column_name

                grid, column_index = self._find_grid_column_index(page, column_name)
                if grid is None or column_index is None:
                    if self._is_kicked_or_unauthorized(page):
                        raise InvalidCredentialsError("Unauthenticated while finding column.")
                    raise RuntimeError(f"Column '{column_name}' was not found in section '{section}'.")

                target_box = self._get_grid_filter_input(page, column_index, column_name, grid=grid)
                def submit_search():
                    target_box.fill("")
                    target_box.fill(str(val))
                    target_box.press("Enter")

                self.log(f"[WAIT] Waiting for C-PRO to load {column_name} search: '{val}'...")
                elapsed = self._trigger_grid_action_and_wait(page, submit_search, f"search {column_name} = '{val}'", require_activity=True)
                self.log(f"[WAIT] C-PRO grid settled after {elapsed:.1f}s for '{val}'.")

                if self._is_kicked_or_unauthorized(page):
                    self.log(f"[!] Session unauthenticated after search for '{val}'. Re-logging...")
                    self._login_and_restore_section(page, acc, section)
                    continue

                rows = grid.locator("tr.dx-data-row:visible")
                matched = []
                for row_idx in range(rows.count()):
                    row = rows.nth(row_idx)
                    cells = row.locator("td, th")
                    if cells.count() <= column_index: continue
                    searched_cell_text = cells.nth(column_index).inner_text().strip()
                    if not self._grid_cell_matches(column_name, searched_cell_text, val): continue
                    row_text = row.inner_text().strip()
                    if row_text and "No data" not in row_text:
                        matched.append(" | ".join(part.strip() for part in row_text.split("\t") if part.strip()))
                return matched

            except InvalidCredentialsError as e:
                if "rejected" in str(e).lower() or "cancelled" in str(e).lower(): raise e
                if attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session unauthenticated (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try: self._login_and_restore_section(page, acc, section)
                    except Exception as login_err: self.log(f"[!] Re-login failed: {login_err}")
                    continue
                raise
            except (GridLoadTimeoutError, GridSearchError) as e:
                if self._is_kicked_or_unauthorized(page) and attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session interrupted (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try: self._login_and_restore_section(page, acc, section)
                    except Exception as login_err: self.log(f"[!] Re-login failed: {login_err}")
                    continue
                raise
            except Exception as e:
                if self._is_kicked_or_unauthorized(page) and attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session interruption (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try: self._login_and_restore_section(page, acc, section)
                    except Exception as login_err: self.log(f"[!] Re-login error: {login_err}")
                    continue
                self.log(f"[ERROR] Search failed for '{column_name}' = '{val}' in {section}: {e}.")
                raise GridSearchError(f"Could not complete '{column_name}' search for '{val}': {e}") from e
        return []

    def _cpro_worker(self, active_accounts, active_sections):
        mode = self.search_mode.get()
        headless = not self.show_browser.get()

        found_tracker = {idx: {"accounts": [], "sections": [], "details": []} for idx in range(len(self.cases_data))}
        for row in self.cases_data:
            row["Availability Status"] = "NOT FOUND"
            row["Found In Account(s)"] = "None"
            row["Found In Section(s)"] = "None"
            row["C-PRO Row Details"] = "-"
        self.refresh_table()

        detailed_log = []
        total_steps = len(active_accounts) * len(active_sections) * len(self.cases_data)
        completed_steps = 0
        self.set_progress(0, total_steps)

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=headless, slow_mo=60)

                for acc in active_accounts:
                    if self.stop_requested: break
                    group_label = acc["group_label"]
                    username = acc["username"]
                    self.log(f"=== Starting Account: {username} ({group_label}) ===")

                    context = browser.new_context()
                    page = context.new_page()
                    current_case = None
                    current_section = "(not selected)"

                    try:
                        self._login_and_restore_section(page, acc, section=None)

                        for section in active_sections:
                            if self.stop_requested: break
                            current_section = section
                            self.log(f"[{username}] Switching to section: {section}")

                            try:
                                if self._is_kicked_or_unauthorized(page):
                                    self._login_and_restore_section(page, acc, section=None)
                                self._switch_section_and_wait(page, section)
                            except InvalidCredentialsError as e: raise e
                            except Exception as section_error:
                                raise GridLoadTimeoutError(f"Could not load the {section} section table: {section_error}") from section_error

                            for idx, case in enumerate(self.cases_data):
                                if self.stop_requested: break
                                
                                # EARLY EXIT: Skip if already found in a previous account/section
                                if case.get("Availability Status") == "AVAILABLE":
                                    completed_steps += 1
                                    self.set_progress(completed_steps, total_steps)
                                    continue

                                current_case = case
                                self.refresh_table(active_idx=idx)

                                acc_no = case.get("Account No", "").strip()
                                search_name = case.get("Search Name", "").strip()
                                matched_data = []
                                matched_by = ""

                                try:
                                    if mode in ("ACCOUNT_THEN_NAME", "ACCOUNT_ONLY") and acc_no:
                                        matched_data = self._search_grid_with_relogin(page, acc, section, "Account No", acc_no)
                                        if matched_data: matched_by = f"Account No ({acc_no})"

                                    if not matched_data and mode in ("ACCOUNT_THEN_NAME", "NAME_ONLY"):
                                        if not search_name:
                                            self.log(f"  [!] Cannot fallback to Name for {acc_no} (Name is empty).")
                                        else:
                                            matched_data = self._search_grid_with_relogin(page, acc, section, "Primary Borrower Name", search_name)
                                            if matched_data: matched_by = f"Name ({search_name})"
                                except (GridSearchError, GridLoadTimeoutError) as e:
                                    self.log(f"  [ERROR] Search failed for '{acc_no}' / '{search_name}': {e}")
                                    case["Availability Status"] = "CHECK ERROR"
                                    case["Found In Account(s)"] = "Check incomplete"
                                    case["Found In Section(s)"] = f"{username}: {section}"
                                    case["C-PRO Row Details"] = str(e)
                                    
                                    completed_steps += 1
                                    self.set_progress(completed_steps, total_steps)
                                    self.refresh_table()
                                    continue
                                except InvalidCredentialsError:
                                    raise

                                if matched_data:
                                    self.log(f"  ✓ FOUND by {matched_by}: {search_name or acc_no} in {username} -> {section}")
                                    found_tracker[idx]["accounts"].append(group_label)
                                    found_tracker[idx]["sections"].append(f"{username}: {section}")
                                    found_tracker[idx]["details"].extend(matched_data)

                                    case["Availability Status"] = "AVAILABLE"
                                    case["Found In Account(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["accounts"]))
                                    case["Found In Section(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["sections"]))
                                    case["C-PRO Row Details"] = " || ".join(found_tracker[idx]["details"])

                                    for r_str in matched_data:
                                        detailed_log.append({
                                            "Account No": acc_no, "Search Name": search_name, "C-PRO Account Group": group_label,
                                            "Login Used": username, "Section": section, "Status": "AVAILABLE", "Matched By": matched_by, "C-PRO Row Details": r_str
                                        })
                                else:
                                    self.log(f"  ✗ NOT FOUND: {search_name or acc_no or '(no search value)'} in {username} -> {section}")
                                    detailed_log.append({
                                        "Account No": acc_no, "Search Name": search_name, "C-PRO Account Group": group_label,
                                        "Login Used": username, "Section": section, "Status": "Not Found", "Matched By": "-", "C-PRO Row Details": "-"
                                    })

                                completed_steps += 1
                                self.set_progress(completed_steps, total_steps)
                                self.refresh_table()

                    except InvalidCredentialsError as e:
                        if "rejected" in str(e).lower() or "cancelled" in str(e).lower():
                            for pending_case in self.cases_data:
                                if pending_case.get("Availability Status") != "AVAILABLE":
                                    pending_case["Availability Status"] = "CHECK ERROR"
                                    pending_case["Found In Account(s)"] = "Login not verified"
                                    pending_case["Found In Section(s)"] = f"{username}: authentication"
                                    pending_case["C-PRO Row Details"] = str(e)
                            self.refresh_table()
                            self.log(f"[AUTH STOP] {e} Automation paused to protect account.")
                            self.root.after(0, lambda msg=str(e): messagebox.showwarning("C-PRO Login Stopped", msg + "\n\nCorrect the password in the left panel, then restart.", parent=self.root))
                            self.stop_requested = True
                        else:
                            for pending_case in self.cases_data:
                                if pending_case.get("Availability Status") != "AVAILABLE":
                                    pending_case["Availability Status"] = "CHECK ERROR"
                                    pending_case["Found In Account(s)"] = "Check incomplete"
                                    pending_case["Found In Section(s)"] = f"{username}: {current_section}"
                                    pending_case["C-PRO Row Details"] = str(e)
                            self.refresh_table()
                            self.log(f"[AUTH ERROR] {username} could not re-authenticate: {e}. Continuing to next account.")
                    except (GridLoadTimeoutError, GridSearchError) as e:
                        for pending_case in self.cases_data:
                            if pending_case.get("Availability Status") != "AVAILABLE":
                                pending_case["Availability Status"] = "CHECK ERROR"
                                pending_case["Found In Account(s)"] = "Check incomplete"
                                pending_case["Found In Section(s)"] = f"{username}: {current_section}"
                                pending_case["C-PRO Row Details"] = str(e)
                        self.log(f"[ERROR] {username} could not complete a search in {current_section}: {e}.")
                        self.refresh_table()
                    except Exception as e:
                        self.log(f"[ERROR] Account {username} encountered an error: {e}")
                    finally:
                        context.close()

                browser.close()

            df_summary = pd.DataFrame(self.cases_data)
            df_detailed = pd.DataFrame(detailed_log)

            def _write_report(path):
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    df_summary.to_excel(writer, index=False, sheet_name="Summary_Report")
                    df_detailed.to_excel(writer, index=False, sheet_name="All_Checks_Detail")
                    for sheet in writer.sheets.values():
                        for col in sheet.columns:
                            max_len = max(len(str(cell.value or "")) for cell in col) + 4
                            sheet.column_dimensions[col[0].column_letter].width = min(max(max_len, 14), 50)

            report_path = self._write_with_fallback(REPORT_EXCEL, _write_report, "C-PRO availability report")
            self.last_report_path = report_path
            self.log(f"Done! Full availability report saved to '{report_path}'.")
            
            # Open the folder automatically so the user sees the file
            self.root.after(0, lambda: self._open_folder(os.path.abspath(os.path.dirname(report_path))))

            incomplete_count = sum(1 for row in self.cases_data if row.get("Availability Status") == "CHECK ERROR")
            if incomplete_count:
                self.log(f"[WARNING] {incomplete_count} case(s) are CHECK ERROR.")
                self.set_status("Completed with errors", bg="#FFF4D6", fg="#8A5A00")
                completion_title = "C-PRO Check Finished with Incomplete Searches"
                completion_message = f"Finished processing {len(self.cases_data)} case(s), but {incomplete_count} case(s) have CHECK ERROR.\n\nReport saved to:\n{report_path}"
            else:
                self.set_status("Completed", bg="#E8F8EE", fg=THEME["green"])
                completion_title = "C-PRO Check Complete"
                completion_message = f"Finished checking {len(self.cases_data)} case(s)!\n\nReport saved to:\n{report_path}"
            
            self.root.after(0, lambda saved_path=report_path, title=completion_title, msg=completion_message: messagebox.showinfo(title, msg, parent=self.root))

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Automation failed: {e}")
            self.set_status("Error", bg="#FFF2F2", fg=THEME["red"])
        finally:
            self.is_running = False
            self.stop_requested = False
            self._last_searched_column = None
            self.refresh_table(active_idx=None)
            self.root.after(0, lambda: self.btn_run_cpro.config(state="normal"))
            self.root.after(0, lambda: self.btn_upload_pdf.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

if __name__ == "__main__":
    root = tk.Tk()
    app = CProSuiteApp(root)
    root.mainloop()