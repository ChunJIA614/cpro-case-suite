import os
import re
import io
import sys
import json
import time
import shutil
import datetime
import threading
import subprocess
import ctypes
import tempfile
from decimal import Decimal, InvalidOperation
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
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

try:
    from PIL import Image
except ImportError:
    Image = None

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
MAX_INLINE_SEARCH_RETRIES = 2
MAX_ERROR_RECOVERY_ROUNDS = 3
GRID_SETTLE_TIMEOUT_MS = 30000
CHECKED_GLYPH = "☑"
UNCHECKED_GLYPH = "☐"

WINRT_BATCH_OCR_PS = r'''
param([string]$Folder)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Graphics, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
Function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
$ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($null -eq $ocr) {
    $ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("en-US"))
}
if ($null -eq $ocr) {
    Write-Error "Windows OCR engine could not be initialized."
    exit 2
}
$files = Get-ChildItem -Path $Folder -Filter "page_*.png" | Sort-Object Name
foreach ($f in $files) {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($f.FullName)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    if ($bitmap.BitmapPixelFormat -ne [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8) {
        $converted = [Windows.Graphics.Imaging.SoftwareBitmap]::Convert($bitmap, [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8)
        $bitmap.Dispose()
        $bitmap = $converted
    }
    $res = Await ($ocr.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $lines = @()
    foreach ($ln in $res.Lines) {
        $words = @()
        foreach ($w in $ln.Words) {
            $r = $w.BoundingRect
            $words += @{
                t = $w.Text
                x = [math]::Round($r.X, 1)
                y = [math]::Round($r.Y, 1)
                w = [math]::Round($r.Width, 1)
                h = [math]::Round($r.Height, 1)
            }
        }
        $lines += @{ text = $ln.Text; words = $words }
    }
    $obj = @{
        file = $f.Name
        width = $bitmap.PixelWidth
        height = $bitmap.PixelHeight
        lines = $lines
    }
    Write-Output ($obj | ConvertTo-Json -Depth 5 -Compress)
    $bitmap.Dispose()
    $stream.Dispose()
}
'''


class GridLoadTimeoutError(RuntimeError):
    pass


class GridSearchError(RuntimeError):
    pass


class InvalidCredentialsError(RuntimeError):
    pass


def _find_writable_output_dirs():
    """Find the best place to save files, prioritizing the script's folder, then Desktop/Documents."""
    candidates = []

    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(script_dir)

    user_home = os.path.expanduser("~")
    if user_home:
        candidates.append(os.path.join(user_home, "Desktop", "CPro Case Suite Output"))
        candidates.append(os.path.join(user_home, "Documents", "CPro Case Suite Output"))

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(os.path.join(local_app_data, "CProCaseSuite", "Output"))

    candidates.append(os.path.join(tempfile.gettempdir(), "CProCaseSuite", "Output"))
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
    OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "CProCaseSuite", "Output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    WRITABLE_OUTPUT_DIRS = [OUTPUT_DIR]
else:
    OUTPUT_DIR = WRITABLE_OUTPUT_DIRS[0]

EXTRACTED_EXCEL = os.path.join(OUTPUT_DIR, "extracted_cases.xlsx")
REPORT_EXCEL = os.path.join(OUTPUT_DIR, "cpro_availability_report.xlsx")
EXTRACTED_INVOICES_DIR = os.path.join(OUTPUT_DIR, "extracted_invoices")


# ==============================================================================
# DATE, AMOUNT & BILLING MATCHING HELPERS
# ==============================================================================
_NUMBER_WORDS = {
    "ZERO": 0, "ONE": 1, "TWO": 2, "THREE": 3, "FOUR": 4, "FIVE": 5, "SIX": 6,
    "SEVEN": 7, "EIGHT": 8, "NINE": 9, "TEN": 10, "ELEVEN": 11, "TWELVE": 12,
    "THIRTEEN": 13, "FOURTEEN": 14, "FIFTEEN": 15, "SIXTEEN": 16, "SEVENTEEN": 17,
    "EIGHTEEN": 18, "NINETEEN": 19, "TWENTY": 20, "THIRTY": 30, "FORTY": 40,
    "FIFTY": 50, "SIXTY": 60, "SEVENTY": 70, "EIGHTY": 80, "NINETY": 90,
}

_OCR_WORD_FIXES = {
    "SRXRY": "SIXTY", "SRXTY": "SIXTY", "SXTY": "SIXTY", "SLXTY": "SIXTY",
    "SIXRY": "SIXTY", "FCRTY": "FORTY", "F0RTY": "FORTY", "FORTV": "FORTY",
    "NINETEE": "NINETEEN", "EIGTHY": "EIGHTY", "EIGHTV": "EIGHTY",
    "THLRTY": "THIRTY", "THIRTV": "THIRTY", "TWENTV": "TWENTY",
    "FIFT": "FIFTY", "FIFTV": "FIFTY", "SEVENTV": "SEVENTY",
    "HUNDREO": "HUNDRED", "HUNORED": "HUNDRED", "THOUSANO": "THOUSAND",
    "ONLV": "ONLY", "CENIS": "CENTS", "CENIS.": "CENTS", "CENT": "CENTS",
}


def parse_date_flexible(value):
    """Parse common C-PRO, invoice, and Excel date strings into datetime.date."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    raw = re.sub(r"\s+", " ", str(value).strip())
    if not raw or raw.lower() in ("nan", "nat", "-", "n/a", "none"):
        return None

    # Strip trailing time like "17 Jan 2026 3:36 PM" or "2026-04-09 00:00:00"
    date_candidates = [raw]
    m_ymd = re.match(r"^(\d{4}[-/]\d{1,2}[-/]\d{1,2})\b", raw)
    if m_ymd:
        date_candidates.insert(0, m_ymd.group(1))
    m_dmy_mon = re.match(r"^(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})\b", raw)
    if m_dmy_mon:
        date_candidates.insert(0, m_dmy_mon.group(1))
    m_dmy_num = re.match(r"^(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b", raw)
    if m_dmy_num:
        date_candidates.insert(0, m_dmy_num.group(1))

    formats = (
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%d/%m/%y", "%d-%m-%y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d %b %Y", "%d %B %Y",
        "%d-%b-%Y", "%d-%B-%Y",
    )
    for cand in date_candidates:
        for fmt in formats:
            try:
                return datetime.datetime.strptime(cand, fmt).date()
            except ValueError:
                continue
    return None


def format_date_dmy(dt):
    """Format a date object or string consistently as DD/MM/YYYY."""
    parsed = parse_date_flexible(dt)
    return parsed.strftime("%d/%m/%Y") if parsed else str(dt or "").strip()


def parse_amount_cents(value):
    """Parse an RM currency string/float into exact integer cents."""
    if value is None:
        return None
    raw = str(value).strip()
    if not raw or raw.lower() in ("nan", "none", "-", "n/a"):
        return None
    cleaned = re.sub(r"(?i)rm|\s|,", "", raw)
    if not re.match(r"^-?\d+(?:\.\d+)?$", cleaned):
        return None
    try:
        dec = Decimal(cleaned)
        return int((dec * 100).quantize(Decimal("1")))
    except (InvalidOperation, ValueError):
        return None


def format_cents(cents):
    if cents is None:
        return ""
    return f"{cents / 100:.2f}"


def _words_to_int(tokens):
    total = 0
    cur = 0
    for tok in tokens:
        t = re.sub(r"[^A-Z0-9]", "", tok.upper())
        if not t or t == "AND":
            continue
        t = _OCR_WORD_FIXES.get(t, t)
        if t in _NUMBER_WORDS:
            cur += _NUMBER_WORDS[t]
        elif t == "HUNDRED":
            cur = (cur or 1) * 100
        elif t == "THOUSAND":
            total += (cur or 1) * 1000
            cur = 0
        else:
            return None
    return total + cur


def parse_ringgit_words(text):
    """Extract numeric amount in RM from 'RINGGIT MALAYSIA ...' text (handles wrapped lines)."""
    if not text:
        return None
    m = re.search(
        r"RINGGIT\s+MALAYSIA\s+(.+?)(?=\bONLY\b|\bNotes\b|\bAuthorised\b|\bSub\s*Total\b|\bService\s*Tax\b|$)",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None
    phrase = re.sub(r"\s+", " ", m.group(1)).strip().upper()
    if not phrase:
        return None

    # Normalize common OCR corruption of CENTS and hyphenated OCR glitches before splitting hyphens
    phrase = re.sub(r"\bCEN[TI1]S?\b", "CENTS", phrase)
    phrase = phrase.replace("SRX-RY", "SIXTY").replace("SRX-TY", "SIXTY").replace("SlX-TY", "SIXTY")
    if "CENTS" in phrase:
        rm_part, sen_part = phrase.split("CENTS", 1)
    else:
        rm_part, sen_part = phrase, ""

    rm_val = _words_to_int(rm_part.replace("-", " ").split())
    if rm_val is None:
        return None
    sen_val = 0
    if sen_part.strip():
        s_val = _words_to_int(sen_part.replace("-", " ").split())
        if s_val is None or s_val >= 100:
            return None
        sen_val = s_val
    return rm_val * 100 + sen_val


def find_billing_match(billing_rows, target_date, target_cents):
    """
    Verify target_date and target_cents against C-PRO Billing History rows.
    Excludes 'Rejected' rows. Matches either a single row or a subset sum of rows
    on target_date equal to target_cents.
    Returns (matched: bool, detail_str: str).
    """
    if target_date is None or target_cents is None:
        return False, "Missing invoice date or amount for comparison"

    same_date_rows = []
    other_dates_preview = []
    for row in billing_rows:
        status_low = row.get("status", "").strip().casefold()
        if "rejected" in status_low:
            continue
        r_date = row.get("date_obj")
        r_cents = row.get("amount_cents")
        if r_date is None or r_cents is None:
            continue
        if r_date == target_date:
            if r_cents > 0:
                same_date_rows.append(row)
        else:
            if len(other_dates_preview) < 4:
                other_dates_preview.append(
                    f"{r_date.strftime('%d/%m/%Y')}: RM {format_cents(r_cents)} ({row.get('status', '')})"
                )

    if not same_date_rows:
        hint = "; ".join(other_dates_preview) if other_dates_preview else "no non-rejected billing rows"
        return False, f"No non-rejected billing on {target_date.strftime('%d/%m/%Y')} ({hint})"

    # 1. Exact single row match
    for row in same_date_rows:
        if row["amount_cents"] == target_cents:
            ref = row.get("billing_ref") or "-"
            st = row.get("status") or "-"
            return True, f"Single row match on {target_date.strftime('%d/%m/%Y')}: RM {format_cents(target_cents)} [{ref}, {st}]"

    # 2. Subset-sum match (2+ rows on the same date summing to target_cents)
    reachable = {0: ()}
    for idx, row in enumerate(same_date_rows):
        amt = row["amount_cents"]
        if amt <= 0 or amt > target_cents:
            continue
        new_entries = {}
        for prev_sum, prev_idxs in reachable.items():
            nxt_sum = prev_sum + amt
            if nxt_sum <= target_cents and nxt_sum not in reachable and nxt_sum not in new_entries:
                new_entries[nxt_sum] = prev_idxs + (idx,)
        reachable.update(new_entries)
        if target_cents in reachable:
            chosen = [same_date_rows[i] for i in reachable[target_cents]]
            parts = " + ".join(
                f"RM {format_cents(r['amount_cents'])} ({r.get('status', '-')})" for r in chosen
            )
            return True, f"Combined rows match on {target_date.strftime('%d/%m/%Y')}: {parts} = RM {format_cents(target_cents)}"

    found_list = ", ".join(
        f"RM {format_cents(r['amount_cents'])} ({r.get('status', '-')})" for r in same_date_rows
    )
    return False, f"Date {target_date.strftime('%d/%m/%Y')} found [{found_list}], none/sum equals RM {format_cents(target_cents)}"


def preview_matches_case(preview_info, case):
    """
    Compare extracted fields from an existing E-INVOICE preview against the target case.
    Returns (is_same_invoice: bool, reason: str).
    """
    if not preview_info:
        return False, "No readable text from preview"

    def _norm_inv(v):
        s = re.sub(r"[\s–—]+", "-", str(v or "").strip().upper())
        return re.sub(r"-+", "-", s)

    target_inv = _norm_inv(case.get("Invoice No"))
    prev_inv = _norm_inv(preview_info.get("inv_no"))
    target_date = parse_date_flexible(case.get("Invoice Date"))
    prev_date = parse_date_flexible(preview_info.get("inv_date"))
    target_cents = parse_amount_cents(case.get("Invoice Amount"))
    prev_cents = parse_amount_cents(preview_info.get("inv_amt"))

    if target_inv and prev_inv:
        if target_inv == prev_inv:
            return True, f"Invoice No matched ({prev_inv})"
        return False, f"Preview is {prev_inv}, expected {target_inv}"

    if target_date and prev_date:
        if target_cents is not None and prev_cents is not None:
            if target_date == prev_date and target_cents == prev_cents:
                return True, f"Date & Amount matched ({prev_date.strftime('%d/%m/%Y')}, RM {format_cents(prev_cents)})"
            return False, f"Preview is {prev_date.strftime('%d/%m/%Y')} RM {format_cents(prev_cents)}, expected {target_date.strftime('%d/%m/%Y')} RM {format_cents(target_cents)}"
        if target_date != prev_date:
            return False, f"Preview date {prev_date.strftime('%d/%m/%Y')} != {target_date.strftime('%d/%m/%Y')}"
        return True, f"Date matched ({prev_date.strftime('%d/%m/%Y')})"

    return False, "Could not identify Invoice No or Date from preview"


def parse_page_numbers_spec(spec):
    """Parse '1' or '13-14' or '1, 2' into sorted unique 1-based page ints."""
    if not spec:
        return []
    pages = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            if bounds[0].strip().isdigit() and bounds[1].strip().isdigit():
                start_p, end_p = int(bounds[0].strip()), int(bounds[1].strip())
                if start_p <= end_p:
                    pages.extend(range(start_p, end_p + 1))
        elif part.isdigit():
            pages.append(int(part))
    return sorted(set(p for p in pages if p >= 1))


def format_page_numbers_spec(pages):
    """Format list of 1-based page ints cleanly as '1' or '13-14'."""
    nums = sorted(set(int(p) for p in pages if int(p) >= 1))
    if not nums:
        return ""
    ranges = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            ranges.append(str(start) if start == prev else f"{start}-{prev}")
            start = prev = n
    ranges.append(str(start) if start == prev else f"{start}-{prev}")
    return ", ".join(ranges)


def extract_invoice_pdf_pages(source_pdf_path, page_numbers, dest_pdf_path):
    """Extract only the specified 1-based page_numbers into a standalone PDF."""
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError("Missing 'pypdf' package. Run: python -m pip install pypdf")
    if not source_pdf_path or not os.path.isfile(source_pdf_path):
        raise FileNotFoundError(f"Source PDF not found: {source_pdf_path}")
    pages = parse_page_numbers_spec(page_numbers) if isinstance(page_numbers, str) else list(page_numbers or [])
    if not pages:
        raise ValueError("No invoice page numbers specified to extract.")

    reader = PdfReader(source_pdf_path)
    writer = PdfWriter()
    total_pages = len(reader.pages)
    added = 0
    for pno in pages:
        if 1 <= pno <= total_pages:
            writer.add_page(reader.pages[pno - 1])
            added += 1
    if added == 0:
        raise ValueError(f"Requested page(s) {pages} out of range for '{os.path.basename(source_pdf_path)}' ({total_pages} pages).")

    os.makedirs(os.path.dirname(os.path.abspath(dest_pdf_path)), exist_ok=True)
    with open(dest_pdf_path, "wb") as f:
        writer.write(f)
    return dest_pdf_path


# ==============================================================================
# OPTIMIZED PDF & INVOICE EXTRACTION ENGINE (DIGITAL + SCANNED WINRT / RAPIDOCR)
# ==============================================================================
class PDFExtractorEngine:
    def __init__(self):
        self.ocr = None

    @staticmethod
    def _clean_borrower_short_name(full_name):
        """Preserve the full borrower name (whitespace-normalized) without reducing BIN/BINTI/A/L/A/P."""
        return re.sub(r"\s+", " ", str(full_name or "")).strip()

    @staticmethod
    def _is_billing_printout_text(upper_text):
        billing_hits = sum(1 for k in (
            "BILLING HISTORY",
            "PENDING APPROVAL",
            "AUTO-BILLING",
            "DOUBLE CLICK TO VIEW BILLING",
            "CURRENT LEGAL STATUS",
            "COURT SUIT NO.",
            "LAW FIRM REF NO.",
        ) if k in upper_text)
        return billing_hits >= 2 or "BILLING HISTORY" in upper_text or "DOUBLE CLICK TO VIEW BILLING" in upper_text

    @staticmethod
    def _is_invoice_page_text(upper_text):
        if PDFExtractorEngine._is_billing_printout_text(upper_text):
            return False
        inv_hits = sum(1 for k in (
            "INVOICE",
            "SST-20",
            "THOMAS ONG LEONG",
            "RINGGIT MALAYSIA",
            "SUB TOTAL",
            "LHDN",
            "INCLUSIVE OF TAX",
        ) if k in upper_text)
        return inv_hits >= 2 or bool(re.search(r"\bSST\s*[-–—]\s*\d{4}\s*[-–—]\s*\d{4,6}", upper_text))

    @staticmethod
    def _run_winrt_ocr_folder(folder):
        """Run Windows WinRT OCR on all page_*.png files inside folder and return parsed JSON dicts."""
        ps_path = os.path.join(folder, "batch_ocr.ps1")
        with open(ps_path, "w", encoding="utf-8") as f:
            f.write(WINRT_BATCH_OCR_PS)

        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path, "-Folder", folder],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if proc.returncode != 0 and not proc.stdout.strip():
            raise RuntimeError(f"Windows OCR failed: {proc.stderr.strip()}")

        results = []
        for line in proc.stdout.splitlines():
            line = line.strip()
            if not line.startswith("{"):
                continue
            results.append(json.loads(line))
        return results

    def _run_winrt_ocr_pdf(self, pdf_path, progress_cb=None):
        """Render/extract images from each page of pdf_path and OCR in one fast Windows batch."""
        if sys.platform != "win32" or PdfReader is None or Image is None:
            return None
        reader = PdfReader(pdf_path)
        total = len(reader.pages)
        page_results = []

        with tempfile.TemporaryDirectory() as td:
            rendered_count = 0
            for idx, pg in enumerate(reader.pages):
                if progress_cb and idx % 20 == 0:
                    progress_cb(f"  Preparing page {idx + 1}/{total} for OCR...")
                best_pil = None
                best_area = -1
                try:
                    for im in pg.images:
                        try:
                            cand = Image.open(io.BytesIO(im.data))
                            area = cand.width * cand.height
                            if area > best_area:
                                best_area = area
                                best_pil = cand
                        except Exception:
                            continue
                except Exception:
                    best_pil = None

                if best_pil is None and pymupdf is not None:
                    try:
                        doc = pymupdf.open(pdf_path)
                        pix = doc[idx].get_pixmap(dpi=200)
                        best_pil = Image.open(io.BytesIO(pix.tobytes("png")))
                        doc.close()
                    except Exception:
                        best_pil = None

                if best_pil is None:
                    continue

                rot = int(pg.rotation or 0) % 360
                if rot == 90:
                    best_pil = best_pil.rotate(-90, expand=True)
                elif rot == 180:
                    best_pil = best_pil.rotate(180, expand=True)
                elif rot == 270:
                    best_pil = best_pil.rotate(90, expand=True)
                if best_pil.mode != "RGB":
                    best_pil = best_pil.convert("RGB")
                best_pil.save(os.path.join(td, f"page_{idx + 1:04d}.png"))
                rendered_count += 1

            if rendered_count == 0:
                return []

            if progress_cb:
                progress_cb(f"  Running Windows OCR on {rendered_count} page(s)...")

            for obj in self._run_winrt_ocr_folder(td):
                m_pno = re.search(r"page_(\d+)\.png", obj.get("file", ""))
                if not m_pno:
                    continue
                obj["page_num"] = int(m_pno.group(1))
                page_results.append(obj)

        page_results.sort(key=lambda x: x["page_num"])
        return page_results

    def extract_preview_info_from_text(self, text):
        """Parse Invoice No, Date, Amount, and Account No from OCR/preview text."""
        inv_m = re.search(r"\b(SST\s*[-–—]\s*\d{4}\s*[-–—]\s*\d{4,6}[A-Z]?)\b", text or "", re.IGNORECASE)
        inv_no = re.sub(r"[\s–—]+", "-", inv_m.group(1).upper()) if inv_m else ""
        inv_no = re.sub(r"-+", "-", inv_no)

        inv_date = ""
        date_m = re.search(r"\b(\d{2}[/-]\d{2}[/-]\d{4})\b", text or "")
        if date_m:
            inv_date = format_date_dmy(date_m.group(1))

        amt_cents = parse_ringgit_words(text or "")
        acc_m = re.search(
            r"(?:^|[\s|])([A-Z0-9()/+]{1,18})\s*[-–—:]\s*(\d{6,8})\s*[-–—:.]\s*(\d{7,20})",
            text or "",
            re.IGNORECASE | re.MULTILINE,
        )
        acc_no = acc_m.group(3).strip() if acc_m else ""
        return {
            "inv_no": inv_no,
            "inv_date": inv_date,
            "inv_amt": format_cents(amt_cents) if amt_cents is not None else "",
            "acc_no": acc_no,
            "raw_text": text or "",
        }

    def extract_preview_info_from_image(self, image_path):
        """OCR a single PNG/image screenshot of Preview Document and extract invoice info."""
        if sys.platform == "win32" and Image is not None:
            with tempfile.TemporaryDirectory() as td:
                im = Image.open(image_path)
                if im.mode != "RGB":
                    im = im.convert("RGB")
                im.save(os.path.join(td, "page_0001.png"))
                objs = self._run_winrt_ocr_folder(td)
                if objs:
                    txt = "\n".join(ln.get("text", "") for ln in objs[0].get("lines", []))
                    return self.extract_preview_info_from_text(txt)
        if RapidOCR is not None:
            if self.ocr is None:
                self.ocr = RapidOCR()
            with open(image_path, "rb") as f:
                ocr_res, _ = self.ocr(f.read())
            txt = "\n".join(item[1] for item in (ocr_res or []))
            return self.extract_preview_info_from_text(txt)
        return {}

    def extract_preview_info_from_pdf_bytes(self, pdf_bytes):
        """Extract invoice identifiers from raw PDF bytes captured from Preview or Download."""
        with tempfile.TemporaryDirectory() as td:
            pdf_path = os.path.join(td, "preview.pdf")
            with open(pdf_path, "wb") as f:
                f.write(pdf_bytes)
            cases = self.extract_pdf(pdf_path)
            if cases:
                first = cases[0]
                return {
                    "inv_no": first.get("Invoice No", ""),
                    "inv_date": first.get("Invoice Date", ""),
                    "inv_amt": first.get("Invoice Amount", ""),
                    "acc_no": first.get("Account No", ""),
                    "raw_text": "",
                }
            pages = self._read_pdf_pages_ocr_or_text(pdf_path)
            txt = "\n".join(p.get("text", "") for p in pages)
            return self.extract_preview_info_from_text(txt)

    def _read_pdf_pages_ocr_or_text(self, pdf_path, progress_cb=None):
        """
        Return list of page dicts:
        {'page_num': int, 'width': int, 'height': int, 'text': str, 'lines': [{'text': str, 'words': [...]}]}
        """
        pages_out = []

        # 1. Try fast digital text layer first (PyMuPDF or pypdf)
        digital_texts = []
        if pymupdf is not None:
            doc = pymupdf.open(pdf_path)
            for idx in range(len(doc)):
                digital_texts.append((idx + 1, doc[idx].get_text("text").strip()))
            doc.close()
        elif PdfReader is not None:
            reader = PdfReader(pdf_path)
            for idx, pg in enumerate(reader.pages):
                digital_texts.append((idx + 1, (pg.extract_text() or "").strip()))

        has_digital_text = digital_texts and sum(1 for _, t in digital_texts if len(t) >= 30) >= max(1, len(digital_texts) // 2)
        if has_digital_text:
            for pno, txt in digital_texts:
                lines = [{"text": ln, "words": []} for ln in txt.splitlines() if ln.strip()]
                pages_out.append({"page_num": pno, "width": 0, "height": 0, "text": txt, "lines": lines})
            return pages_out

        # 2. Scanned PDF: prefer fast zero-dependency Windows WinRT OCR
        if sys.platform == "win32" and PdfReader is not None and Image is not None:
            win_pages = self._run_winrt_ocr_pdf(pdf_path, progress_cb=progress_cb)
            if win_pages is not None:
                for wp in win_pages:
                    lines = wp.get("lines", [])
                    wp["text"] = "\n".join(ln.get("text", "") for ln in lines)
                    pages_out.append(wp)
                return pages_out

        # 3. Fallback to PyMuPDF + RapidOCR (for non-Windows or if WinRT unavailable)
        if pymupdf is None or RapidOCR is None:
            raise RuntimeError(
                "Scanned PDF OCR requires Windows OCR (with pypdf + Pillow) or pymupdf + rapidocr-onnxruntime."
            )
        if self.ocr is None:
            self.ocr = RapidOCR()
        doc = pymupdf.open(pdf_path)
        for idx in range(len(doc)):
            if progress_cb and idx % 10 == 0:
                progress_cb(f"  Running RapidOCR page {idx + 1}/{len(doc)}...")
            page = doc[idx]
            text = page.get_text("text").strip()
            if len(text) < 30:
                pix = page.get_pixmap(dpi=200)
                ocr_res, _ = self.ocr(pix.tobytes("png"))
                if ocr_res:
                    text = "\n".join(item[1] for item in ocr_res)
            lines = [{"text": ln, "words": []} for ln in text.splitlines() if ln.strip()]
            pages_out.append({"page_num": idx + 1, "width": 0, "height": 0, "text": text, "lines": lines})
        doc.close()
        return pages_out

    @staticmethod
    def _extract_fields_from_invoice_group(page_objs):
        """Parse Invoice No, Date, Amount, Account No(s), and Borrower Name from grouped invoice pages."""
        combined_text = "\n".join(p.get("text", "") for p in page_objs)

        # 1. Invoice No (e.g., SST-2026-04873)
        inv_m = re.search(r"\b(SST\s*[-–—]\s*\d{4}\s*[-–—]\s*\d{4,6}[A-Z]?)\b", combined_text, re.IGNORECASE)
        inv_no = re.sub(r"[\s–—]+", "-", inv_m.group(1).upper()) if inv_m else ""
        inv_no = re.sub(r"-+", "-", inv_no)

        # 2. Invoice Date (prefer top Date field DD/MM/YYYY)
        inv_date = ""
        date_m = re.search(r"\b(\d{2}[/-]\d{2}[/-]\d{4})\b", combined_text)
        if date_m:
            inv_date = format_date_dmy(date_m.group(1))

        # 3. Account Number(s)
        # Supports MG-2611838-2681668098, BN (HP)-2610975-36693180757, BN (+)-2610974-33183484324,
        # POD (CC)-2621994-080616873126, FC-2611496-191008557, and paired "& 26791000564"
        acc_pattern = re.compile(
            r"(?:^|[\s|])([A-Z0-9()/+]{1,18})\s*[-–—:]\s*(\d{6,8})\s*[-–—:.]\s*(\d{7,20})"
            r"(?:\s*(?:&|/|and)\s*(\d{7,20}))?",
            re.IGNORECASE | re.MULTILINE,
        )

        acc_num = ""
        alt_acc_nums = []
        acc_full = ""
        acc_match_page = None
        acc_match_obj = None
        for p_obj in page_objs:
            for m in acc_pattern.finditer(p_obj.get("text", "")):
                prefix = re.sub(r"\s+", " ", m.group(1)).strip().upper()
                mid = m.group(2).strip()
                primary = m.group(3).strip()
                paired = (m.group(4) or "").strip()
                if not acc_num:
                    acc_num = primary
                    acc_full = f"{prefix}-{mid}-{acc_num}"
                    acc_match_page = p_obj
                    acc_match_obj = m
                elif primary != acc_num and primary not in alt_acc_nums:
                    alt_acc_nums.append(primary)
                if paired and paired != acc_num and paired not in alt_acc_nums:
                    alt_acc_nums.append(paired)

        # Capture 16-digit Credit Card line "c: 4687862801722252" as primary (if missing) or alternate
        for cc_m in re.finditer(r"\b[cC]\s*[:;]\s*(\d{15,19})\b", combined_text):
            cc_num = cc_m.group(1).strip()
            if not acc_num:
                acc_num = cc_num
                acc_full = cc_num
            elif cc_num != acc_num and cc_num not in alt_acc_nums:
                alt_acc_nums.append(cc_num)
        alt_acc_num = ", ".join(alt_acc_nums)

        # 4. Borrower Name (layout-aware from OCR words first, regex fallback second)
        borrower_name = ""
        if acc_match_page is not None and acc_match_page.get("lines") and acc_num:
            lines = acc_match_page["lines"]
            all_words = [w for ln in lines for w in ln.get("words", [])]
            anchor_words = [w for w in all_words if acc_num in w.get("t", "") or acc_num[-7:] in w.get("t", "")]
            if anchor_words:
                ay = anchor_words[0]["y"]
                ax = anchor_words[0]["x"]
                candidates = []
                for ln in lines:
                    words = ln.get("words", [])
                    if not words:
                        continue
                    ly = words[0]["y"]
                    lx = words[0]["x"]
                    if 15 <= (ly - ay) <= 230 and abs(lx - ax) < 230:
                        txt = ln.get("text", "").strip()
                        if not txt:
                            continue
                        # Skip card number line "c: 4687..." and numeric columns
                        if re.match(r"^[cC]\s*[:;.]?\s*\d{10,}", txt):
                            continue
                        if re.match(r"^[\d\s.,/-]+$", txt):
                            continue
                        if any(stop in txt.upper() for stop in ("RINGGIT", "SUB TOTAL", "SERVICE TAX", "NOTES", "QTY", "U/ PRICE")):
                            continue
                        if re.search(r"[A-Za-z]{3,}", txt):
                            candidates.append((ly, txt))
                if candidates:
                    candidates.sort(key=lambda item: item[0])
                    borrower_name = candidates[0][1]

        if not borrower_name and acc_match_obj is not None and acc_match_page is not None:
            snippet = acc_match_page.get("text", "")[acc_match_obj.end():acc_match_obj.end() + 200]
            snippet = re.sub(r"^[cC]\s*[:;.]?\s*\d{10,}\s*", "", snippet.strip())
            snippet = re.sub(r"^[\d.,\s\-|]+", "", re.sub(r"\s+", " ", snippet))
            nm = re.search(r"^([A-Z][A-Z\s@/.'()-]{2,80})", snippet)
            if nm:
                cand = re.sub(
                    r"\s+(SV-8|Legal fee|Qty|UOM|RINGGIT|Sub Total|Service Tax|Total).*$",
                    "",
                    nm.group(1).strip(),
                    flags=re.IGNORECASE,
                ).strip()
                if len(cand) > 2 and not any(k in cand.upper() for k in ("NO.", "DATE", "PAGE")):
                    borrower_name = cand

        # 5. Invoice Total Amount (Inclusive of Tax) via multi-signal consensus
        amt_words_cents = parse_ringgit_words(combined_text)

        # Look at last page in the group for totals block
        last_page = page_objs[-1]
        page_w = float(last_page.get("width") or 2480)
        x_thresh = page_w * 0.72
        right_nums = []
        sub_y = tax_y = tot_y = None
        for ln in last_page.get("lines", []):
            words = ln.get("words", [])
            if not words:
                continue
            lt = ln.get("text", "").upper()
            ly = words[0]["y"]
            lx = words[0]["x"]
            if page_w * 0.45 <= lx <= page_w * 0.85:
                if "SUB TOTAL" in lt or "EXCLUDING TAX" in lt:
                    sub_y = ly
                elif "SERVICE TAX" in lt:
                    tax_y = ly
                elif "INCLUSIVE OF TAX" in lt:
                    tot_y = ly
            for w in words:
                if w["x"] >= x_thresh:
                    cleaned = w["t"].replace(",", "").strip()
                    if re.match(r"^\d+\.\d{2}$", cleaned):
                        c_val = parse_amount_cents(cleaned)
                        if c_val is not None:
                            right_nums.append((w["y"], c_val))

        sub_cents = tax_cents = tot_box_cents = None
        if right_nums and (sub_y or tax_y or tot_y):
            for ny, nval in right_nums:
                if sub_y is not None and abs(ny - sub_y) <= 45 and sub_cents is None:
                    sub_cents = nval
                elif tax_y is not None and abs(ny - tax_y) <= 45 and tax_cents is None:
                    tax_cents = nval
                elif tot_y is not None and abs(ny - tot_y) <= 45 and tot_box_cents is None:
                    tot_box_cents = nval

        # Also check "Service Tax @ 8% on <taxable>" if tax_cents wasn't aligned
        if tax_cents is None:
            st_m = re.search(r"Service\s*Tax\s*@\s*8\s*(?:%|0/0)\s*on\s*([\d,]+\.\d{2})", combined_text, re.IGNORECASE)
            if st_m:
                base_c = parse_amount_cents(st_m.group(1))
                if base_c is not None:
                    tax_cents = int(round(base_c * 0.08))

        amt_calc_cents = None
        if sub_cents is not None and tax_cents is not None:
            amt_calc_cents = sub_cents + tax_cents
        elif tot_box_cents is not None:
            amt_calc_cents = tot_box_cents

        # Choose final consensus amount
        final_cents = None
        if amt_words_cents is not None and amt_calc_cents is not None:
            if abs(amt_words_cents - amt_calc_cents) <= 1:
                final_cents = amt_words_cents
            elif tot_box_cents is not None and abs(tot_box_cents - amt_words_cents) <= 1:
                final_cents = amt_words_cents
            else:
                # Prefer Ringgit words if valid (much less prone to clipped/dropped decimals)
                final_cents = amt_words_cents
        elif amt_words_cents is not None:
            final_cents = amt_words_cents
        elif amt_calc_cents is not None:
            final_cents = amt_calc_cents

        # Fallback: if totals page was missing (e.g. truncated 1-of-2), sum Qty * U/Price lines
        if final_cents is None:
            item_matches = re.findall(
                r"(?m)^\s*\d+\.\s+(?:SV-8|D2-Disb)?.*?(\d+)\s+1\s+([\d,]+\.\d{2})\b",
                combined_text,
            )
            if item_matches:
                pass  # Layout varies horizontally; leave blank for manual review rather than guessing wrong

        return {
            "Invoice No": inv_no,
            "Invoice Date": inv_date,
            "Invoice Amount": format_cents(final_cents) if final_cents is not None else "",
            "Account No": acc_num,
            "Alt Account No": alt_acc_num,
            "Full Account No": acc_full,
            "Full Borrower Name": borrower_name,
            "Search Name": PDFExtractorEngine._clean_borrower_short_name(borrower_name),
        }

    def extract_pdf(self, pdf_path, progress_cb=None):
        """
        Extract per-invoice records (with exact PDF page numbers, Invoice No, Date, Amount,
        Account No, and Borrower Name) from a single or multi-invoice PDF.
        Excludes C-PRO Billing History printout pages.
        """
        pages = self._read_pdf_pages_ocr_or_text(pdf_path, progress_cb=progress_cb)
        if not pages:
            return []

        filename = os.path.basename(pdf_path)
        abs_pdf = os.path.abspath(pdf_path)

        # Group invoice pages (skipping Billing History printouts)
        invoice_groups = []
        current_group = []
        current_inv_no = ""

        def _flush_group():
            nonlocal current_group, current_inv_no
            if current_group:
                invoice_groups.append(current_group)
                current_group = []
                current_inv_no = ""

        for p_idx, p_obj in enumerate(pages):
            txt = p_obj.get("text", "")
            upper = txt.upper()
            if self._is_billing_printout_text(upper):
                # Recover Borrower Name for the preceding invoice if it lacked one
                if invoice_groups or current_group:
                    target_grp = current_group if current_group else invoice_groups[-1]
                    b_m = re.search(r"Borrower'?s?\s*Name\s*[:\n]\s*([A-Z][A-Z\s@/.'()-]{2,80})", txt)
                    if b_m:
                        target_grp[0]["_fallback_borrower"] = b_m.group(1).strip()
                _flush_group()
                continue

            if not self._is_invoice_page_text(upper):
                continue

            inv_m = re.search(r"\b(SST\s*[-–—]\s*\d{4}\s*[-–—]\s*\d{4,6}[A-Z]?)\b", txt, re.IGNORECASE)
            page_inv_no = re.sub(r"[\s–—]+", "-", inv_m.group(1).upper()) if inv_m else ""
            has_total_end = ("RINGGIT MALAYSIA" in upper) or ("INCLUSIVE OF TAX" in upper)

            if not current_group:
                current_group = [p_obj]
                current_inv_no = page_inv_no
            elif page_inv_no and current_inv_no and page_inv_no == current_inv_no:
                current_group.append(p_obj)
            elif not page_inv_no and not ("RINGGIT MALAYSIA" in current_group[-1].get("text", "").upper()):
                current_group.append(p_obj)
            else:
                _flush_group()
                current_group = [p_obj]
                current_inv_no = page_inv_no

            if has_total_end and re.search(r"\b1\s*of\s*1\b", txt, re.IGNORECASE):
                _flush_group()

        _flush_group()

        # If structured invoice detection found nothing (e.g., non-invoice summons PDF), use legacy generic parser
        if not invoice_groups:
            combined_text = "\n".join(p.get("text", "") for p in pages)
            legacy_cases = self.parse_cases_from_text(combined_text, filename)
            for c in legacy_cases:
                c["Source PDF Path"] = abs_pdf
                c["Invoice Pages"] = format_page_numbers_spec([p["page_num"] for p in pages])
            return legacy_cases

        cases = []
        seen_keys = set()
        for grp in invoice_groups:
            fields = self._extract_fields_from_invoice_group(grp)
            # Filter out non-HLB invoices that have no bank account number AND no borrower
            if not fields["Account No"] and not fields["Full Borrower Name"]:
                continue
            if not fields["Full Borrower Name"] and grp[0].get("_fallback_borrower"):
                fb = grp[0]["_fallback_borrower"]
                fields["Full Borrower Name"] = fb
                fields["Search Name"] = self._clean_borrower_short_name(fb)

            dedup_key = (fields["Invoice No"], fields["Account No"])
            if fields["Invoice No"] and fields["Account No"] and dedup_key in seen_keys:
                continue
            seen_keys.add(dedup_key)

            page_spec = format_page_numbers_spec([p["page_num"] for p in grp])
            cases.append({
                "Source File": filename,
                "Source PDF Path": abs_pdf,
                "Invoice Pages": page_spec,
                "Invoice No": fields["Invoice No"],
                "Invoice Date": fields["Invoice Date"],
                "Invoice Amount": fields["Invoice Amount"],
                "Account No": fields["Account No"],
                "Alt Account No": fields["Alt Account No"],
                "Search Name": fields["Search Name"],
                "Full Borrower Name": fields["Full Borrower Name"],
                "Availability Status": "Pending",
                "Found In Account(s)": "-",
                "Found In Section(s)": "-",
                "Billing Check": "-",
                "Upload Status": "-",
                "Extracted Invoice PDF": "-",
                "C-PRO Row Details": "-",
                "_checked": True,
            })

        return cases

    @staticmethod
    def parse_cases_from_text(text, filename):
        cases = []
        seen_acc = set()

        acc_pattern = re.compile(
            r"(?:^|[\s|])([A-Z0-9()/+]{1,18})\s*[-–—:]\s*(\d{6,8})\s*[-–—:.]\s*(\d{7,20})",
            re.IGNORECASE | re.MULTILINE,
        )

        for match in acc_pattern.finditer(text):
            acc_no = match.group(3)
            if acc_no in seen_acc:
                continue

            start_idx = match.end()
            snippet = text[start_idx:start_idx + 180]
            snippet = re.sub(r"\s+", " ", snippet)
            snippet = re.sub(r"^[cC]\s*[:;.]?\s*\d{10,}\s*", "", snippet)
            snippet = re.sub(r"^[\d\.\,\s\-\|]+", "", snippet)

            name_match = re.search(r"^([A-Z][A-Z\s@\/\.\']{2,80})", snippet, re.IGNORECASE)
            name = ""
            if name_match:
                raw_name = name_match.group(1).strip()
                raw_name = re.sub(
                    r"\s+(SV-8|Legal fee|Notification|Letter|Unsealed|Writ|Summons|Billing|History|All|Pending|Approval|Approved|Rejected|Description|Qty|UOM).*$",
                    "",
                    raw_name,
                    flags=re.IGNORECASE,
                ).strip()
                if len(raw_name) > 2 and not any(k in raw_name.upper() for k in ["NO.", "DATE", "PAGE"]):
                    name = raw_name

            if name:
                short_name = PDFExtractorEngine._clean_borrower_short_name(name)
                cases.append({
                    "Source File": filename,
                    "Source PDF Path": "",
                    "Invoice Pages": "",
                    "Invoice No": "",
                    "Invoice Date": "",
                    "Invoice Amount": "",
                    "Account No": acc_no,
                    "Alt Account No": "",
                    "Search Name": short_name,
                    "Full Borrower Name": name,
                    "Availability Status": "Pending",
                    "Found In Account(s)": "-",
                    "Found In Section(s)": "-",
                    "Billing Check": "-",
                    "Upload Status": "-",
                    "Extracted Invoice PDF": "-",
                    "C-PRO Row Details": "-",
                    "_checked": True,
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
        self.root.geometry("1500x880")
        self.root.minsize(1200, 740)
        self.root.configure(bg=THEME["bg"])

        self.pdf_engine = PDFExtractorEngine()
        self.cases_data = []
        self.is_running = False
        self.stop_requested = False
        self.last_extracted_path = EXTRACTED_EXCEL
        self.last_report_path = REPORT_EXCEL
        self._last_searched_column = None

        self.accounts = []
        self.show_pwd = tk.BooleanVar(value=False)
        self.accounts_config_path = os.path.join(OUTPUT_DIR, "cpro_saved_accounts.json")

        self.sec_in_progress = tk.BooleanVar(value=True)
        self.sec_withhold = tk.BooleanVar(value=True)
        self.sec_withdrawn = tk.BooleanVar(value=True)

        self.search_mode = tk.StringVar(value="ACCOUNT_THEN_NAME")
        self.verify_and_upload = tk.BooleanVar(value=True)
        self.show_browser = tk.BooleanVar(value=True)

        self._configure_styles()
        self._build_layout()
        self._load_saved_accounts()
        self.log(f"Output folder: {OUTPUT_DIR}")
        self._auto_load_existing_excel()

    def _configure_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Apple.Treeview",
            background=THEME["card"],
            foreground=THEME["text"],
            fieldbackground=THEME["card"],
            rowheight=34,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Apple.Treeview.Heading",
            background="#F9F9FB",
            foreground=THEME["subtext"],
            relief="flat",
            borderwidth=0,
            padding=(8, 8),
            font=("Segoe UI Semibold", 8),
        )
        style.map("Apple.Treeview", background=[("selected", "#D8EAFE")], foreground=[("selected", THEME["text"])])
        style.configure(
            "Apple.Horizontal.TProgressbar",
            troughcolor="#E5E5EA",
            background=THEME["blue"],
            bordercolor="#E5E5EA",
            lightcolor=THEME["blue"],
            darkcolor=THEME["blue"],
            thickness=6,
        )

    def _create_card(self, parent, padx=14, pady=12):
        outer = tk.Frame(parent, bg=THEME["border"], padx=1, pady=1)
        inner = tk.Frame(outer, bg=THEME["card"], padx=padx, pady=pady)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def _create_button(self, parent, text, command, bg, fg="#FFFFFF", hover_bg=None, width=None, padx=10, pady=7, font_size=9):
        btn = tk.Button(
            parent, text=text, command=command, bg=bg, fg=fg,
            activebackground=hover_bg or bg, activeforeground=fg,
            relief="flat", bd=0, padx=padx, pady=pady, cursor="hand2",
            font=("Segoe UI Semibold", font_size),
        )
        if width:
            btn.config(width=width)
        if hover_bg:
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg) if btn["state"] != "disabled" else None)
            btn.bind("<Leave>", lambda e: btn.config(bg=bg) if btn["state"] != "disabled" else None)
        return btn

    def _build_layout(self):
        header = tk.Frame(self.root, bg=THEME["bg"], padx=24, pady=14)
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=THEME["bg"])
        title_box.pack(side="left")
        tk.Label(title_box, text="C-PRO Case Suite", font=("Segoe UI", 18, "bold"), bg=THEME["bg"], fg=THEME["text"]).pack(anchor="w")
        tk.Label(
            title_box,
            text="Extracts invoices from PDFs, verifies C-PRO Billing Date & Amount, and uploads E-INVOICE automatically.",
            font=("Segoe UI", 9),
            bg=THEME["bg"],
            fg=THEME["subtext"],
        ).pack(anchor="w", pady=(2, 0))

        status_frame = tk.Frame(header, bg=THEME["bg"])
        status_frame.pack(side="right", anchor="center")
        self.status_badge = tk.Label(status_frame, text="●  Ready", font=("Segoe UI Semibold", 9), bg="#E8E8ED", fg=THEME["text"], padx=12, pady=5)
        self.status_badge.pack(side="right")

        workspace = tk.Frame(self.root, bg=THEME["bg"], padx=20, pady=4)
        workspace.pack(fill="both", expand=True)

        sidebar_container = tk.Frame(workspace, bg=THEME["bg"], width=390)
        sidebar_container.pack(side="left", fill="y", padx=(0, 16))
        sidebar_container.pack_propagate(False)

        self.sidebar_canvas = tk.Canvas(sidebar_container, bg=THEME["bg"], highlightthickness=0)
        self.sidebar_vsb = ttk.Scrollbar(sidebar_container, orient="vertical", command=self.sidebar_canvas.yview)
        sidebar = tk.Frame(self.sidebar_canvas, bg=THEME["bg"])

        sidebar.bind(
            "<Configure>",
            lambda e: self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all")),
        )
        self.sidebar_window = self.sidebar_canvas.create_window((0, 0), window=sidebar, anchor="nw")
        self.sidebar_canvas.bind(
            "<Configure>",
            lambda e: self.sidebar_canvas.itemconfig(self.sidebar_window, width=e.width),
        )
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_vsb.set)

        self.sidebar_canvas.pack(side="left", fill="both", expand=True)
        self.sidebar_vsb.pack(side="right", fill="y")

        def _on_mousewheel(event):
            self.sidebar_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        sidebar_container.bind("<Enter>", lambda e: self.sidebar_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        sidebar_container.bind("<Leave>", lambda e: self.sidebar_canvas.unbind_all("<MouseWheel>"))

        self._build_sidebar(sidebar)

        content = tk.Frame(workspace, bg=THEME["bg"])
        content.pack(side="left", fill="both", expand=True)
        self._build_content_area(content)

    def _build_sidebar(self, sidebar):
        card1_outer, card1 = self._create_card(sidebar)
        card1_outer.pack(fill="x", pady=(0, 12))
        tk.Label(card1, text="C-PRO ACCOUNTS", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w")
        tk.Label(card1, text="Enter your C-PRO usernames & session passwords.", font=("Segoe UI", 8), bg=THEME["card"], fg=THEME["subtext"], wraplength=330, justify="left").pack(anchor="w", pady=(2, 8))

        self.accounts_list_frame = tk.Frame(card1, bg=THEME["card"])
        self.accounts_list_frame.pack(fill="x", expand=True)

        def toggle_pwd_visibility():
            char = "" if self.show_pwd.get() else "•"
            for acc in self.accounts:
                acc["ent_pwd"].config(show=char)

        tk.Checkbutton(
            card1, text="Show Passwords", variable=self.show_pwd, command=toggle_pwd_visibility,
            bg=THEME["card"], fg=THEME["subtext"], activebackground=THEME["card"], font=("Segoe UI", 8), anchor="w",
        ).pack(fill="x", pady=(6, 4))

        self.btn_add_acc = self._create_button(
            card1, "＋ Add Account", lambda: self._add_account_row(),
            bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"],
        )
        self.btn_add_acc.pack(fill="x")

        card2_outer, card2 = self._create_card(sidebar)
        card2_outer.pack(fill="x", pady=(0, 12))
        tk.Label(card2, text="SECTIONS TO CHECK", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 4))
        for sec_label, sec_var in [("Case In Progress", self.sec_in_progress), ("Case Withhold", self.sec_withhold), ("Case Withdrawn", self.sec_withdrawn)]:
            tk.Checkbutton(card2, text=sec_label, variable=sec_var, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9), anchor="w").pack(fill="x", pady=1)

        tk.Frame(card2, bg=THEME["border"], height=1).pack(fill="x", pady=8)
        tk.Label(card2, text="SEARCH STRATEGY", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 4))
        modes = [
            ("Account No → Fallback to Name", "ACCOUNT_THEN_NAME"),
            ("Account No Only", "ACCOUNT_ONLY"),
            ("Borrower Name Only", "NAME_ONLY"),
        ]
        for text, val in modes:
            tk.Radiobutton(card2, text=text, variable=self.search_mode, value=val, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9), anchor="w", justify="left").pack(fill="x", pady=1)

        tk.Frame(card2, bg=THEME["border"], height=1).pack(fill="x", pady=8)
        tk.Label(card2, text="BILLING & UPLOAD WORKFLOW", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 4))
        tk.Checkbutton(
            card2,
            text="Verify Billings & Upload E-INVOICE",
            variable=self.verify_and_upload,
            bg=THEME["card"],
            fg=THEME["text"],
            activebackground=THEME["card"],
            font=("Segoe UI Semibold", 9),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            card2,
            text="Checks Billings (single/sum on date, excl. Rejected); if matched, extracts invoice page(s) and uploads E-INVOICE to Legal Details → Additional Documents (upload skipped for Case Withdrawn).",
            font=("Segoe UI", 7),
            bg=THEME["card"],
            fg=THEME["subtext"],
            wraplength=330,
            justify="left",
        ).pack(anchor="w", padx=(20, 0), pady=(1, 6))

        tk.Frame(card2, bg=THEME["border"], height=1).pack(fill="x", pady=8)
        tk.Checkbutton(card2, text="Show Live Chrome Browser", variable=self.show_browser, bg=THEME["card"], fg=THEME["text"], activebackground=THEME["card"], font=("Segoe UI", 9), anchor="w").pack(fill="x")

        card3_outer, card3 = self._create_card(sidebar)
        card3_outer.pack(fill="x", pady=(0, 12))
        tk.Label(card3, text="EXCEL FILES & OUTPUTS", font=("Segoe UI Semibold", 8), bg=THEME["card"], fg=THEME["subtext"]).pack(anchor="w", pady=(0, 8))

        self._create_button(card3, "Open Extracted Cases Excel", lambda: self._open_file(getattr(self, "last_extracted_path", EXTRACTED_EXCEL)), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x", pady=(0, 6))
        self._create_button(card3, "Open Final C-PRO Report Excel", lambda: self._open_file(getattr(self, "last_report_path", REPORT_EXCEL)), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x", pady=(0, 6))
        self._create_button(card3, "📂 Open Output Folder", lambda: self._open_folder(OUTPUT_DIR), bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"]).pack(fill="x")

    def _add_account_row(self, username="", enabled=True):
        idx = len(self.accounts) + 1
        user_var = tk.StringVar(value=username)
        pwd_var = tk.StringVar()
        enable_var = tk.BooleanVar(value=enabled)

        card = tk.Frame(self.accounts_list_frame, bg=THEME["card"])
        card.pack(fill="x", expand=True, pady=(0, 6))

        box = tk.Frame(card, bg="#F9F9FB", padx=10, pady=8, highlightbackground=THEME["border"], highlightthickness=1)
        box.pack(fill="x", expand=True)

        top_line = tk.Frame(box, bg="#F9F9FB")
        top_line.pack(fill="x", pady=(0, 4))

        chk = tk.Checkbutton(
            top_line, text=f"Account #{idx}", variable=enable_var,
            bg="#F9F9FB", fg=THEME["text"], activebackground="#F9F9FB",
            font=("Segoe UI Semibold", 8), anchor="w",
        )
        chk.pack(side="left")

        del_btn = tk.Label(top_line, text="✕ Remove", bg="#F9F9FB", fg=THEME["subtext"], cursor="hand2", font=("Segoe UI", 8))
        del_btn.pack(side="right")

        u_box = tk.Frame(box, bg="#F9F9FB")
        u_box.pack(fill="x", pady=(1, 3))
        tk.Label(u_box, text="Username:", font=("Segoe UI", 8), bg="#F9F9FB", fg=THEME["subtext"]).pack(anchor="w")
        u_border = tk.Frame(u_box, bg=THEME["border"], padx=1, pady=1)
        u_border.pack(fill="x", pady=(1, 0))
        ent_user = tk.Entry(u_border, textvariable=user_var, bg=THEME["card"], fg=THEME["text"], relief="flat", bd=4, font=("Segoe UI", 9))
        ent_user.pack(fill="x")

        p_box = tk.Frame(box, bg="#F9F9FB")
        p_box.pack(fill="x", pady=(1, 2))
        tk.Label(p_box, text="Session Password:", font=("Segoe UI", 8), bg="#F9F9FB", fg=THEME["subtext"]).pack(anchor="w")
        p_border = tk.Frame(p_box, bg=THEME["border"], padx=1, pady=1)
        p_border.pack(fill="x", pady=(1, 0))
        show_char = "" if self.show_pwd.get() else "•"
        ent_pwd = tk.Entry(p_border, textvariable=pwd_var, show=show_char, bg=THEME["card"], fg=THEME["text"], relief="flat", bd=4, font=("Segoe UI", 9))
        ent_pwd.pack(fill="x")

        acc_obj = {
            "user": user_var,
            "pwd": pwd_var,
            "enabled": enable_var,
            "card": card,
            "chk": chk,
            "ent_pwd": ent_pwd,
            "del_btn": del_btn,
        }

        del_btn.bind("<Button-1>", lambda e, a=acc_obj: self._remove_account_row(a))
        del_btn.bind("<Enter>", lambda e, b=del_btn: b.config(fg=THEME["red"]))
        del_btn.bind("<Leave>", lambda e, b=del_btn: b.config(fg=THEME["subtext"]))

        self.accounts.append(acc_obj)
        self._refresh_account_indices()
        self.sidebar_canvas.update_idletasks()
        self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all"))

    def _remove_account_row(self, acc_obj):
        if len(self.accounts) <= 1:
            acc_obj["user"].set("")
            acc_obj["pwd"].set("")
            acc_obj["enabled"].set(True)
            return
        acc_obj["card"].destroy()
        self.accounts.remove(acc_obj)
        self._refresh_account_indices()
        self.sidebar_canvas.update_idletasks()
        self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all"))
        self._save_accounts_config()

    def _refresh_account_indices(self):
        for idx, acc in enumerate(self.accounts, start=1):
            acc["chk"].config(text=f"Account #{idx}")

    def _load_saved_accounts(self):
        saved_list = []
        if os.path.isfile(self.accounts_config_path):
            try:
                with open(self.accounts_config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        saved_list = data
            except Exception as e:
                self.log(f"Could not load saved accounts: {e}")

        if not saved_list:
            saved_list = [{"username": "", "enabled": True}, {"username": "", "enabled": False}]

        for item in saved_list:
            u = item.get("username", "") if isinstance(item, dict) else str(item)
            en = item.get("enabled", True) if isinstance(item, dict) else True
            self._add_account_row(username=u, enabled=en)

    def _save_accounts_config(self):
        data = []
        for acc in self.accounts:
            u = acc["user"].get().strip()
            if u:
                data.append({"username": u, "enabled": acc["enabled"].get()})
        try:
            os.makedirs(os.path.dirname(self.accounts_config_path), exist_ok=True)
            with open(self.accounts_config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.log(f"Notice: Could not save accounts configuration: {e}")

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
        table_header = tk.Frame(table_card, bg=THEME["card"], padx=16)
        table_header.pack(fill="x", pady=(10, 4))
        self.lbl_case_count = tk.Label(table_header, text="Cases Queue (0 records)", font=("Segoe UI Semibold", 10), bg=THEME["card"], fg=THEME["text"])
        self.lbl_case_count.pack(side="left")

        sel_bar = tk.Frame(table_header, bg=THEME["card"])
        sel_bar.pack(side="right")
        self.btn_select_all = self._create_button(sel_bar, "Select All", self.on_select_all_rows, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"], padx=8, pady=3, font_size=8)
        self.btn_select_all.pack(side="left", padx=(0, 5))
        self.btn_deselect_all = self._create_button(sel_bar, "Deselect All", self.on_deselect_all_rows, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"], padx=8, pady=3, font_size=8)
        self.btn_deselect_all.pack(side="left", padx=(0, 5))
        self.btn_select_pending = self._create_button(sel_bar, "Select Pending", self.on_select_pending_rows, bg=THEME["neutral_btn"], fg=THEME["text"], hover_bg=THEME["neutral_hover"], padx=8, pady=3, font_size=8)
        self.btn_select_pending.pack(side="left", padx=(0, 5))
        self.btn_select_errors = self._create_button(sel_bar, "Select Errors", self.on_select_error_rows, bg="#FFF4D6", fg="#8A5A00", hover_bg="#FFE8AD", padx=8, pady=3, font_size=8)
        self.btn_select_errors.pack(side="left")

        tip_row = tk.Frame(table_card, bg=THEME["card"], padx=16)
        tip_row.pack(fill="x", pady=(0, 8))
        tk.Label(
            tip_row,
            text="Tip: Click ☑ to select rows · Double-click a row to edit Account / Date / Amount / Pages · Space to toggle · Delete to remove",
            font=("Segoe UI", 8),
            bg=THEME["card"],
            fg=THEME["subtext"],
        ).pack(anchor="w")
        tk.Frame(table_card, bg=THEME["border"], height=1).pack(fill="x")

        tree_frame = tk.Frame(table_card, bg=THEME["card"])
        tree_frame.pack(fill="both", expand=True)

        cols = ("check", "no", "inv_no", "acc_no", "search_name", "inv_date", "inv_amt", "pages", "status", "found_acc", "billing_check", "upload_status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", style="Apple.Treeview", selectmode="extended")
        headings = [
            ("check", CHECKED_GLYPH, 40, "center", False),
            ("no", "#", 40, "center", False),
            ("inv_no", "Invoice No", 118, "w", False),
            ("acc_no", "Account No", 122, "w", False),
            ("search_name", "Borrower Name", 175, "w", True),
            ("inv_date", "Inv Date", 88, "center", False),
            ("inv_amt", "Inv Amt", 76, "e", False),
            ("pages", "Pages", 56, "center", False),
            ("status", "Status", 130, "center", False),
            ("found_acc", "Account(s)", 98, "w", False),
            ("billing_check", "Billing Check", 150, "w", True),
            ("upload_status", "Upload Status", 135, "w", True),
        ]
        for col_id, title, width, anchor, stretch in headings:
            self.tree.heading(col_id, text=title, anchor=anchor)
            self.tree.column(col_id, width=width, minwidth=min(width, 38), anchor=anchor, stretch=stretch)

        self.tree.tag_configure("AVAILABLE", background=THEME["row_found"])
        self.tree.tag_configure("COMPLETE", background=THEME["row_found"])
        self.tree.tag_configure("NOT FOUND", background=THEME["row_missing"])
        self.tree.tag_configure("NONE_MISMATCH", background=THEME["row_missing"])
        self.tree.tag_configure("CHECKING", background=THEME["row_active"])
        self.tree.tag_configure("CHECK ERROR", background=THEME["row_error"])
        self.tree.tag_configure("PARTIAL", background=THEME["row_error"])

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_double_click_row)
        self.tree.bind("<space>", self.on_toggle_highlighted_rows)
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

    def _set_controls_running(self, running, allow_stop=False):
        state = "disabled" if running else "normal"
        stop_state = "normal" if (running and allow_stop) else "disabled"

        def _apply():
            for btn in (
                self.btn_upload_pdf,
                self.btn_add_manual,
                self.btn_load_excel,
                self.btn_clear,
                self.btn_run_cpro,
                self.btn_select_all,
                self.btn_deselect_all,
                self.btn_select_pending,
                self.btn_select_errors,
            ):
                btn.config(state=state)
            self.btn_stop.config(state=stop_state)

        self.root.after(0, _apply)

    def _update_selection_summary(self):
        total = len(self.cases_data)
        checked = sum(1 for row in self.cases_data if row.get("_checked", True))
        error_count = sum(1 for row in self.cases_data if row.get("Availability Status") == "CHECK ERROR")
        if total == 0:
            self.lbl_case_count.config(text="Cases Queue (0 records)")
            self.tree.heading("check", text=UNCHECKED_GLYPH)
        else:
            self.lbl_case_count.config(text=f"Cases Queue ({total} records · {checked} selected)")
            self.tree.heading("check", text=CHECKED_GLYPH if checked == total else UNCHECKED_GLYPH)
        err_label = f"Select Errors ({error_count})" if error_count else "Select Errors"
        self.btn_select_errors.config(text=err_label)

    def _set_all_checked(self, checked):
        for row in self.cases_data:
            row["_checked"] = bool(checked)
        glyph = CHECKED_GLYPH if checked else UNCHECKED_GLYPH
        for iid in self.tree.get_children():
            self.tree.set(iid, "check", glyph)
        self._update_selection_summary()

    def _toggle_row_checkbox(self, idx):
        if not (0 <= idx < len(self.cases_data)):
            return
        new_val = not self.cases_data[idx].get("_checked", True)
        self.cases_data[idx]["_checked"] = new_val
        iid = str(idx)
        if self.tree.exists(iid):
            self.tree.set(iid, "check", CHECKED_GLYPH if new_val else UNCHECKED_GLYPH)
        self._update_selection_summary()

    def _toggle_all_checkboxes(self):
        if not self.cases_data:
            return
        all_checked = all(row.get("_checked", True) for row in self.cases_data)
        self._set_all_checked(not all_checked)

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        if self.is_running:
            return "break"
        if region == "heading":
            self._toggle_all_checkboxes()
            return "break"
        if region in ("cell", "tree"):
            iid = self.tree.identify_row(event.y)
            if iid:
                self._toggle_row_checkbox(int(iid))
                self.tree.focus(iid)
            return "break"

    def on_toggle_highlighted_rows(self, event):
        if self.is_running:
            return "break"
        sel = self.tree.selection()
        if not sel:
            return
        any_unchecked = any(not self.cases_data[int(iid)].get("_checked", True) for iid in sel)
        target_val = bool(any_unchecked)
        for iid in sel:
            idx = int(iid)
            self.cases_data[idx]["_checked"] = target_val
            if self.tree.exists(iid):
                self.tree.set(iid, "check", CHECKED_GLYPH if target_val else UNCHECKED_GLYPH)
        self._update_selection_summary()
        return "break"

    def on_select_all_rows(self):
        if self.is_running:
            return
        self._set_all_checked(True)

    def on_deselect_all_rows(self):
        if self.is_running:
            return
        self._set_all_checked(False)

    def _select_rows_by_status(self, target_status, empty_title, empty_msg):
        if self.is_running:
            return
        matching = [i for i, row in enumerate(self.cases_data) if row.get("Availability Status") == target_status]
        if not matching:
            messagebox.showinfo(empty_title, empty_msg, parent=self.root)
            return
        match_set = set(matching)
        for i, row in enumerate(self.cases_data):
            checked = i in match_set
            row["_checked"] = checked
            iid = str(i)
            if self.tree.exists(iid):
                self.tree.set(iid, "check", CHECKED_GLYPH if checked else UNCHECKED_GLYPH)
        self._update_selection_summary()
        self.log(f"Selected {len(matching)} '{target_status}' row(s) for C-PRO check.")

    def on_select_pending_rows(self):
        self._select_rows_by_status("Pending", "No Pending Rows", "There are currently no Pending rows in the queue.")

    def on_select_error_rows(self):
        self._select_rows_by_status("CHECK ERROR", "No Error Rows", "There are currently no CHECK ERROR rows in the queue.")

    def refresh_table(self, active_idx=None):
        def _rebuild():
            try:
                y_pos = self.tree.yview()[0]
            except Exception:
                y_pos = 0.0
            self.tree.delete(*self.tree.get_children())
            for i, row in enumerate(self.cases_data):
                status = row.get("Availability Status", "Pending")
                checked = row.get("_checked", True)
                tag = ""
                if i == active_idx:
                    tag = "CHECKING"
                elif status in ("AVAILABLE", "COMPLETE"):
                    tag = "COMPLETE"
                elif status == "NOT FOUND":
                    tag = "NOT FOUND"
                elif status.startswith("NONE"):
                    tag = "NONE_MISMATCH"
                elif status == "CHECK ERROR":
                    tag = "CHECK ERROR"
                elif "BILLING MATCHED" in status:
                    tag = "PARTIAL"

                self.tree.insert(
                    "",
                    "end",
                    iid=str(i),
                    values=(
                        CHECKED_GLYPH if checked else UNCHECKED_GLYPH,
                        i + 1,
                        row.get("Invoice No", ""),
                        row.get("Account No", ""),
                        row.get("Search Name", ""),
                        row.get("Invoice Date", ""),
                        row.get("Invoice Amount", ""),
                        row.get("Invoice Pages", ""),
                        status,
                        row.get("Found In Account(s)", "-"),
                        row.get("Billing Check", "-"),
                        row.get("Upload Status", "-"),
                    ),
                    tags=(tag,) if tag else (),
                )
            self._update_selection_summary()
            if active_idx is not None and self.tree.exists(str(active_idx)):
                self.tree.see(str(active_idx))
            elif y_pos > 0.0:
                self.tree.yview_moveto(y_pos)

        self.root.after(0, _rebuild)

    @staticmethod
    def _normalize_imported_row(rec, fallback_source=""):
        inv_date_raw = rec.get("Invoice Date", rec.get("Date", "")).strip()
        inv_date = format_date_dmy(inv_date_raw) if inv_date_raw else ""

        inv_amt_raw = rec.get("Invoice Amount", rec.get("Amount", rec.get("Total", ""))).strip()
        inv_amt_cents = parse_amount_cents(inv_amt_raw)
        inv_amt = format_cents(inv_amt_cents) if inv_amt_cents is not None else inv_amt_raw

        src_file = rec.get("Source File", "").strip() or fallback_source
        src_path = rec.get("Source PDF Path", "").strip()
        if not src_path and src_file and src_file.lower().endswith(".pdf"):
            # Try auto-resolving in local Scanned Invoice folder or script dir
            base_dir = os.path.dirname(os.path.abspath(__file__))
            for cand in (
                os.path.join(base_dir, src_file),
                os.path.join(base_dir, "Scanned Invoice", src_file),
                os.path.join(OUTPUT_DIR, src_file),
                os.path.join(OUTPUT_DIR, "Scanned Invoice", src_file),
            ):
                if os.path.isfile(cand):
                    src_path = cand
                    break

        full_name = re.sub(
            r"\s+", " ", str(rec.get("Full Borrower Name") or rec.get("Search Name") or rec.get("Borrower Name") or "")
        ).strip()
        raw_search = re.sub(
            r"\s+", " ", str(rec.get("Search Name") or rec.get("Borrower Name") or "")
        ).strip()
        # Upgrade previously truncated Search Name (e.g. "NOH") back to Full Borrower Name ("NOH BIN AWANG")
        if full_name and (
            not raw_search
            or (len(full_name) > len(raw_search) and full_name.casefold().startswith(raw_search.casefold()))
        ):
            search_name = full_name
        else:
            search_name = raw_search or full_name

        return {
            "Source File": src_file,
            "Source PDF Path": src_path,
            "Invoice Pages": rec.get("Invoice Pages", rec.get("Pages", "")).strip(),
            "Invoice No": rec.get("Invoice No", "").strip(),
            "Invoice Date": inv_date,
            "Invoice Amount": inv_amt,
            "Account No": rec.get("Account No", "").strip(),
            "Alt Account No": rec.get("Alt Account No", "").strip(),
            "Search Name": search_name,
            "Full Borrower Name": full_name,
            "Availability Status": rec.get("Availability Status", "").strip() or "Pending",
            "Found In Account(s)": rec.get("Found In Account(s)", "").strip() or "-",
            "Found In Section(s)": rec.get("Found In Section(s)", "").strip() or "-",
            "Billing Check": rec.get("Billing Check", "").strip() or "-",
            "Upload Status": rec.get("Upload Status", "").strip() or "-",
            "Extracted Invoice PDF": rec.get("Extracted Invoice PDF", "").strip() or "-",
            "C-PRO Row Details": rec.get("C-PRO Row Details", "").strip() or "-",
            "_checked": True,
        }

    def _auto_load_existing_excel(self):
        candidates = [
            EXTRACTED_EXCEL,
            os.path.join(os.getcwd(), "extracted_cases.xlsx"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_cases.xlsx"),
        ]
        source_path = next((path for path in candidates if os.path.isfile(path)), None)
        if not source_path:
            return
        try:
            df = pd.read_excel(source_path, dtype=str).fillna("")
            for rec in df.to_dict("records"):
                self.cases_data.append(self._normalize_imported_row(rec, os.path.basename(source_path)))
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
                probe.write(b"ok")
                probe.flush()
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
            if normalized in checked:
                continue
            checked.add(normalized)
            if not self._directory_is_writable(folder):
                continue
            for suffix in range(10):
                extra = f"_{suffix + 1}" if suffix else ""
                candidate = os.path.join(folder, f"{base}_{timestamp}{extra}{ext}")
                if os.path.exists(candidate):
                    continue
                try:
                    write_callback(candidate)
                    self.log(f"[!] {description} saved to fallback: '{candidate}'.")
                    return candidate
                except Exception as e:
                    last_error = e
                    break
        raise PermissionError(f"Could not save {description}. Last error: {last_error or first_error_text}")

    def _save_extracted_excel(self):
        if not self.cases_data:
            return None
        df = pd.DataFrame(self.cases_data)
        desired_cols = [
            "Source File", "Source PDF Path", "Invoice Pages", "Invoice No",
            "Invoice Date", "Invoice Amount", "Account No", "Alt Account No",
            "Search Name", "Full Borrower Name",
        ]
        export_cols = [c for c in desired_cols if c in df.columns]

        def _write_workbook(path):
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df[export_cols].to_excel(writer, index=False, sheet_name="To_Check")
                ws = writer.sheets["To_Check"]
                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col) + 4
                    ws.column_dimensions[col[0].column_letter].width = min(max(max_len, 14), 50)

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
            if sys.platform == "win32":
                os.startfile(os.path.abspath(filepath))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", os.path.abspath(filepath)])
            else:
                subprocess.Popen(["xdg-open", os.path.abspath(filepath)])
        except Exception as e:
            messagebox.showerror("Error Opening File", str(e))

    def _open_folder(self, folder):
        if not os.path.exists(folder):
            messagebox.showinfo("Folder Not Found", f"The folder '{folder}' does not exist yet.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(os.path.abspath(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", os.path.abspath(folder)])
            else:
                subprocess.Popen(["xdg-open", os.path.abspath(folder)])
        except Exception as e:
            messagebox.showerror("Error Opening Folder", str(e))

    def on_upload_pdfs(self):
        if self.is_running:
            return
        initial_dir = os.path.join(OUTPUT_DIR, "Scanned Invoice")
        if not os.path.isdir(initial_dir):
            initial_dir = OUTPUT_DIR
        pdf_files = filedialog.askopenfilenames(
            parent=self.root,
            title="Select Invoice / Case PDF Files",
            initialdir=initial_dir,
            filetypes=[("PDF Files", "*.pdf")],
        )
        if not pdf_files:
            return

        self.is_running = True
        self._set_controls_running(True, allow_stop=False)
        self.set_status("Extracting PDFs...", bg="#EBF5FF", fg=THEME["blue"])

        def _worker():
            try:
                total = len(pdf_files)
                for idx, path in enumerate(pdf_files, start=1):
                    fname = os.path.basename(path)
                    self.log(f"Scanning ({idx}/{total}): {fname} ...")
                    try:
                        extracted_cases = self.pdf_engine.extract_pdf(path, progress_cb=self.log)
                        if not extracted_cases:
                            self.log(f"  [!] No Invoices / Account Numbers found in {fname}")
                        for case in extracted_cases:
                            self.cases_data.append(case)
                            inv_info = f"{case.get('Invoice No') or '-'} | {case.get('Invoice Date') or '-'} | RM {case.get('Invoice Amount') or '-'}"
                            self.log(f"  -> p.{case.get('Invoice Pages') or '?'} | {inv_info} | {case['Account No']} | {case['Search Name']}")
                        self.refresh_table()
                    except Exception as e:
                        self.log(f"  [ERROR] Failed to extract {fname}: {e}")
                    self.set_progress(idx, total)

                saved_path = self._save_extracted_excel()
                if saved_path:
                    self.log(f"Saved {len(self.cases_data)} total record(s) to '{saved_path}'.")
                self.set_status("Ready", bg="#E8E8ED", fg=THEME["text"])
            finally:
                self.is_running = False
                self._set_controls_running(False)

        threading.Thread(target=_worker, daemon=True).start()

    def on_import_excel(self):
        if self.is_running:
            return
        path = filedialog.askopenfilename(parent=self.root, title="Select Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path, dtype=str).fillna("")
            fname = os.path.basename(path)
            for rec in df.to_dict("records"):
                self.cases_data.append(self._normalize_imported_row(rec, fname))
            self.refresh_table()
            self._save_extracted_excel()
            self.log(f"Imported {len(df)} row(s) from {fname}.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def _open_case_form_dialog(self, title, field_specs, submit_label, on_submit):
        """Open a DPI-aware, auto-sized, scrollable modal form dialog centered over the app."""
        dialog = tk.Toplevel(self.root)
        dialog.withdraw()
        dialog.title(title)
        dialog.configure(bg=THEME["card"])
        dialog.resizable(True, True)
        dialog.transient(self.root)

        footer = tk.Frame(dialog, bg=THEME["card"], padx=20, pady=12)
        footer.pack(side="bottom", fill="x")

        body_wrap = tk.Frame(dialog, bg=THEME["card"])
        body_wrap.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(body_wrap, bg=THEME["card"], highlightthickness=0, bd=0)
        v_scroll = ttk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas.pack(side="left", fill="both", expand=True)

        form_frame = tk.Frame(canvas, bg=THEME["card"], padx=20, pady=14)
        canvas_window = canvas.create_window((0, 0), window=form_frame, anchor="nw")

        entries = {}
        for key, label_text, initial_val in field_specs:
            tk.Label(
                form_frame, text=label_text, bg=THEME["card"], fg=THEME["text"], font=("Segoe UI Semibold", 9)
            ).pack(anchor="w")
            ent = tk.Entry(form_frame, bg=THEME["input_bg"], relief="flat", bd=5, font=("Segoe UI", 9))
            if initial_val:
                ent.insert(0, str(initial_val))
            ent.pack(fill="x", pady=(2, 8))
            entries[key] = ent

        def _sync_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_form_width(event):
            canvas.itemconfigure(canvas_window, width=event.width)

        form_frame.bind("<Configure>", _sync_scroll_region)
        canvas.bind("<Configure>", _sync_form_width)

        def _on_mousewheel(event):
            if canvas.winfo_height() < form_frame.winfo_reqheight():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        dialog.bind("<MouseWheel>", _on_mousewheel)
        dialog.bind("<Escape>", lambda _e: dialog.destroy())

        def _handle_submit():
            values = {k: ent.get().strip() for k, ent in entries.items()}
            on_submit(values, dialog)

        self._create_button(
            footer, submit_label, _handle_submit, bg=THEME["blue"], hover_bg=THEME["blue_hover"]
        ).pack(fill="x")

        dialog.update_idletasks()
        screen_w = dialog.winfo_screenwidth()
        screen_h = dialog.winfo_screenheight()
        req_w = max(500, form_frame.winfo_reqwidth() + 28, footer.winfo_reqwidth() + 28)
        req_form_h = form_frame.winfo_reqheight()
        req_footer_h = footer.winfo_reqheight()
        total_req_h = req_form_h + req_footer_h + 12

        win_w = min(req_w, int(screen_w * 0.90))
        max_h = int(screen_h * 0.88)
        win_h = min(total_req_h, max_h)
        if total_req_h > max_h:
            v_scroll.pack(side="right", fill="y")
            win_w = min(win_w + 18, int(screen_w * 0.90))

        self.root.update_idletasks()
        parent_x = self.root.winfo_rootx()
        parent_y = self.root.winfo_rooty()
        parent_w = self.root.winfo_width() or screen_w
        parent_h = self.root.winfo_height() or screen_h
        pos_x = max(0, min(screen_w - win_w, parent_x + (parent_w - win_w) // 2))
        pos_y = max(0, min(screen_h - win_h, parent_y + (parent_h - win_h) // 2))

        dialog.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        dialog.minsize(min(440, win_w), min(360, win_h))
        dialog.deiconify()
        dialog.grab_set()
        if field_specs:
            entries[field_specs[0][0]].focus_set()
        return dialog

    def on_add_manual_row(self):
        if self.is_running:
            return
        specs = [
            ("Account No", "Account No:", ""),
            ("Search Name", "Borrower Name:", ""),
            ("Invoice No", "Invoice No (optional):", ""),
            ("Invoice Date", "Invoice Date (DD/MM/YYYY):", ""),
            ("Invoice Amount", "Invoice Amount RM (e.g. 113.40):", ""),
            ("Source PDF Path", "Source PDF Path (optional):", ""),
            ("Invoice Pages", "Invoice Pages in PDF (e.g. 1 or 13-14):", ""),
        ]

        def _save(values, dialog):
            acc = values["Account No"]
            name = values["Search Name"]
            if not acc and not name:
                messagebox.showwarning("Input Required", "Please enter an Account No or Borrower Name.", parent=dialog)
                return
            rec = self._normalize_imported_row({
                "Source File": os.path.basename(values["Source PDF Path"]) or "Manual Entry",
                "Source PDF Path": values["Source PDF Path"],
                "Invoice Pages": values["Invoice Pages"],
                "Invoice No": values["Invoice No"],
                "Invoice Date": values["Invoice Date"],
                "Invoice Amount": values["Invoice Amount"],
                "Account No": acc,
                "Search Name": name,
                "Full Borrower Name": name,
            })
            self.cases_data.append(rec)
            self.refresh_table()
            self._save_extracted_excel()
            dialog.destroy()

        self._open_case_form_dialog("Add Case Manually", specs, "Add to Queue", _save)

    def on_double_click_row(self, event):
        if self.is_running:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        if self.tree.identify_column(event.x) == "#1":
            return "break"
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        idx = int(iid)
        row_data = self.cases_data[idx]

        specs = [
            ("Account No", "Account No:", row_data.get("Account No", "")),
            ("Search Name", "Search Name:", row_data.get("Search Name", "")),
            ("Invoice No", "Invoice No:", row_data.get("Invoice No", "")),
            ("Invoice Date", "Invoice Date (DD/MM/YYYY):", row_data.get("Invoice Date", "")),
            ("Invoice Amount", "Invoice Amount RM (e.g. 113.40):", row_data.get("Invoice Amount", "")),
            ("Source PDF Path", "Source PDF Path:", row_data.get("Source PDF Path", "")),
            ("Invoice Pages", "Invoice Pages (e.g. 1 or 13-14):", row_data.get("Invoice Pages", "")),
        ]

        def _update(values, dialog):
            row_data["Account No"] = values["Account No"]
            row_data["Search Name"] = values["Search Name"]
            row_data["Full Borrower Name"] = values["Search Name"]
            row_data["Invoice No"] = values["Invoice No"]
            d_raw = values["Invoice Date"]
            row_data["Invoice Date"] = format_date_dmy(d_raw) if d_raw else ""
            a_raw = values["Invoice Amount"]
            a_cents = parse_amount_cents(a_raw)
            row_data["Invoice Amount"] = format_cents(a_cents) if a_cents is not None else a_raw
            row_data["Source PDF Path"] = values["Source PDF Path"]
            if row_data["Source PDF Path"]:
                row_data["Source File"] = os.path.basename(row_data["Source PDF Path"])
            row_data["Invoice Pages"] = values["Invoice Pages"]
            self.refresh_table()
            self._save_extracted_excel()
            dialog.destroy()

        self._open_case_form_dialog(f"Edit Row #{idx + 1}", specs, "Save Changes", _update)

    def on_delete_selected(self):
        if self.is_running:
            return
        sel = self.tree.selection()
        if not sel:
            return
        indices = sorted([int(i) for i in sel], reverse=True)
        for i in indices:
            del self.cases_data[i]
        self.refresh_table()
        self._save_extracted_excel()

    def on_clear_list(self):
        if self.is_running or not self.cases_data:
            return
        if messagebox.askyesno("Clear Queue", "Remove all cases from the current list?"):
            self.cases_data.clear()
            self.refresh_table()

    def on_stop_automation(self):
        if self.is_running:
            self.stop_requested = True
            self.log("[!] Stop requested by user. Finishing current action...")
            self.set_status("Stopping...", bg="#FFF2F2", fg=THEME["red"])

    def on_start_cpro_check(self):
        if self.is_running:
            return
        if not self.cases_data:
            messagebox.showwarning("Empty Queue", "Please upload PDFs or add cases before starting the C-PRO check.")
            return
        target_indices = [i for i, row in enumerate(self.cases_data) if row.get("_checked", True)]
        if not target_indices:
            messagebox.showwarning("No Rows Selected", "Please check (☑) at least one row in the Cases Queue to check in C-PRO.", parent=self.root)
            return
        if sync_playwright is None:
            messagebox.showerror("Playwright Missing", "Playwright is not installed. Run:\npython -m pip install playwright\npython -m playwright install chromium")
            return

        verify_mode = bool(self.verify_and_upload.get())
        if verify_mode:
            missing_billing_meta = [
                i + 1 for i in target_indices
                if parse_date_flexible(self.cases_data[i].get("Invoice Date")) is None
                or parse_amount_cents(self.cases_data[i].get("Invoice Amount")) is None
            ]
            if missing_billing_meta:
                preview = ", ".join(f"#{n}" for n in missing_billing_meta[:8])
                if len(missing_billing_meta) > 8:
                    preview += f" (+{len(missing_billing_meta) - 8} more)"
                proceed = messagebox.askyesno(
                    "Missing Invoice Date / Amount",
                    f"{len(missing_billing_meta)} selected row(s) ({preview}) are missing Invoice Date or Invoice Amount.\n\n"
                    "Without Date & Amount, billing verification will fail for those rows.\n\n"
                    "Click Yes to continue anyway, or No to cancel and upload PDFs / edit those rows.",
                    parent=self.root,
                )
                if not proceed:
                    return

        active_accounts = []
        for idx, acc in enumerate(self.accounts, start=1):
            if acc["enabled"].get():
                username = acc["user"].get().strip()
                password = acc["pwd"].get().strip()
                if not username:
                    messagebox.showwarning("Missing Username", f"Account #{idx} is enabled but has no username.\n\nPlease enter a username or uncheck/remove it.", parent=self.root)
                    return
                active_accounts.append({
                    "username": username,
                    "password": password,
                    "password_var": acc["pwd"],
                })

        if not active_accounts:
            messagebox.showwarning("No Accounts Selected", "Please enter and enable at least one C-PRO account in the left sidebar.", parent=self.root)
            return

        for acc in active_accounts:
            if not acc["password"]:
                pwd = simpledialog.askstring("Password Required", f"Enter C-PRO password for '{acc['username']}':", show="•", parent=self.root)
                if not pwd:
                    return
                acc["password"] = pwd.strip()
                acc["password_var"].set(acc["password"])

        self._save_accounts_config()

        active_sections = []
        if self.sec_in_progress.get():
            active_sections.append("Case In Progress")
        if self.sec_withhold.get():
            active_sections.append("Case Withhold")
        if self.sec_withdrawn.get():
            active_sections.append("Case Withdrawn")

        if not active_sections:
            messagebox.showwarning("No Sections Selected", "Please select at least one section to check.")
            return

        mode = self.search_mode.get()
        headless = not self.show_browser.get()

        self.is_running = True
        self.stop_requested = False
        self._set_controls_running(True, allow_stop=True)
        self.set_status("Checking C-PRO...", bg="#E8F8EE", fg=THEME["green"])
        self.log(f"Starting C-PRO automation ({'Verify Billing & Upload E-INVOICE' if verify_mode else 'Availability Check Only'}) for {len(target_indices)}/{len(self.cases_data)} row(s)...")

        threading.Thread(
            target=self._cpro_worker,
            args=(active_accounts, active_sections, target_indices, mode, headless, verify_mode),
            daemon=True,
        ).start()

    # --------------------------------------------------------------------------
    # PLAYWRIGHT PORTAL HELPERS
    # --------------------------------------------------------------------------
    def _dismiss_session_timeout_warning(self, page, from_handler=False):
        """If C-PRO shows the 'Session Timeout Warning' popup, click 'Stay' to prevent kickout."""
        if getattr(self, "_handling_timeout_popup", False):
            return False
        try:
            self._handling_timeout_popup = True
            if not from_handler:
                # Use page.evaluate during polling so Locator checks/clicks never deadlock with add_locator_handler
                clicked = page.evaluate(r"""() => {
                    const containers = Array.from(document.querySelectorAll(".swal2-popup, .modal, [role='dialog'], .dx-popup-wrapper, div"));
                    for (const el of containers) {
                        const style = window.getComputedStyle(el);
                        if (!style || style.display === "none" || style.visibility === "hidden" || el.offsetParent === null) {
                            if (style.position !== "fixed") continue;
                        }
                        if (el.offsetWidth === 0 && el.offsetHeight === 0) continue;
                        const txt = (el.innerText || "").trim();
                        if (!/Session\s*Timeout\s*Warning|Your\s+session\s+will\s+expire/i.test(txt)) continue;
                        const btns = Array.from(el.querySelectorAll("button, a, [role='button']"));
                        const stayBtn = btns.find(b => /^\s*Stay\s*$/i.test((b.innerText || "").trim()) && b.offsetWidth > 0 && b.offsetHeight > 0);
                        if (stayBtn) {
                            stayBtn.click();
                            return true;
                        }
                    }
                    return false;
                }""")
                if clicked:
                    self.log("[SESSION] Detected 'Session Timeout Warning' popup — clicked 'Stay' to keep session active.")
                    page.wait_for_timeout(300)
                    self._wait_for_network_quiet(page, timeout=2000)
                return bool(clicked)

            warn_txt = re.compile(r"Session\s*Timeout\s*Warning|Your\s+session\s+will\s+expire", re.I)
            popup = page.locator(
                ".swal2-popup:visible, .modal:visible, [role='dialog']:visible, .dx-popup-wrapper:visible"
            ).filter(has_text=warn_txt).first
            has_popup = popup.count() > 0 and popup.is_visible()
            if not has_popup:
                txt_loc = page.locator("text=/Session\\s*Timeout\\s*Warning|Your\\s+session\\s+will\\s+expire/i").first
                if txt_loc.count() == 0 or not txt_loc.is_visible():
                    return False
                stay_btn = page.locator(
                    "button:has-text('Stay'):visible, a:has-text('Stay'):visible, [role='button']:has-text('Stay'):visible"
                ).first
            else:
                stay_btn = popup.locator(
                    "button:has-text('Stay'), a:has-text('Stay'), [role='button']:has-text('Stay'), button.swal2-confirm:has-text('Stay')"
                ).first

            if stay_btn.count() > 0 and stay_btn.is_visible():
                self.log("[SESSION] Detected 'Session Timeout Warning' popup — clicking 'Stay' to keep session active...")
                stay_btn.click(timeout=3000)
                try:
                    if has_popup:
                        popup.wait_for(state="hidden", timeout=4000)
                    else:
                        page.wait_for_timeout(400)
                except Exception:
                    pass
                self._wait_for_network_quiet(page, timeout=2000)
                return True
        except Exception:
            return False
        finally:
            self._handling_timeout_popup = False
        return False

    def _install_session_timeout_handler(self, page):
        """Register Playwright locator handler so any action blocked by Session Timeout Warning auto-clicks 'Stay'."""
        if not hasattr(page, "add_locator_handler"):
            return
        warn_loc = page.locator(
            ".swal2-popup, .modal, [role='dialog'], .dx-popup-wrapper"
        ).filter(has_text=re.compile(r"Session\s*Timeout\s*Warning|Your\s+session\s+will\s+expire", re.I)).first
        page.add_locator_handler(warn_loc, lambda: self._dismiss_session_timeout_warning(page, from_handler=True))

    @staticmethod
    def _is_portal_role_select_page(page):
        """Return True when C-PRO is showing the 'Welcome Back' portal selector (LAW FIRM, MG VALUER, etc.)."""
        try:
            return bool(page.evaluate("""() => {
                const text = (document.body && (document.body.innerText || document.body.textContent) || '').trim();
                if (!text) return false;
                const upper = text.toUpperCase();
                if (!/\\bLAW\\s*FIRM\\b/.test(upper)) return false;
                const hasRoleList = (
                    upper.includes('MG VALUER')
                    || upper.includes('MG AUCTIONEER')
                    || upper.includes('HP AUCTIONEER')
                );
                const hasWelcome = (
                    upper.includes('WELCOME BACK')
                    && upper.includes('SIGN IN TO GET STARTED')
                );
                return hasRoleList || hasWelcome;
            }"""))
        except Exception:
            return False

    def _click_law_firm_portal_if_present(self, page):
        """Click the 'LAW FIRM' option on the C-PRO portal selector page to enter the login page."""
        if not self._is_portal_role_select_page(page):
            return False
        pattern = re.compile(r"^\s*LAW\s*FIRM\s*$", re.I)
        candidates = (
            page.locator("a, button, [role='button'], [tabindex]").filter(has_text=pattern),
            page.get_by_text(pattern),
        )
        clicked = False
        for loc in candidates:
            try:
                count = min(loc.count(), 10)
            except Exception:
                count = 0
            for idx in range(count):
                el = loc.nth(idx)
                try:
                    if not el.is_visible():
                        continue
                    el.scroll_into_view_if_needed(timeout=2000)
                    el.click(timeout=5000)
                    clicked = True
                    break
                except Exception:
                    continue
            if clicked:
                break
        if not clicked:
            # Fallback JS click in case an Angular/SPA click handler is bound on a non-standard wrapper
            try:
                clicked = bool(page.evaluate("""() => {
                    const els = Array.from(document.querySelectorAll('a, button, [role="button"], li, div, span, p, h1, h2, h3, h4'));
                    for (const el of els) {
                        const txt = (el.innerText || el.textContent || '').trim();
                        if (/^LAW\\s*FIRM$/i.test(txt) && el.offsetParent !== null) {
                            el.click();
                            return true;
                        }
                    }
                    return false;
                }"""))
            except Exception:
                clicked = False
        if not clicked:
            return False

        self.log("[AUTH] Portal role-selection page detected; clicked 'LAW FIRM' to enter login page.")
        try:
            page.wait_for_load_state("domcontentloaded", timeout=10000)
        except Exception:
            pass
        self._wait_for_network_quiet(page, timeout=4000)

        deadline = time.monotonic() + 8.0
        while time.monotonic() < deadline:
            if self.stop_requested:
                break
            try:
                if (
                    page.locator("button:has-text('Sign In')").first.is_visible()
                    or page.locator("input[type='password']").first.is_visible()
                ):
                    break
                if "case-listing" in page.url.lower() and not self._is_portal_role_select_page(page):
                    break
            except Exception:
                pass
            page.wait_for_timeout(250)
        return True

    def _is_kicked_or_unauthorized(self, page):
        try:
            self._dismiss_session_timeout_warning(page)
            if self._is_portal_role_select_page(page):
                return True
            if "login" in page.url.lower() or "auth" in page.url.lower():
                return True
            if page.locator("button:has-text('Sign In')").first.is_visible():
                return True
            toast_error = page.locator("text=/Error! Unauthenticated|Unauthenticated|Session expired|Logged in from another/i")
            if toast_error.count() > 0 and toast_error.first.is_visible():
                return True
            unauth = page.locator("text=/unauthorized|session expired|logged in from another|invalid session|please login again/i")
            if unauth.count() > 0 and unauth.first.is_visible():
                return True
        except Exception:
            return True
        return False

    @staticmethod
    def _wait_for_network_quiet(page, timeout=5000):
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except Exception:
            pass

    @staticmethod
    def _is_fatal_auth_error(exc):
        msg = str(exc or "").lower()
        return "rejected" in msg or "cancelled" in msg

    @staticmethod
    def _section_tab_candidates(page, section):
        exact = re.compile(rf"^\s*{re.escape(section)}(?:\s*\(?\s*[\d,]+\+?\s*\)?)?\s*$", re.IGNORECASE)
        selectors = ("[role='tab']", ".nav-tabs a", ".nav-tabs button", ".nav-link", ".dx-tab", "button", "a", "li")
        seen = set()
        for selector in selectors:
            locator = page.locator(selector)
            try:
                count = min(locator.count(), 200)
            except Exception:
                continue
            for index in range(count):
                item = locator.nth(index)
                try:
                    if not item.is_visible():
                        continue
                    text = re.sub(r"\s+", " ", item.inner_text(timeout=500)).strip()
                    if not exact.match(text):
                        continue
                    key = item.evaluate("el => el.tagName + '|' + el.className + '|' + el.textContent")
                    if key in seen:
                        continue
                    seen.add(key)
                    yield item
                except Exception:
                    continue
        try:
            labels = page.get_by_text(section, exact=True)
            for index in range(min(labels.count(), 50)):
                item = labels.nth(index)
                try:
                    if item.is_visible():
                        yield item
                except Exception:
                    continue
        except Exception:
            return

    def _wait_for_section_table(self, page, section, timeout_ms=45000, action=None):
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
            except Exception:
                pass

        def on_done(request):
            nonlocal last_network
            try:
                if request.resource_type in ("xhr", "fetch"):
                    pending.discard(id(getattr(request, "_impl_obj", request)))
                    last_network = time.monotonic()
            except Exception:
                pass

        page.on("request", on_request)
        page.on("requestfinished", on_done)
        page.on("requestfailed", on_done)
        try:
            if action is not None:
                action()
            while time.monotonic() < deadline:
                if self.stop_requested:
                    raise GridLoadTimeoutError(f"Stopped while waiting for {section}.")
                if self._is_kicked_or_unauthorized(page):
                    raise InvalidCredentialsError(f"Session not authenticated while loading {section}.")

                signature, loading = self._grid_snapshot(page)
                try:
                    headers_ready = page.locator(".dx-datagrid-headers:visible, tr.dx-header-row:visible, [role='columnheader']:visible").count() > 0
                    filter_ready = page.locator("tr.dx-datagrid-filter-row:visible").count() > 0
                    rows_ready = page.locator("tr.dx-data-row:visible").count() > 0
                    empty_ready = page.locator(".dx-datagrid-nodata:visible, .dx-datagrid-rowsview .dx-empty-message:visible").count() > 0
                except Exception:
                    headers_ready = filter_ready = rows_ready = empty_ready = False

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

                if table_ready and not loading and elapsed >= min_wait and stable_for >= 1.25 and (network_quiet if saw_xhr else True):
                    self.log(f"[WAIT] {section} table ready after {elapsed:.1f}s")
                    return
                page.wait_for_timeout(250)

            raise GridLoadTimeoutError(f"Timed out waiting for {section} table.")
        finally:
            for event, handler in (("request", on_request), ("requestfinished", on_done), ("requestfailed", on_done)):
                try:
                    page.remove_listener(event, handler)
                except Exception:
                    pass

    def _switch_section_and_wait(self, page, section):
        self._dismiss_session_timeout_warning(page)
        tab = next(self._section_tab_candidates(page, section), None)
        if tab is None:
            if self._is_kicked_or_unauthorized(page):
                raise InvalidCredentialsError(f"Session not authenticated before switching to '{section}'.")
            raise GridLoadTimeoutError(f"Could not find a visible '{section}' tab.")
        self.log(f"[WAIT] Clicking section tab '{section}'...")
        self._last_searched_column = None

        def _click_tab():
            try:
                tab.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            tab.click(timeout=8000)

        self._wait_for_section_table(page, section, timeout_ms=45000, action=_click_tab)

    def _prompt_for_correct_password(self, acc, reason):
        result = {"password": None}
        ready = threading.Event()

        def ask():
            try:
                result["password"] = simpledialog.askstring(
                    "C-PRO Password Rejected",
                    f"C-PRO rejected the password for {acc['username']}.\n{reason}\n\nEnter correct password:",
                    show="•",
                    parent=self.root,
                )
            finally:
                ready.set()

        self.root.after(0, ask)
        while not ready.wait(0.2):
            if self.stop_requested:
                return None
        password = (result["password"] or "").strip()
        if password:
            acc["password"] = password
            if acc.get("password_var") is not None:
                self.root.after(0, lambda v=acc["password_var"], p=password: v.set(p))
        return password or None

    @staticmethod
    def _login_rejection_text(page):
        try:
            body = page.locator("body").inner_text(timeout=1500).casefold()
        except Exception:
            body = ""
        keywords = (
            "invalid password", "incorrect password", "invalid username", "invalid credentials",
            "login failed", "authentication failed", "password is incorrect", "wrong password",
        )
        for keyword in keywords:
            if keyword in body:
                return keyword
        return ""

    def _login_and_restore_section(self, page, acc, section=None):
        username = acc["username"]
        password = acc["password"]
        self.log(f"[AUTH] Logging into {username}...")
        self.set_status(f"Authenticating {username}...", bg="#EBF5FF", fg=THEME["blue"])

        try:
            close_btns = page.locator(".toast-close-button, .close, button[aria-label='Close'], .swal2-close")
            for i in range(min(close_btns.count(), 20)):
                if close_btns.nth(i).is_visible():
                    close_btns.nth(i).click(timeout=500)
        except Exception:
            pass

        for login_attempt in range(2):
            self._click_law_firm_portal_if_present(page)
            if not page.locator("button:has-text('Sign In')").first.is_visible():
                for nav_try in range(2):
                    try:
                        page.goto(CASE_LISTING_URL, wait_until="domcontentloaded", timeout=45000)
                        break
                    except PlaywrightTimeoutError:
                        if nav_try == 1:
                            raise
                        self.log(f"[AUTH] Navigation timed out for {username}; retrying once...")
                        page.wait_for_timeout(1000)
                self._wait_for_network_quiet(page, timeout=4000)

                auth_settle_deadline = time.monotonic() + 6.0
                while time.monotonic() < auth_settle_deadline:
                    if self.stop_requested:
                        break
                    if self._click_law_firm_portal_if_present(page):
                        continue
                    try:
                        if page.locator("button:has-text('Sign In')").first.is_visible():
                            break
                        if (
                            "case-listing" in page.url.lower()
                            and not self._is_kicked_or_unauthorized(page)
                            and (
                                page.locator(".dx-datagrid:visible").count() > 0
                                or page.locator("text=/Case In Progress|Case Withdrawn|Case Withhold/i").count() > 0
                            )
                        ):
                            break
                    except Exception:
                        pass
                    page.wait_for_timeout(250)

            sign_in = page.locator("button:has-text('Sign In')").first
            user_input = page.locator("input[placeholder*='Username' i], input[name*='user' i], input[type='text']").first
            if sign_in.is_visible():
                if not user_input.is_visible():
                    user_input = page.locator("input:not([type]), input[type='email']").first
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
                    if rejection:
                        break
                    if (
                        "case-listing" in page.url.lower()
                        and not self._is_portal_role_select_page(page)
                        and not page.locator("input[type='password']").first.is_visible()
                    ):
                        break
                    page.wait_for_timeout(250)

                rejection = rejection or self._login_rejection_text(page)
                if not rejection:
                    try:
                        if (
                            page.locator("button:has-text('Sign In')").first.is_visible()
                            and page.locator("input[type='password']").first.is_visible()
                        ):
                            rejection = "C-PRO remained on the sign-in form"
                    except Exception:
                        pass

                if rejection:
                    self.log(f"[AUTH ERROR] C-PRO rejected login for {username}: {rejection}.")
                    if login_attempt == 1:
                        raise InvalidCredentialsError(f"Corrected password was also rejected for {username}.")
                    corrected = self._prompt_for_correct_password(acc, rejection)
                    if not corrected:
                        raise InvalidCredentialsError(f"Login cancelled for {username}.")
                    password = corrected
                    continue

                try:
                    page.wait_for_url("**/case-listing**", timeout=8000)
                except PlaywrightTimeoutError:
                    pass
                self._wait_for_network_quiet(page, timeout=4000)
                if self._is_kicked_or_unauthorized(page):
                    raise InvalidCredentialsError(f"Could not confirm authenticated session for {username}.")
                break
            else:
                if self._is_kicked_or_unauthorized(page):
                    raise InvalidCredentialsError(f"Sign-in form not available for {username}.")
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
        aliases = {
            "accountno": {"accountno", "accountnumber", "account"},
            "primaryborrowername": {"primaryborrowername", "borrowername", "primaryname"},
        }
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
                    except Exception:
                        continue
            header_cells = grid.locator(".dx-datagrid-headers [role='columnheader']")
            for col_idx in range(header_cells.count()):
                cell = header_cells.nth(col_idx)
                try:
                    text = re.sub(r"[^a-z0-9]+", "", cell.inner_text(timeout=700).casefold())
                    aria = re.sub(r"[^a-z0-9]+", "", (cell.get_attribute("aria-label") or "").casefold())
                    candidates = {x for x in (text, aria) if x}
                    if any(candidate in wanted_set or any(candidate.endswith(alias) for alias in wanted_set) for candidate in candidates):
                        return grid, col_idx
                except Exception:
                    continue
        return None, None

    def _get_grid_filter_input(self, page, column_index, column_name, grid=None):
        filter_rows = (grid if grid is not None else page).locator("tr.dx-datagrid-filter-row:visible")
        if filter_rows.count() == 0:
            raise RuntimeError("The current C-PRO grid has no filter row")

        filter_row = None
        for row_idx in range(filter_rows.count()):
            candidate = filter_rows.nth(row_idx)
            if candidate.locator("td, th").count() > column_index:
                filter_row = candidate
                break
        if filter_row is None:
            raise RuntimeError(f"Filter row does not contain column '{column_name}'")

        cell = filter_row.locator("td, th").nth(column_index)

        def visible_input_in_cell():
            inputs = cell.locator("input.dx-texteditor-input, input[type='text'], input:not([type])")
            for input_idx in range(inputs.count()):
                candidate = inputs.nth(input_idx)
                try:
                    if candidate.is_visible() and candidate.is_enabled():
                        return candidate
                except Exception:
                    continue
            return None

        target = visible_input_in_cell()
        if target is not None:
            return target

        for selector in (".dx-icon-search", "button", "[role='button']"):
            controls = cell.locator(selector)
            for control_idx in range(controls.count()):
                control = controls.nth(control_idx)
                try:
                    if control.is_visible() and control.is_enabled():
                        control.click(timeout=1200)
                        page.wait_for_timeout(150)
                        target = visible_input_in_cell()
                        if target is not None:
                            return target
                        overlay_inputs = page.locator(".dx-overlay-content input.dx-texteditor-input:visible, .dx-popup-content input.dx-texteditor-input:visible")
                        if overlay_inputs.count() > 0:
                            return overlay_inputs.last
                except Exception:
                    continue

        try:
            cell.click(timeout=1200)
            page.wait_for_timeout(150)
        except Exception:
            pass
        target = visible_input_in_cell()
        if target is not None:
            return target

        overlay_inputs = page.locator(".dx-overlay-content input.dx-texteditor-input:visible, .dx-popup-content input.dx-texteditor-input:visible")
        if overlay_inputs.count() > 0:
            return overlay_inputs.last
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
                    except Exception:
                        pass
            if cleared:
                first_input = page.locator("tr.dx-datagrid-filter-row:visible input.dx-texteditor-input:visible").first
                if first_input.count() > 0:
                    first_input.press("Enter")
                    self._wait_for_network_quiet(page, timeout=2000)
                    page.wait_for_timeout(350)
        except Exception:
            pass

    @staticmethod
    def _grid_cell_matches(column_name, cell_text, search_value):
        cell = re.sub(r"\s+", " ", str(cell_text or "")).strip()
        query = re.sub(r"\s+", " ", str(search_value or "")).strip()
        if not cell or not query:
            return False

        if column_name.casefold() == "account no":
            cell_digits = re.sub(r"\D", "", cell)
            query_digits = re.sub(r"\D", "", query)
            if cell_digits and query_digits:
                return cell_digits.lstrip("0") == query_digits.lstrip("0") or cell_digits == query_digits
            return cell.casefold() == query.casefold()
        return query.casefold() in cell.casefold()

    def _grid_snapshot(self, page):
        signature_parts = []
        try:
            rows = page.locator("tr.dx-data-row:visible")
            for row_idx in range(min(rows.count(), 100)):
                try:
                    signature_parts.append(rows.nth(row_idx).inner_text(timeout=1000).strip())
                except Exception:
                    signature_parts.append("<row-refreshing>")
        except Exception:
            signature_parts.append("<grid-refreshing>")

        for selector in (".dx-datagrid-nodata:visible", ".dx-datagrid-rowsview .dx-empty-message:visible"):
            try:
                labels = page.locator(selector)
                for i in range(min(labels.count(), 3)):
                    text = labels.nth(i).inner_text(timeout=500).strip()
                    if text:
                        signature_parts.append("EMPTY:" + text)
            except Exception:
                pass

        loading = False
        for selector in (".dx-loadpanel:visible", ".dx-loadindicator:visible", ".dx-datagrid-loading:visible", ".dx-datagrid[aria-busy='true']"):
            try:
                loc = page.locator(selector)
                for i in range(min(loc.count(), 5)):
                    try:
                        if loc.nth(i).is_visible():
                            loading = True
                            break
                    except Exception:
                        continue
                if loading:
                    break
            except Exception:
                continue
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
            except Exception:
                pass

        def on_request_done(request):
            try:
                if request.resource_type in ("xhr", "fetch"):
                    pending_xhr.discard(id(getattr(request, "_impl_obj", request)))
                    last_network_activity[0] = time.monotonic()
            except Exception:
                pass

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
                if self.stop_requested:
                    raise RuntimeError("Stopped by user while waiting for the C-PRO grid")
                if self._is_kicked_or_unauthorized(page):
                    raise InvalidCredentialsError("Session expired or unauthenticated during grid refresh.")

                signature, loading = self._grid_snapshot(page)
                now = time.monotonic()
                elapsed = now - started

                if signature != last_signature:
                    last_signature = signature
                    stable_samples = 0
                else:
                    stable_samples += 1
                if signature != before_signature:
                    grid_change_seen = True
                if loading:
                    stable_samples = 0
                    loading_seen = True

                xhr_seen = bool(seen_xhr)
                network_quiet = not pending_xhr and (now - last_network_activity[0]) >= 0.9
                changed = signature != before_signature

                if xhr_seen:
                    min_elapsed = 1.0
                    network_ready = network_quiet
                    required_stable_samples = 4
                else:
                    min_elapsed = 1.5 if changed else 3.5
                    network_ready = True
                    required_stable_samples = 5

                if elapsed >= min_elapsed and stable_samples >= required_stable_samples and not loading and network_ready:
                    if require_activity and not (xhr_seen or grid_change_seen or loading_seen):
                        current_sig, _ = self._grid_snapshot(page)
                        if any("EMPTY:" in s or "No data" in s for s in current_sig):
                            return elapsed
                        raise GridSearchError(f"No XHR/loading/row change observed after submitting {description}.")
                    return elapsed
                page.wait_for_timeout(250)

            raise GridLoadTimeoutError(f"C-PRO grid did not finish '{description}' within {timeout_ms // 1000}s.")
        finally:
            for event, handler in (("request", on_request), ("requestfinished", on_request_done), ("requestfailed", on_request_done)):
                try:
                    page.remove_listener(event, handler)
                except Exception:
                    pass

    @staticmethod
    def _extract_account_candidates(acc_no, alt_acc_no=""):
        """Extract all distinct account numbers from primary and alternate account fields."""
        candidates = []
        seen = set()

        def _add_from_raw(raw_val):
            raw = str(raw_val or "").strip()
            if not raw or raw == "-":
                return
            digit_runs = re.findall(r"\d{7,20}", raw)
            if digit_runs:
                for num in digit_runs:
                    key = num.lstrip("0") or num
                    if key not in seen:
                        seen.add(key)
                        candidates.append(num)
            else:
                for part in re.split(r"[,;/&]+", raw):
                    p = part.strip()
                    if p and p.casefold() not in seen:
                        seen.add(p.casefold())
                        candidates.append(p)

        _add_from_raw(acc_no)
        _add_from_raw(alt_acc_no)
        return candidates

    def _go_to_grid_page(self, page, target_page_num):
        """Advance DevExtreme Case Listing pager from page 1 to target_page_num (1-based)."""
        if target_page_num <= 1:
            return
        for _ in range(target_page_num - 1):
            grid = page.locator(".dx-datagrid:visible").first
            next_btn = grid.locator(".dx-pager .dx-next-button:not(.dx-button-disable):visible")
            if next_btn.count() == 0:
                break
            before_sig, _ = self._grid_snapshot(page)
            next_btn.first.click(timeout=3000)
            self._wait_for_network_quiet(page, timeout=3000)
            self._wait_for_any_grid_quiet(page, timeout_ms=8000)
            after_sig, _ = self._grid_snapshot(page)
            if after_sig == before_sig:
                break

    def _search_grid_with_relogin(self, page, acc, section, column_name, val):
        """Return list of (row_index, row_summary_str, grid_page_num) for all matching rows across grid pages."""
        if not val:
            return []
        non_kickout_errors = 0
        for attempt in range(1, MAX_RETRIES_ON_KICKOUT + 1):
            if self.stop_requested:
                return []
            try:
                if self._is_kicked_or_unauthorized(page):
                    self.log(f"[!] Session unauthenticated before search for '{val}'. Re-logging...")
                    self._login_and_restore_section(page, acc, section)
                    non_kickout_errors = 0

                if getattr(self, "_last_searched_column", None) != column_name:
                    self._clear_all_filters(page)
                    self._last_searched_column = column_name

                grid, column_index = self._find_grid_column_index(page, column_name)
                if grid is None or column_index is None:
                    if self._is_kicked_or_unauthorized(page):
                        raise InvalidCredentialsError("Unauthenticated while finding column.")
                    raise RuntimeError(f"Column '{column_name}' was not found in section '{section}'.")

                target_box = self._get_grid_filter_input(page, column_index, column_name, grid=grid)
                try:
                    existing_val = target_box.input_value()
                except Exception:
                    existing_val = ""
                if existing_val.strip() == str(val).strip():
                    target_box.fill("")
                    target_box.press("Enter")
                    self._wait_for_network_quiet(page, timeout=1500)
                    page.wait_for_timeout(300)

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
                    non_kickout_errors = 0
                    continue

                matched = []
                max_grid_pages = 10
                for grid_page in range(1, max_grid_pages + 1):
                    rows = grid.locator("tr.dx-data-row:visible")
                    for row_idx in range(rows.count()):
                        row = rows.nth(row_idx)
                        cells = row.locator("td, th")
                        if cells.count() <= column_index:
                            continue
                        searched_cell_text = cells.nth(column_index).inner_text().strip()
                        if not self._grid_cell_matches(column_name, searched_cell_text, val):
                            continue
                        row_text = row.inner_text().strip()
                        if row_text and "No data" not in row_text:
                            summary = " | ".join(part.strip() for part in row_text.split("\t") if part.strip())
                            matched.append((row_idx, summary, grid_page))

                    next_btn = grid.locator(".dx-pager .dx-next-button:not(.dx-button-disable):visible")
                    if next_btn.count() == 0:
                        break
                    before_sig, _ = self._grid_snapshot(page)
                    next_btn.first.click(timeout=3000)
                    self._wait_for_network_quiet(page, timeout=3000)
                    self._wait_for_any_grid_quiet(page, timeout_ms=8000)
                    after_sig, _ = self._grid_snapshot(page)
                    if after_sig == before_sig:
                        break

                # If we advanced past page 1 and found matches on page 1, return to page 1 if needed
                if matched and matched[0][2] == 1 and grid.locator(".dx-pager .dx-page").count() > 1:
                    first_page_btn = grid.locator(".dx-pager .dx-page").filter(has_text=re.compile(r"^\s*1\s*$")).first
                    if first_page_btn.count() > 0 and "dx-selection" not in (first_page_btn.get_attribute("class") or ""):
                        first_page_btn.click(timeout=3000)
                        self._wait_for_network_quiet(page, timeout=3000)
                        self._wait_for_any_grid_quiet(page, timeout_ms=6000)

                return matched

            except InvalidCredentialsError as e:
                if "rejected" in str(e).lower() or "cancelled" in str(e).lower():
                    raise e
                if attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session unauthenticated (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try:
                        self._login_and_restore_section(page, acc, section)
                        non_kickout_errors = 0
                    except Exception as login_err:
                        self.log(f"[!] Re-login failed: {login_err}")
                    continue
                raise
            except (GridLoadTimeoutError, GridSearchError) as e:
                if self._is_kicked_or_unauthorized(page) and attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session interrupted (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try:
                        self._login_and_restore_section(page, acc, section)
                        non_kickout_errors = 0
                    except Exception as login_err:
                        self.log(f"[!] Re-login failed: {login_err}")
                    continue
                non_kickout_errors += 1
                if non_kickout_errors < MAX_INLINE_SEARCH_RETRIES and not self.stop_requested:
                    self.log(f"[!] Search hiccup for '{val}' in {section} ({e}). Resetting grid & retrying ({non_kickout_errors}/{MAX_INLINE_SEARCH_RETRIES - 1})...")
                    self._last_searched_column = None
                    self._clear_all_filters(page)
                    self._switch_section_and_wait(page, section)
                    continue
                raise
            except Exception as e:
                if self._is_kicked_or_unauthorized(page) and attempt < MAX_RETRIES_ON_KICKOUT:
                    self.log(f"[!] Session interruption (Attempt {attempt}/{MAX_RETRIES_ON_KICKOUT}). Re-logging...")
                    try:
                        self._login_and_restore_section(page, acc, section)
                        non_kickout_errors = 0
                    except Exception as login_err:
                        self.log(f"[!] Re-login error: {login_err}")
                    continue
                non_kickout_errors += 1
                if non_kickout_errors < MAX_INLINE_SEARCH_RETRIES and not self.stop_requested:
                    self.log(f"[!] Unexpected grid error for '{val}' in {section} ({e}). Resetting grid & retrying...")
                    self._last_searched_column = None
                    try:
                        self._clear_all_filters(page)
                        self._switch_section_and_wait(page, section)
                    except Exception as reset_err:
                        raise GridSearchError(f"Could not complete '{column_name}' search for '{val}': {reset_err}") from reset_err
                    continue
                self.log(f"[ERROR] Search failed for '{column_name}' = '{val}' in {section}: {e}.")
                raise GridSearchError(f"Could not complete '{column_name}' search for '{val}': {e}") from e
        return []

    # --------------------------------------------------------------------------
    # CASE DETAILS -> BILLINGS VERIFICATION -> LEGAL DETAILS -> UPLOAD E-INVOICE
    # --------------------------------------------------------------------------
    def _is_on_case_details(self, page):
        try:
            url = page.url.lower()
            if re.search(r"/case-listing/\d+", url) or "case-details" in url:
                return True
            has_heading = page.locator("text=/^Case Details$/i").count() > 0 and page.locator("text=/^Case Details$/i").first.is_visible()
            has_billings_tab = next(self._section_tab_candidates(page, "Billings"), None) is not None
            has_legal_tab = next(self._section_tab_candidates(page, "Legal Details"), None) is not None
            return has_heading or (has_billings_tab and has_legal_tab)
        except Exception:
            return False

    def _wait_for_case_details_ready(self, page, timeout_ms=15000):
        deadline = time.monotonic() + timeout_ms / 1000.0
        while time.monotonic() < deadline:
            if self.stop_requested:
                raise GridLoadTimeoutError("Stopped while opening Case Details.")
            if self._is_kicked_or_unauthorized(page):
                raise InvalidCredentialsError("Session expired while opening Case Details.")
            if self._is_on_case_details(page):
                self._wait_for_network_quiet(page, timeout=3000)
                page.wait_for_timeout(350)
                return True
            page.wait_for_timeout(200)
        return False

    def _open_case_from_grid_row(self, page, row_index):
        grid = page.locator(".dx-datagrid:visible").first
        rows = grid.locator("tr.dx-data-row:visible")
        if row_index >= rows.count():
            raise RuntimeError(f"Row #{row_index} is no longer visible in Case Listing grid.")
        row = rows.nth(row_index)
        row.scroll_into_view_if_needed(timeout=3000)

        # Strategy 1: Click explicit link/anchor or clickable cell inside the row
        triggers = []
        links = row.locator("a, [role='link']")
        for i in range(min(links.count(), 6)):
            triggers.append(("click", links.nth(i)))
        cells = row.locator("td")
        if cells.count() > 1:
            triggers.append(("click", cells.nth(1)))
        if cells.count() > 4:
            triggers.append(("click", cells.nth(4)))
        triggers.append(("dblclick", row))
        triggers.append(("click", row))

        for mode, loc in triggers:
            try:
                if not loc.is_visible():
                    continue
                if mode == "dblclick":
                    loc.dblclick(timeout=3000)
                else:
                    loc.click(timeout=3000)
                if self._wait_for_case_details_ready(page, timeout_ms=4500):
                    return
            except InvalidCredentialsError:
                raise
            except Exception:
                continue

        raise RuntimeError("Could not enter Case Details page from the matched Case Listing row.")

    def _switch_case_tab(self, page, tab_name):
        self._dismiss_session_timeout_warning(page)
        tab = next(self._section_tab_candidates(page, tab_name), None)
        if tab is None:
            if self._is_kicked_or_unauthorized(page):
                raise InvalidCredentialsError(f"Session expired before switching to '{tab_name}' tab.")
            raise RuntimeError(f"Could not find '{tab_name}' tab on Case Details page.")
        tab.scroll_into_view_if_needed(timeout=2000)
        tab.click(timeout=5000)
        self._wait_for_network_quiet(page, timeout=4000)
        page.wait_for_timeout(450)
        if self._is_kicked_or_unauthorized(page):
            raise InvalidCredentialsError(f"Session expired switching to '{tab_name}' tab.")

    def _wait_for_any_grid_quiet(self, page, timeout_ms=15000):
        deadline = time.monotonic() + timeout_ms / 1000.0
        last_sig = None
        stable = 0
        while time.monotonic() < deadline:
            if self.stop_requested:
                raise GridLoadTimeoutError("Stopped while waiting for table.")
            if self._is_kicked_or_unauthorized(page):
                raise InvalidCredentialsError("Session expired while waiting for table.")
            sig, loading = self._grid_snapshot(page)
            try:
                grid_visible = page.locator(".dx-datagrid:visible .dx-datagrid-headers:visible").count() > 0
            except Exception:
                grid_visible = False
            if grid_visible and sig == last_sig and not loading:
                stable += 1
            else:
                last_sig = sig
                stable = 0
            if grid_visible and stable >= 3 and not loading:
                return
            page.wait_for_timeout(200)

    def _read_billing_history_rows(self, page):
        """Read all non-empty rows across pages from the Billing History DevExtreme grid."""
        self._wait_for_any_grid_quiet(page, timeout_ms=15000)

        # Ensure 'All' sub-tab under Billing History is active
        all_tab = next(self._section_tab_candidates(page, "All"), None)
        if all_tab is not None:
            try:
                all_tab.click(timeout=3000)
                self._wait_for_network_quiet(page, timeout=3000)
                self._wait_for_any_grid_quiet(page, timeout_ms=10000)
            except Exception:
                pass

        # Locate the Billing History grid
        target_grid = None
        col_map = {}
        grids = page.locator(".dx-datagrid:visible")
        for g_idx in range(grids.count()):
            g = grids.nth(g_idx)
            hdrs = g.locator(".dx-datagrid-headers [role='columnheader'], .dx-datagrid-headers tr.dx-header-row td, .dx-datagrid-headers tr.dx-header-row th")
            temp_map = {}
            for c_idx in range(hdrs.count()):
                try:
                    raw_h = re.sub(r"[^a-z0-9]+", "", hdrs.nth(c_idx).inner_text(timeout=500).casefold())
                    if raw_h:
                        temp_map[raw_h] = c_idx
                except Exception:
                    continue
            if "billingdate" in temp_map or "totalamt" in temp_map:
                target_grid = g
                col_map = temp_map
                break

        if target_grid is None:
            raise RuntimeError("Billing History grid was not found on the Billings tab.")

        idx_ref = col_map.get("billingref")
        idx_date = col_map.get("billingdate")
        idx_amt = col_map.get("totalamt")
        idx_status = col_map.get("status")
        if idx_date is None or idx_amt is None:
            raise RuntimeError(f"Billing History missing expected columns (found keys: {sorted(col_map.keys())}).")

        collected = []
        max_pages = 25
        for page_turn in range(max_pages):
            rows = target_grid.locator("tr.dx-data-row:visible")
            for r_i in range(rows.count()):
                cells = rows.nth(r_i).locator("td")
                c_count = cells.count()
                if c_count <= max(idx_date, idx_amt):
                    continue
                try:
                    date_txt = (cells.nth(idx_date).inner_text(timeout=800) or cells.nth(idx_date).text_content(timeout=800) or "").strip()
                    amt_txt = (cells.nth(idx_amt).inner_text(timeout=800) or cells.nth(idx_amt).text_content(timeout=800) or "").strip()
                    status_txt = (cells.nth(idx_status).inner_text(timeout=800) or cells.nth(idx_status).text_content(timeout=800) or "").strip() if idx_status is not None and idx_status < c_count else ""
                    ref_txt = (cells.nth(idx_ref).inner_text(timeout=800) or cells.nth(idx_ref).text_content(timeout=800) or "").strip() if idx_ref is not None and idx_ref < c_count else ""
                except Exception:
                    continue
                collected.append({
                    "billing_ref": ref_txt,
                    "billing_date_raw": date_txt,
                    "date_obj": parse_date_flexible(date_txt),
                    "amount_raw": amt_txt,
                    "amount_cents": parse_amount_cents(amt_txt),
                    "status": status_txt,
                })

            # Check if DevExtreme pager has an enabled Next button or next numbered page
            next_btn = target_grid.locator(".dx-pager .dx-next-button:not(.dx-button-disable):visible")
            if next_btn.count() == 0:
                break
            try:
                before_sig, _ = self._grid_snapshot(page)
                next_btn.first.click(timeout=3000)
                self._wait_for_network_quiet(page, timeout=3000)
                self._wait_for_any_grid_quiet(page, timeout_ms=8000)
                after_sig, _ = self._grid_snapshot(page)
                if after_sig == before_sig:
                    break
            except Exception:
                break

        return collected

    def _billing_grid_is_visible(self, page):
        try:
            grids = page.locator(".dx-datagrid:visible")
            for g_idx in range(grids.count()):
                g = grids.nth(g_idx)
                hdr_text = re.sub(r"[^a-z0-9]+", "", g.locator(".dx-datagrid-headers:visible").inner_text(timeout=500).casefold())
                if "billingdate" in hdr_text or "totalamt" in hdr_text:
                    return True
        except Exception:
            pass
        return False

    def _verify_case_billings(self, page, case):
        """Switch to Billings tab and verify invoice date & amount."""
        target_date = parse_date_flexible(case.get("Invoice Date"))
        target_cents = parse_amount_cents(case.get("Invoice Amount"))
        if target_date is None or target_cents is None:
            return False, "Missing Invoice Date or Invoice Amount"

        billings_ready = False
        for tab_try in range(2):
            self._switch_case_tab(page, "Billings")
            deadline = time.monotonic() + (6.0 if tab_try == 0 else 10.0)
            while time.monotonic() < deadline:
                if self.stop_requested:
                    raise GridLoadTimeoutError("Stopped while waiting for Billing History.")
                if self._is_kicked_or_unauthorized(page):
                    raise InvalidCredentialsError("Session expired while loading Billing History.")
                if self._billing_grid_is_visible(page):
                    billings_ready = True
                    break
                try:
                    vis_heading = page.locator("h1:visible, h2:visible, h3:visible, h4:visible, h5:visible, span:visible, div:visible").filter(
                        has_text=re.compile(r"^\s*Billing\s+History\s*$", re.I)
                    )
                    if vis_heading.count() > 0 and page.locator(".dx-datagrid:visible").count() > 0:
                        billings_ready = True
                        break
                except Exception:
                    pass
                page.wait_for_timeout(250)
            if billings_ready:
                break
            self.log("  [!] Billings tab did not display Billing History yet; retrying tab click...")

        if not billings_ready:
            raise RuntimeError("Billing History grid did not become visible on the Billings tab.")

        billing_rows = self._read_billing_history_rows(page)
        return find_billing_match(billing_rows, target_date, target_cents)

    def _close_preview_modal(self, page, modal):
        """Close the Preview Document modal cleanly."""
        try:
            close_btn = modal.locator(
                "button.close, .close, button[aria-label='Close'], .dx-closebutton, button:has-text('×'), a:has-text('×'), span:has-text('×')"
            ).first
            if close_btn.count() > 0 and close_btn.is_visible():
                close_btn.click(timeout=3000)
                modal.wait_for(state="hidden", timeout=4000)
                return
        except Exception:
            pass
        try:
            page.keyboard.press("Escape")
            modal.wait_for(state="hidden", timeout=3000)
        except Exception:
            pass

    def _preview_additional_doc_row(self, page, row_index_in_page):
        """
        Click the Eye (Preview) icon on an Additional Documents row, extract invoice identifiers
        from the Preview Document modal (PDF network response, iframe/embed fetch, or modal screenshot OCR),
        fall back to Download icon if needed, and return parsed preview_info dict.
        """
        import base64

        captured_pdfs = []

        def _on_response(resp):
            try:
                if resp.status != 200:
                    return
                ctype = (resp.headers.get("content-type") or "").lower()
                if "pdf" in ctype or resp.url.lower().split("?")[0].endswith(".pdf"):
                    body = resp.body()
                    if body and body.startswith(b"%PDF"):
                        captured_pdfs.append(body)
            except Exception:
                pass

        page.on("response", _on_response)
        preview_info = {}
        try:
            self._dismiss_session_timeout_warning(page)
            # Locate and click the Eye button inside the target row
            clicked_eye = page.evaluate(r"""(trIdx) => {
                const headingNodes = Array.from(document.querySelectorAll("*"))
                    .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
                let container = null;
                for (const h of headingNodes) {
                    let cur = h.parentElement;
                    for (let d = 0; d < 8 && cur && cur !== document.body; d++) {
                        if ((cur.querySelector("table, .dx-datagrid") || /Document Name/i.test(cur.innerText || '')) && !/Court Suit Number/i.test(cur.innerText || '')) {
                            container = cur;
                            break;
                        }
                        cur = cur.parentElement;
                    }
                    if (container) break;
                }
                if (!container) return false;
                const dataRows = Array.from(container.querySelectorAll("tbody tr, tr.dx-data-row"))
                    .filter(tr => tr.querySelectorAll("td").length >= 3);
                const tr = dataRows[trIdx];
                if (!tr) return false;

                const tds = tr.querySelectorAll("td");
                const actionTd = tds[tds.length - 1];
                const explicit = actionTd.querySelector(
                    "[class*='eye' i], [title*='preview' i], [title*='view' i], [aria-label*='preview' i], [aria-label*='view' i]"
                );
                if (explicit) {
                    const clickable = explicit.closest("button, a, [role='button']") || explicit;
                    clickable.click();
                    return true;
                }
                const fallback = actionTd.querySelector("button, a, [role='button'], i, svg");
                if (fallback) {
                    fallback.click();
                    return true;
                }
                return false;
            }""", row_index_in_page)

            if not clicked_eye:
                return {}

            modal = page.locator(".modal:visible, [role='dialog']:visible, .dx-popup-wrapper:visible").filter(
                has_text=re.compile(r"Preview\s*Document", re.I)
            ).first
            try:
                modal.wait_for(state="visible", timeout=7000)
            except Exception:
                # Fallback if modal has no "Preview Document" heading
                modal = page.locator(".modal:visible, [role='dialog']:visible").filter(
                    has=page.locator("iframe, embed, object, canvas, img")
                ).first
                modal.wait_for(state="visible", timeout=3000)

            self._wait_for_network_quiet(page, timeout=3500)
            page.wait_for_timeout(500)

            # Layer 1: PDF bytes captured directly from network response
            if captured_pdfs:
                preview_info = self.pdf_engine.extract_preview_info_from_pdf_bytes(captured_pdfs[-1])

            # Layer 2: Read iframe / embed / object src (data URL, blob/http fetch, or text layer)
            if not (preview_info.get("inv_no") or preview_info.get("inv_date")):
                embed_srcs = modal.evaluate(r"""async (modalEl) => {
                    const out = [];
                    const els = Array.from(modalEl.querySelectorAll("iframe, embed, object"));
                    for (const el of els) {
                        const rect = el.getBoundingClientRect();
                        if (rect.width === 0 && rect.height === 0) continue;
                        const src = el.getAttribute("src") || el.getAttribute("data") || "";
                        let b64 = null;
                        let frameText = "";
                        try {
                            if (el.contentDocument && el.contentDocument.body) {
                                frameText = el.contentDocument.body.innerText || "";
                            }
                        } catch (e) {}
                        if (src && !src.startsWith("data:")) {
                            try {
                                const r = await fetch(src);
                                if (r.ok) {
                                    const buf = await r.arrayBuffer();
                                    const bytes = new Uint8Array(buf);
                                    if (bytes.length > 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
                                        let bin = "";
                                        for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
                                        b64 = btoa(bin);
                                    }
                                }
                            } catch (e) {}
                        }
                        out.push({ src, b64, frameText });
                    }
                    return out;
                }""")
                for item in embed_srcs:
                    src = item.get("src") or ""
                    if src.startswith("data:application/pdf;base64,"):
                        pdf_bytes = base64.b64decode(src.split(",", 1)[1])
                        preview_info = self.pdf_engine.extract_preview_info_from_pdf_bytes(pdf_bytes)
                        break
                    if item.get("b64"):
                        pdf_bytes = base64.b64decode(item["b64"])
                        preview_info = self.pdf_engine.extract_preview_info_from_pdf_bytes(pdf_bytes)
                        break
                    if item.get("frameText"):
                        parsed = self.pdf_engine.extract_preview_info_from_text(item["frameText"])
                        if parsed.get("inv_no") or parsed.get("inv_date"):
                            preview_info = parsed
                            break

            # Layer 3: Screenshot the Preview Document modal and run OCR
            if not (preview_info.get("inv_no") or preview_info.get("inv_date")):
                with tempfile.TemporaryDirectory() as td:
                    shot_path = os.path.join(td, "preview_modal.png")
                    modal.screenshot(path=shot_path)
                    preview_info = self.pdf_engine.extract_preview_info_from_image(shot_path)

            self._close_preview_modal(page, modal)
        except Exception as e:
            self.log(f"  [!] Preview inspection notice: {e}")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
        finally:
            try:
                page.remove_listener("response", _on_response)
            except Exception:
                pass

        # Layer 4: Fallback to Download icon if Preview yielded no Invoice No or Date
        if not (preview_info.get("inv_no") or preview_info.get("inv_date")):
            try:
                dl_btn = page.evaluate_handle(r"""(trIdx) => {
                    const headingNodes = Array.from(document.querySelectorAll("*"))
                        .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
                    for (const h of headingNodes) {
                        let cur = h.parentElement;
                        for (let d = 0; d < 8 && cur && cur !== document.body; d++) {
                            if ((cur.querySelector("table, .dx-datagrid") || /Document Name/i.test(cur.innerText || '')) && !/Court Suit Number/i.test(cur.innerText || '')) {
                                const dataRows = Array.from(cur.querySelectorAll("tbody tr, tr.dx-data-row"))
                                    .filter(tr => tr.querySelectorAll("td").length >= 3);
                                const tr = dataRows[trIdx];
                                if (!tr) return null;
                                const tds = tr.querySelectorAll("td");
                                const actionTd = tds[tds.length - 1];
                                const explicit = actionTd.querySelector("[class*='download' i], [title*='download' i], [aria-label*='download' i]");
                                if (explicit) return explicit.closest("button, a, [role='button']") || explicit;
                                const btns = Array.from(actionTd.querySelectorAll("button, a, [role='button'], i, svg"));
                                if (btns.length >= 2) return btns[1];
                            }
                            cur = cur.parentElement;
                        }
                    }
                    return null;
                }""", row_index_in_page).as_element()
                if dl_btn is not None:
                    with page.expect_download(timeout=8000) as dl_info:
                        dl_btn.click()
                    download = dl_info.value
                    with tempfile.TemporaryDirectory() as td:
                        dl_path = os.path.join(td, download.suggested_filename or "doc.pdf")
                        download.save_as(dl_path)
                        with open(dl_path, "rb") as f:
                            dl_bytes = f.read()
                        if dl_bytes.startswith(b"%PDF"):
                            preview_info = self.pdf_engine.extract_preview_info_from_pdf_bytes(dl_bytes)
            except Exception:
                pass

        return preview_info

    def _read_additional_docs_rows_and_check_einvoice(self, page, case):
        """
        Inspect Additional Documents table (including pagination) for existing E-INVOICE rows.
        When an E-INVOICE row exists, clicks its Eye icon to preview the PDF and verify whether
        it is this specific invoice or a different invoice.
        Returns (exists_for_this_invoice: bool, info_str: str).
        """
        def _inspect_visible_rows():
            return page.evaluate(r"""() => {
                const headingNodes = Array.from(document.querySelectorAll("*"))
                    .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
                let container = null;
                for (const h of headingNodes) {
                    let cur = h.parentElement;
                    for (let d = 0; d < 8 && cur && cur !== document.body; d++) {
                        if ((cur.querySelector("table, .dx-datagrid") || /Document Name/i.test(cur.innerText || '')) && !/Court Suit Number/i.test(cur.innerText || '')) {
                            container = cur;
                            break;
                        }
                        cur = cur.parentElement;
                    }
                    if (container) break;
                }
                if (!container) return { foundContainer: false, rows: [], nextPageBtnText: null };

                const rowsOut = [];
                const trs = Array.from(container.querySelectorAll("tbody tr, tr.dx-data-row"));
                for (const tr of trs) {
                    const tds = Array.from(tr.querySelectorAll("td"));
                    if (tds.length < 3) continue;
                    const texts = tds.map(td => (td.innerText || '').trim());
                    rowsOut.push(texts);
                }

                let activePageNum = null;
                let hasNextPage = false;
                const pageItems = Array.from(container.querySelectorAll(".pagination li, .pagination a, .pagination button, .dx-pager .dx-page, ul li a, ul li button, nav a, nav button"));
                for (const el of pageItems) {
                    const t = (el.innerText || '').trim();
                    const cls = (el.className || '') + ' ' + ((el.parentElement && el.parentElement.className) || '');
                    if (/^\d+$/.test(t) && /active|selected|dx-selection/i.test(cls)) {
                        activePageNum = parseInt(t, 10);
                    }
                }
                if (activePageNum !== null) {
                    const targetNext = String(activePageNum + 1);
                    for (const el of pageItems) {
                        if ((el.innerText || '').trim() === targetNext) {
                            hasNextPage = true;
                            break;
                        }
                    }
                }
                return { foundContainer: true, rows: rowsOut, activePageNum, hasNextPage };
            }""")

        for _ in range(15):
            self._dismiss_session_timeout_warning(page)
            info = _inspect_visible_rows()
            if not info.get("foundContainer"):
                break
            for row_idx_on_page, cells in enumerate(info.get("rows", [])):
                doc_name = (cells[0] if len(cells) > 0 else "").strip()
                doc_type = (cells[1] if len(cells) > 1 else "").strip()
                created_str = (cells[2] if len(cells) > 2 else "").strip()
                if doc_name.casefold() in ("e-invoice", "einvoice", "e invoice"):
                    self.log(f"  [PREVIEW] Found E-INVOICE ({doc_type}, {created_str or '-'}) -> clicking eye icon to inspect...")
                    preview_info = self._preview_additional_doc_row(page, row_idx_on_page)
                    is_same, reason = preview_matches_case(preview_info, case)
                    if is_same:
                        return True, f"Preview confirmed E-INVOICE ({reason}, {created_str or '-'})"
                    self.log(f"  [PREVIEW] Existing E-INVOICE does not match current invoice ({reason}); checking next...")

            if not info.get("hasNextPage"):
                break

            next_num = str((info.get("activePageNum") or 1) + 1)
            clicked = page.evaluate(r"""(nextNum) => {
                const headingNodes = Array.from(document.querySelectorAll("*"))
                    .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
                for (const h of headingNodes) {
                    let cur = h.parentElement;
                    for (let d = 0; d < 8 && cur && cur !== document.body; d++) {
                        if ((cur.querySelector("table, .dx-datagrid") || /Document Name/i.test(cur.innerText || '')) && !/Court Suit Number/i.test(cur.innerText || '')) {
                            const btns = Array.from(cur.querySelectorAll(".pagination a, .pagination button, .dx-pager .dx-page, ul li a, ul li button, nav a, nav button"));
                            const target = btns.find(b => (b.innerText || '').trim() === nextNum);
                            if (target) { target.click(); return true; }
                        }
                        cur = cur.parentElement;
                    }
                }
                return false;
            }""", next_num)
            if not clicked:
                break
            self._wait_for_network_quiet(page, timeout=3000)
            page.wait_for_timeout(450)

        return False, ""

    def _open_additional_docs_add_modal(self, page):
        """Scroll to Additional Documents, expand if collapsed, and click its '+' button."""
        heading = page.locator("text=/^Additional Documents$/i").first
        heading.wait_for(state="visible", timeout=10000)
        heading.scroll_into_view_if_needed(timeout=4000)
        page.wait_for_timeout(250)
        self._dismiss_session_timeout_warning(page)

        # Expand section if its table/empty message isn't visible yet
        container_visible_table = page.evaluate(r"""() => {
            const headingNodes = Array.from(document.querySelectorAll("*"))
                .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
            for (const h of headingNodes) {
                let cur = h.parentElement;
                for (let d = 0; d < 8 && cur && cur !== document.body; d++) {
                    if (/Document Name/i.test(cur.innerText || '') && !/Court Suit Number/i.test(cur.innerText || '')) {
                        const tbl = cur.querySelector("table, .dx-datagrid");
                        if (tbl && tbl.offsetParent !== null) return true;
                    }
                    cur = cur.parentElement;
                }
            }
            return false;
        }""")
        if not container_visible_table:
            page.evaluate(r"""() => {
                const headingNodes = Array.from(document.querySelectorAll("*"))
                    .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
                for (const h of headingNodes) {
                    let cur = h.parentElement;
                    for (let d = 0; d < 4 && cur && cur !== document.body; d++) {
                        const chev = cur.querySelector(".fa-chevron-down, .fa-angle-down, .dx-icon-chevrondown, [aria-expanded='false']");
                        if (chev) { chev.click(); return; }
                        cur = cur.parentElement;
                    }
                }
            }""")
            page.wait_for_timeout(350)

        clicked_plus = page.evaluate(r"""() => {
            const headingNodes = Array.from(document.querySelectorAll("*"))
                .filter(el => /^Additional Documents$/i.test((el.innerText || '').trim()));
            for (const h of headingNodes) {
                let cur = h.parentElement;
                for (let d = 0; d < 6 && cur && cur !== document.body; d++) {
                    if (/Court Suit Number/i.test(cur.innerText || '')) break;
                    const btns = Array.from(cur.querySelectorAll("button, a, [role='button'], i, span"));
                    for (const b of btns) {
                        const cls = String(b.className || '');
                        if (/chevron|angle|arrow|collapse/i.test(cls)) continue;
                        const txt = (b.innerText || '').trim();
                        const hasPlusIcon = b.querySelector(".fa-plus, .mdi-plus, .dx-icon-plus, [class*='plus']") !== null || /fa-plus|mdi-plus|dx-icon-plus|\bplus\b/i.test(cls);
                        const isPlusBtn = txt === '+' || hasPlusIcon || /add/i.test(b.getAttribute('title') || b.getAttribute('aria-label') || '');
                        if (isPlusBtn && b.offsetParent !== null) {
                            b.click();
                            return true;
                        }
                    }
                    // Fallback: check the header row right next to Additional Documents
                    const rect = h.getBoundingClientRect();
                    const rightBtns = btns.filter(b => {
                        const r = b.getBoundingClientRect();
                        return r.width > 16 && r.height > 16 && Math.abs(r.top - rect.top) < 40 && r.left > rect.right;
                    });
                    if (rightBtns.length > 0) {
                        rightBtns.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
                        rightBtns[0].click();
                        return true;
                    }
                    cur = cur.parentElement;
                }
            }
            return false;
        }""")
        if not clicked_plus:
            raise RuntimeError("Could not click the '+' button inside Additional Documents.")

        modal = page.locator(".modal:visible, [role='dialog']:visible, .dx-popup-wrapper:visible").filter(has_text="Add New Document").first
        modal.wait_for(state="visible", timeout=8000)
        return modal

    def _fill_and_submit_add_document_modal(self, page, modal, pdf_file_path):
        """Fill Document Name='E-INVOICE', Document Type='Others', Upload File, empty Remark, and click Add now."""
        # 1. Document Name input
        doc_name_input = modal.locator("xpath=.//*[contains(normalize-space(.), 'Document Name')]/following::input[1]").first
        if doc_name_input.count() == 0:
            doc_name_input = modal.locator("input[type='text']:visible").first
        doc_name_input.wait_for(state="visible", timeout=4000)
        doc_name_input.fill("E-INVOICE")

        # 2. Document Type = 'Others'
        select_loc = modal.locator("select").first
        if select_loc.count() > 0 and select_loc.is_visible():
            try:
                select_loc.select_option(label="Others")
            except Exception:
                # Try matching option text case-insensitively
                opt_val = select_loc.evaluate(r"""sel => {
                    const opt = Array.from(sel.options).find(o => /others/i.test((o.text || '').trim()));
                    return opt ? opt.value : null;
                }""")
                if opt_val is not None:
                    select_loc.select_option(value=opt_val)
                else:
                    raise RuntimeError("Could not find 'Others' in Document Type dropdown.")
        else:
            # Custom dropdown / DevExtreme selectbox fallback
            type_ctrl = modal.locator("xpath=.//*[contains(normalize-space(.), 'Document Type')]/following::*[self::input or @role='combobox' or contains(@class,'dx-selectbox')][1]").first
            type_ctrl.click(timeout=3000)
            opt = page.locator(".dx-list-item:visible, [role='option']:visible, .dropdown-item:visible").filter(has_text=re.compile(r"^\s*Others\s*$", re.I)).first
            opt.wait_for(state="visible", timeout=4000)
            opt.click(timeout=3000)

        # 3. Upload File (hidden input[type='file'] or Browse button file chooser)
        file_inputs = modal.locator("input[type='file']")
        if file_inputs.count() > 0:
            file_inputs.first.set_input_files(pdf_file_path)
        else:
            browse_btn = modal.locator("button:has-text('Browse'), label:has-text('Browse')").first
            browse_btn.wait_for(state="visible", timeout=4000)
            with page.expect_file_chooser(timeout=5000) as fc_info:
                browse_btn.click()
            fc_info.value.set_files(pdf_file_path)
        page.wait_for_timeout(300)

        # 4. Remark blank
        remark_loc = modal.locator("textarea").first
        if remark_loc.count() > 0 and remark_loc.is_visible():
            remark_loc.fill("")

        # 5. Click 'Add now'
        add_btn = modal.locator("button:has-text('Add now'), button:has-text('Add Now')").first
        add_btn.wait_for(state="visible", timeout=4000)
        if not add_btn.is_enabled():
            raise RuntimeError("'Add now' button remained disabled after filling E-INVOICE form.")
        add_btn.click(timeout=5000)

        # Wait for modal to close and upload request to finish
        modal.wait_for(state="hidden", timeout=25000)
        self._wait_for_network_quiet(page, timeout=6000)
        page.wait_for_timeout(500)

        if self._is_kicked_or_unauthorized(page):
            raise InvalidCredentialsError("Session expired during document upload.")

        err_toast = page.locator(".toast-error:visible, .alert-danger:visible").first
        if err_toast.count() > 0 and err_toast.is_visible():
            raise RuntimeError(f"Portal reported upload error: {err_toast.inner_text(timeout=1000).strip()}")

    def _build_extracted_invoice_pdf_path(self, case):
        inv_no = re.sub(r"[^A-Za-z0-9_-]+", "_", case.get("Invoice No", "").strip()).strip("_")
        acc_no = re.sub(r"[^A-Za-z0-9_-]+", "_", case.get("Account No", "").strip()).strip("_")
        pages_slug = re.sub(r"[^0-9-]+", "_", case.get("Invoice Pages", "").strip()).strip("_")
        parts = [p for p in (inv_no, acc_no, f"p{pages_slug}" if pages_slug else "") if p]
        base = "_".join(parts) or f"invoice_{int(time.time())}"
        return os.path.join(EXTRACTED_INVOICES_DIR, f"{base}.pdf")

    def _upload_invoice_in_legal_details(self, page, case):
        """Navigate back to Legal Details, check Additional Documents, extract invoice pages, and upload."""
        self._switch_case_tab(page, "Legal Details")
        page.locator("text=/^Additional Documents$/i").first.wait_for(state="visible", timeout=10000)
        page.locator("text=/^Additional Documents$/i").first.scroll_into_view_if_needed(timeout=4000)
        page.wait_for_timeout(300)

        exists, existing_info = self._read_additional_docs_rows_and_check_einvoice(page, case)
        if exists:
            return "SKIPPED_EXISTS", existing_info, case.get("Extracted Invoice PDF", "-")

        src_pdf = case.get("Source PDF Path", "").strip()
        page_spec = case.get("Invoice Pages", "").strip()
        if not src_pdf or not os.path.isfile(src_pdf) or not parse_page_numbers_spec(page_spec):
            return "NO_PDF", "Billing matched, but no PDF/page range linked for extraction", "-"

        dest_pdf = self._build_extracted_invoice_pdf_path(case)
        extract_invoice_pdf_pages(src_pdf, page_spec, dest_pdf)

        modal = self._open_additional_docs_add_modal(page)
        self._fill_and_submit_add_document_modal(page, modal, dest_pdf)
        return "UPLOADED", f"Uploaded {os.path.basename(dest_pdf)} (p.{page_spec})", dest_pdf

    def _return_to_case_listing_and_restore_section(self, page, acc, section):
        """Return from Case Details back to Case Listing grid and ensure active section tab."""
        self._last_searched_column = None
        returned = False
        try:
            back_btn = page.locator("a:has-text('BACK'), button:has-text('BACK'), [class*='back']:has-text('BACK')").first
            if back_btn.count() > 0 and back_btn.is_visible():
                back_btn.click(timeout=5000)
                self._wait_for_network_quiet(page, timeout=4000)
                page.wait_for_timeout(400)
                if "/case-listing" in page.url.lower() and not self._is_on_case_details(page):
                    returned = True
        except Exception:
            returned = False

        if self._is_kicked_or_unauthorized(page):
            self._login_and_restore_section(page, acc, section=section)
            return

        if not returned:
            page.goto(CASE_LISTING_URL, wait_until="domcontentloaded", timeout=45000)
            self._wait_for_network_quiet(page, timeout=4000)
            if self._is_kicked_or_unauthorized(page):
                self._login_and_restore_section(page, acc, section=section)
                return

        try:
            self._switch_section_and_wait(page, section)
        except InvalidCredentialsError as e:
            if self._is_fatal_auth_error(e):
                raise
            self._login_and_restore_section(page, acc, section=section)
        except Exception:
            if self._is_kicked_or_unauthorized(page):
                self._login_and_restore_section(page, acc, section=section)
            else:
                raise

    def _process_matched_rows_for_case(self, page, acc, section, search_col, search_val, matched_rows, case, skip_summaries=None):
        """
        Enter each matched row in Case Listing, verify Billings Date & Amount, and if matched,
        upload the extracted E-INVOICE in Legal Details -> Additional Documents (unless Case Withdrawn).
        Returns dict with keys: outcome ('COMPLETE'|'BILLING_NO_PDF'|'BILLING_MISMATCH'|'ALL_SKIPPED'),
        billing_detail, upload_status, extracted_pdf, matched_row_detail.
        """
        if skip_summaries is None:
            skip_summaries = set()
        last_mismatch_detail = ""
        checked_any_row = False
        first_checked_summary = ""
        need_reapply_filter = False

        for idx_in_matches, match_item in enumerate(matched_rows):
            if self.stop_requested:
                break

            if len(match_item) >= 3:
                row_idx, row_summary, grid_page = match_item[0], match_item[1], match_item[2]
            else:
                row_idx, row_summary, grid_page = match_item[0], match_item[1], 1

            if row_summary in skip_summaries:
                continue

            if not first_checked_summary:
                first_checked_summary = row_summary

            # Re-apply filter whenever returning from an earlier Case Details check
            if need_reapply_filter:
                refreshed_matches = self._search_grid_with_relogin(page, acc, section, search_col, search_val)
                found_match = next(
                    (m for m in refreshed_matches if m[1] == row_summary),
                    None,
                )
                if found_match is not None:
                    row_idx = found_match[0]
                    grid_page = found_match[2] if len(found_match) >= 3 else 1
                elif idx_in_matches < len(refreshed_matches):
                    m_fb = refreshed_matches[idx_in_matches]
                    row_idx = m_fb[0]
                    grid_page = m_fb[2] if len(m_fb) >= 3 else 1
                else:
                    continue

            if grid_page > 1:
                self._go_to_grid_page(page, grid_page)

            checked_any_row = True
            need_reapply_filter = True
            self.log(f"  [OPEN] Entering case row #{row_idx + 1} (page {grid_page}) to verify Billings...")
            try:
                self._open_case_from_grid_row(page, row_idx)
                billing_ok, billing_detail = self._verify_case_billings(page, case)
                if not billing_ok:
                    skip_summaries.add(row_summary)
                    self.log(f"  [BILLING MISMATCH] {billing_detail} -> checking next account / treating as None.")
                    last_mismatch_detail = billing_detail
                    continue

                self.log(f"  [BILLING OK] {billing_detail}")
                if section.strip().casefold() == "case withdrawn":
                    skip_summaries.add(row_summary)
                    msg = "Skipped upload (Case Withdrawn)"
                    self.log(f"  [COMPLETE] {msg}")
                    return {
                        "outcome": "COMPLETE",
                        "billing_detail": billing_detail,
                        "upload_status": msg,
                        "extracted_pdf": "-",
                        "matched_row_detail": row_summary,
                    }
                upload_outcome, upload_msg, pdf_out = self._upload_invoice_in_legal_details(page, case)
                skip_summaries.add(row_summary)
                if upload_outcome in ("UPLOADED", "SKIPPED_EXISTS"):
                    self.log(f"  [COMPLETE] {upload_msg}")
                    return {
                        "outcome": "COMPLETE",
                        "billing_detail": billing_detail,
                        "upload_status": upload_msg,
                        "extracted_pdf": pdf_out,
                        "matched_row_detail": row_summary,
                    }
                else:
                    self.log(f"  [PARTIAL] {upload_msg}")
                    return {
                        "outcome": "BILLING_NO_PDF",
                        "billing_detail": billing_detail,
                        "upload_status": upload_msg,
                        "extracted_pdf": "-",
                        "matched_row_detail": row_summary,
                    }
            finally:
                if not self.stop_requested:
                    try:
                        self._return_to_case_listing_and_restore_section(page, acc, section)
                    except InvalidCredentialsError as ret_err:
                        if self._is_fatal_auth_error(ret_err):
                            raise
                        self.log(f"  [!] Session interrupted while returning to {section}; will recover on next step.")
                    except Exception as ret_err:
                        self.log(f"  [!] Notice returning to {section}: {ret_err}")

        if not checked_any_row:
            return {
                "outcome": "ALL_SKIPPED",
                "billing_detail": "",
                "upload_status": "",
                "extracted_pdf": "-",
                "matched_row_detail": "-",
            }

        return {
            "outcome": "BILLING_MISMATCH",
            "billing_detail": last_mismatch_detail or "Billing date/amount did not match",
            "upload_status": "Not uploaded (billing mismatch)",
            "extracted_pdf": "-",
            "matched_row_detail": first_checked_summary or (matched_rows[0][1] if matched_rows else "-"),
        }

    # --------------------------------------------------------------------------
    # AUTOMATION PASS & WORKER THREAD
    # --------------------------------------------------------------------------
    @staticmethod
    def _is_case_done(case, verify_mode):
        status = case.get("Availability Status", "")
        if verify_mode:
            return status == "COMPLETE"
        return status == "AVAILABLE"

    @staticmethod
    def _mark_case_error(case, checked_pairs, error_reasons, acc_idx, username, section, reason):
        if case.get("Availability Status") in ("AVAILABLE", "COMPLETE"):
            return
        checked_pairs[(acc_idx, section)] = "ERROR"
        error_reasons.append((username, section, reason))
        failed_locs = ", ".join(dict.fromkeys(f"{u}: {s}" for u, s, _ in error_reasons))
        case["Availability Status"] = "CHECK ERROR"
        case["Found In Account(s)"] = "Check incomplete"
        case["Found In Section(s)"] = failed_locs
        case["C-PRO Row Details"] = reason

    @staticmethod
    def _resolve_case_after_clean_pair(case, checked_pairs, error_reasons, acc_idx, username, section, active_accounts, active_sections, pair_state="NOT_FOUND", mismatch_note=""):
        if case.get("Availability Status") in ("AVAILABLE", "COMPLETE", "BILLING MATCHED (No PDF)"):
            return
        checked_pairs[(acc_idx, section)] = pair_state
        if mismatch_note and pair_state == "BILLING_MISMATCH":
            case["Billing Check"] = f"{username}: {section} -> {mismatch_note}"
        if error_reasons:
            return
        all_checked = all(
            checked_pairs.get((a_i, sec)) in ("NOT_FOUND", "BILLING_MISMATCH")
            for a_i in range(len(active_accounts))
            for sec in active_sections
        )
        if all_checked:
            had_mismatch = any(
                checked_pairs.get((a_i, sec)) == "BILLING_MISMATCH"
                for a_i in range(len(active_accounts))
                for sec in active_sections
            )
            if had_mismatch:
                case["Availability Status"] = "NONE (Billing Mismatch)"
                case["Found In Account(s)"] = "None"
                case["Found In Section(s)"] = "None"
                case["Upload Status"] = "Not uploaded (consider None)"
            else:
                case["Availability Status"] = "NOT FOUND"
                case["Found In Account(s)"] = "None"
                case["Found In Section(s)"] = "None"
                case["Billing Check"] = "-"
                case["Upload Status"] = "-"
                case["C-PRO Row Details"] = "-"

    def _run_check_pass(self, browser, active_accounts, active_sections, pass_indices, found_tracker, detailed_log, mode, verify_mode, round_label):
        checked_pairs = {idx: {} for idx in pass_indices}
        error_reasons = {idx: [] for idx in pass_indices}
        for idx in pass_indices:
            found_tracker[idx] = {"accounts": [], "sections": [], "details": []}
            case = self.cases_data[idx]
            case["Availability Status"] = "Pending"
            case["Found In Account(s)"] = "-"
            case["Found In Section(s)"] = "-"
            case["Billing Check"] = "-"
            case["Upload Status"] = "-"
            case["Extracted Invoice PDF"] = "-"
            case["C-PRO Row Details"] = "-"
        self.refresh_table()

        total_steps = len(active_accounts) * len(active_sections) * len(pass_indices)
        completed_steps = 0
        self.set_progress(0, total_steps)

        def _all_done():
            return all(self._is_case_done(self.cases_data[i], verify_mode) for i in pass_indices)

        def _mark_sections_error(acc_idx, username, sections_slice, reason, log_details=True):
            for idx in pass_indices:
                case = self.cases_data[idx]
                if self._is_case_done(case, verify_mode):
                    continue
                acc_no = case.get("Account No", "").strip()
                search_name = (case.get("Search Name") or case.get("Full Borrower Name") or "").strip()
                for sec in sections_slice:
                    if checked_pairs[idx].get((acc_idx, sec)) in ("NOT_FOUND", "MATCHED", "BILLING_MISMATCH"):
                        continue
                    self._mark_case_error(case, checked_pairs[idx], error_reasons[idx], acc_idx, username, sec, reason)
                    if log_details:
                        detailed_log.append({
                            "Invoice No": case.get("Invoice No", ""),
                            "Account No": acc_no,
                            "Search Name": search_name,
                            "C-PRO Account": username,
                            "Login Used": username,
                            "Section": sec,
                            "Status": "CHECK ERROR",
                            "Matched By": "-",
                            "Billing Check": "-",
                            "Upload Status": "-",
                            "C-PRO Row Details": reason,
                            "Round": round_label,
                        })

        for acc_idx, acc in enumerate(active_accounts):
            if self.stop_requested:
                break
            if _all_done():
                completed_steps = total_steps
                self.set_progress(completed_steps, total_steps)
                break

            username = acc["username"]
            self.log(f"=== [{round_label}] Starting Account: {username} ===")

            context = browser.new_context()
            page = context.new_page()
            self._install_session_timeout_handler(page)
            current_sec_idx = 0
            current_section = active_sections[0] if active_sections else "(not selected)"

            try:
                self._login_and_restore_section(page, acc, section=None)

                for sec_idx, section in enumerate(active_sections):
                    if self.stop_requested:
                        break
                    if _all_done():
                        completed_steps = total_steps
                        self.set_progress(completed_steps, total_steps)
                        break

                    current_sec_idx = sec_idx
                    current_section = section
                    self.log(f"[{username}] Switching to section: {section}")

                    section_loaded = False
                    last_section_err = None
                    for sec_attempt in range(1, MAX_RETRIES_ON_KICKOUT + 1):
                        if self.stop_requested:
                            break
                        try:
                            needs_restore = (
                                sec_attempt > 1
                                or self._is_kicked_or_unauthorized(page)
                                or self._is_on_case_details(page)
                                or "/case-listing" not in page.url.lower()
                            )
                            if needs_restore:
                                self._login_and_restore_section(page, acc, section=section)
                            else:
                                self._switch_section_and_wait(page, section)
                            section_loaded = True
                            break
                        except InvalidCredentialsError as sec_auth_err:
                            if self._is_fatal_auth_error(sec_auth_err):
                                raise
                            last_section_err = sec_auth_err
                            self.log(
                                f"[!] Session expired/kicked switching to '{section}' "
                                f"(Attempt {sec_attempt}/{MAX_RETRIES_ON_KICKOUT}): {sec_auth_err}. Re-logging..."
                            )
                            page.wait_for_timeout(800)
                        except Exception as section_error:
                            if self.stop_requested:
                                break
                            last_section_err = section_error
                            self.log(
                                f"[!] Section '{section}' failed to load "
                                f"(Attempt {sec_attempt}/{MAX_RETRIES_ON_KICKOUT}: {section_error}). Re-authenticating & retrying..."
                            )
                            page.wait_for_timeout(800)

                    if self.stop_requested:
                        break
                    if not section_loaded:
                        err_msg = f"Could not load the {section} section table after {MAX_RETRIES_ON_KICKOUT} attempt(s): {last_section_err}"
                        self.log(f"[ERROR] [{username}] {err_msg}")
                        _mark_sections_error(acc_idx, username, [section], err_msg)
                        completed_steps = (acc_idx * len(active_sections) + sec_idx + 1) * len(pass_indices)
                        self.set_progress(completed_steps, total_steps)
                        self.refresh_table()
                        continue

                    for idx in pass_indices:
                        if self.stop_requested:
                            break
                        case = self.cases_data[idx]

                        if self._is_case_done(case, verify_mode):
                            completed_steps += 1
                            self.set_progress(completed_steps, total_steps)
                            continue

                        self.refresh_table(active_idx=idx)
                        acc_no = case.get("Account No", "").strip()
                        alt_acc_no = case.get("Alt Account No", "").strip()
                        search_name = (case.get("Search Name") or case.get("Full Borrower Name") or "").strip()

                        search_attempts = []
                        if mode in ("ACCOUNT_THEN_NAME", "ACCOUNT_ONLY"):
                            for cand_acc in self._extract_account_candidates(acc_no, alt_acc_no):
                                search_attempts.append(("Account No", cand_acc, f"Account No ({cand_acc})"))
                        if mode in ("ACCOUNT_THEN_NAME", "NAME_ONLY"):
                            if search_name:
                                search_attempts.append(("Primary Borrower Name", search_name, f"Name ({search_name})"))
                            elif mode == "ACCOUNT_THEN_NAME":
                                self.log(f"  [!] Cannot fallback to Name for {acc_no} (Name is empty).")

                        checked_row_summaries = set()
                        found_any_row = False
                        case_resolved_in_pair = False
                        pair_had_error = False
                        last_mismatch_detail = ""

                        for s_col, s_val, s_label in search_attempts:
                            if self.stop_requested or case_resolved_in_pair:
                                break
                            try:
                                matched_rows = self._search_grid_with_relogin(page, acc, section, s_col, s_val)
                            except InvalidCredentialsError as auth_err:
                                if self._is_fatal_auth_error(auth_err):
                                    raise
                                if self.stop_requested:
                                    break
                                self.log(f"  [ERROR] Session error searching '{s_val}' after retries: {auth_err}. Marking this row only & recovering section...")
                                self._mark_case_error(case, checked_pairs[idx], error_reasons[idx], acc_idx, username, section, str(auth_err))
                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": "CHECK ERROR",
                                    "Matched By": s_label,
                                    "Billing Check": "-",
                                    "Upload Status": "-",
                                    "C-PRO Row Details": str(auth_err),
                                    "Round": round_label,
                                })
                                try:
                                    self._login_and_restore_section(page, acc, section=section)
                                except InvalidCredentialsError as rec_err:
                                    if self._is_fatal_auth_error(rec_err):
                                        raise
                                except Exception:
                                    pass
                                pair_had_error = True
                                break
                            except Exception as e:
                                if self.stop_requested:
                                    break
                                self.log(f"  [ERROR] Search failed for '{s_val}': {e}")
                                self._mark_case_error(case, checked_pairs[idx], error_reasons[idx], acc_idx, username, section, str(e))
                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": "CHECK ERROR",
                                    "Matched By": s_label,
                                    "Billing Check": "-",
                                    "Upload Status": "-",
                                    "C-PRO Row Details": str(e),
                                    "Round": round_label,
                                })
                                pair_had_error = True
                                break

                            if self.stop_requested:
                                break

                            new_rows = [m for m in matched_rows if m[1] not in checked_row_summaries]
                            if not new_rows:
                                continue

                            found_any_row = True
                            self.log(f"  ✓ FOUND by {s_label}: {search_name or acc_no} in {username} -> {section} ({len(new_rows)} new row(s))")

                            if not verify_mode:
                                checked_pairs[idx][(acc_idx, section)] = "MATCHED"
                                found_tracker[idx]["accounts"].append(username)
                                found_tracker[idx]["sections"].append(f"{username}: {section}")
                                found_tracker[idx]["details"].extend(m[1] for m in new_rows)

                                case["Availability Status"] = "AVAILABLE"
                                case["Found In Account(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["accounts"]))
                                case["Found In Section(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["sections"]))
                                case["C-PRO Row Details"] = " || ".join(found_tracker[idx]["details"])
                                for m in new_rows:
                                    checked_row_summaries.add(m[1])
                                    detailed_log.append({
                                        "Invoice No": case.get("Invoice No", ""),
                                        "Account No": acc_no,
                                        "Search Name": search_name,
                                        "C-PRO Account": username,
                                        "Login Used": username,
                                        "Section": section,
                                        "Status": "AVAILABLE",
                                        "Matched By": s_label,
                                        "Billing Check": "Skipped (Availability Only)",
                                        "Upload Status": "Skipped",
                                        "C-PRO Row Details": m[1],
                                        "Round": round_label,
                                    })
                                case_resolved_in_pair = True
                                break

                            wf = None
                            wf_last_err = None
                            current_matches = matched_rows
                            for wf_attempt in range(1, MAX_RETRIES_ON_KICKOUT + 1):
                                if self.stop_requested:
                                    break
                                try:
                                    wf = self._process_matched_rows_for_case(
                                        page, acc, section, s_col, s_val, current_matches, case, skip_summaries=checked_row_summaries
                                    )
                                    wf_last_err = None
                                    break
                                except InvalidCredentialsError as wf_auth_err:
                                    if self._is_fatal_auth_error(wf_auth_err):
                                        raise
                                    wf_last_err = wf_auth_err
                                    if wf_attempt < MAX_RETRIES_ON_KICKOUT and not self.stop_requested:
                                        self.log(
                                            f"  [!] Session expired inside Case Details for '{acc_no}' "
                                            f"(Attempt {wf_attempt}/{MAX_RETRIES_ON_KICKOUT}: {wf_auth_err}). Re-logging & retrying..."
                                        )
                                        try:
                                            self._login_and_restore_section(page, acc, section=section)
                                            current_matches = self._search_grid_with_relogin(page, acc, section, s_col, s_val)
                                            if not current_matches:
                                                wf = {
                                                    "outcome": "ALL_SKIPPED",
                                                    "billing_detail": "",
                                                    "upload_status": "",
                                                    "extracted_pdf": "-",
                                                    "matched_row_detail": "-",
                                                }
                                                wf_last_err = None
                                                break
                                        except InvalidCredentialsError as rec_err:
                                            if self._is_fatal_auth_error(rec_err):
                                                raise
                                            wf_last_err = rec_err
                                        except Exception as rec_err:
                                            wf_last_err = rec_err
                                        continue
                                    break
                                except Exception as wf_err:
                                    if self.stop_requested:
                                        break
                                    err_low = str(wf_err).lower()
                                    is_recoverable_nav = (
                                        "execution context was destroyed" in err_low
                                        or self._is_kicked_or_unauthorized(page)
                                    )
                                    wf_last_err = wf_err
                                    if is_recoverable_nav and wf_attempt < MAX_RETRIES_ON_KICKOUT:
                                        self.log(
                                            f"  [!] Navigation/session hiccup inside Case Details for '{acc_no}' "
                                            f"(Attempt {wf_attempt}/{MAX_RETRIES_ON_KICKOUT}: {wf_err}). Re-logging & retrying..."
                                        )
                                        try:
                                            self._login_and_restore_section(page, acc, section=section)
                                            current_matches = self._search_grid_with_relogin(page, acc, section, s_col, s_val)
                                            if not current_matches:
                                                wf = {
                                                    "outcome": "ALL_SKIPPED",
                                                    "billing_detail": "",
                                                    "upload_status": "",
                                                    "extracted_pdf": "-",
                                                    "matched_row_detail": "-",
                                                }
                                                wf_last_err = None
                                                break
                                        except InvalidCredentialsError as rec_err:
                                            if self._is_fatal_auth_error(rec_err):
                                                raise
                                            wf_last_err = rec_err
                                        except Exception as rec_err:
                                            wf_last_err = rec_err
                                        continue
                                    break

                            if self.stop_requested:
                                break
                            if wf is None:
                                self.log(f"  [ERROR] Workflow failed inside Case Details for '{acc_no}': {wf_last_err}")
                                self._mark_case_error(case, checked_pairs[idx], error_reasons[idx], acc_idx, username, section, str(wf_last_err))
                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": "CHECK ERROR",
                                    "Matched By": s_label,
                                    "Billing Check": "Error",
                                    "Upload Status": "Error",
                                    "C-PRO Row Details": str(wf_last_err),
                                    "Round": round_label,
                                })
                                try:
                                    self._return_to_case_listing_and_restore_section(page, acc, section)
                                except InvalidCredentialsError as rec_err:
                                    if self._is_fatal_auth_error(rec_err):
                                        raise
                                    try:
                                        self._login_and_restore_section(page, acc, section=section)
                                    except InvalidCredentialsError as rec_err2:
                                        if self._is_fatal_auth_error(rec_err2):
                                            raise
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                                pair_had_error = True
                                break

                            outcome = wf["outcome"]
                            if outcome == "ALL_SKIPPED":
                                continue
                            if outcome in ("COMPLETE", "BILLING_NO_PDF"):
                                checked_pairs[idx][(acc_idx, section)] = "MATCHED"
                                found_tracker[idx]["accounts"].append(username)
                                found_tracker[idx]["sections"].append(f"{username}: {section}")
                                found_tracker[idx]["details"].append(wf["matched_row_detail"])

                                case["Availability Status"] = "COMPLETE" if outcome == "COMPLETE" else "BILLING MATCHED (No PDF)"
                                case["Found In Account(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["accounts"]))
                                case["Found In Section(s)"] = ", ".join(dict.fromkeys(found_tracker[idx]["sections"]))
                                case["Billing Check"] = wf["billing_detail"]
                                case["Upload Status"] = wf["upload_status"]
                                case["Extracted Invoice PDF"] = wf["extracted_pdf"]
                                case["C-PRO Row Details"] = " || ".join(found_tracker[idx]["details"])

                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": case["Availability Status"],
                                    "Matched By": s_label,
                                    "Billing Check": wf["billing_detail"],
                                    "Upload Status": wf["upload_status"],
                                    "C-PRO Row Details": wf["matched_row_detail"],
                                    "Round": round_label,
                                })
                                case_resolved_in_pair = True
                                break
                            else:
                                # Billing date/amount did not match -> record detail and continue checking remaining accounts/name fallback
                                last_mismatch_detail = wf["billing_detail"]
                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": "NONE (Billing Mismatch)",
                                    "Matched By": s_label,
                                    "Billing Check": wf["billing_detail"],
                                    "Upload Status": wf["upload_status"],
                                    "C-PRO Row Details": wf["matched_row_detail"],
                                    "Round": round_label,
                                })

                        if not self.stop_requested and not pair_had_error and not case_resolved_in_pair:
                            if found_any_row:
                                self._resolve_case_after_clean_pair(
                                    case, checked_pairs[idx], error_reasons[idx],
                                    acc_idx, username, section, active_accounts, active_sections,
                                    pair_state="BILLING_MISMATCH",
                                    mismatch_note=last_mismatch_detail,
                                )
                            else:
                                self.log(f"  ✗ NOT FOUND: {search_name or acc_no or '(no search value)'} in {username} -> {section}")
                                self._resolve_case_after_clean_pair(
                                    case, checked_pairs[idx], error_reasons[idx],
                                    acc_idx, username, section, active_accounts, active_sections,
                                    pair_state="NOT_FOUND",
                                )
                                detailed_log.append({
                                    "Invoice No": case.get("Invoice No", ""),
                                    "Account No": acc_no,
                                    "Search Name": search_name,
                                    "C-PRO Account": username,
                                    "Login Used": username,
                                    "Section": section,
                                    "Status": "Not Found",
                                    "Matched By": "-",
                                    "Billing Check": "-",
                                    "Upload Status": "-",
                                    "C-PRO Row Details": "-",
                                    "Round": round_label,
                                })

                        completed_steps += 1
                        self.set_progress(completed_steps, total_steps)
                        self.refresh_table()

            except InvalidCredentialsError as e:
                if "rejected" in str(e).lower() or "cancelled" in str(e).lower():
                    for idx in pass_indices:
                        pending_case = self.cases_data[idx]
                        if not self._is_case_done(pending_case, verify_mode):
                            pending_case["Availability Status"] = "CHECK ERROR"
                            pending_case["Found In Account(s)"] = "Login not verified"
                            pending_case["Found In Section(s)"] = f"{username}: authentication"
                            pending_case["C-PRO Row Details"] = str(e)
                    self.refresh_table()
                    self.log(f"[AUTH STOP] {e} Automation paused to protect account.")
                    self.root.after(0, lambda msg=str(e): messagebox.showwarning("C-PRO Login Stopped", msg + "\n\nCorrect the password in the left panel, then restart.", parent=self.root))
                    self.stop_requested = True
                else:
                    _mark_sections_error(acc_idx, username, active_sections[current_sec_idx:], str(e))
                    completed_steps = (acc_idx + 1) * len(active_sections) * len(pass_indices)
                    self.set_progress(min(completed_steps, total_steps), total_steps)
                    self.refresh_table()
                    self.log(f"[AUTH ERROR] {username} could not re-authenticate: {e}. Continuing to next account.")
            except (GridLoadTimeoutError, GridSearchError) as e:
                _mark_sections_error(acc_idx, username, active_sections[current_sec_idx:], str(e))
                completed_steps = (acc_idx + 1) * len(active_sections) * len(pass_indices)
                self.set_progress(min(completed_steps, total_steps), total_steps)
                self.log(f"[ERROR] {username} could not complete a search in {current_section}: {e}.")
                self.refresh_table()
            except Exception as e:
                _mark_sections_error(acc_idx, username, active_sections[current_sec_idx:], str(e))
                completed_steps = (acc_idx + 1) * len(active_sections) * len(pass_indices)
                self.set_progress(min(completed_steps, total_steps), total_steps)
                self.log(f"[ERROR] Account {username} encountered an error: {e}")
                self.refresh_table()
            finally:
                try:
                    context.close()
                except Exception:
                    pass

    def _cpro_worker(self, active_accounts, active_sections, target_indices, mode, headless, verify_mode):
        found_tracker = {}
        detailed_log = []
        recovery_resolved = False

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=headless, slow_mo=60)
                try:
                    self._run_check_pass(
                        browser=browser,
                        active_accounts=active_accounts,
                        active_sections=active_sections,
                        pass_indices=target_indices,
                        found_tracker=found_tracker,
                        detailed_log=detailed_log,
                        mode=mode,
                        verify_mode=verify_mode,
                        round_label="Initial",
                    )

                    initial_errors = sum(1 for i in target_indices if self.cases_data[i].get("Availability Status") == "CHECK ERROR")
                    for recovery_round in range(1, MAX_ERROR_RECOVERY_ROUNDS + 1):
                        if self.stop_requested:
                            break
                        error_indices = [
                            idx for idx in target_indices
                            if self.cases_data[idx].get("Availability Status") == "CHECK ERROR"
                        ]
                        if not error_indices:
                            if initial_errors > 0:
                                recovery_resolved = True
                            break

                        self.log(
                            f"[RECOVERY] Found {len(error_indices)} CHECK ERROR case(s). "
                            f"Re-checking all enabled account(s) & section(s) (Round {recovery_round}/{MAX_ERROR_RECOVERY_ROUNDS})..."
                        )
                        self.set_status(
                            f"Re-checking errors ({recovery_round}/{MAX_ERROR_RECOVERY_ROUNDS})...",
                            bg="#FFF4D6",
                            fg="#8A5A00",
                        )
                        self._run_check_pass(
                            browser=browser,
                            active_accounts=active_accounts,
                            active_sections=active_sections,
                            pass_indices=error_indices,
                            found_tracker=found_tracker,
                            detailed_log=detailed_log,
                            mode=mode,
                            verify_mode=verify_mode,
                            round_label=f"Recovery {recovery_round}/{MAX_ERROR_RECOVERY_ROUNDS}",
                        )
                    else:
                        if initial_errors > 0 and not any(
                            self.cases_data[i].get("Availability Status") == "CHECK ERROR" for i in target_indices
                        ):
                            recovery_resolved = True
                finally:
                    browser.close()

            summary_rows = [
                {k: v for k, v in row.items() if not str(k).startswith("_")}
                for row in self.cases_data
            ]
            df_summary = pd.DataFrame(summary_rows)
            if detailed_log:
                df_detailed = pd.DataFrame(detailed_log)
            else:
                df_detailed = pd.DataFrame(columns=[
                    "Invoice No", "Account No", "Search Name", "C-PRO Account", "Login Used",
                    "Section", "Status", "Matched By", "Billing Check", "Upload Status",
                    "C-PRO Row Details", "Round",
                ])

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
            self.log(f"Done! Full report saved to '{report_path}'.")

            self.root.after(0, lambda: self._open_folder(os.path.abspath(os.path.dirname(report_path))))

            incomplete_count = sum(1 for i in target_indices if self.cases_data[i].get("Availability Status") == "CHECK ERROR")
            complete_count = sum(1 for i in target_indices if self.cases_data[i].get("Availability Status") in ("COMPLETE", "AVAILABLE"))
            if self.stop_requested:
                self.set_status("Stopped", bg="#FFF2F2", fg=THEME["red"])
                completion_title = "C-PRO Check Stopped"
                completion_message = (
                    f"C-PRO run was stopped ({len(target_indices)} selected case(s) queued).\n"
                    + (f"{incomplete_count} case(s) currently have CHECK ERROR.\n" if incomplete_count else "")
                    + f"\nPartial report saved to:\n{report_path}"
                )
            elif incomplete_count:
                self.log(f"[WARNING] {incomplete_count} case(s) still have CHECK ERROR after {MAX_ERROR_RECOVERY_ROUNDS} recovery round(s).")
                self.set_status("Completed with errors", bg="#FFF4D6", fg="#8A5A00")
                completion_title = "C-PRO Check Finished with Incomplete Searches"
                completion_message = (
                    f"Finished processing {len(target_indices)} selected case(s) ({complete_count} complete), "
                    f"but {incomplete_count} case(s) still have CHECK ERROR.\n\n"
                    f"Click 'Select Errors' to re-check those rows.\n\nReport saved to:\n{report_path}"
                )
            else:
                self.set_status("Completed", bg="#E8F8EE", fg=THEME["green"])
                completion_title = "C-PRO Automation Complete"
                extra = " All initial error cases were automatically re-checked and resolved!" if recovery_resolved else ""
                completion_message = (
                    f"Finished processing {len(target_indices)} selected case(s)! "
                    f"({complete_count} complete){extra}\n\nReport saved to:\n{report_path}"
                )

            self.root.after(0, lambda title=completion_title, msg=completion_message: messagebox.showinfo(title, msg, parent=self.root))

        except Exception as e:
            self.log(f"[CRITICAL ERROR] Automation failed: {e}")
            self.set_status("Error", bg="#FFF2F2", fg=THEME["red"])
        finally:
            self.is_running = False
            self.stop_requested = False
            self._last_searched_column = None
            self.refresh_table(active_idx=None)
            self._set_controls_running(False)


if __name__ == "__main__":
    root = tk.Tk()
    app = CProSuiteApp(root)
    root.mainloop()
