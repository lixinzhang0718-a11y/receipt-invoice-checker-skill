#!/usr/bin/env python3
"""Daily Gmail receipt/invoice checker.

This script is intentionally conservative. It downloads or reads cases, keeps
original files unchanged, extracts invoice and receipt evidence, and writes
review files. Optional dependencies unlock Gmail, OCR, Excel, and PDF features.
"""

from __future__ import annotations

import argparse
import base64
import csv
import datetime as dt
import email.utils
import html
import json
import re
import shutil
import sys
import textwrap
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


CONFIG_DIR = Path.home() / ".codex" / "receipt-invoice-checker"
CONFIG_PATH = CONFIG_DIR / "config.json"
DEFAULT_OUTPUT_ROOT = Path.home() / "Desktop" / "Receipt_Check"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}
EXCEL_EXTS = {".xlsx", ".xlsm"}
PDF_EXTS = {".pdf"}
SUPPORTED_EXTS = IMAGE_EXTS | EXCEL_EXTS | PDF_EXTS
GENERATED_PREFIXES = ("00_", "01_", "daily_review_summary", "case_result", "extracted_data", "email_info")

TKE_RE = re.compile(r"\bTKE[\s_-]*(\d{4,})\b", re.I)
DATE_TOKEN_RE = re.compile(r"(?<!\d)(20\d{6}|\d{6}|\d{4})(?!\d)")
AMOUNT_RE = re.compile(
    r"(?:JPY|YEN|yen|¥|￥)\s*([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{3,7})|"
    r"([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{3,7})\s*(?:JPY|YEN|yen|円)"
)


class SetupError(RuntimeError):
    pass


@dataclass
class EvidenceAmount:
    amount: int
    source: str
    context: str


@dataclass
class InvoiceItem:
    category: str
    label: str
    amount: int
    source: str
    row: Optional[int] = None


@dataclass
class CheckRow:
    item: str
    invoice_amount: str
    receipt_evidence: str
    result: str
    note: str


def sanitize(value: str, fallback: str = "unknown") -> str:
    value = (value or "").strip()
    value = re.sub(r"[<>:\"/\\|?*\r\n\t]+", "_", value)
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return value[:100] or fallback


def parse_day(value: str) -> dt.date:
    try:
        return dt.datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise SystemExit("Use date format YYYY-MM-DD, for example 2026-05-08.") from exc


def gmail_date(day: dt.date) -> str:
    return f"{day.year}/{day.month}/{day.day}"


def load_config() -> Dict[str, Any]:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    return {
        "output_root": str(DEFAULT_OUTPUT_ROOT),
        "gmail_credentials": str(CONFIG_DIR / "gmail_credentials.json"),
        "gmail_token": str(CONFIG_DIR / "gmail_token.json"),
        "tesseract_cmd": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        "ocr_languages": "jpn+eng",
        "gmail_extra_query": "",
        "known_sender_query": "",
        "manual_send_only": True,
    }


def init_config() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if CONFIG_PATH.exists():
        print(f"Config already exists: {CONFIG_PATH}")
        return
    config = load_config()
    CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")
    print(f"Created config: {CONFIG_PATH}")
    print("Put Gmail OAuth desktop credentials at:")
    print(config["gmail_credentials"])


def build_query(day: dt.date, config: Dict[str, Any]) -> str:
    next_day = day + dt.timedelta(days=1)
    base = (
        f"after:{gmail_date(day)} before:{gmail_date(next_day)} has:attachment "
        '(receipt OR invoice OR "tour ref" OR TKE OR filename:xlsx OR filename:xlsm '
        "OR filename:pdf OR filename:jpg OR filename:jpeg OR filename:png)"
    )
    known_sender = (config.get("known_sender_query") or "").strip()
    extra = (config.get("gmail_extra_query") or "").strip()
    parts = [p for p in [known_sender, base, extra] if p]
    return " ".join(parts)


def get_gmail_service(config: Dict[str, Any]):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except Exception as exc:  # pragma: no cover - dependency message
        raise SetupError(
            "Missing Gmail packages. Install google-api-python-client, "
            "google-auth-httplib2, and google-auth-oauthlib."
        ) from exc

    credentials_path = Path(config["gmail_credentials"]).expanduser()
    token_path = Path(config["gmail_token"]).expanduser()
    if not credentials_path.exists():
        raise SetupError(f"Missing Gmail credentials file: {credentials_path}")

    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def gmail_headers(payload: Dict[str, Any]) -> Dict[str, str]:
    headers = {}
    for header in payload.get("headers", []):
        headers[header.get("name", "").lower()] = header.get("value", "")
    return headers


def walk_parts(part: Dict[str, Any]) -> Iterable[Dict[str, Any]]:
    yield part
    for child in part.get("parts", []) or []:
        yield from walk_parts(child)


def decode_body(data: str) -> str:
    try:
        raw = base64.urlsafe_b64decode(data.encode("ascii"))
        return raw.decode("utf-8", errors="replace")
    except Exception:
        return ""


def message_text(payload: Dict[str, Any]) -> str:
    chunks = []
    for part in walk_parts(payload):
        mime = part.get("mimeType", "")
        body = part.get("body", {})
        data = body.get("data")
        if data and mime in {"text/plain", "text/html"}:
            chunks.append(decode_body(data))
    return "\n".join(chunks)


def extract_tour_ref(text: str) -> Optional[str]:
    match = TKE_RE.search(text or "")
    if match:
        return "TKE" + match.group(1)
    return None


def extract_date_token(filename: str, selected_day: Optional[dt.date] = None) -> Optional[str]:
    """Return a normalized attachment date token, favoring MMDD such as 0607."""
    candidates = []
    for match in DATE_TOKEN_RE.finditer(Path(filename).stem):
        mmdd = normalize_date_token(match.group(1))
        if mmdd:
            candidates.append(mmdd)
    if not candidates:
        return None
    if selected_day:
        selected_mmdd = selected_day.strftime("%m%d")
        if selected_mmdd in candidates:
            return selected_mmdd
    return candidates[0]


def normalize_date_token(token: str) -> Optional[str]:
    token = re.sub(r"\D", "", token)
    if len(token) == 8 and token.startswith("20"):
        mmdd = token[4:8]
    elif len(token) == 6:
        yy_mmdd = token[2:6]
        mmdd_yy = token[0:4]
        mmdd = yy_mmdd if valid_mmdd(yy_mmdd) else mmdd_yy
    elif len(token) == 4:
        mmdd = token
    else:
        return None
    return mmdd if valid_mmdd(mmdd) else None


def valid_mmdd(mmdd: str) -> bool:
    if not re.fullmatch(r"\d{4}", mmdd):
        return False
    month = int(mmdd[:2])
    day = int(mmdd[2:])
    return 1 <= month <= 12 and 1 <= day <= 31


def case_label_for_group(date_token: str, tour_ref: Optional[str], subject: str, msg_id: str) -> str:
    base = tour_ref or sanitize(subject[:45], "message_" + msg_id[:8])
    if date_token and date_token != "undated":
        return sanitize(f"{date_token}_{base}")
    return sanitize(base)


def supported_input_file(path: Path) -> bool:
    if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTS:
        return False
    return not path.stem.lower().startswith(GENERATED_PREFIXES)


def download_gmail(day: dt.date, out_root: Path, config: Dict[str, Any]) -> Path:
    service = get_gmail_service(config)
    day_folder = out_root / day.isoformat()
    day_folder.mkdir(parents=True, exist_ok=True)
    query = build_query(day, config)
    print(f"Gmail query: {query}")

    message_ids: List[str] = []
    request = service.users().messages().list(userId="me", q=query)
    while request is not None:
        response = request.execute()
        message_ids.extend([m["id"] for m in response.get("messages", [])])
        request = service.users().messages().list_next(request, response)

    for msg_id in message_ids:
        msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
        payload = msg.get("payload", {})
        headers = gmail_headers(payload)
        subject = headers.get("subject", "")
        sender_raw = headers.get("from", "")
        sender_name, sender_email = email.utils.parseaddr(sender_raw)
        sender_folder_name = sanitize(sender_email or sender_name or "unknown_sender")
        body = message_text(payload)
        evidence_text = "\n".join([subject, sender_raw, body])
        tour_ref = extract_tour_ref(evidence_text)
        received = headers.get("date", "")

        saved = 0
        for part in walk_parts(payload):
            filename = part.get("filename") or ""
            body_info = part.get("body", {})
            attachment_id = body_info.get("attachmentId")
            if not filename or not attachment_id:
                continue
            suffix = Path(filename).suffix.lower()
            if suffix not in SUPPORTED_EXTS:
                continue
            att = (
                service.users()
                .messages()
                .attachments()
                .get(userId="me", messageId=msg_id, id=attachment_id)
                .execute()
            )
            data = base64.urlsafe_b64decode(att["data"].encode("ascii"))
            date_token = extract_date_token(filename, day) or "undated"
            case_label = case_label_for_group(date_token, tour_ref, subject, msg_id)
            case_folder = day_folder / sender_folder_name / case_label
            original = case_folder / "original"
            original.mkdir(parents=True, exist_ok=True)
            (case_folder / "annotated").mkdir(parents=True, exist_ok=True)
            info = {
                "gmail_message_id": msg_id,
                "sender": sender_raw,
                "sender_email": sender_email,
                "subject": subject,
                "date_header": received,
                "tour_ref": tour_ref,
                "attachment_date_token": date_token,
                "query": query,
            }
            (case_folder / "email_info.json").write_text(json.dumps(info, indent=2), encoding="utf-8")
            target = original / sanitize(filename, f"attachment_{saved}{suffix}")
            if target.exists():
                target = original / f"{target.stem}_{saved}{target.suffix}"
            target.write_bytes(data)
            saved += 1

        if saved == 0:
            case_label = case_label_for_group("undated", tour_ref, subject, msg_id)
            case_folder = day_folder / sender_folder_name / case_label
            case_folder.mkdir(parents=True, exist_ok=True)
            (case_folder / "case_result.md").write_text(
                "# Case Result\n\nNo supported attachments were downloaded.\n",
                encoding="utf-8",
            )

    print(f"Downloaded {len(message_ids)} messages into {day_folder}")
    return day_folder


def rows_from_xlsx(path: Path) -> List[List[Any]]:
    try:
        import openpyxl

        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        rows: List[List[Any]] = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                rows.append(list(row))
        return rows
    except Exception:
        return rows_from_xlsx_zip(path)


def rows_from_xlsx_zip(path: Path) -> List[List[Any]]:
    rows: List[List[Any]] = []
    with zipfile.ZipFile(path) as zf:
        shared: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            import xml.etree.ElementTree as ET

            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            for si in root.findall(".//x:si", ns):
                texts = [t.text or "" for t in si.findall(".//x:t", ns)]
                shared.append("".join(texts))
        sheet_names = [n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        import xml.etree.ElementTree as ET

        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        for sheet in sheet_names:
            root = ET.fromstring(zf.read(sheet))
            for r in root.findall(".//x:sheetData/x:row", ns):
                vals: List[Any] = []
                for c in r.findall("x:c", ns):
                    v = c.find("x:v", ns)
                    if v is None:
                        vals.append("")
                    elif c.attrib.get("t") == "s":
                        vals.append(shared[int(v.text or "0")])
                    else:
                        text = v.text or ""
                        try:
                            vals.append(float(text) if "." in text else int(text))
                        except ValueError:
                            vals.append(text)
                rows.append(vals)
    return rows


def to_int_amount(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) and value > 0:
        return int(round(value))
    text = str(value)
    amounts = extract_amounts(text)
    return amounts[0] if amounts else None


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").lower()).strip()


def categorize_label(label: str) -> Optional[str]:
    s = normalize(label)
    if not s:
        return None
    if "consumption tax" in s or "消費税" in label:
        return None
    if "total invoice amount" in s or "grand total" in s:
        return "invoice_total"
    if "total fee" in s or ("guide fee" in s and "tour" in s):
        return "guide_fee"
    if "guide fee" in s or "fd tour" in s or "hd tour" in s:
        return "guide_fee"
    if "total entrance" in s:
        return "entrance_total"
    if "entrance" in s or "museum" in s or "temple" in s or "shrine" in s or "castle" in s or "green tea" in s:
        return "entrance"
    if "total transportation" in s:
        return "transportation_total"
    if "transportation" in s or "jr" in s or "metro" in s or "taxi" in s or "bus" in s or "pasmo" in s:
        return "transportation"
    if "total lunch" in s:
        return "lunch_total"
    if "lunch" in s or "cafe" in s or "café" in s or "restaurant" in s or "food & drink" in s:
        return "lunch"
    return None


def extract_invoice_items(path: Path) -> List[InvoiceItem]:
    rows = rows_from_xlsx(path)
    items: List[InvoiceItem] = []
    for idx, row in enumerate(rows, start=1):
        cells = [str(x).strip() for x in row if x not in (None, "")]
        if not cells:
            continue
        label = " | ".join(cells)
        category = categorize_label(label)
        if not category:
            continue
        amounts = [to_int_amount(x) for x in row]
        amounts = [x for x in amounts if x and x > 0]
        if not amounts:
            continue
        amount = amounts[-1]
        items.append(InvoiceItem(category=category, label=label[:240], amount=amount, source=str(path), row=idx))
    return dedupe_invoice_items(items)


def dedupe_invoice_items(items: List[InvoiceItem]) -> List[InvoiceItem]:
    guide_totals = [item for item in items if item.category == "guide_fee" and "total fee" in normalize(item.label)]
    if guide_totals:
        items = [
            item
            for item in items
            if item.category != "guide_fee" or "total fee" in normalize(item.label)
        ]
    seen = set()
    result = []
    for item in items:
        key = (item.category, item.label, item.amount)
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def extract_amounts(text: str) -> List[int]:
    amounts: List[int] = []
    for match in AMOUNT_RE.finditer(text or ""):
        token = match.group(1) or match.group(2)
        if not token:
            continue
        value = int(token.replace(",", ""))
        if 100 <= value <= 9999999:
            amounts.append(value)
    for token in re.findall(r"(?<!\d)([0-9]{1,3}(?:,[0-9]{3})+|[1-9][0-9]{2,6})(?!\d)", text or ""):
        value = int(token.replace(",", ""))
        if 100 <= value <= 9999999 and value not in amounts:
            amounts.append(value)
    return amounts


def extract_receipt_evidence_amounts(text: str, source: str) -> List[EvidenceAmount]:
    """Extract conservative receipt evidence amounts for matching.

    OCR sees many non-money numbers on Japanese receipts: phone numbers,
    registration numbers, ticket numbers, dates, and terminal IDs. For receipt
    matching, prefer money-looking values and derived tax-included totals.
    """
    candidates: List[Tuple[int, str, int]] = []
    seen = set()
    money_re = re.compile(
        r"(?P<prefix>JPY|YEN|yen|¥|￥|\\)?\s*"
        r"(?P<num>[0-9]{1,3}(?:,[0-9]{3})+|[1-9][0-9]{2,5})\s*"
        r"(?P<suffix>JPY|YEN|yen|円)?",
        re.I,
    )
    for match in money_re.finditer(text or ""):
        token = match.group("num")
        amount = int(token.replace(",", ""))
        has_money_signal = bool(match.group("prefix") or match.group("suffix") or "," in token)
        if not has_money_signal:
            continue
        strong_money_signal = "," in token or match.group("prefix") in {"\\", "¥", "￥"} or bool(match.group("suffix"))
        if not strong_money_signal and is_bad_receipt_number(text, match.start("num"), match.end("num"), amount):
            continue
        context = context_for_span(text, match.start("num"), match.end("num"))
        key = (amount, context)
        if key not in seen:
            candidates.append((amount, context, 80 if "," in token else 65))
            seen.add(key)

    derived = derive_tax_included_totals(candidates)
    for amount, context, score in derived:
        key = (amount, context)
        if key not in seen:
            candidates.append((amount, context, score))
            seen.add(key)

    candidates.sort(key=lambda x: (-x[2], -x[0], x[1]))
    return [EvidenceAmount(amount, source, context) for amount, context, _score in candidates]


def is_bad_receipt_number(text: str, start: int, end: int, amount: int) -> bool:
    if amount < 100 or amount > 999999:
        return True
    if 1900 <= amount <= 2099:
        return True
    before = text[max(0, start - 1):start]
    after = text[end:end + 1]
    if before == "-" or after == "-":
        return True
    context = context_for_span(text, start, end).lower()
    if re.search(r"\d{2,4}[- ]\d{2,4}[- ]\d{2,4}", context):
        return True
    bad_words = [
        "tel", "phone", "no.", "no ", "number", "id", "terminal", "register",
        "伝票", "番号", "登録", "取引", "端末", "電話",
    ]
    if any(word in context for word in bad_words) and amount >= 1000:
        return True
    return False


def context_for_span(text: str, start: int, end: int) -> str:
    left = max(0, start - 55)
    right = min(len(text), end + 85)
    return re.sub(r"\s+", " ", text[left:right]).strip()


def derive_tax_included_totals(candidates: List[Tuple[int, str, int]]) -> List[Tuple[int, str, int]]:
    amounts = [amount for amount, _context, _score in candidates]
    results: List[Tuple[int, str, int]] = []
    for base in amounts:
        for tax in amounts:
            if base <= tax:
                continue
            # Japanese receipts often show subtotal and 10% tax separately.
            if abs(tax - round(base * 0.10)) <= 3:
                total = base + tax
                if 100 <= total <= 999999:
                    results.append((total, f"Derived receipt total: JPY {base:,} + tax JPY {tax:,} = JPY {total:,}", 95))
    return results


def context_for_amount(text: str, amount: int) -> str:
    plain = str(amount)
    comma = f"{amount:,}"
    pos = -1
    for token in [comma, plain]:
        pos = text.find(token)
        if pos >= 0:
            break
    if pos < 0:
        return ""
    start = max(0, pos - 50)
    end = min(len(text), pos + 80)
    return re.sub(r"\s+", " ", text[start:end]).strip()


def ocr_image(path: Path, annotated_dir: Path, languages: str, issue_banner: str = "") -> Tuple[str, List[EvidenceAmount], Optional[str]]:
    try:
        from PIL import Image, ImageDraw, ImageFont
        import pytesseract
    except Exception:
        return "", [], "OCR not configured: install pillow, pytesseract, and Tesseract OCR."

    config = load_config()
    tesseract_cmd = config.get("tesseract_cmd")
    if tesseract_cmd and Path(tesseract_cmd).exists():
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    image = Image.open(path)
    ocr_image_obj = image.convert("RGB")
    try:
        text = pytesseract.image_to_string(ocr_image_obj, lang=languages)
    except Exception:
        text = pytesseract.image_to_string(ocr_image_obj)

    amounts = extract_receipt_evidence_amounts(text, str(path))

    annotated = ocr_image_obj.copy()
    draw = ImageDraw.Draw(annotated)
    try:
        data = pytesseract.image_to_data(ocr_image_obj, lang=languages, output_type=pytesseract.Output.DICT)
    except Exception:
        data = pytesseract.image_to_data(ocr_image_obj, output_type=pytesseract.Output.DICT)

    highlight_amounts = set()
    for ev in amounts:
        highlight_amounts.add(ev.amount)
        for derived_part in re.findall(r"JPY ([0-9,]+)", ev.context):
            try:
                highlight_amounts.add(int(derived_part.replace(",", "")))
            except ValueError:
                pass

    for i, word in enumerate(data.get("text", [])):
        word_amounts = set(extract_amounts(word))
        if word_amounts and word_amounts.intersection(highlight_amounts):
            x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
            draw.rectangle([x, y, x + w, y + h], outline=(230, 0, 0), width=4)
    if issue_banner:
        banner = issue_banner[:120]
        draw.rectangle([0, 0, annotated.width, 42], fill=(255, 255, 180))
        draw.text((10, 10), banner, fill=(160, 0, 0))

    target = annotated_dir / f"{path.stem}_annotated{path.suffix}"
    annotated.save(target)
    return text, amounts, None


def extract_pdf(path: Path) -> Tuple[str, List[EvidenceAmount], Optional[str]]:
    try:
        import fitz  # PyMuPDF
    except Exception:
        return "", [], "PDF extraction not configured: install pymupdf."
    text_parts = []
    with fitz.open(path) as doc:
        for page in doc:
            text_parts.append(page.get_text())
    text = "\n".join(text_parts)
    amounts = extract_receipt_evidence_amounts(text, str(path))
    return text, amounts, None


def find_case_dirs(day_folder: Path) -> List[Path]:
    case_dirs: List[Path] = []
    for original in day_folder.rglob("original"):
        if original.is_dir() and any(supported_input_file(p) for p in original.iterdir()):
            case_dirs.append(original.parent)
    if not case_dirs:
        # Split manually downloaded loose attachments by filename date token.
        direct_files = [p for p in day_folder.iterdir() if supported_input_file(p)]
        if not direct_files:
            return []
        grouped: Dict[str, List[Path]] = {}
        for file in direct_files:
            token = extract_date_token(file.name) or "undated"
            grouped.setdefault(token, []).append(file)
        for token, files in grouped.items():
            case_folder = day_folder / "local_files" / sanitize(token)
            original = case_folder / "original"
            original.mkdir(parents=True, exist_ok=True)
            (case_folder / "annotated").mkdir(exist_ok=True)
            for file in files:
                shutil.move(str(file), str(original / file.name))
            (case_folder / "email_info.json").write_text(
                json.dumps({"source": "local-input", "attachment_date_token": token}, indent=2),
                encoding="utf-8",
            )
            case_dirs.append(case_folder)
    return sorted(set(case_dirs))


def reconcile(items: List[InvoiceItem], evidence: List[EvidenceAmount], extraction_warnings: List[str]) -> Tuple[str, List[CheckRow]]:
    evidence_by_amount: Dict[int, List[EvidenceAmount]] = {}
    for ev in evidence:
        evidence_by_amount.setdefault(ev.amount, []).append(ev)

    rows: List[CheckRow] = []
    worst = "Passed"
    order = {"Passed": 0, "Mostly matched": 1, "Needs review": 2, "Incomplete": 3, "Mismatch": 4}

    def raise_status(status: str) -> None:
        nonlocal worst
        if order[status] > order[worst]:
            worst = status

    if extraction_warnings:
        rows.append(CheckRow("Extraction", "", "", "Needs review", "; ".join(sorted(set(extraction_warnings)))))
        raise_status("Needs review")

    if not items:
        rows.append(CheckRow("Invoice", "", "", "Incomplete", "No invoice spreadsheet/PDF data was extracted."))
        raise_status("Incomplete")

    receipt_required = {"entrance", "transportation", "lunch"}
    for item in items:
        if item.category in {"invoice_total", "entrance_total", "transportation_total", "lunch_total"}:
            continue
        label_amounts = [a for a in extract_amounts(item.label) if a != item.amount]
        if label_amounts:
            rows.append(
                CheckRow(
                    item.label,
                    f"JPY {item.amount:,}",
                    ", ".join(f"JPY {a:,}" for a in label_amounts),
                    "Needs review",
                    "The description contains a different amount from the entered invoice amount.",
                )
            )
            raise_status("Needs review")
            continue
        if item.category == "guide_fee":
            rows.append(CheckRow(item.label, f"JPY {item.amount:,}", "Fixed fee", "Mostly matched", "Guide/FD/HD fee is treated as fixed unless policy requires receipt."))
            raise_status("Mostly matched")
            continue
        matches = evidence_by_amount.get(item.amount, [])
        if matches:
            source = Path(matches[0].source).name
            rows.append(CheckRow(item.label, f"JPY {item.amount:,}", source, "Passed", matches[0].context))
        elif item.category in receipt_required:
            rows.append(CheckRow(item.label, f"JPY {item.amount:,}", "", "Needs review", "No matching receipt amount was found."))
            raise_status("Needs review")

    totals = {item.category: item.amount for item in items if item.category.endswith("_total") or item.category == "guide_fee"}
    invoice_total = totals.get("invoice_total")
    parts = [
        totals.get("guide_fee"),
        totals.get("entrance_total"),
        totals.get("transportation_total"),
        totals.get("lunch_total"),
    ]
    if invoice_total and all(x is not None for x in parts):
        computed = sum(x for x in parts if x is not None)
        if computed == invoice_total:
            rows.append(CheckRow("Invoice total calculation", f"JPY {invoice_total:,}", f"JPY {computed:,}", "Passed", "Category totals add up to invoice total."))
        else:
            rows.append(CheckRow("Invoice total calculation", f"JPY {invoice_total:,}", f"JPY {computed:,}", "Mismatch", "Category totals do not add up to invoice total."))
            raise_status("Mismatch")
    elif invoice_total:
        rows.append(CheckRow("Invoice total calculation", f"JPY {invoice_total:,}", "", "Needs review", "Not all category totals were extracted for recalculation."))
        raise_status("Needs review")

    return worst, rows


def filter_evidence_for_invoice(evidence: List[EvidenceAmount], items: List[InvoiceItem]) -> List[EvidenceAmount]:
    if not items:
        return evidence
    target_amounts = {
        item.amount
        for item in items
        if item.category not in {"guide_fee", "invoice_total"}
    }
    filtered: List[EvidenceAmount] = []
    seen = set()
    for ev in evidence:
        keep = ev.amount in target_amounts or ev.context.startswith("Derived receipt total:")
        if not keep:
            continue
        key = (ev.amount, ev.source, ev.context)
        if key not in seen:
            filtered.append(ev)
            seen.add(key)
    return filtered or evidence


def write_case_outputs(case_dir: Path, rows: List[CheckRow], status: str, extracted: Dict[str, Any]) -> None:
    md = ["# Case Result", "", f"Status: **{status}**", ""]
    md.append("| Item | Invoice Amount | Receipt Evidence | Result | Note |")
    md.append("|---|---:|---|---|---|")
    for row in rows:
        md.append(
            f"| {escape_md(row.item)} | {escape_md(row.invoice_amount)} | "
            f"{escape_md(row.receipt_evidence)} | {escape_md(row.result)} | {escape_md(row.note)} |"
        )
    (case_dir / "case_result.md").write_text("\n".join(md) + "\n", encoding="utf-8")

    with (case_dir / "case_result.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["item", "invoice_amount", "receipt_evidence", "result", "note"])
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))

    (case_dir / "extracted_data.json").write_text(json.dumps(extracted, indent=2, ensure_ascii=False), encoding="utf-8")
    create_case_summary_workbook(case_dir, rows, status, extracted)
    create_case_review_html(case_dir, rows, status, extracted)


def escape_md(value: str) -> str:
    return (value or "").replace("|", "\\|").replace("\n", " ")


def h(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def rel_link(from_dir: Path, target: Path) -> str:
    try:
        return Path(target).resolve().relative_to(from_dir.resolve()).as_posix()
    except Exception:
        return Path(target).as_posix()


def status_class(status: str) -> str:
    return {
        "Passed": "passed",
        "Mostly matched": "mostly",
        "Needs review": "review",
        "Incomplete": "incomplete",
        "Mismatch": "mismatch",
    }.get(status, "review")


def main_issue(rows: List[CheckRow]) -> str:
    for severity in ["Mismatch", "Incomplete", "Needs review"]:
        for row in rows:
            if row.result == severity:
                return row.note or row.item
    return "No major issue found."


def status_fill(status: str) -> str:
    return {
        "Passed": "C6EFCE",
        "Mostly matched": "D9EAD3",
        "Needs review": "FFF2CC",
        "Incomplete": "D9D9D9",
        "Mismatch": "F4CCCC",
    }.get(status, "FFF2CC")


def style_worksheet(ws) -> None:
    try:
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except Exception:
        return
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            widths[cell.column] = min(max(widths.get(cell.column, 8), len(str(cell.value or "")) + 2), 55)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def create_case_review_html(case_dir: Path, rows: List[CheckRow], status: str, extracted: Dict[str, Any]) -> Path:
    annotated = case_dir / "annotated"
    original = case_dir / "original"
    annotated_files = sorted([p for p in annotated.iterdir() if p.is_file()]) if annotated.exists() else []
    original_files = sorted([p for p in original.iterdir() if p.is_file()]) if original.exists() else []
    issue = main_issue(rows)
    html_rows = []
    for row in rows:
        html_rows.append(
            "<tr>"
            f"<td>{h(row.result)}</td>"
            f"<td>{h(row.item)}</td>"
            f"<td>{h(row.invoice_amount)}</td>"
            f"<td>{h(row.receipt_evidence)}</td>"
            f"<td>{h(row.note)}</td>"
            "</tr>"
        )
    images = []
    for file in annotated_files:
        if file.suffix.lower() in IMAGE_EXTS:
            images.append(
                f'<figure><a href="{h(rel_link(case_dir, file))}">'
                f'<img src="{h(rel_link(case_dir, file))}" alt="{h(file.name)}"></a>'
                f"<figcaption>{h(file.name)}</figcaption></figure>"
            )
    file_links = "".join(
        f'<li><a href="{h(rel_link(case_dir, file))}">{h(file.name)}</a></li>'
        for file in original_files
    )
    doc = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Case Review - {h(case_dir.name)}</title>
<style>
body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2933; }}
.status {{ display:inline-block; padding:8px 12px; border-radius:6px; font-weight:700; }}
.passed {{ background:#c6efce; }} .mostly {{ background:#d9ead3; }}
.review {{ background:#fff2cc; }} .incomplete {{ background:#d9d9d9; }} .mismatch {{ background:#f4cccc; }}
.issue {{ margin:16px 0; padding:12px 14px; background:#fff7ed; border-left:5px solid #f97316; }}
table {{ border-collapse: collapse; width:100%; margin-top:16px; }}
th, td {{ border:1px solid #d0d7de; padding:8px; vertical-align:top; font-size:14px; }}
th {{ background:#1f4e78; color:white; text-align:left; position:sticky; top:0; }}
.grid {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-top:16px; }}
figure {{ margin:0; border:1px solid #d0d7de; padding:8px; background:#fafafa; }}
img {{ max-width:100%; height:auto; display:block; }}
figcaption {{ font-size:12px; margin-top:6px; color:#52616b; word-break:break-all; }}
a {{ color:#0969da; }}
</style>
</head>
<body>
<h1>Case Review: {h(case_dir.name)}</h1>
<p><span class="status {status_class(status)}">{h(status)}</span></p>
<div class="issue"><strong>Main issue:</strong> {h(issue)}</div>
<p><a href="01_CASE_SUMMARY.xlsx">Open 01_CASE_SUMMARY.xlsx</a> | <a href="case_result.csv">Open CSV backup</a></p>
<h2>Line Check</h2>
<table>
<thead><tr><th>Result</th><th>Item</th><th>Invoice Amount</th><th>Receipt Evidence</th><th>Note</th></tr></thead>
<tbody>{''.join(html_rows)}</tbody>
</table>
<h2>Annotated Evidence</h2>
<div class="grid">{''.join(images) if images else '<p>No annotated image generated.</p>'}</div>
<h2>Original Attachments</h2>
<ul>{file_links}</ul>
</body>
</html>
"""
    target = case_dir / "00_OPEN_CASE_REVIEW.html"
    target.write_text(doc, encoding="utf-8")
    return target


def create_case_summary_workbook(case_dir: Path, rows: List[CheckRow], status: str, extracted: Dict[str, Any]) -> None:
    try:
        import openpyxl
        from openpyxl.styles import PatternFill
    except Exception:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    info = {}
    info_path = case_dir / "email_info.json"
    if info_path.exists():
        try:
            info = json.loads(info_path.read_text(encoding="utf-8"))
        except Exception:
            info = {}
    ws.append(["Field", "Value"])
    for key, value in [
        ("Status", status),
        ("Main Issue", main_issue(rows)),
        ("Case Folder", str(case_dir)),
        ("Sender", info.get("sender") or info.get("sender_email", "")),
        ("Subject", info.get("subject", "")),
        ("Tour Ref", info.get("tour_ref", "")),
        ("Attachment Date Token", info.get("attachment_date_token", "")),
    ]:
        ws.append([key, value])
    style_worksheet(ws)
    ws["B2"].fill = PatternFill("solid", fgColor=status_fill(status))

    ws2 = wb.create_sheet("Line_Check")
    ws2.append(["Result", "Item", "Invoice Amount", "Receipt Evidence", "Note"])
    for row in rows:
        ws2.append([row.result, row.item, row.invoice_amount, row.receipt_evidence, row.note])
        ws2.cell(ws2.max_row, 1).fill = PatternFill("solid", fgColor=status_fill(row.result))
    style_worksheet(ws2)

    ws3 = wb.create_sheet("Evidence")
    ws3.append(["Amount", "Source", "Context"])
    for ev in extracted.get("evidence_amounts", []):
        ws3.append([ev.get("amount"), Path(ev.get("source", "")).name, ev.get("context", "")])
    style_worksheet(ws3)
    wb.save(case_dir / "01_CASE_SUMMARY.xlsx")


def create_checked_workbook(case_dir: Path, invoice_files: List[Path], rows: List[CheckRow]) -> None:
    if not invoice_files:
        return
    try:
        import openpyxl
    except Exception:
        return
    src = invoice_files[0]
    try:
        wb = openpyxl.load_workbook(src)
    except Exception:
        return
    if "CHECK_RESULT" in wb.sheetnames:
        del wb["CHECK_RESULT"]
    ws = wb.create_sheet("CHECK_RESULT")
    ws.append(["Item", "Invoice Amount", "Receipt Evidence", "Result", "Note"])
    for row in rows:
        ws.append([row.item, row.invoice_amount, row.receipt_evidence, row.result, row.note])
    target = case_dir / "annotated" / "invoice_checked.xlsx"
    wb.save(target)


def process_case(case_dir: Path, config: Dict[str, Any]) -> Dict[str, Any]:
    original = case_dir / "original"
    annotated = case_dir / "annotated"
    annotated.mkdir(exist_ok=True)
    files = [p for p in original.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]
    invoice_files = [p for p in files if p.suffix.lower() in EXCEL_EXTS]
    image_files = [p for p in files if p.suffix.lower() in IMAGE_EXTS]
    pdf_files = [p for p in files if p.suffix.lower() in PDF_EXTS]

    warnings: List[str] = []
    items: List[InvoiceItem] = []
    evidence: List[EvidenceAmount] = []
    ocr_texts: Dict[str, str] = {}

    for xlsx in invoice_files:
        try:
            items.extend(extract_invoice_items(xlsx))
        except Exception as exc:
            warnings.append(f"Excel extraction failed for {xlsx.name}: {exc}")

    if not invoice_files and not pdf_files:
        warnings.append("No Excel/PDF invoice file found.")
    if not image_files and not pdf_files:
        warnings.append("No image/PDF receipt evidence found.")

    for pdf in pdf_files:
        text, amounts, warning = extract_pdf(pdf)
        if warning:
            warnings.append(warning)
        ocr_texts[pdf.name] = text
        evidence.extend(amounts)

    languages = config.get("ocr_languages", "jpn+eng")
    for image in image_files:
        text, amounts, warning = ocr_image(image, annotated, languages)
        if warning:
            warnings.append(warning)
        ocr_texts[image.name] = text
        evidence.extend(amounts)

    evidence = filter_evidence_for_invoice(evidence, items)
    status, rows = reconcile(items, evidence, warnings)
    create_checked_workbook(case_dir, invoice_files, rows)
    extracted = {
        "case_dir": str(case_dir),
        "invoice_items": [asdict(x) for x in items],
        "evidence_amounts": [asdict(x) for x in evidence],
        "warnings": warnings,
        "ocr_text": ocr_texts,
        "status": status,
    }
    write_case_outputs(case_dir, rows, status, extracted)
    return {
        "case": case_dir.name,
        "sender": case_dir.parent.name if case_dir.parent != case_dir else "",
        "status": status,
        "issue_count": sum(1 for r in rows if r.result in {"Needs review", "Mismatch", "Incomplete"}),
        "main_issue": main_issue(rows),
        "review_path": str(case_dir / "00_OPEN_CASE_REVIEW.html"),
        "path": str(case_dir),
    }


def write_daily_summary(day_folder: Path, summaries: List[Dict[str, Any]]) -> None:
    md = ["# Daily Review Summary", "", f"Folder: `{day_folder}`", ""]
    md.append("| Sender | Case | Status | Issues | Folder |")
    md.append("|---|---|---|---:|---|")
    for s in summaries:
        md.append(f"| {escape_md(s['sender'])} | {escape_md(s['case'])} | {s['status']} | {s['issue_count']} | {escape_md(s['path'])} |")
    (day_folder / "Daily_Review_Summary.md").write_text("\n".join(md) + "\n", encoding="utf-8")

    csv_path = day_folder / "Daily_Review_Summary.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["sender", "case", "status", "issue_count", "main_issue", "review_path", "path"],
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(summaries)

    create_daily_summary_workbook(day_folder, summaries)
    create_daily_dashboard_html(day_folder, summaries)


def create_daily_summary_workbook(day_folder: Path, summaries: List[Dict[str, Any]]) -> None:
    try:
        import openpyxl
        from openpyxl.styles import PatternFill
    except Exception:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Review"
    ws.append(["Status", "Priority", "Sender", "Case", "Issues", "Main Issue", "Open Review", "Folder"])
    priority = {"Mismatch": 1, "Incomplete": 2, "Needs review": 3, "Mostly matched": 4, "Passed": 5}
    for s in sorted(summaries, key=lambda x: (priority.get(x["status"], 9), x["sender"], x["case"])):
        review_path = Path(s["path"]) / "00_OPEN_CASE_REVIEW.html"
        ws.append([
            s["status"],
            priority.get(s["status"], 9),
            s["sender"],
            s["case"],
            s["issue_count"],
            s.get("main_issue", ""),
            "Open case review",
            s["path"],
        ])
        row_idx = ws.max_row
        ws.cell(row_idx, 1).fill = PatternFill("solid", fgColor=status_fill(s["status"]))
        if review_path.exists():
            ws.cell(row_idx, 7).hyperlink = str(review_path)
            ws.cell(row_idx, 7).style = "Hyperlink"
    style_worksheet(ws)
    target = day_folder / "00_DAILY_SUMMARY.xlsx"
    wb.save(target)
    wb.save(day_folder / "Daily_Review_Summary.xlsx")


def create_daily_dashboard_html(day_folder: Path, summaries: List[Dict[str, Any]]) -> Path:
    counts: Dict[str, int] = {}
    for row in summaries:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    cards = "".join(
        f'<div class="card {status_class(status)}"><strong>{h(status)}</strong><span>{counts.get(status, 0)}</span></div>'
        for status in ["Mismatch", "Incomplete", "Needs review", "Mostly matched", "Passed"]
    )
    priority = {"Mismatch": 1, "Incomplete": 2, "Needs review": 3, "Mostly matched": 4, "Passed": 5}
    rows = []
    for s in sorted(summaries, key=lambda x: (priority.get(x["status"], 9), x["sender"], x["case"])):
        review = Path(s["path"]) / "00_OPEN_CASE_REVIEW.html"
        review_link = rel_link(day_folder, review) if review.exists() else s["path"]
        rows.append(
            "<tr>"
            f'<td><span class="pill {status_class(s["status"])}">{h(s["status"])}</span></td>'
            f"<td>{h(s.get('issue_count', 0))}</td>"
            f"<td>{h(s.get('sender', ''))}</td>"
            f"<td>{h(s.get('case', ''))}</td>"
            f"<td>{h(s.get('main_issue', ''))}</td>"
            f'<td><a href="{h(review_link)}">Open review</a></td>'
            "</tr>"
        )
    doc = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Daily Review Dashboard</title>
<style>
body {{ font-family: Segoe UI, Arial, sans-serif; margin:24px; color:#1f2933; }}
.cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin:16px 0; }}
.card {{ padding:14px; border-radius:8px; border:1px solid #d0d7de; }}
.card span {{ display:block; font-size:28px; margin-top:6px; }}
.pill {{ display:inline-block; padding:5px 8px; border-radius:999px; font-weight:700; }}
.passed {{ background:#c6efce; }} .mostly {{ background:#d9ead3; }}
.review {{ background:#fff2cc; }} .incomplete {{ background:#d9d9d9; }} .mismatch {{ background:#f4cccc; }}
table {{ border-collapse:collapse; width:100%; }}
th,td {{ border:1px solid #d0d7de; padding:8px; vertical-align:top; font-size:14px; }}
th {{ background:#1f4e78; color:white; text-align:left; position:sticky; top:0; }}
a {{ color:#0969da; }}
</style>
</head>
<body>
<h1>Daily Review Dashboard</h1>
<p>Folder: {h(day_folder)}</p>
<p><a href="00_DAILY_SUMMARY.xlsx">Open 00_DAILY_SUMMARY.xlsx</a></p>
<div class="cards">{cards}</div>
<table>
<thead><tr><th>Status</th><th>Issues</th><th>Sender</th><th>Case</th><th>Main Issue</th><th>Review</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
</body>
</html>
"""
    target = day_folder / "00_OPEN_DAILY_DASHBOARD.html"
    target.write_text(doc, encoding="utf-8")
    return target


def process_day(day_folder: Path, config: Dict[str, Any]) -> List[Dict[str, Any]]:
    case_dirs = find_case_dirs(day_folder)
    summaries = []
    for case_dir in case_dirs:
        summaries.append(process_case(case_dir, config))
    write_daily_summary(day_folder, summaries)
    return summaries


def print_summary(summaries: List[Dict[str, Any]], day_folder: Path) -> None:
    counts: Dict[str, int] = {}
    for row in summaries:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    print(f"Output folder: {day_folder}")
    print(f"Cases processed: {len(summaries)}")
    for status in ["Passed", "Mostly matched", "Needs review", "Incomplete", "Mismatch"]:
        if counts.get(status):
            print(f"{status}: {counts[status]}")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Download/classify Gmail receipt cases and reconcile invoice attachments.",
        epilog=textwrap.dedent(
            """
            Examples:
              daily_receipt_checker.py --init-config
              daily_receipt_checker.py --date 2026-05-08 --gmail
              daily_receipt_checker.py --date 2026-05-08 --local-input C:\\Users\\Administrator\\Desktop\\Receipt_Check\\2026-05-08
            """
        ),
    )
    parser.add_argument("--date", help="Processing date in YYYY-MM-DD format.")
    parser.add_argument("--gmail", action="store_true", help="Search Gmail and download supported attachments.")
    parser.add_argument("--local-input", help="Process an existing day/case folder instead of Gmail.")
    parser.add_argument("--output-root", help="Override output root folder.")
    parser.add_argument("--init-config", action="store_true", help="Create default config and exit.")
    args = parser.parse_args(argv)

    if args.init_config:
        init_config()
        return 0

    if not args.date:
        raise SystemExit("--date is required unless --init-config is used.")

    day = parse_day(args.date)
    config = load_config()
    out_root = Path(args.output_root or config.get("output_root") or DEFAULT_OUTPUT_ROOT).expanduser()

    try:
        if args.gmail:
            day_folder = download_gmail(day, out_root, config)
        elif args.local_input:
            day_folder = Path(args.local_input).expanduser()
            if not day_folder.exists():
                raise SystemExit(f"Local input folder does not exist: {day_folder}")
        else:
            day_folder = out_root / day.isoformat()
            if not day_folder.exists():
                raise SystemExit("Use --gmail to download from Gmail or --local-input for an existing folder.")
        summaries = process_day(day_folder, config)
        print_summary(summaries, day_folder)
        return 0
    except SetupError as exc:
        print("Setup required:")
        print(exc)
        print()
        print(f"Run: python {Path(__file__).name} --init-config")
        print("Then read references/gmail-setup.md in this skill.")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
