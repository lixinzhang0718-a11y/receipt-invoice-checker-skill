#!/usr/bin/env python3
"""Download and classify Gmail receipt/invoice attachments.

This helper intentionally does not reconcile amounts or generate review files.
Codex performs the review directly from the downloaded originals and reports the
result in chat.
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import email.utils
import json
import re
import sys
import tempfile
import textwrap
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


CONFIG_DIR = Path.home() / ".codex" / "receipt-invoice-checker"
CONFIG_PATH = CONFIG_DIR / "config.json"
DEFAULT_OUTPUT_ROOT = Path.home() / "Desktop" / "Receipt_Check"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
PDF_EXTS = {".pdf"}
SUPPORTED_EXTS = IMAGE_EXTS | EXCEL_EXTS | PDF_EXTS

TKE_RE = re.compile(r"\bTKE[\s_-]*(\d{4,})\b", re.I)
DATE_TOKEN_RE = re.compile(r"(?<!\d)(20\d{6}|\d{6}|\d{4})(?!\d)")


class SetupError(RuntimeError):
    pass


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
        "gmail_extra_query": "",
        "known_sender_query": "",
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


def build_query(day: dt.date, config: Dict[str, Any], broad: bool = False) -> str:
    next_day = day + dt.timedelta(days=1)
    if broad:
        base = f"after:{gmail_date(day)} before:{gmail_date(next_day)} has:attachment"
    else:
        base = (
            f"after:{gmail_date(day)} before:{gmail_date(next_day)} has:attachment "
            '(receipt OR invoice OR "tour ref" OR TKE OR filename:xlsx OR filename:xls '
            "OR filename:xlsm OR filename:pdf OR filename:jpg OR filename:jpeg OR filename:png)"
        )
    known_sender = (config.get("known_sender_query") or "").strip()
    extra = (config.get("gmail_extra_query") or "").strip()
    return " ".join(p for p in [known_sender, base, extra] if p)


def get_gmail_service(config: Dict[str, Any]):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except Exception as exc:
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
        data = part.get("body", {}).get("data")
        if data and mime in {"text/plain", "text/html"}:
            chunks.append(decode_body(data))
    return "\n".join(chunks)


def extract_tour_ref(text: str) -> Optional[str]:
    match = TKE_RE.search(text or "")
    return "TKE" + match.group(1) if match else None


def extract_date_token(filename: str, selected_day: Optional[dt.date] = None) -> Optional[str]:
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
    return path.is_file() and not path.name.startswith("~$") and path.suffix.lower() in SUPPORTED_EXTS


def list_message_ids(service: Any, query: str) -> List[str]:
    message_ids: List[str] = []
    request = service.users().messages().list(userId="me", q=query)
    while request is not None:
        response = request.execute()
        message_ids.extend([m["id"] for m in response.get("messages", [])])
        request = service.users().messages().list_next(request, response)
    return message_ids


def unique_target(original: Path, filename: str, index: int) -> Path:
    suffix = Path(filename).suffix.lower()
    target = original / sanitize(filename, f"attachment_{index}{suffix}")
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    counter = 2
    while True:
        candidate = original / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def download_gmail(day: dt.date, out_root: Path, config: Dict[str, Any]) -> Path:
    service = get_gmail_service(config)
    day_folder = out_root / day.isoformat()
    day_folder.mkdir(parents=True, exist_ok=True)

    query = build_query(day, config)
    print(f"Gmail query: {query}")
    message_ids = list_message_ids(service, query)
    if not message_ids:
        query = build_query(day, config, broad=True)
        print(f"No narrow-query messages found. Retrying: {query}")
        message_ids = list_message_ids(service, query)

    downloaded = 0
    for msg_id in message_ids:
        msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
        payload = msg.get("payload", {})
        headers = gmail_headers(payload)
        subject = headers.get("subject", "")
        sender_raw = headers.get("from", "")
        sender_name, sender_email = email.utils.parseaddr(sender_raw)
        sender_folder_name = sanitize(sender_email or sender_name or "unknown_sender")
        evidence_text = "\n".join([subject, sender_raw, message_text(payload)])
        tour_ref = extract_tour_ref(evidence_text)

        attachment_index = 0
        for part in walk_parts(payload):
            filename = part.get("filename") or ""
            attachment_id = part.get("body", {}).get("attachmentId")
            if not filename or not attachment_id:
                continue
            if Path(filename).suffix.lower() not in SUPPORTED_EXTS:
                continue

            attachment = (
                service.users()
                .messages()
                .attachments()
                .get(userId="me", messageId=msg_id, id=attachment_id)
                .execute()
            )
            data = base64.urlsafe_b64decode(attachment["data"].encode("ascii"))
            date_token = extract_date_token(filename, day) or "undated"
            case_label = case_label_for_group(date_token, tour_ref, subject, msg_id)
            original = day_folder / sender_folder_name / case_label / "original"
            original.mkdir(parents=True, exist_ok=True)
            target = unique_target(original, filename, attachment_index)
            target.write_bytes(data)
            downloaded += 1
            attachment_index += 1

    print(f"Messages found: {len(message_ids)}")
    print(f"Attachments downloaded: {downloaded}")
    return day_folder


def case_from_original(original: Path, day_folder: Path) -> Dict[str, Any]:
    files = sorted(p.name for p in original.iterdir() if supported_input_file(p))
    sender = original.parent.parent.name if original.parent.parent != day_folder else ""
    return {
        "sender": sender,
        "case": original.parent.name,
        "path": str(original.parent),
        "files": files,
    }


def list_downloaded_cases(day_folder: Path) -> List[Dict[str, Any]]:
    cases: List[Dict[str, Any]] = []
    for original in sorted(day_folder.rglob("original")):
        if original.is_dir() and any(supported_input_file(p) for p in original.iterdir()):
            cases.append(case_from_original(original, day_folder))

    if cases:
        return cases

    loose_files = sorted(p for p in day_folder.iterdir() if supported_input_file(p)) if day_folder.exists() else []
    grouped: Dict[str, List[str]] = {}
    for file in loose_files:
        token = extract_date_token(file.name) or "undated"
        grouped.setdefault(token, []).append(file.name)
    for token, files in grouped.items():
        cases.append(
            {
                "sender": "local_files",
                "case": sanitize(token),
                "path": str(day_folder),
                "files": files,
            }
        )
    return cases


def print_download_summary(day_folder: Path) -> None:
    cases = list_downloaded_cases(day_folder)
    print(f"Download folder: {day_folder}")
    print(f"Cases ready for direct review: {len(cases)}")
    for case in cases:
        print(f"- {case['sender']} / {case['case']}")
        print(f"  Folder: {case['path']}")
        for file in case["files"]:
            print(f"  File: {file}")
    print("Generated review files: none")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Download/classify Gmail receipt cases for direct Codex review.",
        epilog=textwrap.dedent(
            """
            Examples:
              daily_receipt_checker.py --init-config
              daily_receipt_checker.py --date 2026-05-08 --gmail
              daily_receipt_checker.py --date 2026-05-08 --gmail --keep-output
              daily_receipt_checker.py --date 2026-05-08 --local-input C:\\Users\\Administrator\\Desktop\\Receipt_Check\\2026-05-08
            """
        ),
    )
    parser.add_argument("--date", help="Processing date in YYYY-MM-DD format.")
    parser.add_argument("--gmail", action="store_true", help="Search Gmail and download supported attachments.")
    parser.add_argument("--local-input", help="List an existing day/case folder instead of Gmail.")
    parser.add_argument("--output-root", help="Override output root folder.")
    parser.add_argument(
        "--keep-output",
        action="store_true",
        help="Use the configured Receipt_Check output folder instead of a temporary folder.",
    )
    parser.add_argument("--init-config", action="store_true", help="Create default config and exit.")
    args = parser.parse_args(argv)

    if args.init_config:
        init_config()
        return 0

    if not args.date:
        raise SystemExit("--date is required unless --init-config is used.")

    day = parse_day(args.date)
    config = load_config()
    if args.keep_output or args.output_root:
        out_root = Path(args.output_root or config.get("output_root") or DEFAULT_OUTPUT_ROOT).expanduser()
    else:
        out_root = Path(tempfile.mkdtemp(prefix=f"receipt_review_{day.isoformat()}_"))

    try:
        if args.gmail:
            day_folder = download_gmail(day, out_root, config)
        elif args.local_input:
            day_folder = Path(args.local_input).expanduser()
            if not day_folder.exists():
                raise SystemExit(f"Local input folder does not exist: {day_folder}")
        else:
            raise SystemExit("Use --gmail to download from Gmail or --local-input for an existing folder.")
        print_download_summary(day_folder)
        return 0
    except SetupError as exc:
        print("Setup required:")
        print(exc)
        print()
        print(f"Run: python {Path(__file__).name} --init-config")
        print("Then read references/gmail-setup.md in this skill.")
        return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
