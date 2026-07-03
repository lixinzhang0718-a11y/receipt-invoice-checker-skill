---
name: receipt-invoice-checker
description: Download, classify, and reconcile Gmail receipt/invoice cases for a selected date. Use when the user asks to check daily Gmail emails with Excel, PDF, and image attachments; compare receipt photos/PDFs against invoice spreadsheets; create sender-based folders; annotate suspicious attachment copies; and generate per-case plus daily review summaries.
---

# Receipt Invoice Checker

## Core Behavior

Use this skill to process daily Gmail receipt checking work. The user normally invokes it with a short command such as "use receipt-invoice-checker" or "check Gmail receipts for 2026-05-08". If no date is given, ask for one date in `YYYY-MM-DD` format before running any download or reconciliation step.

Protect evidence integrity:

- Never modify original Gmail attachments.
- Save raw files under `original/`.
- Save marked-up copies under `annotated/`.
- Save extracted text, OCR output, reconciliation tables, HTML dashboards, and Excel summaries as separate result files.
- If OCR, Gmail access, or dependencies are not configured, stop with a clear setup note instead of guessing.

## Standard Workflow

1. Confirm the processing date.
2. Read `references/gmail-setup.md` only if Gmail credentials, token, or dependency setup is missing.
3. Use `scripts/daily_receipt_checker.py` for the automated pass:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date YYYY-MM-DD --gmail
```

4. Review the generated daily folder:

```text
Receipt_Check/
  YYYY-MM-DD/
    00_OPEN_DAILY_DASHBOARD.html
    00_DAILY_SUMMARY.xlsx
    sender@example.com/
      0607_TKE1262583/
        00_OPEN_CASE_REVIEW.html
        01_CASE_SUMMARY.xlsx
        original/
        annotated/
        case_result.md
        case_result.csv
        extracted_data.json
        email_info.json
```

5. Give the user a concise Chinese summary:

- total cases processed
- passed cases
- review/mismatch/incomplete cases
- folder path
- highest priority issues

## Gmail Search Rules

For a selected date, search Gmail for messages with attachments and likely receipt/invoice meaning. Use a date-bounded query similar to:

```text
after:YYYY/MM/DD before:YYYY/MM/DD has:attachment (receipt OR invoice OR "tour ref" OR TKE OR filename:xlsx OR filename:xls OR filename:pdf OR filename:jpg OR filename:jpeg OR filename:png)
```

If Gmail search returns too many unrelated messages, narrow by known sender or subject keywords. If it returns too few, search only:

```text
after:YYYY/MM/DD before:YYYY/MM/DD has:attachment
```

Then filter locally by attachment type and text evidence.

## Classification Rules

Create one sender folder per sender email address. Create one case folder per likely attachment group. If one Gmail message contains several cases, group attachments by filename date token first. For example, all attachments containing `0607` belong together, and all attachments containing `0608` belong together.

Determine the case name from this priority order:

1. Attachment filename date token such as `0607` or `20260607`, normalized to `0607`.
2. Tour reference such as `TKE1262583`.
3. Guest name or clear subject token.
4. Message received time plus short subject.
5. Gmail message id if no other stable name exists.

Use only filesystem-safe names. Keep a copy of the original subject, sender, received time, and Gmail message id inside `email_info.json`.

## Reconciliation Rules

Read `references/matching-rules.md` when the result is ambiguous or when changing matching logic.

Default expense categories:

- `guide_fee`: guide fee or FD/HD tour fee. Usually no receipt required; mark as fixed fee unless the user's rules say otherwise.
- `entrance`: museum, temple, shrine, castle, green tea/food sampling if shown under entrance fees.
- `transportation`: JR, metro, bus, water bus, taxi, Pasmo charge, guide transport.
- `lunch`: cafe, restaurant, lunch budget, food and drink.
- `invoice_total`: final total invoice amount.

For receipt photos, do not treat every OCR number as a receipt amount. Use total-focused matching:

- Ignore obvious phone numbers, ticket numbers, registration numbers, years, terminal IDs, and long identifiers.
- Prefer amounts with currency signs, commas, or total-like receipt context.
- If a Japanese receipt shows subtotal and 10 percent tax separately, derive the tax-included total. Example: `JPY 3,380 + tax JPY 338 = JPY 3,718`.
- Show and match only amounts that connect to an invoice line or a derived receipt total. Keep raw OCR text in `extracted_data.json` for audit, but do not clutter the review page with irrelevant numbers.

Status meanings:

- `Passed`: invoice calculation is correct and all receipt-required items are matched.
- `Mostly matched`: main receipt amounts match, but fixed guide fee or minor note needs human awareness.
- `Needs review`: OCR is weak, receipt missing, note conflicts with entered amount, or category matching is uncertain.
- `Mismatch`: Excel/PDF invoice amount clearly differs from receipt amount.
- `Incomplete`: required Excel/PDF or image/PDF evidence is missing.

## Annotation Rules

Annotate copies only:

- For images, draw boxes around OCR-detected amounts and add a short issue banner if the case is not passed.
- For invoice spreadsheets, create `invoice_checked.xlsx` with a `CHECK_RESULT` sheet when `openpyxl` is available.
- For PDFs, extract text and save findings; only create annotated PDF copies when a PDF annotation library is available and reliable.

If annotation is not possible, still create `case_result.md`, `case_result.csv`, and `extracted_data.json`.

## Review Output Rules

Make `00_OPEN_DAILY_DASHBOARD.html` the main daily entry point. It must summarize all cases with status counts, issue counts, sender, case name, main issue, and links to each case review.

Make `00_OPEN_CASE_REVIEW.html` the main case entry point. It must show the case status, main issue, line-check table, annotated evidence images, and original attachment links.

Make Excel outputs usable for daily work:

- `00_DAILY_SUMMARY.xlsx`: colored daily dashboard with filter, frozen header, priority column, main issue, and hyperlinks to case review pages.
- `01_CASE_SUMMARY.xlsx`: per-case workbook with `Summary`, `Line_Check`, and `Evidence` sheets.

Keep `.md` and `.csv` only as backup outputs, not as the primary review format.

## Manual Review Expectations

Do not overclaim full automation. Japanese receipt OCR, handwritten notes, and scanned PDFs can be uncertain. Always surface uncertain items rather than forcing a pass.

For each questionable case, include an English note suitable for email follow-up, for example:

```text
Please confirm the guide transportation amount. The note says 2,276 yen, but the invoice amount entered is 1,500 yen.
```

## Useful Commands

Run the Gmail daily workflow:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --gmail
```

Run against an already downloaded local folder:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --local-input "C:\Users\Administrator\Desktop\Receipt_Check\2026-05-08"
```

Create a default config template:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --init-config
```
