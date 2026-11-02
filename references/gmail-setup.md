# Gmail Setup

Use this reference only when the Gmail workflow cannot run because credentials, token, or dependencies are missing.

## Required Python Packages

Install these packages in the Python environment used by Codex:

```powershell
python -m pip install -r "$env:USERPROFILE\.codex\skills\receipt-invoice-checker\requirements.txt"
```

The downloader needs only the Gmail packages in `requirements.txt`. To inspect spreadsheets, PDFs, or images locally, install suitable readers such as `openpyxl`, `pymupdf`, and `pillow` in the environment performing the review. OCR is optional; `pytesseract` requires the Tesseract program and Japanese language data for `jpn+eng`.

## Gmail OAuth Files

Create a Google Cloud OAuth desktop client for the Gmail API, then download the client secret JSON file. Save it as:

```text
C:\Users\Administrator\.codex\receipt-invoice-checker\gmail_credentials.json
```

Run:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --init-config
```

Then run the daily command. On first use, the script opens a Google consent flow and saves:

```text
C:\Users\Administrator\.codex\receipt-invoice-checker\gmail_token.json
```

## Gmail Permission Scope

Use read-only Gmail access:

```text
https://www.googleapis.com/auth/gmail.readonly
```

The workflow downloads attachments and metadata only. It should not send, delete, archive, or mark messages unless the user explicitly asks for a future version that does so.

## Common Failures

- Missing `gmail_credentials.json`: create/download OAuth desktop credentials.
- Browser consent fails: rerun the command and complete the URL shown by Google.
- No messages found: verify the selected date, Gmail account, and any filters in `config.json`.
- Receipt text is hard to read: inspect the original image; OCR tools are optional aids to the manual review.
