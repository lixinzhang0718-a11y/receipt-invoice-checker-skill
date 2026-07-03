# Gmail Setup

Use this reference only when the Gmail workflow cannot run because credentials, token, or dependencies are missing.

## Required Python Packages

Install these packages in the Python environment used by Codex:

```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl pillow pytesseract pymupdf
```

`pytesseract` also requires the Tesseract OCR program installed on Windows. For Japanese receipts, install the Japanese language data and use `jpn+eng` when possible.

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
- No messages found: broaden the query to `after:YYYY/MM/DD before:YYYY/MM/DD has:attachment`.
- OCR weak or missing: install Tesseract and Japanese language data, then rerun local processing.
