# Receipt Invoice Checker

A Codex skill for reviewing Gmail receipt and invoice attachments by case. The helper script downloads original Excel, PDF, and image attachments; Codex then reads the files and compares each invoice line with the available receipt evidence. The script does not decide whether amounts match.

The review is designed for travel expense cases, including Japanese receipts and invoices. Review output defaults to English and links each case's original files alongside a four-column comparison table.

## What it does

- Searches one selected Gmail date for messages with attachments using read-only Gmail access.
- Downloads supported Excel (`.xlsx`, `.xls`, `.xlsm`), PDF, and image (`.jpg`, `.jpeg`, `.png`, `.webp`, `.tif`, `.tiff`, `.bmp`) files.
- Groups files by sender, filename date token, tour reference, and subject.
- Gives Codex the original files to inspect directly, including combined receipt amounts and missing evidence.

It does not send or change email, modify attachments, upload credentials, or create a separate review report. The helper does not run OCR or perform automatic reconciliation.

## Install

Clone the repository into your Codex skills directory and install the Gmail dependencies in the Python environment used by Codex:

```powershell
git clone https://github.com/lixinzhang0718-a11y/receipt-invoice-checker-skill.git "$env:USERPROFILE\.codex\skills\receipt-invoice-checker"
python -m pip install -r "$env:USERPROFILE\.codex\skills\receipt-invoice-checker\requirements.txt"
```

For an existing clone, update it with `git -C "$env:USERPROFILE\.codex\skills\receipt-invoice-checker" pull`.

## Connect Gmail

1. Enable the Gmail API in a Google Cloud project and create an **OAuth desktop app** client.
2. Save its downloaded JSON file as `%USERPROFILE%\.codex\receipt-invoice-checker\gmail_credentials.json`.
3. Initialize the local configuration:

   ```powershell
   python "$env:USERPROFILE\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --init-config
   ```

4. On the first Gmail run, complete the Google consent flow. The token is stored locally at `%USERPROFILE%\.codex\receipt-invoice-checker\gmail_token.json`.

The requested Gmail scope is `gmail.readonly`. Keep the credentials, token, and downloaded attachments private. See [Gmail setup](references/gmail-setup.md) for troubleshooting.

## Run

Ask Codex to use `$receipt-invoice-checker` for a specific date, or run the downloader directly:

```powershell
python "$env:USERPROFILE\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --gmail
```

The script prints the folder and cases ready for review. By default, it writes to a newly created system temporary folder. **The script does not delete that folder automatically.** Remove it after review when no longer needed, or use `--keep-output` to save cases under the configured `output_root` (initially `Desktop\Receipt_Check`).

To list a previously downloaded folder without contacting Gmail:

```powershell
python "$env:USERPROFILE\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --local-input "$env:USERPROFILE\Desktop\Receipt_Check\2026-05-08"
```

The date search includes all emails with attachments for that Gmail search day. The script then keeps supported file types; Codex determines which cases are relevant. For a busy inbox, set `known_sender_query` or `gmail_extra_query` in `%USERPROFILE%\.codex\receipt-invoice-checker\config.json` to narrow the Gmail search. Gmail applies its own date-search timezone, so check adjacent dates if an email near midnight is missing.

## Review result

For each case, Codex links the original files and compares invoice items, subtotals, and the final total with the visible receipt evidence. It distinguishes confirmed mismatches from missing or unclear receipts, and notes fixed guide fees separately. The expected format is documented in [SKILL.md](SKILL.md) and [matching rules](references/matching-rules.md).

## Development

Run the self-contained tests from the repository root:

```powershell
python -m unittest discover -s tests -v
```

The tests use a fake Gmail service and do not require credentials or network access.
