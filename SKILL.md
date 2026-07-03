---
name: receipt-invoice-checker
description: Retrieve Gmail receipt/invoice attachments for a selected date and perform direct reasoning-based reconciliation. Use when the user asks to check daily Gmail emails with Excel, PDF, and image attachments; download attachments to a temporary workspace; compare receipts against invoice spreadsheets manually with model reasoning; and output one section per case with attached file links/images, a four-column comparison table, and a short explanation without generating annotation, HTML dashboard, or Excel review files.
---

# Receipt Invoice Checker

## Default Behavior

Use this skill to process daily Gmail receipt checking work. If no date is given, ask for one date in `YYYY-MM-DD` format.

Default to **reasoning review mode**, not automated rule-report mode:

1. Search Gmail for the selected date.
2. Download relevant Excel/PDF/image attachments to a temporary workspace.
3. Group attachments by sender, subject, Tour Ref, and filename date token such as `0607`.
4. Directly inspect the Excel/PDF/image contents.
5. Reconcile amounts using human-style reasoning.
6. Output one clean chat section per case, with the case's attachments, a four-column comparison table, and a short explanation.
7. Do not generate annotation images, HTML dashboards, Excel review files, markdown review files, or CSV review files.
8. Delete or ignore the temporary workspace after the final answer unless the user asks to keep it.

Never modify Gmail attachments. Never upload credentials, tokens, downloaded emails, or attachments.

## Language Rules

Default all chat-facing output to English, including overview text, case headings, tables, short explanations, and suggested replies.

Remain compatible with Chinese and Japanese:

- Understand Chinese, Japanese, and English user requests.
- Read Chinese/Japanese/English text in email subjects, Excel/PDF files, and receipt images when visible.
- Preserve original place names, item names, and notes when they matter for evidence, but explain the reconciliation result in English by default.
- Reply in Chinese or Japanese only if the user explicitly asks for that language.

## Gmail Download Command

Use this command for the default lightweight workflow:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date YYYY-MM-DD --gmail
```

The command prints a temporary folder path. Use that path as the working source for direct analysis. It downloads only original attachments and creates no review artifacts. The folder structure remains simple:

```text
receipt_review_YYYY-MM-DD_xxx/
  YYYY-MM-DD/
    sender@example.com/
      0607_Subject_or_TKE/
        original/
          invoice.xlsx
          receipt.jpg
```

If Gmail access or dependencies fail, read `references/gmail-setup.md`.

## Gmail Search Rules

For a selected date, search Gmail with a date-bounded attachment query similar to:

```text
after:YYYY/MM/DD before:YYYY/MM/DD has:attachment (receipt OR invoice OR "tour ref" OR TKE OR filename:xlsx OR filename:xls OR filename:pdf OR filename:jpg OR filename:jpeg OR filename:png)
```

If too few messages are found, broaden to:

```text
after:YYYY/MM/DD before:YYYY/MM/DD has:attachment
```

Then filter locally by attachment types and visible case relevance.

## Grouping Rules

Group first by sender. Inside each sender, group attachments by:

1. Filename date token, such as `0607`, `20260607`, or `260607`, normalized to `0607`.
2. Tour reference such as `TKE1262583`.
3. Subject/guest name.
4. Gmail message id if no better label exists.

If one email contains multiple cases, filename date token controls grouping. For example, all `0607` Excel/images belong together, while all `0608` files belong to another case.

## Direct Reconciliation Method

Do not rely only on OCR amount matching. Use model reasoning from the visible receipt image, Excel rows, filenames, and email context.

For each case:

1. Read the Excel/PDF invoice and list every meaningful amount:
   - guide fee
   - entrance/food sampling/coffee break
   - transportation
   - lunch
   - category totals
   - final invoice total
2. Inspect receipt images visually. When one image contains multiple receipts, separate them mentally by region.
3. Match exact amounts where visible.
4. Support combination matching:
   - `1,000 + 1,000 = 2,000`
   - `178 + 252 + 155 + 178 = 763`
   - `3,380 + 338 = 3,718`
5. Mark fixed guide fees separately if no receipt is expected.
6. Mark missing or unclear receipt evidence as needing confirmation.
7. Identify description conflicts, such as Excel text saying `1,900 x 2` while receipts show `1,000 x 2`.

## Output Format

Return a concise English answer unless the user explicitly asks for Chinese or Japanese. Keep the internal template ASCII-safe.

Use this shape. Separate cases clearly in the conversation so each case can be checked independently.

```text
Overview:
- Total cases:
- Matched:
- Needs confirmation:
- Mismatch:

## Case [number]: [case name]
Status: [Matched / Mostly matched / Needs confirmation / Mismatch / Incomplete]

Attachments for manual check:
- Excel/PDF: [filename](absolute local path)
- Image: [filename](absolute local path)

![filename](absolute local image path)

| Item | Amount shown in photos | Amount in Excel | Result |
|---|---:|---:|---|
| Nezu Museum | JPY 6,000 | JPY 6,000 | Matched |
| Entrance fees subtotal | JPY 6,000 + JPY 2,250 = JPY 8,250 | JPY 8,250 | Matched |
| Guide transportation: Yokosuka-Shimbashi | No matching receipt was found in the photos; Excel note says JPY 2,276, but the amount entered is JPY 1,500 | JPY 1,500 | Needs confirmation |
| Total amount | Based on Excel calculation: JPY 30,000 + JPY 8,250 + JPY 5,520 + JPY 5,450 | JPY 49,220 | Calculation is correct |

Short explanation:
[2-4 English sentences explaining the main matched evidence, missing evidence, calculation result, and any description conflict.]

Suggested reply:
[short English follow-up if needed]
```

Table requirements:

- Always use exactly these four columns: `Item`, `Amount shown in photos`, `Amount in Excel`, `Result`.
- Include every meaningful Excel line item and subtotal in the table, even if no receipt is found.
- Use `Matched`, `Needs confirmation`, `Mismatch`, `Calculation is correct`, or `Calculation mismatch` in the `Result` column.
- Put combination evidence directly in `Amount shown in photos`, for example `JPY 1,000 + JPY 1,000 = JPY 2,000`.
- If the receipt is missing, write a short reason in `Amount shown in photos`, not only `N/A`.
- Keep table wording and the short explanation in English unless the user explicitly asks for another language.

Attachment display requirements:

- At the start of each case section, list the exact local paths of that case's Excel/PDF/image attachments as clickable Markdown links.
- For local images, also embed them with Markdown image syntax using absolute paths when there are only a few images and the chat will remain readable.
- If a case has many images, embed the most relevant receipt images and link the rest.
- Do not mix attachments from different cases. The attachments shown in a case section must come from that case's folder only.

Prefer practical conclusions over raw OCR dumps. Do not show long extracted text unless the user asks.

## Useful Commands

Initialize config:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --init-config
```

Download to temporary folder:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --gmail
```

List an already downloaded local folder without generating review files:

```bash
python "%USERPROFILE%\.codex\skills\receipt-invoice-checker\scripts\daily_receipt_checker.py" --date 2026-05-08 --local-input "C:\Users\Administrator\Desktop\Receipt_Check\2026-05-08"
```
