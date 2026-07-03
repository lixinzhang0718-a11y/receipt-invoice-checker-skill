# Matching Rules

Use these rules for direct reasoning-based receipt/invoice review. Do not rely only on automated OCR matching.

## Language Handling

- Default all final review output to English.
- Understand Chinese, Japanese, and English text in user requests, email subjects, spreadsheets, PDFs, and receipt photos.
- Preserve source-language item names when they are needed for evidence, but explain the result in English by default.
- Switch the final answer to Chinese or Japanese only when the user explicitly asks.

## Folder Logic

- Group by sender first.
- Inside each sender, group attachments by filename date token first. For example, `0607_receipt.jpg` and `0607_invoice.xlsx` belong to the same case, while `0608_receipt.jpg` belongs to another case.
- Normalize `YYYYMMDD`, `YYMMDD`, and `MMDD` filename tokens to `MMDD`.
- Group cases by `TKE` reference when no filename date token is available or as the second part of the case name.
- If no `TKE` reference exists, use received time plus subject.
- Default workflow uses temporary folders and does not create persistent review packs unless requested.

## Direct Review Logic

For each case, inspect Excel/PDF and receipt images together.

1. Extract all meaningful invoice lines:
   - guide fee
   - entrance / food sampling / coffee break
   - transportation
   - lunch
   - category totals
   - final invoice total
2. Inspect receipt images visually. If one image contains several receipts, treat each visible receipt area separately.
3. Match exact amounts when visible.
4. Match combinations when business logic is clear:
   - `1,000 + 1,000 = 2,000`
   - `178 + 252 + 155 + 178 = 763`
   - `3,380 + 338 = 3,718`
5. Recalculate category totals and final invoice total.
6. Separate amount mismatch from missing evidence and description conflict.

## Receipt Evidence

Usually require receipt evidence:

- museum, temple, shrine, castle, attraction admission
- green tea / food sampling / coffee break if entered as claimable item
- lunch, cafe, restaurant
- JR, metro, taxi, bus, water bus, Pasmo/Suica charge

Usually can be fixed fee without receipt:

- guide fee
- FD Tour / HD Tour fee

If company policy requires guide fee proof, treat guide fee as receipt-required.

## Common Reasoning Patterns

- A Japanese receipt may show subtotal and consumption tax separately; the payable total can be subtotal plus tax.
- Two stored-value card charge receipts can support one Excel line, such as `1,000 + 1,000 = 2,000`.
- A small transport history strip can support one guide/guest transport line by summing each ride.
- A handwritten receipt can support an Excel claim if the amount, date, and store/person context are visible enough.
- If Excel text says one quantity but the amount and receipt show another, mark it as description conflict.

## Status Labels

Use simple human labels:

- `Matched`: evidence supports the item.
- `Mostly matched`: main amounts are supported but minor items/descriptions need confirmation.
- `Needs confirmation`: evidence is missing, unclear, or description conflicts.
- `Mismatch`: a confirmed receipt amount differs from the invoice amount.
- `Incomplete`: required invoice or receipt attachment is missing.

## Output Focus

Final output should be concise and divided by case. Each case section must include the case's own attachments, a four-column comparison table, and a short explanation.

Use exactly this table structure for every case:

| Item | Amount shown in photos | Amount in Excel | Result |
|---|---:|---:|---|

Rules for the table:

- Include every meaningful Excel line item, subtotal, and final total.
- Put receipt combination evidence directly in the photo amount column.
- If evidence is missing or unclear, explain the issue in the photo amount column.
- Use clear result labels: `Matched`, `Needs confirmation`, `Mismatch`, `Calculation is correct`, or `Calculation mismatch`.
- Keep table text in English so it can be copied into replies or audit notes.

Before each table, list the exact case attachments:

- Link Excel and PDF files with clickable local Markdown links.
- Embed the most relevant local receipt images using Markdown image syntax with absolute paths.
- If there are many images, embed only the key images and link the rest.
- Never mix attachments from different case folders.

After each table, write a short English explanation:

- Main matched evidence
- Missing or unclear receipt evidence
- Total calculation result
- Any description conflict

Avoid raw OCR dumps unless the user asks for technical debugging.
