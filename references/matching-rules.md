# Matching Rules

Use these rules when interpreting receipt/invoice results.

## Folder Logic

- Group by sender first.
- Inside each sender, group attachments by filename date token first. For example, `0607_receipt.jpg` and `0607_invoice.xlsx` belong to the same case, while `0608_receipt.jpg` belongs to another case.
- Normalize `YYYYMMDD`, `YYMMDD`, and `MMDD` filename tokens to `MMDD`.
- Group cases by `TKE` reference when no filename date token is available or as the second part of the case name.
- If no `TKE` reference exists, use received time plus subject.
- Keep originals unchanged in `original/`.
- Put marked-up copies in `annotated/`.
- Use `00_OPEN_DAILY_DASHBOARD.html` and `00_OPEN_CASE_REVIEW.html` as the main human review entry points.

## Amount Logic

Treat invoice category totals separately from detail lines:

- `invoice_total`: final total payable.
- `guide_fee`: FD/HD tour fee or guide service fee.
- `entrance_total`: total entrance fees.
- `transportation_total`: total transportation.
- `lunch_total`: total lunch.

The daily checker should recalculate:

```text
guide fee + entrance total + transportation total + lunch total = invoice total
```

If this calculation fails, mark `Mismatch`.

## Receipt OCR Amount Logic

Use conservative evidence amounts for receipt images and PDFs:

- Do not use every OCR number as evidence. Japanese receipts contain phone numbers, registration numbers, ticket numbers, dates, and terminal IDs.
- Keep money-looking amounts first: currency sign, yen symbol, comma-formatted number, or total context.
- Exclude clear identifiers such as years, phone-number fragments, and long ID-like numbers.
- For Japanese receipts with subtotal and 10 percent tax, create a derived total when the tax is about 10 percent of the subtotal. For example, `3,380 + 338 = 3,718`.
- After extracting receipt evidence, filter displayed evidence to amounts that can match Excel invoice lines or derived receipt totals. Raw OCR text can remain in `extracted_data.json`.

## Receipt-Required Items

Usually require receipt evidence:

- museum, temple, shrine, castle, attraction admission
- green tea / food sampling if entered as entrance or food item
- lunch, cafe, restaurant
- JR, metro, taxi, bus, water bus, Pasmo charge

Usually can be fixed fee without receipt:

- guide fee
- FD Tour / HD Tour fee

If the user later gives a company policy that guide fees need proof, update this rule.

## Notes and Conflicts

Mark `Needs review` when:

- a description contains one amount but the entered amount differs
- handwritten note contains a different amount
- OCR detects the right store but not the right amount
- Excel has a category amount but no receipt/PDF/image supports it
- a receipt date falls outside the selected tour date

Example:

```text
Transportation for guide (Yokosuka-Shimbashi) says 2276yen, but amount entered is 1500.
```

## Status Choice

Use the strictest status found inside the case:

```text
Mismatch > Incomplete > Needs review > Mostly matched > Passed
```

Use `Mostly matched` only when the math is correct and all receipt-required line items are matched, but a fixed fee or minor note still deserves human awareness.

## Line Check Display

The human-facing case review must include every invoice item extracted from Excel/PDF:

- Detail lines such as entrance, transportation, lunch and bottled water.
- Category totals such as total entrance fees, total transportation and total lunch.
- Final invoice total.
- Fixed guide fee.

Use row-level wording that is easy to scan:

- `matched` for receipt-supported amounts.
- `not matched` for missing or conflicting receipt evidence.
- `fixed fee` for accepted guide/FD/HD fees.
- `calculation matched` or `calculation mismatch` for the invoice total calculation.
