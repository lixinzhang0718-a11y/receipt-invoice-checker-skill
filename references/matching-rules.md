# Matching Rules

Use these rules when interpreting receipt/invoice results.

## Folder Logic

- Group by sender first.
- Group cases by `TKE` reference when available.
- If no `TKE` reference exists, use received time plus subject.
- Keep originals unchanged in `original/`.
- Put marked-up copies in `annotated/`.

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
