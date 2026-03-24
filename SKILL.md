---
name: pdf-invoice-to-excel
description: Extracts structured line-item data from invoice PDFs — both text-based and scanned image PDFs — into clean Excel files. One Excel file per invoice. Use when someone uploads invoice PDFs and needs the data in Excel, says "extract line items from this invoice", "convert this invoice to Excel", or explicitly calls /pdf-invoice-to-excel.
---

## Extraction Strategy

Use pdfplumber FIRST to try extracting text (it's fast and works on most digital PDFs - 10x faster than OCR)
Fall back to Tesseract OCR only if pdfplumber returns empty or garbled text.

When using OCR:
- Use Tesseract OCR with proper preprocessing (deskew, denoise, contrast enhancement) when needed
- Double-check numbers - common OCR errors: 0/O, 8/3, 1/l, 5/S, 6/b, 9/g
- Verify extracted values make logical sense (prices, quantities, dates)


## Input

Find invoice PDFs in this order:
1. Any PDFs uploaded directly in the conversation
2. Any mounted/accessible local folder
3. If neither is clear, ask the user where the invoices are


## What to Extract
Create a standardized line items table with these columns (in this order):

1. **Invoice No** - Extract from invoice header/top (usually labeled "Invoice #", "Invoice No", "INV#")
2. **Tariff Code / HS Code** - Extract from line items (look for "HS:", "HS Code:", "Tariff:", numbers like "6802 9100" or "9401390000")
   - If embedded in description (e.g., "Cube in marble (HS: 6802 9100)"), extract it separately
   - Clean format: remove "HS:" prefix, remove ALL spaces (both prefix, suffix, and middle)
   - Example: "HS: 6802 9100" → "68029100"
   - Example: " 9401 3900 00 " → "9401390000"
3. **DESCRIPTION** - Product/service description WITHOUT the HS code
   - Remove HS codes, tariff codes from description
   - Keep core product details (material, size, color, specifications)
4. **UNIT** - Unit of measurement if present (PCS, KG, M, CTN, SET, etc.)
   - If quantity shows "33 PCS", extract "PCS" here
5. **WT** - Weight if present (optional - mark "[NOT FOUND]" if missing)
6. **AMOUNT** - Total price for this line item (not unit price)
7. **COO** - Country of Origin
   - Look at bottom of invoice, in line items, or notes section
   - Common labels: "Origin:", "Country of Origin:", "Made in:", "COO:"
   - If same for all items, repeat for each row

**CRITICAL rules:**
- **Preserve EVERY line item exactly as shown** - do NOT consolidate or merge items even if they look similar
- **Different SKU/model/size/color = different line item, always**
- Skip rows that are clearly headers like "Packaging", "Bank charges", "TOTAL" - only extract actual product/service lines


## Examples of Field Extraction

**HS Code extraction:**
- Input: "Cube in Pink marble 24×19×H5 cm (HS: 6802 9100)"
- Tariff Code column: `68029100` (no spaces)
- Description column: `Cube in Pink marble 24×19×H5 cm`

**Country of Origin:**
- Look for "Made in China", "Origin: Italy", "COO: India" anywhere in the invoice
- If all items are from same country, repeat it for every row

**Unit extraction:**
- Input: "33 PCS" in quantity column
- UNIT column: `PCS`

**Skip non-product rows:**
- "Packaging (for Kit Marble Cubes)" → Skip this row
- "Bank charges" → Skip this row
- "TOTAL" → Skip this row


## Excel Formatting
1. **Single sheet only** - don't create multiple sheets
2. **Mark unclear data** with "[UNCLEAR]" - don't guess
3. **Professional formatting**:
   - Blue headers with white bold text
   - Currency: `$#,##0.00` format for Amount column
   - Dates: consistent format (YYYY-MM-DD or MM/DD/YYYY)
   - Percentages: 0.0% format
   - Thin borders around data cells
   - Auto-size columns for readability
4. **Confidence color coding (cell background):**:
   - 🟢 Green background: High confidence extraction
   - 🟡 Yellow background: Verify recommended
   - 🔴 Red background: Needs manual review
5. **No metadata or banking details** - no "extracted by", no confidence legends, no source notes unless specifically requested by user
6. **Simple layout**: Just the line items table with column headers at top


## Math Verification
- If invoice shows both quantity and unit price: verify Quantity × Unit Price = AMOUNT
- If the math doesn't match, flag the row with YELLOW background
- Add a comment in the AMOUNT cell: "Math error: {qty} × {unit_price} = {calculated} but shows {extracted}"
- If invoice only shows total amounts without unit prices, skip verification for that row
- Common OCR errors: 0 vs O, 8 vs 3, 1 vs l, 5 vs S, 6 vs b - check these first before flagging


## Output File Naming
`InvoiceData_{InvoiceNumber}_{VendorName}_{Date}.xlsx`

Use `UNKNOWN` as placeholder if invoice number or vendor name cannot be found.

Examples:
- `InvoiceData_INV2024001_ABCCorp_20240315.xlsx` (ideal)
- `InvoiceData_UNKNOWN_ABCCorp_20240315.xlsx` (no invoice number)
- `InvoiceData_INV2024001_UNKNOWN_20240315.xlsx` (no vendor name)


## Batch Processing & Output Requirements
When multiple invoices are provided, process each one independently.

For every invoice — no exceptions:
- Always produce a `.xlsx` file, never just text or JSON
- One Excel file per invoice
- The 7-column table must exist even if some cells are `[UNCLEAR]` or `[NOT FOUND]`

After completing a batch, provide a brief summary:
- How many invoices were processed
- Filenames created
- Any failures or issues encountered



