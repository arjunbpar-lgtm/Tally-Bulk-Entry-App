# HR Feeds Purchase Format (v2.0 Reference)

## Source File
- `C:\Users\bparj\Downloads\PURCHASE REPORT -FEBRUARY 2026.xlsx`

## Workbook Layout
- Primary sheet: `Sheet1`
- Header row: `6`
- Data starts: `10` in current sample
- Non-data narrative rows seen at `7`, `8` (must skip)

## Columns (Row 6)
1. `SlNo`
2. `Date`
3. `Inv.Dt`
4. `VCh.No`
5. `AcName`
6. `City`
7. `State`
8. `State Code`
9. `GST No`
10. `ItemName`
11. `HSN`
12. `Item group`
13. `Qty`
14. `Unit`
15. `Amount`
16. `SGST%`
17. `SGST`
18. `CGST%`
19. `CGST`
20. `IGST%`
21. `IGSTA`
22. `Total`

## Required Mapping for Purchase Import
- Voucher Date: `Date`
- Voucher Number: `VCh.No`
- Supplier Ledger: `AcName`
- Stock Item: `ItemName`
- HSN: `HSN`
- Quantity: `Qty`
- Unit: `Unit`
- Line Amount: prefer `Total` (fallback `Amount`)

## Grouping Rules
- Group stock lines into one Purchase voucher by:
  - `Date + VCh.No + AcName`

## Skip Rules
- Skip rows where `SlNo` is not numeric (e.g., report headings)
- Skip empty rows
- Skip summary rows like `ItemName = Total`

## Sample Stats (February 2026 file)
- Data rows: `13`
- Purchase vouchers after grouping: `6`
- Suppliers in sample: `1` (`SHANTHI FEEDS PRIVATE LIMITED`)
- Main stock items: 3 variants (Finisher / Starter / Pre Starter)
