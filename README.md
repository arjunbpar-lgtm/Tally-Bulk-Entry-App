# VoucherGen Suite

Welcome to the **VoucherGen Suite**! This is a powerful, modular accounting toolkit designed to bridge the gap between raw Excel data and Tally-compliant XML files.

## 📦 Modules Included

### 1. Excel to Tally Converter (from Totals)
*(Formerly Rubber Purchase Reconstruction)*
This module is built for historical bookkeeping and reconstruction. 
- **What it does**: Takes daily total cash purchase figures (e.g., from an Excel column) and generates a randomized, statistically natural distribution of individual small-value purchase vouchers.
- **Features**: Customizable min/max limits, preset statistical distribution curves (e.g., Heavy Collection Day, Small Farmer Market), and dynamic Ledger assignment.
- **Workflow**: 
  1. Import Excel with daily totals.
  2. Click **Generate All** to build natural distributions.
  3. Export to Excel (for manual numbering) or directly to Tally XML.

### 2. Bulk Tally XML Converter
This is a universal passthrough converter for **any** voucher type.
- **What it does**: Takes a strictly formatted Excel sheet and converts it directly into a Tally-compliant XML file, enforcing strict accounting view semantics.
- **Features**: Dynamically detects `Purchase` and `Sales` voucher types and automatically injects the critical `<ISPARTYLEDGER>Yes</ISPARTYLEDGER>` tag required by Tally. It handles `Date`, `Reference`, `Amount`, and `Narration` cleanly.
- **Workflow**:
  1. Click **Download Excel Template**.
  2. Fill out your exact vouchers in Excel.
  3. Click **Upload Excel & Generate XML** to instantly get your Tally file.

## 🚀 How to Run
Run the central dashboard using Python:

```bash
python app.py
```

The Dashboard allows you to launch both modules simultaneously, keeping their logic fully independent while utilizing shared Tally generation rules.

## 💡 Using Tooltips
Throughout both modules, you will see small `[?]` buttons. Hover over these with your mouse to receive context-specific help and tips regarding the adjacent setting or button.

## 🛠 Architecture ("Hub & Spoke")
The suite uses a central launcher (`app.py`) that opens individual app modules housed in the `apps/` directory. All shared utilities (like the Universal Tally XML engine and Tooltip generators) are kept in the `core/` directory to ensure future apps can easily plug into the suite.

---
*Created with the help of Google Antigravity*
*GitHub: [arjunbpar-lgtm](https://github.com/arjunbpar-lgtm)*
