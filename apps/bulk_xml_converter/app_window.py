import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from xml.etree.ElementTree import Element, SubElement, ElementTree
import webbrowser
from core.ui_utils import create_tooltip

class BulkXMLScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.github_url = "https://github.com/arjunbpar-lgtm"
        self.setup_ui()

    def setup_ui(self):
        title = tk.Label(self, text="Bulk Tally XML Converter", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        subtitle = tk.Label(self, text="Excel → Tally XML Converter", font=("Arial", 10))
        subtitle.pack(pady=5)
        
        # Action Buttons wrapped in frames for tooltips
        f1 = tk.Frame(self)
        f1.pack(pady=5)
        tk.Button(f1, text="Download Excel Template", command=self.download_template, width=35).pack(side=tk.LEFT)
        create_tooltip(f1, "Downloads a blank Excel sheet with the exact columns required for conversion.", side=tk.LEFT, padx=5)

        f2 = tk.Frame(self)
        f2.pack(pady=5)
        tk.Button(f2, text="Upload Excel & Generate XML", command=self.process_file, width=35).pack(side=tk.LEFT)
        create_tooltip(f2, "Converts your filled Excel template directly into a Tally-compliant XML file.", side=tk.LEFT, padx=5)

        tk.Button(self, text="Help / How to Use", command=self.show_help, width=35).pack(pady=5)

        footer_frame = tk.Frame(self)
        footer_frame.pack(side="bottom", pady=10)

        version_label = tk.Label(footer_frame, text="Version 1.0", font=("Arial", 8))
        version_label.pack(side="left", padx=5)

        antigravity_label = tk.Label(footer_frame, text="| Created with Google Antigravity |", font=("Arial", 8, "italic"), fg="gray")
        antigravity_label.pack(side="left", padx=5)

        github_label = tk.Label(
            footer_frame,
            text="GitHub Repo",
            fg="blue",
            cursor="hand2",
            font=("Arial", 8, "underline")
        )
        github_label.pack(side="left", padx=5)
        github_label.bind("<Button-1>", lambda e: webbrowser.open(self.github_url))

    # ---------- TEMPLATE DOWNLOAD ----------
    def download_template(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")],
                                                 initialfile="BulkXMLTemplate.xlsx")
        if not file_path:
            return

        df = pd.DataFrame(columns=[
            "Date", "Reference", "VoucherType", "DebitLedger",
            "CreditLedger", "Amount", "Narration"
        ])

        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "Template downloaded successfully.")

    # ---------- XML GENERATION ----------
    def create_xml(self, df, output_file):
        envelope = Element("ENVELOPE")
        header = SubElement(envelope, "HEADER")
        SubElement(header, "TALLYREQUEST").text = "Import Data"
        body = SubElement(envelope, "BODY")
        importdata = SubElement(body, "IMPORTDATA")
        requestdesc = SubElement(importdata, "REQUESTDESC")
        SubElement(requestdesc, "REPORTNAME").text = "Vouchers"
        requestdata = SubElement(importdata, "REQUESTDATA")

        for _, row in df.iterrows():
            tallymsg = SubElement(requestdata, "TALLYMESSAGE")
            tallymsg.set("xmlns:UDF", "TallyUDF")

            vch_type = str(row["VoucherType"]).strip()
            voucher = SubElement(tallymsg, "VOUCHER", VCHTYPE=vch_type, ACTION="Create")

            date_value = row["Date"]
            if pd.isna(date_value) or date_value == "":
                continue

            if isinstance(date_value, pd.Timestamp):
                formatted_date = date_value.strftime("%Y%m%d")
            else:
                date_str = str(date_value).strip()
                if " " in date_str:
                    date_str = date_str.split(" ")[0]
                if "-" in date_str:
                    parts = date_str.split("-")
                    formatted_date = parts[0] + parts[1] + parts[2]
                else:
                    formatted_date = date_str

            SubElement(voucher, "DATE").text = formatted_date

            ref_value = row.get("Reference", "")
            reference = "" if pd.isna(ref_value) else str(ref_value).strip()
            if reference:
                SubElement(voucher, "VOUCHERNUMBER").text = reference
                SubElement(voucher, "REFERENCE").text = reference

            SubElement(voucher, "VOUCHERTYPENAME").text = vch_type

            narration_val = row.get("Narration", "")
            narration = "" if pd.isna(narration_val) else str(narration_val).strip()
            if narration:
                SubElement(voucher, "NARRATION").text = narration

            SubElement(voucher, "OBJVIEW").text = "Accounting Voucher View"
            SubElement(voucher, "PERSISTEDVIEW").text = "Accounting Voucher View"
            SubElement(voucher, "ISINVOICE").text = "No"

            amount = float(row["Amount"])

            debit_entry = SubElement(voucher, "LEDGERENTRIES.LIST")
            SubElement(debit_entry, "LEDGERNAME").text = str(row["DebitLedger"]).strip()
            SubElement(debit_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            SubElement(debit_entry, "AMOUNT").text = str(-abs(amount))

            credit_entry = SubElement(voucher, "LEDGERENTRIES.LIST")
            SubElement(credit_entry, "LEDGERNAME").text = str(row["CreditLedger"]).strip()
            SubElement(credit_entry, "ISDEEMEDPOSITIVE").text = "No"
            
            if vch_type.lower() in ['purchase', 'sales']:
                SubElement(credit_entry, "ISPARTYLEDGER").text = "Yes"
                
            SubElement(credit_entry, "AMOUNT").text = str(abs(amount))

        tree = ElementTree(envelope)
        tree.write(output_file, encoding="utf-8", xml_declaration=True)

    def process_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            df = df.dropna(how='all')
            required_cols = ["Date", "Reference", "VoucherType", "DebitLedger",
                             "CreditLedger", "Amount", "Narration"]
            for col in required_cols:
                if col not in df.columns:
                    raise Exception(f"Missing column: {col}")

            output_file = filedialog.asksaveasfilename(defaultextension=".xml",
                                                       filetypes=[("XML Files", "*.xml")],
                                                       initialfile="BulkGeneratedVouchers.xml")
            if not output_file:
                return

            self.create_xml(df, output_file)
            messagebox.showinfo("Success", "XML generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def show_help(self):
        help_text = """Tally XML Generator - How to Use

1. Download Excel Template
2. Fill the data:
Columns: Date (YYYYMMDD), Reference, VoucherType, DebitLedger, CreditLedger, Amount, Narration
3. Upload Excel
4. Output XML ready for Tally!"""
        messagebox.showinfo("Help / Instructions", help_text)
