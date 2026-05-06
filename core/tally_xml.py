import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

class TallyXMLEngine:
    def __init__(self, config):
        self.config = config
        
    def generate_xml_tree(self, export_days, export_batch_id):
        """
        export_days: list of DayReconstruction objects
        """
        envelope = ET.Element("ENVELOPE")
        header = ET.SubElement(envelope, "HEADER")
        tally_req = ET.SubElement(header, "TALLYREQUEST")
        tally_req.text = "Import Data"
        
        body = ET.SubElement(envelope, "BODY")
        import_data = ET.SubElement(body, "IMPORTDATA")
        req_desc = ET.SubElement(import_data, "REQUESTDESC")
        req_desc_report = ET.SubElement(req_desc, "REPORTNAME")
        req_desc_report.text = "Vouchers"
        
        req_desc_static = ET.SubElement(req_desc, "STATICVARIABLES")
        svc_company = ET.SubElement(req_desc_static, "SVCURRENTCOMPANY")
        svc_company.text = "##Company Name##" # Placeholder
        
        req_data = ET.SubElement(import_data, "REQUESTDATA")
        voucher_type = self.config.get('voucher_type', 'Journal')
        debit_ledger = self.config.get('debit_ledger_name', 'Local Rubber Purchase')
        credit_ledger = self.config.get('credit_ledger_name', 'Cash')
        
        for day in export_days:
            # Parse date to YYYYMMDD
            try:
                date_obj = datetime.strptime(day.date_str, "%d-%m-%Y")
                tally_date = date_obj.strftime("%Y%m%d")
            except ValueError:
                tally_date = day.date_str.replace("-", "")
                
            for entry in day.entries:
                amt = entry.get('Amount', 0)
                voucher_no = entry.get('Voucher No', '')
                narration = entry.get('Narration', '')
                
                tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", attrib={"xmlns:UDF": "TallyUDF"})
                voucher = ET.SubElement(tally_msg, "VOUCHER", attrib={"VCHTYPE": voucher_type, "ACTION": "Create"})
                
                v_date = ET.SubElement(voucher, "DATE")
                v_date.text = tally_date
                
                # Semantic Accounting Fields
                ET.SubElement(voucher, "OBJVIEW").text = "Accounting Voucher View"
                ET.SubElement(voucher, "PERSISTEDVIEW").text = "Accounting Voucher View"
                ET.SubElement(voucher, "ISINVOICE").text = "No"
                
                v_type_name = ET.SubElement(voucher, "VOUCHERTYPENAME")
                v_type_name.text = voucher_type
                
                if voucher_no:
                    v_number = ET.SubElement(voucher, "VOUCHERNUMBER")
                    v_number.text = voucher_no
                    
                    v_reference = ET.SubElement(voucher, "REFERENCE")
                    v_reference.text = voucher_no
                
                if narration:
                    v_narration = ET.SubElement(voucher, "NARRATION")
                    v_narration.text = narration
                    
                # Credit Entry (Cash/Supplier) -> ISDEEMEDPOSITIVE = No, Amount = Positive
                ledger_entry_cr = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                l_name_cr = ET.SubElement(ledger_entry_cr, "LEDGERNAME")
                l_name_cr.text = credit_ledger
                ET.SubElement(ledger_entry_cr, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(ledger_entry_cr, "ISPARTYLEDGER").text = "Yes"
                amt_cr = ET.SubElement(ledger_entry_cr, "AMOUNT")
                amt_cr.text = str(abs(amt)) # Credit is positive
                
                # Debit Entry (Purchase A/c) -> ISDEEMEDPOSITIVE = Yes, Amount = Negative
                ledger_entry_dr = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                l_name_dr = ET.SubElement(ledger_entry_dr, "LEDGERNAME")
                l_name_dr.text = debit_ledger
                ET.SubElement(ledger_entry_dr, "ISDEEMEDPOSITIVE").text = "Yes"
                amt_dr = ET.SubElement(ledger_entry_dr, "AMOUNT")
                amt_dr.text = str(-abs(amt)) # Debit is negative in Tally XML
                
        return envelope
        
    def validate_xml_structure(self, envelope):
        """
        Validates the XML structure and voucher totals before saving.
        """
        vouchers = envelope.findall(".//VOUCHER")
        if not vouchers:
            return False, "No vouchers found to export."
            
        seen_vouchers = set()
        
        for v in vouchers:
            v_no = v.find("VOUCHERNUMBER")
            v_no_text = v_no.text if (v_no is not None and v_no.text) else "Blank"
            
            # Simple Duplicate Check within this batch (only if not Blank)
            if v_no_text != "Blank":
                if v_no_text in seen_vouchers:
                    return False, f"Duplicate Voucher Number in batch: {v_no_text}"
                seen_vouchers.add(v_no_text)
                
            dr_total = 0
            cr_total = 0
            
            ledgers = v.findall("LEDGERENTRIES.LIST")
            if len(ledgers) < 2:
                return False, f"Voucher is missing ledger entries."
                
            for l in ledgers:
                amt_node = l.find("AMOUNT")
                if amt_node is None:
                    return False, f"Voucher {v_no.text} ledger missing AMOUNT."
                    
                try:
                    amt_val = float(amt_node.text)
                except ValueError:
                    return False, f"Voucher {v_no.text} has invalid AMOUNT: {amt_node.text}"
                    
                is_dr = l.find("ISDEEMEDPOSITIVE")
                if is_dr is not None and is_dr.text == "Yes":
                    dr_total += abs(amt_val)
                else:
                    cr_total += abs(amt_val)
                    
            # In Tally, the sum of debits must equal sum of credits
            if abs(dr_total - cr_total) > 0.01:
                return False, f"Voucher {v_no.text} unbalanced. Dr: {dr_total}, Cr: {cr_total}"
                
        return True, "XML Structure Validated Successfully."

    def generate_and_save(self, export_days, export_batch_id, output_path):
        envelope = self.generate_xml_tree(export_days, export_batch_id)
        
        is_valid, msg = self.validate_xml_structure(envelope)
        if not is_valid:
            raise ValueError(f"XML Validation Failed: {msg}")
            
        xml_str = ET.tostring(envelope, encoding='utf-8')
        reparsed = minidom.parseString(xml_str)
        pretty_xml = reparsed.toprettyxml(indent="    ")
        
        # Remove empty lines introduced by toprettyxml
        pretty_xml = '\n'.join([line for line in pretty_xml.split('\n') if line.strip()])
        
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(pretty_xml)
            
        return len(export_days)
