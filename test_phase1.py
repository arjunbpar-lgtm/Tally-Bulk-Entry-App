from core.generator import generate_entries_for_day
from core.settings_manager import get_active_profile
from core.audit_logger import log_generation
from core.excel_handler import write_output_excel

def run_test():
    print("Testing Generation Engine...")
    config = get_active_profile()
    
    # Simulate a daily total
    target = 183000
    date_str = "01-03-2018"
    
    print(f"Target Total: {target}")
    
    try:
        entries, avg = generate_entries_for_day(date_str, target, config)
        print(f"Generated {len(entries)} entries. Avg: {avg:.2f}")
        
        amounts = [e['Amount'] for e in entries]
        print(f"Amounts: {amounts[:5]}... (first 5)")
        print(f"Sum: {sum(amounts)}")
        print(f"Confidence: {entries[0]['_Confidence']}")
        
        if sum(amounts) != target:
            print("ERROR: Total mismatch!")
        else:
            print("SUCCESS: Exact match.")
            
        write_output_excel(entries, "exports/test_output.xlsx")
        print("Test output written to exports/test_output.xlsx")
        
        log_generation("test_input.xlsx", len(entries), config.get("preset_mode"))
        print("Audit log updated.")
        
    except Exception as e:
        print(f"Test failed: {e}")

if __name__ == "__main__":
    import os
    os.makedirs("exports", exist_ok=True)
    run_test()
