import os
import datetime

AUDIT_LOG_PATH = "audit.log"

def log_generation(input_file, rows_generated, settings_profile, additional_notes=""):
    """
    Appends a generation event to the audit log.
    """
    timestamp = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    log_entry = (
        f"[{timestamp}] GENERATION EVENT\n"
        f"Input File: {input_file}\n"
        f"Rows Generated: {rows_generated}\n"
        f"Settings Profile: {settings_profile}\n"
    )
    if additional_notes:
        log_entry += f"Notes: {additional_notes}\n"
    log_entry += "-" * 40 + "\n"

    try:
        with open(AUDIT_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Failed to write to audit log: {e}")
