import json
import os

SETTINGS_PATH = os.path.join("config", "settings.json")

def load_settings():
    if not os.path.exists(SETTINGS_PATH):
        # Provide defaults if file missing
        return {
            "active_profile": "Default Office",
            "profiles": {
                "Default Office": {
                    "preset_mode": "Mixed Rural Collection",
                    "safe_mode": False,
                    "min_amount": 2500,
                    "max_amount": 9900,
                    "min_entries_per_day": 12,
                    "max_entries_per_day": 40,
                    "voucher_type": "Journal",
                    "narration_text": "Being local rubber purchase paid in cash",
                    "ledger_name": "Local Rubber Purchase"
                }
            }
        }
    with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_settings(settings_data):
    os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings_data, f, indent=4)

def get_active_profile(settings_data=None):
    if settings_data is None:
        settings_data = load_settings()
    active = settings_data.get("active_profile", "Default Office")
    return settings_data.get("profiles", {}).get(active, {})
