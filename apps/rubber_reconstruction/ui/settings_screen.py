import tkinter as tk
from tkinter import ttk, messagebox
from core.ui_utils import create_tooltip

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, project):
        super().__init__(parent)
        self.project = project
        self.title("Settings")
        self.geometry("550x650")
        self.transient(parent)
        self.grab_set()

        self.setup_ui()
        self.load_current_settings()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Generation Settings", font=("TkDefaultFont", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky=tk.W)

        ttk.Label(main_frame, text="Minimum Amount (₹):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.min_amt_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.min_amt_var).grid(row=1, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "The absolute lowest amount a single voucher can be.", use_grid=True, row=1, column=2, padx=5)

        ttk.Label(main_frame, text="Maximum Amount (₹):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.max_amt_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.max_amt_var).grid(row=2, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "The absolute highest amount a single voucher can be.", use_grid=True, row=2, column=2, padx=5)

        ttk.Label(main_frame, text="Min Entries per Day:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.min_entries_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.min_entries_var).grid(row=3, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Max Entries per Day:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.max_entries_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.max_entries_var).grid(row=4, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Preset Mode:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(main_frame, textvariable=self.preset_var, state="readonly")
        self.preset_combo['values'] = ("Mixed Rural Collection", "Small Farmer Market", "Heavy Collection Day", "Rainy/Low Arrival Day")
        self.preset_combo.grid(row=5, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "Changes the statistical distribution curve of the generated amounts.", use_grid=True, row=5, column=2, padx=5)

        ttk.Label(main_frame, text="Multiplication of (e.g., 10, 50):").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.rounding_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.rounding_var).grid(row=6, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "Forces all generated amounts to be divisible by this number.", use_grid=True, row=6, column=2, padx=5)

        ttk.Label(main_frame, text="Debit Ledger Name:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.debit_ledger_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.debit_ledger_var).grid(row=7, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Credit Ledger Name:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.credit_ledger_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.credit_ledger_var).grid(row=8, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Voucher Numbering Mode:").grid(row=9, column=0, sticky=tk.W, pady=5)
        self.numbering_mode_var = tk.StringVar()
        self.numbering_mode_combo = ttk.Combobox(main_frame, textvariable=self.numbering_mode_var, state="readonly")
        self.numbering_mode_combo['values'] = ("Automatic (Continuous)", "Manual (Blank)")
        self.numbering_mode_combo.grid(row=9, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "Automatic creates continuous numbers across all days. Manual leaves them blank for you to fill later.", use_grid=True, row=9, column=2, padx=5)

        ttk.Label(main_frame, text="Voucher Prefix (For Auto):").grid(row=10, column=0, sticky=tk.W, pady=5)
        self.voucher_prefix_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.voucher_prefix_var).grid(row=10, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Starting Number (For Auto):").grid(row=11, column=0, sticky=tk.W, pady=5)
        self.start_num_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.start_num_var).grid(row=11, column=1, sticky=tk.EW, pady=5)

        ttk.Label(main_frame, text="Narration Mode:").grid(row=12, column=0, sticky=tk.W, pady=5)
        self.narration_mode_var = tk.StringVar()
        self.narration_mode_combo = ttk.Combobox(main_frame, textvariable=self.narration_mode_var, state="readonly")
        self.narration_mode_combo['values'] = ("Automatic (Pre-filled)", "Manual (Blank)")
        self.narration_mode_combo.grid(row=12, column=1, sticky=tk.EW, pady=5)
        create_tooltip(main_frame, "Automatic writes the text below into every voucher. Manual leaves it empty.", use_grid=True, row=12, column=2, padx=5)

        ttk.Label(main_frame, text="Narration Text (For Auto):").grid(row=13, column=0, sticky=tk.W, pady=5)
        self.narration_text_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.narration_text_var).grid(row=13, column=1, sticky=tk.EW, pady=5)

        self.safe_mode_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Safe Mode (Smoother Distribution)", variable=self.safe_mode_var).grid(row=14, column=0, columnspan=2, sticky=tk.W, pady=5)
        create_tooltip(main_frame, "Prevents extreme outliers and ensures amounts are tightly packed towards the middle.", use_grid=True, row=14, column=2, padx=5)

        main_frame.grid_columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=15, column=0, columnspan=2, pady=20, sticky=tk.E)

        ttk.Button(btn_frame, text="Save", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def load_current_settings(self):
        config = self.project.config
        self.min_amt_var.set(str(config.get('min_amount', 2500)))
        self.max_amt_var.set(str(config.get('max_amount', 9900)))
        self.min_entries_var.set(str(config.get('min_entries_per_day', 12)))
        self.max_entries_var.set(str(config.get('max_entries_per_day', 40)))
        self.preset_var.set(config.get('preset_mode', 'Mixed Rural Collection'))
        self.rounding_var.set(str(config.get('rounding_multiple', 10)))
        self.debit_ledger_var.set(config.get('debit_ledger_name', 'Local Rubber Purchase'))
        self.credit_ledger_var.set(config.get('credit_ledger_name', 'Cash'))
        self.numbering_mode_var.set(config.get('numbering_mode', 'Automatic (Continuous)'))
        self.voucher_prefix_var.set(config.get('voucher_prefix', 'LP'))
        self.start_num_var.set(str(config.get('voucher_start_num', 1)))
        self.narration_mode_var.set(config.get('narration_mode', 'Automatic (Pre-filled)'))
        self.narration_text_var.set(config.get('narration_text', 'Being cash paid for local rubber purchase'))
        self.safe_mode_var.set(config.get('safe_mode', False))

    def save_settings(self):
        try:
            min_amt = int(self.min_amt_var.get())
            max_amt = int(self.max_amt_var.get())
            min_entries = int(self.min_entries_var.get())
            max_entries = int(self.max_entries_var.get())
            rounding_mult = int(self.rounding_var.get())
            
            if min_amt > max_amt:
                raise ValueError("Minimum amount cannot be greater than maximum amount.")
            if min_entries > max_entries:
                raise ValueError("Min entries cannot be greater than max entries.")
            if rounding_mult < 1:
                raise ValueError("Multiplication factor must be at least 1.")

            self.project.config['min_amount'] = min_amt
            self.project.config['max_amount'] = max_amt
            self.project.config['min_entries_per_day'] = min_entries
            self.project.config['max_entries_per_day'] = max_entries
            self.project.config['preset_mode'] = self.preset_var.get()
            self.project.config['rounding_multiple'] = rounding_mult
            self.project.config['debit_ledger_name'] = self.debit_ledger_var.get()
            self.project.config['credit_ledger_name'] = self.credit_ledger_var.get()
            self.project.config['numbering_mode'] = self.numbering_mode_var.get()
            self.project.config['voucher_prefix'] = self.voucher_prefix_var.get()
            self.project.config['narration_mode'] = self.narration_mode_var.get()
            self.project.config['narration_text'] = self.narration_text_var.get()
            
            try:
                self.project.config['voucher_start_num'] = int(self.start_num_var.get())
            except ValueError:
                raise ValueError("Starting number must be a valid integer.")
                
            self.project.config['safe_mode'] = self.safe_mode_var.get()

            # Apply to core settings manager as well to persist
            from core.settings_manager import load_settings, save_settings
            full_settings = load_settings()
            active = full_settings.get("active_profile", "Default Office")
            if "profiles" not in full_settings:
                full_settings["profiles"] = {}
            if active not in full_settings["profiles"]:
                full_settings["profiles"][active] = {}
                
            full_settings["profiles"][active].update(self.project.config)
            save_settings(full_settings)

            messagebox.showinfo("Success", "Settings saved successfully.")
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
