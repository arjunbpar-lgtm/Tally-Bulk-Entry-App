import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from apps.rubber_reconstruction.project_manager import Project
from core.excel_handler import read_input_excel, create_template, write_output_excel
from core.tally_xml import TallyXMLEngine
from apps.rubber_reconstruction.ui.preview_screen import PreviewScreen
from apps.rubber_reconstruction.ui.settings_screen import SettingsDialog
from core.ui_utils import create_tooltip
from core.theme import configure_window

class ExportPreviewDialog(tk.Toplevel):
    def __init__(self, parent, project):
        super().__init__(parent)
        self.project = project
        self.title("Export Preview & Summary")
        self.geometry("400x300")
        
        self.only_frozen_var = tk.BooleanVar(value=True)
        self.batch_id = None
        self.only_frozen_var = tk.BooleanVar(value=True)
        self.batch_id = None
        self.export_days = []
        
        configure_window(self)
        self.setup_ui()
        self.refresh_summary()
        
    def setup_ui(self):
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Checkbutton(frame, text="Export ONLY Frozen (Locked) Days", variable=self.only_frozen_var, command=self.refresh_summary).pack(anchor=tk.W, pady=5)
        
        self.summary_text = tk.Text(frame, height=8, width=40, state=tk.DISABLED)
        self.summary_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Dry Run (Validate)", command=self.dry_run, style="Info.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Export XML", command=self.export_xml, style="Success.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        
    def refresh_summary(self):
        try:
            self.batch_id, self.export_days = self.project.create_export_snapshot(only_frozen=self.only_frozen_var.get())
            
            total_days = len(self.export_days)
            total_entries = sum(len(d.entries) for d in self.export_days)
            total_amt = sum(d.original_total for d in self.export_days)
            
            summary = (
                f"Export Batch: {self.batch_id}\n\n"
                f"Days Exported: {total_days}\n"
                f"Entries: {total_entries}\n"
                f"Total Amount: {total_amt:,.2f}\n"
                f"Frozen Days Only: {'Yes' if self.only_frozen_var.get() else 'No'}"
            )
            
            self.summary_text.config(state=tk.NORMAL)
            self.summary_text.delete(1.0, tk.END)
            self.summary_text.insert(tk.END, summary)
            self.summary_text.config(state=tk.DISABLED)
        except Exception as e:
            self.summary_text.config(state=tk.NORMAL)
            self.summary_text.delete(1.0, tk.END)
            self.summary_text.insert(tk.END, f"Error generating snapshot:\n{e}")
            self.summary_text.config(state=tk.DISABLED)
            self.export_days = []
            
    def dry_run(self):
        if not self.export_days:
            messagebox.showwarning("Warning", "No data to validate.")
            return
            
        engine = TallyXMLEngine(self.project.config)
        envelope = engine.generate_xml_tree(self.export_days, self.batch_id)
        is_valid, msg = engine.validate_xml_structure(envelope)
        
        if is_valid:
            messagebox.showinfo("Dry Run", f"SUCCESS!\n{msg}\nXML is Tally-compliant.")
        else:
            messagebox.showerror("Dry Run Failed", f"XML Validation Error:\n{msg}")
            
    def export_xml(self):
        if not self.export_days:
            messagebox.showwarning("Warning", "No data to export.")
            return
            
        engine = TallyXMLEngine(self.project.config)
        envelope = engine.generate_xml_tree(self.export_days, self.batch_id)
        is_valid, msg = engine.validate_xml_structure(envelope)
        
        if not is_valid:
            messagebox.showerror("Export Blocked", f"Validation Failed:\n{msg}")
            return
            
        filepath = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")], initialfile=f"TallyExport_{self.batch_id}.xml")
        if filepath:
            try:
                engine.generate_and_save(self.export_days, self.batch_id, filepath)
                messagebox.showinfo("Success", f"Exported {len(self.export_days)} days to {filepath}")
                self.destroy()
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to save XML:\n{e}")

class ErrorLogDialog(tk.Toplevel):
    def __init__(self, parent, errors):
        super().__init__(parent)
        self.title("Generation Error Log")
        self.title("Generation Error Log")
        self.geometry("600x400")
        configure_window(self)
        
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Total Errors: {len(errors)}", font=("TkDefaultFont", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        text_area = tk.Text(frame, wrap=tk.WORD)
        text_area.pack(fill=tk.BOTH, expand=True)
        
        for err in errors:
            text_area.insert(tk.END, err + "\n\n")
            
        text_area.config(state=tk.DISABLED)
        
        ttk.Button(frame, text="Close", command=self.destroy).pack(pady=10)

class HomeScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.project = Project()
        
        self.setup_menu()
        self.setup_ui()
        self.bind_shortcuts()
        
    def setup_menu(self):
        menubar = tk.Menu(self.parent)
        self.parent.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Project", command=self.new_project)
        file_menu.add_command(label="Open Project...", command=self.open_project, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Save Project", command=self.save_project, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As...", command=self.save_project_as)
        file_menu.add_separator()
        file_menu.add_command(label="Download Template", command=self.download_template)
        file_menu.add_command(label="Import Excel (Daily Totals)...", command=self.import_excel, accelerator="Ctrl+I")
        file_menu.add_separator()
        file_menu.add_command(label="Convert Final Excel to XML...", command=self.convert_final_excel_to_xml)
        file_menu.add_separator()
        file_menu.add_command(label="Export Tally XML...", command=self.export_tally_xml, accelerator="Ctrl+E")
        file_menu.add_command(label="Export Excel...", command=self.export_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.parent.quit)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=self.undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=self.redo, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="Settings...", command=self.open_settings)

        process_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Process", menu=process_menu)
        process_menu.add_command(label="Generate All (Unlocked)", command=self.generate_all, accelerator="F5")
        process_menu.add_separator()
        process_menu.add_command(label="View Generation Log...", command=self.view_log)
        
    def bind_shortcuts(self):
        self.parent.bind("<Control-s>", lambda e: self.save_project())
        self.parent.bind("<Control-o>", lambda e: self.open_project())
        self.parent.bind("<Control-i>", lambda e: self.import_excel())
        self.parent.bind("<Control-z>", lambda e: self.undo())
        self.parent.bind("<Control-y>", lambda e: self.redo())
        self.parent.bind("<F5>", lambda e: self.generate_all())

    def setup_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Action Toolbar
        toolbar = ttk.Frame(self, padding="5 5 5 5")
        toolbar.grid(row=0, column=0, sticky="ew")
        
        ttk.Button(toolbar, text="📂 Import Excel", command=self.import_excel, style="TButton").pack(side=tk.LEFT, padx=2)
        create_tooltip(toolbar, "Imports your daily totals from an Excel file.", side=tk.LEFT, padx=(0, 10))

        ttk.Button(toolbar, text="⚙️ Settings", command=self.open_settings, style="Warning.TButton").pack(side=tk.LEFT, padx=2)
        create_tooltip(toolbar, "Configure generation rules like min/max amounts.", side=tk.LEFT, padx=(0, 10))

        ttk.Button(toolbar, text="▶️ Generate All", command=self.generate_all, style="Success.TButton").pack(side=tk.LEFT, padx=2)
        create_tooltip(toolbar, "Generates randomized entries for all unlocked days.", side=tk.LEFT, padx=(0, 10))

        ttk.Button(toolbar, text="💾 Export XML", command=self.export_tally_xml, style="TButton").pack(side=tk.LEFT, padx=2)
        create_tooltip(toolbar, "Exports the final data directly into Tally-compliant XML.", side=tk.LEFT, padx=(0, 10))

        ttk.Button(toolbar, text="📋 View Log", command=self.view_log, style="Info.TButton").pack(side=tk.LEFT, padx=2)
        
        # Legend
        legend_lbl = ttk.Label(toolbar, text="Legend: 'Entries' = Count per day | 'Avg Entry' = Average amount | 'Confidence' = Pattern quality %", foreground="grey")
        legend_lbl.pack(side=tk.RIGHT, padx=5)

        # Main Preview Grid
        self.preview = PreviewScreen(self, self.project)
        self.preview.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        # Summary Dashboard at bottom
        self.dashboard_frame = ttk.LabelFrame(self, text="Summary Dashboard")
        self.dashboard_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        
        self.lbl_days = ttk.Label(self.dashboard_frame, text="Total Days: 0")
        self.lbl_days.pack(side=tk.LEFT, padx=10)
        
        self.lbl_entries = ttk.Label(self.dashboard_frame, text="Total Entries: 0")
        self.lbl_entries.pack(side=tk.LEFT, padx=10)
        
        self.lbl_avg = ttk.Label(self.dashboard_frame, text="Avg Entry: 0")
        self.lbl_avg.pack(side=tk.LEFT, padx=10)
        
        self.lbl_max = ttk.Label(self.dashboard_frame, text="Max Entry: 0")
        self.lbl_max.pack(side=tk.LEFT, padx=10)

    def refresh_dashboard(self):
        total_days = len(self.project.days)
        total_entries = 0
        all_amts = []
        for day in self.project.days.values():
            total_entries += len(day.entries)
            all_amts.extend([e['Amount'] for e in day.entries])
            
        self.lbl_days.config(text=f"Total Days: {total_days}")
        self.lbl_entries.config(text=f"Total Entries: {total_entries}")
        
        avg = sum(all_amts)/len(all_amts) if all_amts else 0
        self.lbl_avg.config(text=f"Avg Entry: {avg:.2f}")
        
        max_amt = max(all_amts) if all_amts else 0
        self.lbl_max.config(text=f"Max Entry: {max_amt}")
        
        # Also refresh preview
        self.preview.refresh_grid()

    def new_project(self):
        if messagebox.askyesno("New Project", "Discard current project?"):
            self.project = Project()
            self.preview.project = self.project
            self.refresh_dashboard()

    def download_template(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="DailyTotalsTemplate.xlsx")
        if filepath:
            create_template(filepath)
            messagebox.showinfo("Success", f"Template saved to {filepath}")

    def import_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            try:
                data = read_input_excel(filepath)
                self.project.load_excel_data(data)
                self.refresh_dashboard()
                messagebox.showinfo("Success", f"Imported {len(data)} days.")
            except Exception as e:
                messagebox.showerror("Error", f"Import failed: {e}")

    def generate_all(self):
        if not self.project.days:
            messagebox.showwarning("Warning", "No data to generate. Import Excel first.")
            return
            
        try:
            self.project.generate_all()
            self.refresh_dashboard()
            if self.project.last_generation_errors:
                messagebox.showwarning("Generation Complete with Errors", f"Generated with {len(self.project.last_generation_errors)} errors. Opening log.")
                self.view_log()
            else:
                messagebox.showinfo("Success", "Generation completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Generation failed completely: {e}")
            
    def view_log(self):
        if not hasattr(self.project, 'last_generation_errors') or not self.project.last_generation_errors:
            messagebox.showinfo("Log", "No errors in the last generation run.")
            return
        dialog = ErrorLogDialog(self, self.project.last_generation_errors)
        self.wait_window(dialog)

    def undo(self):
        if self.project.undo():
            self.refresh_dashboard()

    def redo(self):
        if self.project.redo():
            self.refresh_dashboard()

    def save_project(self):
        if self.project.file_path:
            self.project.save_project(self.project.file_path)
            messagebox.showinfo("Saved", "Project saved successfully.")
        else:
            self.save_project_as()

    def save_project_as(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".rprs", filetypes=[("RPRS Project", "*.rprs")])
        if filepath:
            self.project.save_project(filepath)
            messagebox.showinfo("Saved", f"Project saved to {filepath}")

    def open_project(self):
        filepath = filedialog.askopenfilename(filetypes=[("RPRS Project", "*.rprs")])
        if filepath:
            try:
                self.project = Project.load_project(filepath)
                self.preview.project = self.project
                self.refresh_dashboard()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load project: {e}")

    def export_tally_xml(self):
        dialog = ExportPreviewDialog(self, self.project)
        self.wait_window(dialog)

    def export_excel(self):
        try:
            data = self.project.get_export_data()
            if not data:
                messagebox.showwarning("Warning", "No data to export.")
                return
            filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="ReconstructedPurchases.xlsx")
            if filepath:
                write_output_excel(data, filepath)
                messagebox.showinfo("Success", f"Exported to {filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            
    def open_settings(self):
        dialog = SettingsDialog(self, self.project)
        self.wait_window(dialog)
        
    def convert_final_excel_to_xml(self):
        filepath = filedialog.askopenfilename(title="Select Finalized Excel", filetypes=[("Excel files", "*.xlsx")])
        if not filepath:
            return
            
        try:
            from core.excel_handler import read_final_excel
            from core.tally_xml import TallyXMLEngine
            from datetime import datetime
            
            export_days = read_final_excel(filepath)
            if not export_days:
                messagebox.showerror("Error", "No valid data found in the Excel file.")
                return
                
            batch_id = f"REIMP{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            engine = TallyXMLEngine(self.project.config)
            envelope = engine.generate_xml_tree(export_days, batch_id)
            is_valid, msg = engine.validate_xml_structure(envelope)
            
            if not is_valid:
                messagebox.showerror("Validation Failed", f"The imported data failed XML validation:\n{msg}\n\nPlease check Voucher Numbers or missing entries.")
                return
                
            save_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")], initialfile=f"TallyExport_{batch_id}.xml")
            if save_path:
                engine.generate_and_save(export_days, batch_id, save_path)
                messagebox.showinfo("Success", f"Successfully converted Excel to XML and saved to:\n{save_path}")
                
        except Exception as e:
            messagebox.showerror("Conversion Error", f"Failed to convert Excel to XML:\n{e}")
