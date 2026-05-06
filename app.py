import tkinter as tk
from tkinter import ttk
from core.theme import apply_modern_theme, configure_window
from apps.rubber_reconstruction.ui.home_screen import HomeScreen
import webbrowser

# We will import BulkXMLScreen once we create it.
# from apps.bulk_xml_converter.app_window import BulkXMLScreen

class LauncherScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setup_ui()
        
    def setup_ui(self):
        self.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        title = ttk.Label(self, text="VoucherGen Suite", font=("Arial", 24, "bold"))
        title.pack(pady=(20, 5))
        
        subtitle = ttk.Label(self, text="Central Accounting Dashboard", font=("Arial", 12), foreground="gray")
        subtitle.pack(pady=(0, 30))
        
        # Modules Frame
        modules_frame = ttk.Frame(self)
        modules_frame.pack(fill=tk.BOTH, expand=True)
        
        # App 1 Button
        app1_frame = ttk.LabelFrame(modules_frame, text="Module 1", padding=15)
        app1_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(app1_frame, text="Excel to Tally Converter (from Totals)", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        ttk.Label(app1_frame, text="Generate dynamic purchase entries from daily raw totals.").pack(anchor=tk.W, pady=(5, 10))
        ttk.Button(app1_frame, text="Launch Application", command=self.launch_rubber_app, width=25, style="Success.TButton").pack(anchor=tk.W)
        
        # App 2 Button
        app2_frame = ttk.LabelFrame(modules_frame, text="Module 2", padding=15)
        app2_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(app2_frame, text="Bulk Tally XML Converter", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        ttk.Label(app2_frame, text="Convert structured Excel data of any voucher type directly into Tally XML.").pack(anchor=tk.W, pady=(5, 10))
        ttk.Button(app2_frame, text="Launch Application", command=self.launch_bulk_xml_app, width=25, style="Success.TButton").pack(anchor=tk.W)
        
        
        # Footer Branding
        footer = tk.Frame(self, bg="#ffffff")
        footer.pack(side="bottom", pady=20, fill=tk.X)
        
        branding_lbl = tk.Label(footer, text="Created with the help of Google Antigravity", font=("Arial", 9, "italic"), fg="gray", bg="#ffffff")
        branding_lbl.pack()
        
        github_lbl = tk.Label(footer, text="GitHub: arjunbpar-lgtm", font=("Arial", 9, "underline"), fg="#0d6efd", cursor="hand2", bg="#ffffff")
        github_lbl.pack()
        github_lbl.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/arjunbpar-lgtm"))

    def launch_rubber_app(self):
        new_window = tk.Toplevel(self.parent)
        new_window.title("Excel to Tally Converter")
        new_window.geometry("1024x768")
        configure_window(new_window)
        
        app = HomeScreen(new_window)
        app.pack(fill=tk.BOTH, expand=True)
        
    def launch_bulk_xml_app(self):
        new_window = tk.Toplevel(self.parent)
        new_window.title("Bulk Tally XML Converter")
        new_window.geometry("600x450")
        configure_window(new_window)
        
        from apps.bulk_xml_converter.app_window import BulkXMLScreen
        app = BulkXMLScreen(new_window)
        app.pack(fill=tk.BOTH, expand=True)

def main():
    root = tk.Tk()
    root.title("VoucherGen Suite - Dashboard")
    root.geometry("600x550")
    
    # Apply the global modern theme
    apply_modern_theme(root)
    
    app = LauncherScreen(root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
