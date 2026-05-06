import tkinter as tk
from tkinter import ttk, simpledialog, messagebox

class PreviewScreen(ttk.Frame):
    def __init__(self, parent_screen, project):
        super().__init__(parent_screen)
        self.parent_screen = parent_screen
        self.project = project
        
        self.setup_ui()
        
    def setup_ui(self):
        # Create Treeview
        columns = ("Date", "Total Amount", "Entries Count", "Avg Value", "Confidence", "Locked", "Notes")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        self.tree.heading("Date", text="Date")
        self.tree.heading("Total Amount", text="Target Total")
        self.tree.heading("Entries Count", text="Entries")
        self.tree.heading("Avg Value", text="Avg Entry")
        self.tree.heading("Confidence", text="Confidence %")
        self.tree.heading("Locked", text="Locked")
        self.tree.heading("Notes", text="Notes")
        
        # Column widths
        self.tree.column("Date", width=100, anchor=tk.CENTER)
        self.tree.column("Total Amount", width=120, anchor=tk.E)
        self.tree.column("Entries Count", width=80, anchor=tk.CENTER)
        self.tree.column("Avg Value", width=100, anchor=tk.E)
        self.tree.column("Confidence", width=100, anchor=tk.CENTER)
        self.tree.column("Locked", width=60, anchor=tk.CENTER)
        self.tree.column("Notes", width=250, anchor=tk.W)
        
        # Scrollbars
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Tags for styling
        self.tree.tag_configure('locked', background='#e0e0e0', foreground='#555555')
        self.tree.tag_configure('low_conf', background='#ffcccc')
        self.tree.tag_configure('failed', background='#ff6666', foreground='white')
        
        # Context Menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Regenerate Day", command=self.regenerate_selected)
        self.menu.add_command(label="Toggle Freeze/Lock", command=self.toggle_lock_selected)
        self.menu.add_command(label="Edit Day Note...", command=self.edit_note_selected)
        
        self.tree.bind("<Button-3>", self.show_context_menu)
        # Bind double-click to toggle lock just for speed
        self.tree.bind("<Double-1>", lambda e: self.toggle_lock_selected())
        
    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu.post(event.x_root, event.y_root)
            
    def refresh_grid(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Sort keys
        from datetime import datetime
        sorted_keys = sorted(self.project.days.keys(), key=lambda d: datetime.strptime(d, "%d-%m-%Y"))
        
        for date_str in sorted_keys:
            day_obj = self.project.days[date_str]
            
            entries_count = len(day_obj.entries)
            conf = day_obj.confidence
            locked_str = "🔒 Yes" if day_obj.locked else ""
            
            tags = ()
            if entries_count == 0 and day_obj.day_notes.startswith("Error:"):
                tags = ('failed',)
            elif day_obj.locked:
                tags = ('locked',)
            elif conf < 50:
                tags = ('low_conf',)
                
            self.tree.insert("", tk.END, iid=date_str, values=(
                date_str,
                day_obj.original_total,
                entries_count,
                f"{day_obj.average_val:.2f}",
                f"{conf}%",
                locked_str,
                day_obj.day_notes
            ), tags=tags)
            
    def regenerate_selected(self):
        selected = self.tree.selection()
        if not selected: return
        date_str = selected[0]
        
        day_obj = self.project.days[date_str]
        if day_obj.locked:
            messagebox.showwarning("Locked", "Cannot regenerate a locked day.")
            return
            
        self.project.regenerate_day(date_str)
        self.parent_screen.refresh_dashboard()
        
    def toggle_lock_selected(self):
        selected = self.tree.selection()
        if not selected: return
        date_str = selected[0]
        self.project.toggle_lock(date_str)
        self.refresh_grid()
        
    def edit_note_selected(self):
        selected = self.tree.selection()
        if not selected: return
        date_str = selected[0]
        
        current_note = self.project.days[date_str].day_notes
        new_note = simpledialog.askstring("Edit Note", f"Enter note for {date_str}:", initialvalue=current_note)
        if new_note is not None:
            self.project.update_day_note(date_str, new_note)
            self.refresh_grid()
