import tkinter as tk

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        
    def enter(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True) # Remove window decorations
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"), padding=(5, 5))
        label.pack(ipadx=1)
        
    def leave(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def create_tooltip(parent, text, use_grid=False, **layout_kwargs):
    """Helper to create a small [?] button with a tooltip."""
    lbl = tk.Label(parent, text="[?]", fg="blue", cursor="hand2", font=("Arial", 8))
    if use_grid:
        lbl.grid(**layout_kwargs)
    else:
        lbl.pack(**layout_kwargs)
    ToolTip(lbl, text)
    return lbl
