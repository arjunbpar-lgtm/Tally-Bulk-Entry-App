import tkinter as tk
from tkinter import ttk

def apply_modern_theme(root):
    # Base configuration for standard Tk windows
    root.configure(bg="#ffffff")
    
    style = ttk.Style()
    # Clam allows full customization of backgrounds and colors
    style.theme_use('clam')
    
    bg_color = "#ffffff"
    fg_color = "#212529"
    
    # Base Widget Styling
    style.configure("TFrame", background=bg_color)
    style.configure("TLabel", background=bg_color, foreground=fg_color, font=("Arial", 10))
    style.configure("TCheckbutton", background=bg_color, foreground=fg_color, font=("Arial", 10))
    style.configure("TLabelframe", background=bg_color, foreground=fg_color, font=("Arial", 10, "bold"), bordercolor="#dee2e6")
    style.configure("TLabelframe.Label", background=bg_color, foreground=fg_color, font=("Arial", 10, "bold"))
    
    # Input Styling
    style.configure("TEntry", fieldbackground="#f8f9fa", bordercolor="#ced4da", foreground=fg_color, padding=4)
    style.configure("TCombobox", fieldbackground="#f8f9fa", background="#ffffff")
    
    # --- Buttons ---
    
    # Primary Button (Vibrant Blue - Default)
    style.configure("TButton", 
                    font=("Arial", 10, "bold"), 
                    background="#0d6efd", 
                    foreground="white", 
                    bordercolor="#0d6efd",
                    lightcolor="#0d6efd",
                    darkcolor="#0d6efd",
                    padding=(10, 5))
    style.map("TButton", 
              background=[("active", "#0b5ed7"), ("disabled", "#e9ecef")],
              foreground=[("active", "white"), ("disabled", "#6c757d")])
              
    # Success Button (Green for generation/export)
    style.configure("Success.TButton", 
                    background="#198754", 
                    bordercolor="#198754",
                    lightcolor="#198754",
                    darkcolor="#198754")
    style.map("Success.TButton", background=[("active", "#157347")])
    
    # Warning Button (Orange for settings/tools)
    style.configure("Warning.TButton", 
                    background="#fd7e14", 
                    bordercolor="#fd7e14",
                    lightcolor="#fd7e14",
                    darkcolor="#fd7e14")
    style.map("Warning.TButton", background=[("active", "#e37112")])
    
    # Danger Button (Red for delete/undo/error)
    style.configure("Danger.TButton", 
                    background="#dc3545", 
                    bordercolor="#dc3545",
                    lightcolor="#dc3545",
                    darkcolor="#dc3545")
    style.map("Danger.TButton", background=[("active", "#c82333")])
    
    # Info Button (Teal for logs/info)
    style.configure("Info.TButton", 
                    background="#0dcaf0", 
                    foreground="#000000",
                    bordercolor="#0dcaf0",
                    lightcolor="#0dcaf0",
                    darkcolor="#0dcaf0")
    style.map("Info.TButton", background=[("active", "#31d2f2")])

def configure_window(window):
    """Utility to quickly configure Toplevel windows with the white background."""
    window.configure(bg="#ffffff")
