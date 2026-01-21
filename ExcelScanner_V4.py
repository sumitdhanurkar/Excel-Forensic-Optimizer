import os
import psutil
import platform
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import win32com.client
import threading

# --- MODERN STYLING ---
COLOR_BG = "#f8f9fa"
COLOR_PRIMARY = "#2ecc71"
COLOR_ACCENT = "#34495e"
COLOR_TEXT = "#2c3e50"
COLOR_CONSOLE = "#1e1e1e"

class AdvancedExcelAuditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Performance Engine")
        self.root.geometry("800x600")
        self.root.configure(bg=COLOR_BG)
        
        # Main Container
        self.main_container = tk.Frame(self.root, bg=COLOR_BG)
        self.main_container.pack(fill="both", expand=True, padx=40, pady=20)

        self.show_home_screen()

    def clear_screen(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_home_screen(self):
        self.clear_screen()
        tk.Label(self.main_container, text="Excel Performance Engine", font=("Segoe UI", 24, "bold"), bg=COLOR_BG, fg=COLOR_ACCENT).pack(pady=(50, 10))
        tk.Label(self.main_container, text="Hardware-Aware File Diagnostic Tool", font=("Segoe UI", 12), bg=COLOR_BG, fg="#7f8c8d").pack(pady=5)
        
        self.start_btn = tk.Button(self.main_container, text="SELECT WORKBOOK TO SCAN", command=self.initiate_scan, 
                                   bg=COLOR_PRIMARY, fg="white", font=("Segoe UI", 12, "bold"), 
                                   relief="flat", padx=30, pady=15, cursor="hand2")
        self.start_btn.pack(pady=40)

    def show_loading_screen(self):
        self.clear_screen()
        self.status_title = tk.Label(self.main_container, text="Diagnostic in Progress...", font=("Segoe UI", 16, "bold"), bg=COLOR_BG, fg=COLOR_ACCENT)
        self.status_title.pack(pady=10)
        
        # Console for background tasks
        self.console = tk.Text(self.main_container, bg=COLOR_CONSOLE, fg="#00ff00", font=("Consolas", 10), height=20, borderwidth=0)
        self.console.pack(fill="both", expand=True, padx=10, pady=10)
        self.log("Initializing core engine...")

    def log(self, message):
        self.console.insert(tk.END, f" > {message}\n")
        self.console.see(tk.END)
        self.root.update()

    def initiate_scan(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xlsb")])
        if not file_path: return
        
        self.show_loading_screen()
        # Run in thread to keep UI responsive
        threading.Thread(target=self.run_logic, args=(file_path,), daemon=True).start()

    def run_logic(self, file_path):
        results = {"system": {}, "sheets": [], "issues": []}
        
        try:
            # 1. Hardware Scan
            self.log("Scanning System Resources...")
            results['system'] = {
                "cpu": platform.processor(),
                "ram_total": f"{round(psutil.virtual_memory().total/(1024**3),1)}GB",
                "ram_load": f"{psutil.virtual_memory().percent}%"
            }

            # 2. Excel Deep Audit
            self.log("Connecting to Excel Background Process...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(file_path)
            
            self.log(f"Auditing workbook: {os.path.basename(file_path)}")
            
            for sheet in wb.Sheets:
                self.log(f"Scanning sheet: {sheet.Name}...")
                rows = sheet.UsedRange.Rows.Count
                cols = sheet.UsedRange.Columns.Count
                data_points = excel.WorksheetFunction.CountA(sheet.Cells)
                results['sheets'].append((sheet.Name, rows, cols, data_points))
                
                # Logic for "Ghost Rows" issue
                if rows > (data_points + 5000):
                    results['issues'].append((f"Ghost Rows in '{sheet.Name}'", "Excel is tracking thousands of empty rows. Delete blank rows to shrink file."))

            # VBA & Connections
            self.log("Checking Macros and Data Connections...")
            if wb.HasVBProject: results['issues'].append(("VBA Macros Present", "Heavy macros detected. Disable ScreenUpdating to speed up execution."))
            if wb.Connections.Count > 0: results['issues'].append(("External Connections", "Power Query or SQL detected. Turn off Background Refresh."))

            wb.Close(False)
            excel.Quit()
            
            # Transition to Results
            self.log("Generating final report...")
            self.root.after(1000, lambda: self.show_result_screen(results, os.path.basename(file_path)))

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.show_home_screen()

    def show_result_screen(self, results, filename):
        self.clear_screen()
        
        # Header
        header = tk.Frame(self.main_container, bg=COLOR_BG)
        header.pack(fill="x")
        tk.Label(header, text="Scan Results", font=("Segoe UI", 18, "bold"), bg=COLOR_BG, fg=COLOR_PRIMARY).pack(side="left")
        tk.Button(header, text="New Scan", command=self.show_home_screen, bg=COLOR_ACCENT, fg="white", relief="flat").pack(side="right")
        
        tk.Label(self.main_container, text=f"File: {filename}", font=("Segoe UI", 10, "italic"), bg=COLOR_BG).pack(anchor="w", pady=5)

        # 1. System Info Card
        sys_frame = tk.LabelFrame(self.main_container, text=" System Performance ", font=("Segoe UI", 10, "bold"), bg=COLOR_BG, padx=10, pady=10)
        sys_frame.pack(fill="x", pady=10)
        tk.Label(sys_frame, text=f"Hardware: {results['system']['cpu']} | RAM Load: {results['system']['ram_load']}", bg=COLOR_BG).pack(anchor="w")

        # 2. Data Table
        tk.Label(self.main_container, text="Worksheet Inventory", font=("Segoe UI", 12, "bold"), bg=COLOR_BG).pack(anchor="w", pady=(10, 0))
        table_frame = tk.Frame(self.main_container, bg="white")
        table_frame.pack(fill="both", expand=True, pady=5)
        
        cols = ("Sheet Name", "Total Rows", "Total Cols", "Actual Data")
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=5)
        for col in cols: tree.heading(col, text=col)
        for item in results['sheets']: tree.insert("", "end", values=item)
        tree.pack(fill="both", expand=True)

        # 3. Optimization Issues (The "Detailed Results")
        tk.Label(self.main_container, text="Recommended Fixes", font=("Segoe UI", 12, "bold"), bg=COLOR_BG, fg="#e67e22").pack(anchor="w", pady=(15, 0))
        issue_box = scrolledtext.ScrolledText(self.main_container, height=6, font=("Segoe UI", 10), bg="#fff9f2", borderwidth=0)
        issue_box.pack(fill="x", pady=5)
        
        if not results['issues']:
            issue_box.insert(tk.END, "✅ No major bottlenecks found. Your file is optimized!")
        for issue, solution in results['issues']:
            issue_box.insert(tk.END, f"• {issue.upper()}\n  FIX: {solution}\n\n")
        issue_box.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    # Apply a modern theme to scrollbars and treeview
    style = ttk.Style()
    style.theme_use("clam")
    app = AdvancedExcelAuditor(root)
    root.mainloop()