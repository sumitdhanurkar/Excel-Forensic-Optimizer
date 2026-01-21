import os
import psutil
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import win32com.client # Requires pywin32

class ExcelProfiler:
    def __init__(self, root):
        self.root = root
        self.root.title("Ultimate Excel Performance Profiler")
        self.root.geometry("500x400")
        
        # GUI Setup
        self.label = tk.Label(root, text="Excel Hardware & Logic Auditor", font=("Arial", 14, "bold"))
        self.label.pack(pady=10)

        self.btn = tk.Button(root, text="Select Excel File to Scan", command=self.run_full_diagnostic, 
                             bg="#2c3e50", fg="white", font=("Arial", 10, "bold"), padx=20, pady=10)
        self.btn.pack(pady=20)

        self.output_area = scrolledtext.ScrolledText(root, width=60, height=15)
        self.output_area.pack(pady=10)

    def log(self, text):
        self.output_area.insert(tk.END, text + "\n")
        self.output_area.see(tk.END)

    def scan_system(self):
        self.log("--- SYSTEM HARDWARE ---")
        cpu = platform.processor()
        cores = psutil.cpu_count(logical=True)
        ram = psutil.virtual_memory()
        self.log(f"CPU: {cores} Cores | RAM: {round(ram.total/(1024**3), 2)}GB")
        self.log(f"RAM Usage: {ram.percent}%")
        if ram.percent > 75:
            self.log("!! WARNING: High RAM usage may slow down Excel.")
        self.log("-" * 30)

    def audit_excel_logic(self, file_path):
        self.log("--- EXCEL DEEP AUDIT ---")
        try:
            # Connect to Excel Application
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(file_path)
            
            # 1. VBA Macro Scan
            has_vba = wb.HasVBProject
            self.log(f"VBA Macros Present: {'YES' if has_vba else 'No'}")
            if has_vba:
                self.log("-> Suggestion: Check for 'Worksheet_Change' events; they slow down typing.")

            # 2. Power Query / Connections Audit
            conn_count = wb.Connections.Count
            self.log(f"External Connections/Power Query: {conn_count}")
            for i in range(1, conn_count + 1):
                conn_name = wb.Connections.Item(i).Name
                self.log(f"  - Found: {conn_name}")
            
            if conn_count > 0:
                self.log("-> Suggestion: Set Power Queries to 'Background Refresh = False' for stability.")

            # 3. Hidden Names / Broken Links
            links = wb.LinkSources()
            if links:
                self.log(f"External File Links: {len(links)}")
                self.log("-> Suggestion: Broken network links cause 'Freezing' on startup.")

            # 4. Sheet Count & Size
            self.log(f"Total Sheets: {wb.Sheets.Count}")
            
            wb.Close(False)
            excel.Quit()
            
        except Exception as e:
            self.log(f"Error during audit: {str(e)}")
            self.log("Note: Make sure Excel is not showing a pop-up dialog.")

    def run_full_diagnostic(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xlsb")])
        if not file_path:
            return
            
        self.output_area.delete('1.0', tk.END) # Clear previous results
        self.scan_system()
        self.audit_excel_logic(file_path)
        self.log("\n--- SCAN COMPLETE ---")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProfiler(root)
    root.mainloop()