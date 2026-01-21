import os
import psutil
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import win32com.client
import threading

class TransparentExcelProfiler:
    def __init__(self, root):
        self.root = root
        self.root.title("Deep Excel Auditor & Live Monitor")
        self.root.geometry("700x650")
        
        # UI Elements
        self.label = tk.Label(root, text="Excel Real-Time Performance Monitor", font=("Arial", 14, "bold"))
        self.label.pack(pady=10)

        self.btn = tk.Button(root, text="🚀 START DEEP SCAN", command=self.start_thread, 
                             bg="#0078d7", fg="white", font=("Arial", 10, "bold"), padx=20, pady=10)
        self.btn.pack(pady=5)

        self.status_label = tk.Label(root, text="Status: Idle", fg="blue", font=("Arial", 10, "italic"))
        self.status_label.pack()

        self.output_area = scrolledtext.ScrolledText(root, width=85, height=25, bg="#1e1e1e", fg="#d4d4d4")
        self.output_area.pack(pady=10, padx=10)

    def log(self, text):
        self.output_area.insert(tk.END, text + "\n")
        self.output_area.see(tk.END)

    def update_status(self, msg):
        self.status_label.config(text=f"Status: {msg}")
        self.root.update_idletasks()

    def start_thread(self):
        # Running in a thread prevents the GUI from "Freezing" while scanning
        thread = threading.Thread(target=self.run_diagnostic)
        thread.start()

    def run_diagnostic(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xlsb")])
        if not file_path: return
            
        self.output_area.delete('1.0', tk.END)
        self.log(">>> INITIALIZING BACKGROUND PROCESSES...")
        
        # 1. System Scan
        self.update_status("Scanning Hardware...")
        self.log(f"[STEP 1] Fetching PC Specs...")
        ram = psutil.virtual_memory()
        self.log(f"   - Hardware Found: {platform.processor()}")
        self.log(f"   - RAM Stats: {ram.percent}% utilized of {round(ram.total/(1024**3),1)}GB")

        # 2. Excel COM Connection
        try:
            self.update_status("Opening Excel Instance...")
            self.log(f"[STEP 2] Creating Background Excel.Application Object...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            self.update_status(f"Loading {os.path.basename(file_path)}...")
            self.log(f"[STEP 3] Opening Workbook: {file_path}")
            wb = excel.Workbooks.Open(file_path)

            # 3. Data Inventory (Rows/Cols/Data)
            self.update_status("Auditing Sheets...")
            self.log(f"\n[STEP 4] DATA INVENTORY PER SHEET:")
            self.log("-" * 60)
            self.log(f"{'Sheet Name':<20} | {'Rows':<8} | {'Cols':<8} | {'Data Points'}")
            self.log("-" * 60)

            total_data_points = 0
            for sheet in wb.Sheets:
                rows = sheet.UsedRange.Rows.Count
                cols = sheet.UsedRange.Columns.Count
                # Count only cells that are NOT empty
                data_count = excel.WorksheetFunction.CountA(sheet.Cells)
                total_data_points += data_count
                self.log(f"{sheet.Name[:18]:<20} | {rows:<8} | {cols:<8} | {data_count}")

            # 4. Logic & Performance Scan
            self.update_status("Scanning Logic...")
            self.log(f"\n[STEP 5] SCANNING LOGIC & CONNECTIONS...")
            
            issues = []
            
            # Check Connections
            self.log(f"   - Checking Power Query & Connections...")
            if wb.Connections.Count > 0:
                issues.append((f"Found {wb.Connections.Count} Connections", "Switch to 'Manual Refresh' to stop startup lag."))

            # Check VBA
            self.log(f"   - Inspecting VBA Project Modules...")
            if wb.HasVBProject:
                issues.append(("VBA Macros Present", "If Excel freezes during typing, disable 'Events' in your VBA code."))

            # 5. Summary & Solutions
            self.log(f"\n" + "="*20 + " FINAL DIAGNOSIS " + "="*20)
            if not issues:
                self.log("✅ RESULT: File is structurally healthy.")
            for issue, sol in issues:
                self.log(f"❌ ISSUE: {issue}")
                self.log(f"💡 SOLUTION: {sol}\n")

            self.log(f"Final Report: Total Data Points analyzed: {total_data_points}")
            
            wb.Close(False)
            excel.Quit()
            self.update_status("Scan Complete")

        except Exception as e:
            self.log(f"\n⚠️ BACKGROUND ERROR: {str(e)}")
            self.update_status("Error Occurred")

if __name__ == "__main__":
    root = tk.Tk()
    app = TransparentExcelProfiler(root)
    root.mainloop()