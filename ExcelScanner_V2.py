import os
import psutil
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import win32com.client

class UltimateExcelProfiler:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Doctor: Hardware & File Diagnostic")
        self.root.geometry("600x550")
        self.root.configure(bg="#f0f0f0")
        
        # Header
        self.header = tk.Label(root, text="Excel Performance Optimizer", font=("Segoe UI", 16, "bold"), bg="#f0f0f0", fg="#2c3e50")
        self.header.pack(pady=15)

        # Main Button
        self.btn = tk.Button(root, text="🔍 SCAN EXCEL WORKBOOK", command=self.run_full_diagnostic, 
                             bg="#27ae60", fg="white", font=("Segoe UI", 11, "bold"), padx=30, pady=12, relief="flat")
        self.btn.pack(pady=10)

        # Output Log
        self.output_area = scrolledtext.ScrolledText(root, width=70, height=20, font=("Consolas", 10))
        self.output_area.pack(pady=15, padx=20)

    def log(self, text, color="black"):
        self.output_area.insert(tk.END, text + "\n")
        self.output_area.see(tk.END)

    def run_full_diagnostic(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xlsb")])
        if not file_path: return
            
        self.output_area.delete('1.0', tk.END)
        self.log(f"Starting Scan for: {os.path.basename(file_path)}")
        self.log("="*50)

        # 1. Hardware Check
        self.log("[+] SCANNING HARDWARE...")
        ram = psutil.virtual_memory()
        cpu_cores = psutil.cpu_count()
        self.log(f"  - RAM: {round(ram.total/(1024**3), 1)}GB ({ram.percent}% Used)")
        self.log(f"  - CPU: {cpu_cores} Cores detected.")

        # 2. Deep File Audit via COM
        self.log("\n[+] ANALYZING EXCEL INTERNALS...")
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(file_path)

            problems = []

            # Check for VBA
            if wb.HasVBProject:
                problems.append(("VBA Macros detected", "Macros can trigger on every click. Ensure 'ScreenUpdating' is turned OFF in your code to speed it up."))

            # Check for Power Query / SQL
            if wb.Connections.Count > 0:
                problems.append((f"{wb.Connections.Count} External Connections", "Disable 'Background Refresh' in Data -> Queries & Connections -> Properties to prevent freezing during use."))

            # Check for Object/Shape Bloat (Common cause of lag)
            obj_count = 0
            for sheet in wb.Sheets:
                obj_count += sheet.Shapes.Count
            if obj_count > 50:
                problems.append((f"Found {obj_count} hidden objects", "Excessive shapes/images bloat file size. Press Ctrl+G -> Special -> Objects to find and delete them."))

            # Check for Conditional Formatting Bloat
            # (Heuristic: Check if UsedRange is significantly larger than actual data)
            for sheet in wb.Sheets:
                if sheet.UsedRange.Rows.Count > 5000:
                    problems.append((f"Large Used Range in '{sheet.Name}'", "Excel thinks you have data in thousands of empty cells. Delete all 'empty' rows below your data to reset the file size."))

            # 3. DISPLAY SOLUTIONS
            self.log("\n" + "="*20 + " DIAGNOSIS & SOLUTIONS " + "="*20)
            if not problems:
                self.log("✅ No major performance issues detected in file logic.", "green")
            else:
                for issue, solution in problems:
                    self.log(f"❌ PROBLEM: {issue}")
                    self.log(f"💡 SOLUTION: {solution}\n")

            wb.Close(False)
            excel.Quit()
            
        except Exception as e:
            self.log(f"⚠️ Error: {str(e)}")
            self.log("Ensure the file isn't open or protected by a password.")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateExcelProfiler(root)
    root.mainloop()