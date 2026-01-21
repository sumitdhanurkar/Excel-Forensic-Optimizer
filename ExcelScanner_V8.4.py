import os
import psutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import zipfile

# --- THEME ---
CLR_BG = "#F3F2F1"
CLR_CARD = "#FFFFFF"
CLR_EXCEL = "#107C41"
CLR_ACCENT = "#0078D4"
CLR_ERR = "#D13438"

class SmartOptimizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence Suite v8.4")
        self.root.geometry("1100x900")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_results = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=50, pady=30)
        self.show_home()

    def kill_excel(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Intelligence Suite", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(100, 10))
        tk.Label(self.container, text="Select a workbook to begin automated diagnostic scan", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT & AUTO-SCAN", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 11, "bold"), relief="flat", padx=40, pady=15)
        btn.pack(pady=40)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        # Loading UI
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Performing Deep Diagnostic Scan...", font=("Segoe UI", 16, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=600, mode="determinate")
        self.pb.pack(pady=10)
        self.log_lbl = tk.Label(self.container, text="Initializing Engine...", bg=CLR_BG, font=("Consolas", 10))
        self.log_lbl.pack()
        
        threading.Thread(target=self.run_audit_logic, daemon=True).start()

    def run_audit_logic(self):
        try:
            self.kill_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            self.log_lbl.config(text="Opening Workbook...")
            self.pb['value'] = 20
            wb = excel.Workbooks.Open(self.file_path)
            
            # 1. Scan Usage & Sheets
            self.log_lbl.config(text="Analyzing Sheet Density...")
            sheets = []
            problems = []
            total_data = 0
            for sh in wb.Sheets:
                r = sh.UsedRange.Rows.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                sheets.append((sh.Name, r, data))
                total_data += data
                if r > data + 2000:
                    problems.append(("Ghost Rows", f"Sheet '{sh.Name}' has {r-data} empty formatted rows.", "GHOST"))

            # 2. Check Format
            self.pb['value'] = 60
            self.log_lbl.config(text="Checking File Structure...")
            if not self.file_path.lower().endswith(".xlsb"):
                problems.append(("File Bloat", "Workbook is using XML format instead of Binary.", "BINARY"))

            # 3. Calc Speed
            self.log_lbl.config(text="Testing Calculation Speed...")
            start = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start
            if calc_time > 1.0:
                problems.append(("Slow Logic", f"Full calculation takes {calc_time:.2f}s.", "CALC"))

            self.audit_results = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "sheets": sheets,
                "problems": problems,
                "calc": f"{calc_time:.2f}s"
            }
            
            wb.Close(False)
            excel.Quit()
            self.pb['value'] = 100
            self.root.after(500, self.show_diagnostic_report)
        except Exception as e:
            messagebox.showerror("Audit Error", str(e))
            self.show_home()

    def show_diagnostic_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- USAGE SUMMARY ---
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text="Diagnostic Report", font=("Segoe UI", 20, "bold"), bg=CLR_BG, fg=CLR_EXCEL).pack(side="left")
        tk.Label(header, text=f"File Size: {self.audit_results['size']:.2f} MB", bg=CLR_BG).pack(side="right")

        # Problems Section
        tk.Label(self.container, text="Identified Problems", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        prob_frame = tk.Frame(self.container, bg=CLR_BG)
        prob_frame.pack(fill="x", pady=10)

        if not self.audit_results['problems']:
            tk.Label(prob_frame, text="✅ No major issues found! File is healthy.", fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w")
        else:
            for title, desc, task_id in self.audit_results['problems']:
                f = tk.Frame(prob_frame, bg="white", pady=10, padx=15, highlightbackground="#EDEBE9", highlightthickness=1)
                f.pack(fill="x", pady=3)
                tk.Label(f, text=f"⚠️ {title}:", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(side="left")
                tk.Label(f, text=desc, font=("Segoe UI", 10), bg="white").pack(side="left", padx=10)
                tk.Button(f, text="Fix This", command=lambda t=task_id: self.run_fix(t), bg=CLR_ACCENT, fg="white", font=("Segoe UI", 8)).pack(side="right")

        # --- FINAL SOLUTIONS ---
        tk.Label(self.container, text="Solutions", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20, 5))
        sol_frame = tk.Frame(self.container, bg="white", padx=20, pady=20, highlightbackground="#D2D0CE", highlightthickness=1)
        sol_frame.pack(fill="x")

        tk.Button(sol_frame, text="🚀 APPLY ALL RECOMMENDED CHANGES", command=lambda: self.run_fix("ALL"), 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 11, "bold"), relief="flat", pady=15).pack(fill="x")

    def run_fix(self, mode):
        win = tk.Toplevel(self.root)
        win.title("Repairing...")
        win.geometry("400x200")
        lbl = tk.Label(win, text="Executing Repairs...", pady=40)
        lbl.pack()

        def execute():
            try:
                self.kill_excel()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(self.file_path)
                save_path = self.file_path

                if mode in ["GHOST", "ALL"]:
                    for sh in wb.Sheets:
                        last = sh.Cells.Find("*", SearchOrder=1, SearchDirection=2)
                        if last: sh.Rows(f"{last.Row+1}:{sh.Rows.Count}").Delete()
                
                if mode in ["BINARY", "ALL"]:
                    base, _ = os.path.splitext(self.file_path)
                    save_path = base + "_OPTIMIZED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                wb.Save()
                wb.Close()
                excel.Quit()
                lbl.config(text="✅ Repair Successful!")
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {e}", fg="red")

        threading.Thread(target=execute, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = SmartOptimizer(root)
    root.mainloop()