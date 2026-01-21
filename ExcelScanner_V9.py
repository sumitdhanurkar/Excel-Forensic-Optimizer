import os
import psutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import zipfile

# --- THEME & STYLING ---
CLR_BG = "#F8F9FA"
CLR_CARD = "#FFFFFF"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_TXT = "#323130"

class FinalExcelSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence Suite v9.0")
        self.root.geometry("1050x850")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=30)
        self.show_home()

    def kill_ghost_excel(self):
        """Prevents 'Visible' property errors by clearing stuck background processes."""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Intelligence Suite", font=("Segoe UI", 32, "bold"), bg=CLR_BG, fg=CLR_TXT).pack(pady=(120, 10))
        tk.Label(self.container, text="Advanced Diagnostic & Multi-Task Repair Engine", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT & AUTO-SCAN", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=50, pady=20, cursor="hand2")
        btn.pack(pady=50)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Deep Diagnostic in Progress...", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=700, mode="determinate")
        self.pb.pack(pady=10)
        self.status = tk.Label(self.container, text="Waking Excel Engine...", bg=CLR_BG, font=("Consolas", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_scan, daemon=True).start()

    def perform_scan(self):
        try:
            self.kill_ghost_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            self.status.config(text="Opening Workbook...")
            self.pb['value'] = 25
            wb = excel.Workbooks.Open(self.file_path)
            
            # 1. Audit Sheets & Data
            self.status.config(text="Mapping Data Density...")
            sheets_info = []
            problems = []
            for sh in wb.Sheets:
                rows = sh.UsedRange.Rows.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                sheets_info.append((sh.Name, rows, data))
                if rows > data + 2000:
                    problems.append(("Ghost Rows", f"'{sh.Name}' has {rows-data} bloated empty rows.", "GHOST"))

            # 2. File Format Check
            self.pb['value'] = 60
            if not self.file_path.lower().endswith(".xlsb"):
                problems.append(("Legacy Format", "Using XML structure. Binary is 2x faster.", "BINARY"))

            # 3. Calculation Benchmark
            self.status.config(text="Benchmarking Formulas...")
            start = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start
            if calc_time > 1.0:
                problems.append(("Slow Logic", f"Full calculation takes {calc_time:.2f}s.", "CALC"))

            self.audit_data = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "sheets": sheets_info,
                "problems": problems,
                "calc_speed": f"{calc_time:.2f}s"
            }
            
            wb.Close(False)
            excel.Quit()
            self.pb['value'] = 100
            self.root.after(500, self.display_report)
        except Exception as e:
            messagebox.showerror("Scan Failed", f"Excel Engine Error: {e}")
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- USAGE STATS ---
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text="Diagnostic Report", font=("Segoe UI", 22, "bold"), bg=CLR_BG, fg=CLR_EXCEL).pack(side="left")
        tk.Label(header, text=f"Initial Size: {self.audit_data['size']:.2f} MB", bg=CLR_BG, font=("Segoe UI", 10)).pack(side="right", pady=10)

        # Issues List
        tk.Label(self.container, text="Detected Issues", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        list_frame = tk.Frame(self.container, bg=CLR_BG)
        list_frame.pack(fill="x", pady=10)

        if not self.audit_data['problems']:
            tk.Label(list_frame, text="✅ No critical issues found.", fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w")
        else:
            for title, desc, task_id in self.audit_data['problems']:
                f = tk.Frame(list_frame, bg="white", pady=12, padx=20, highlightbackground="#EDEBE9", highlightthickness=1)
                f.pack(fill="x", pady=4)
                tk.Label(f, text=f"⚠️ {title}:", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(side="left")
                tk.Label(f, text=desc, font=("Segoe UI", 10), bg="white").pack(side="left", padx=15)
                # SINGLE TASK BUTTON
                tk.Button(f, text="Fix Only This", command=lambda t=task_id: self.run_repair(t), 
                          bg=CLR_BLUE, fg="white", font=("Segoe UI", 8, "bold"), padx=15).pack(side="right")

        # --- MASTER CONTROL ---
        tk.Label(self.container, text="Recommended Solution", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(25, 5))
        sol_box = tk.Frame(self.container, bg="white", padx=25, pady=25, highlightbackground="#D2D0CE", highlightthickness=1)
        sol_box.pack(fill="x")

        tk.Label(sol_box, text="Apply All Optimizations", font=("Segoe UI", 11, "bold"), bg="white").pack(anchor="w")
        tk.Label(sol_box, text="This will clean all sheets, convert to Binary, and rebuild the calc chain.", 
                 font=("Segoe UI", 9), bg="white", fg="#605E5C").pack(anchor="w", pady=(0, 20))

        # MASTER TASK BUTTON
        tk.Button(sol_box, text="⚡ EXECUTE MASTER PERFORMANCE OVERHAUL", command=lambda: self.run_repair("ALL"), 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", pady=15).pack(fill="x")

    def run_repair(self, mode):
        rep_win = tk.Toplevel(self.root)
        rep_win.title("Repairing")
        rep_win.geometry("400x250")
        rep_win.configure(bg="white")
        lbl = tk.Label(rep_win, text="Executing Repairs...", font=("Segoe UI", 11), bg="white", pady=50)
        lbl.pack()
        
        def engine():
            try:
                self.kill_ghost_excel()
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
                    save_path = base + "_v9_FIXED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                if mode == "CALC" or mode == "ALL":
                    excel.CalculateFullRebuild()

                wb.Save()
                new_size = os.path.getsize(save_path)/(1024*1024)
                wb.Close()
                excel.Quit()
                
                lbl.config(text=f"✅ Done!\nNew Size: {new_size:.2f} MB", fg=CLR_EXCEL)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Repair Failed: {e}", fg=CLR_ERR)

        threading.Thread(target=engine, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = FinalExcelSuite(root)
    root.mainloop()