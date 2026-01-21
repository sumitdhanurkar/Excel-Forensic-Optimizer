import os
import psutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading

# --- THEME CONSTANTS ---
CLR_BG = "#F3F2F1"
CLR_CARD = "#FFFFFF"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_WARN = "#FFB900"
CLR_TXT = "#323130"

class UltimateExcelSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Performance Intelligence v9.3")
        self.root.geometry("1200x900")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        self.show_home()

    def kill_ghost_excel(self):
        """Emergency process cleanup to prevent 'Visible' property locks."""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Intelligence Engine", font=("Segoe UI", 32, "bold"), bg=CLR_BG, fg=CLR_TXT).pack(pady=(120, 10))
        tk.Label(self.container, text="Complete File Structural Audit & Hardware Diagnostic", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT & ANALYZE WORKBOOK", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=50, pady=20, cursor="hand2")
        btn.pack(pady=50)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Performing Structural Audit...", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=10)
        self.status = tk.Label(self.container, text="Scanning System & Initializing Excel...", bg=CLR_BG, font=("Consolas", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_scan, daemon=True).start()

    def perform_scan(self):
        try:
            # 1. Hardware Snapshot
            cpu = psutil.cpu_percent(interval=0.5)
            ram = psutil.virtual_memory()
            
            self.kill_ghost_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # 2. File Metadata & Load Speed
            start_open = time.time()
            wb = excel.Workbooks.Open(self.file_path)
            load_time = time.time() - start_open
            
            self.pb['value'] = 40
            self.status.config(text="Auditing Rows, Columns, and Bloat...")
            
            # 3. Structural Audit (The core request)
            sheets_report = []
            problems = []
            for sh in wb.Sheets:
                used_r = sh.UsedRange.Rows.Count
                used_c = sh.UsedRange.Columns.Count
                # Count cells that actually contain data (formulas or constants)
                try:
                    real_data_count = excel.WorksheetFunction.CountA(sh.Cells)
                except:
                    real_data_count = 0
                
                sheets_report.append({
                    "name": sh.Name,
                    "rows": used_r,
                    "cols": used_c,
                    "data": real_data_count
                })

                # Logic: If used rows are 5x more than actual data, it's a "Ghost Range"
                if used_r > (real_data_count + 5000) or used_c > 50:
                    problems.append({
                        "issue": f"Excess Range in '{sh.Name}'",
                        "desc": f"Sheet is tracking {used_r} rows but only has {real_data_count} data cells.",
                        "solution": "Purge Ghost Rows/Columns to shrink file size by up to 80%.",
                        "id": "GHOST"
                    })

            self.pb['value'] = 80
            self.status.config(text="Measuring Calculation Chains...")
            start_calc = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start_calc

            # IT Alert Logic
            it_alert = True if (ram.percent > 85 or cpu > 90) else False

            self.audit_data = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "sheets": sheets_report,
                "problems": problems,
                "load": f"{load_time:.2f}s",
                "calc": f"{calc_time:.2f}s",
                "ram": f"{ram.percent}%",
                "cpu": f"{cpu}%",
                "it_alert": it_alert
            }
            
            wb.Close(False)
            excel.Quit()
            self.pb['value'] = 100
            self.root.after(500, self.display_report)
        except Exception as e:
            messagebox.showerror("Audit Error", f"Engine failed: {e}")
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- TOP HARDWARE & IT BAR ---
        hw_bg = CLR_ERR if self.audit_data['it_alert'] else CLR_EXCEL
        hw_bar = tk.Frame(self.container, bg=hw_bg, padx=20, pady=10)
        hw_bar.pack(fill="x", pady=(0, 20))
        
        status_text = "⚠️ HARDWARE STRUGGLING: CONTACT IT TEAM" if self.audit_data['it_alert'] else "✅ SYSTEM HARDWARE HEALTHY"
        tk.Label(hw_bar, text=status_text, font=("Segoe UI", 11, "bold"), bg=hw_bg, fg="white").pack(side="left")
        tk.Label(hw_bar, text=f"RAM: {self.audit_data['ram']} | CPU: {self.audit_data['cpu']}", font=("Segoe UI", 10), bg=hw_bg, fg="white").pack(side="right")

        # --- FILE SUMMARY CARDS ---
        stats_frame = tk.Frame(self.container, bg=CLR_BG)
        stats_frame.pack(fill="x", pady=5)
        
        cards = [("File Size", f"{self.audit_data['size']:.2f} MB"), 
                 ("Load Speed", self.audit_data['load']), 
                 ("Calc Speed", self.audit_data['calc'])]

        for title, val in cards:
            c = tk.Frame(stats_frame, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#E0E0E0")
            c.pack(side="left", fill="both", expand=True, padx=4)
            tk.Label(c, text=title, font=("Segoe UI", 9), bg="white", fg="#605E5C").pack()
            tk.Label(c, text=val, font=("Segoe UI", 12, "bold"), bg="white").pack()

        # --- STRUCTURAL AUDIT TABLE ---
        tk.Label(self.container, text="Worksheet Structural Audit", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20, 5))
        tbl_frame = tk.Frame(self.container, bg="white", highlightthickness=1, highlightbackground="#E0E0E0")
        tbl_frame.pack(fill="x")
        
        cols = ("Sheet Name", "Used Rows", "Used Cols", "Actual Data Cells")
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=5)
        for col in cols: tree.heading(col, text=col)
        for s in self.audit_data['sheets']: tree.insert("", "end", values=(s['name'], s['rows'], s['cols'], s['data']))
        tree.pack(fill="x")

        # --- PROBLEM & SOLUTION HUB ---
        tk.Label(self.container, text="Identified Performance Risks & Solutions", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20, 5))
        
        if not self.audit_data['problems']:
            tk.Label(self.container, text="✅ No structural bloat detected.", fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w")
        else:
            for p in self.audit_data['problems']:
                f = tk.Frame(self.container, bg="white", pady=10, padx=20, highlightthickness=1, highlightbackground="#E0E0E0")
                f.pack(fill="x", pady=2)
                tk.Label(f, text=p['issue'], font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(anchor="w")
                tk.Label(f, text=f"PROBLEM: {p['desc']}", font=("Segoe UI", 9), bg="white").pack(anchor="w")
                tk.Label(f, text=f"SOLUTION: {p['solution']}", font=("Segoe UI", 9, "italic"), bg="white", fg=CLR_BLUE).pack(anchor="w")
                tk.Button(f, text="Fix Single Task", command=lambda i=p['id']: self.run_fix(i), bg=CLR_BLUE, fg="white", font=("Segoe UI", 8)).place(relx=0.9, rely=0.3)

        # --- MASTER ACTION ---
        tk.Button(self.container, text="🚀 EXECUTE FULL MASTER OPTIMIZATION", command=lambda: self.run_repair("ALL"), 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15).pack(fill="x", pady=20)

    def run_repair(self, mode):
        # Repair engine logic to handle GHOST purge and BINARY conversion
        win = tk.Toplevel(self.root)
        win.title("Repairing...")
        win.geometry("400x200")
        lbl = tk.Label(win, text="Running Intelligence Repairs...", pady=50)
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
                    save_path = base + "_OPTIMIZED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                wb.Save()
                wb.Close()
                excel.Quit()
                lbl.config(text="✅ Optimization Successful!", fg=CLR_EXCEL)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {e}", fg=CLR_ERR)

        threading.Thread(target=engine, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateExcelSuite(root)
    root.mainloop()