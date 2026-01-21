import os
import psutil
import platform
import time
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import zipfile

# --- THEME & STYLING ---
CLR_BG = "#F3F2F1"
CLR_CARD = "#FFFFFF"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_TXT = "#323130"
CLR_WARN = "#FFB900"

class EnterpriseExcelSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Enterprise Diagnostic v9.2")
        self.root.geometry("1150x900")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        self.show_home()

    def kill_ghost_excel(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Enterprise Diagnostic", font=("Segoe UI", 32, "bold"), bg=CLR_BG, fg=CLR_TXT).pack(pady=(120, 10))
        tk.Label(self.container, text="Professional Hardware & File Intelligence Engine", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 START SYSTEM AUDIT", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=50, pady=20, cursor="hand2")
        btn.pack(pady=50)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Scanning System & Workbook...", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=10)
        self.status = tk.Label(self.container, text="Initializing Deep Scan...", bg=CLR_BG, font=("Consolas", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_scan, daemon=True).start()

    def perform_scan(self):
        try:
            # --- 1. HARDWARE CHECK ---
            self.status.config(text="Scanning CPU/RAM/Disk...")
            cpu = psutil.cpu_percent(interval=0.5)
            ram = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            self.kill_ghost_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # --- 2. FILE METADATA ---
            self.pb['value'] = 30
            self.status.config(text="Reading File Properties...")
            start_open = time.time()
            wb = excel.Workbooks.Open(self.file_path)
            load_time = time.time() - start_open
            
            props = {
                "Author": wb.BuiltinDocumentProperties("Author").Value,
                "Created": str(wb.BuiltinDocumentProperties("Creation Date").Value)[:19],
                "LastSaved": str(wb.BuiltinDocumentProperties("Last Author").Value)
            }

            # --- 3. PROBLEM DIAGNOSIS ---
            problems = []
            for sh in wb.Sheets:
                rows = sh.UsedRange.Rows.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                if rows > data + 2000:
                    problems.append({
                        "issue": "Ghost Rows",
                        "desc": f"Sheet '{sh.Name}' has excess empty formatting.",
                        "solution": "Delete rows beyond the last data cell and reset the scroll bar.",
                        "id": "GHOST"
                    })

            if not self.file_path.lower().endswith(".xlsb"):
                problems.append({
                    "issue": "Inefficient Format",
                    "desc": "The file is currently XML-based (.xlsx).",
                    "solution": "Convert to Binary (.xlsb) to reduce file size and speed up opening/saving.",
                    "id": "BINARY"
                })

            self.pb['value'] = 75
            self.status.config(text="Testing Formula Calculation Speed...")
            start_calc = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start_calc

            # --- 4. RATING & IT ADVICE ---
            sys_advice = "Normal Operations"
            it_alert = False
            if ram.percent > 85 or cpu > 90:
                sys_advice = "Contact your IT Team for Hardware Upgrade"
                it_alert = True

            self.audit_data = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "props": props,
                "problems": problems,
                "calc": f"{calc_time:.2f}s",
                "load": f"{load_time:.2f}s",
                "cpu": f"{cpu}%",
                "ram": f"{ram.percent}%",
                "advice": sys_advice,
                "it_alert": it_alert
            }
            
            wb.Close(False)
            excel.Quit()
            self.pb['value'] = 100
            self.root.after(500, self.display_report)
        except Exception as e:
            messagebox.showerror("Scan Error", f"Report generation failed: {e}")
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- LEFT PANEL: FILE INFO ---
        main_frame = tk.Frame(self.container, bg=CLR_BG)
        main_frame.pack(fill="both", expand=True)

        info_col = tk.Frame(main_frame, bg=CLR_BG, width=300)
        info_col.pack(side="left", fill="y", padx=(0, 20))

        tk.Label(info_col, text="File Metadata", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        meta_box = tk.Frame(info_col, bg="white", padx=15, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
        meta_box.pack(fill="x", pady=5)
        for k, v in self.audit_data['props'].items():
            tk.Label(meta_box, text=f"{k}:", font=("Segoe UI", 8, "bold"), bg="white").pack(anchor="w")
            tk.Label(meta_box, text=v, font=("Segoe UI", 9), bg="white", fg=CLR_BLUE).pack(anchor="w", pady=(0,5))

        # --- RIGHT PANEL: HARDWARE & SOLUTIONS ---
        content_col = tk.Frame(main_frame, bg=CLR_BG)
        content_col.pack(side="left", fill="both", expand=True)

        # Hardware Card
        hw_bg = CLR_ERR if self.audit_data['it_alert'] else CLR_EXCEL
        hw_card = tk.Frame(content_col, bg=hw_bg, padx=20, pady=15)
        hw_card.pack(fill="x", pady=(0, 20))
        tk.Label(hw_card, text=f"SYSTEM STATUS: {self.audit_data['advice']}", font=("Segoe UI", 11, "bold"), bg=hw_bg, fg="white").pack(side="left")
        tk.Label(hw_card, text=f"RAM: {self.audit_data['ram']} | CPU: {self.audit_data['cpu']}", font=("Segoe UI", 10), bg=hw_bg, fg="white").pack(side="right")

        # Problems & Solutions Table
        tk.Label(content_col, text="Analysis & Solutions", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        
        for p in self.audit_data['problems']:
            p_box = tk.Frame(content_col, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#E0E0E0")
            p_box.pack(fill="x", pady=5)
            
            top = tk.Frame(p_box, bg="white")
            top.pack(fill="x")
            tk.Label(top, text=f"⚠️ {p['issue']}", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(side="left")
            tk.Button(top, text="Fix Problem", bg=CLR_BLUE, fg="white", font=("Segoe UI", 8), command=lambda i=p['id']: self.run_fix(i)).pack(side="right")
            
            tk.Label(p_box, text=p['desc'], font=("Segoe UI", 9), bg="white").pack(anchor="w")
            tk.Label(p_box, text=f"SOLUTION: {p['solution']}", font=("Segoe UI", 9, "italic"), bg="white", fg="#605E5C").pack(anchor="w", pady=(5,0))

        # Master Action
        tk.Button(content_col, text="⚡ APPLY ALL SOLUTIONS & OPTIMIZE PERFORMANCE", 
                  command=lambda: self.run_fix("ALL"), bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15).pack(fill="x", pady=20)

    def run_fix(self, mode):
        win = tk.Toplevel(self.root)
        win.title("Repair Process")
        win.geometry("450x250")
        win.configure(bg="white")
        lbl = tk.Label(win, text="Optimizing... Please Wait.", font=("Segoe UI", 11), bg="white", pady=50)
        lbl.pack()

        def engine():
            start_t = time.time()
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
                
                duration = time.time() - start_t
                lbl.config(text=f"✅ Process Complete!\nDuration: {duration:.2f}s", fg=CLR_EXCEL)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Process Failed: {e}", fg=CLR_ERR)

        threading.Thread(target=engine, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = EnterpriseExcelSuite(root)
    root.mainloop()