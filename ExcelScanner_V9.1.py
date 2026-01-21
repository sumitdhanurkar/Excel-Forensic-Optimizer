import os
import psutil
import platform
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

class FinalHardwareSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence Suite v9.1")
        self.root.geometry("1100x900")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=30)
        self.show_home()

    def kill_ghost_excel(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Intelligence Suite", font=("Segoe UI", 32, "bold"), bg=CLR_BG, fg=CLR_TXT).pack(pady=(120, 10))
        tk.Label(self.container, text="Hardware Diagnostic & Multi-Task Repair Engine", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT & AUTO-SCAN", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=50, pady=20, cursor="hand2")
        btn.pack(pady=50)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Analyzing Hardware & File Sync...", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=700, mode="determinate")
        self.pb.pack(pady=10)
        self.status = tk.Label(self.container, text="Measuring System Load...", bg=CLR_BG, font=("Consolas", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_scan, daemon=True).start()

    def perform_scan(self):
        try:
            # --- HARDWARE SCAN ---
            cpu_usage = psutil.cpu_percent(interval=0.5)
            ram = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            self.kill_ghost_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            self.status.config(text="Opening Workbook...")
            self.pb['value'] = 25
            
            start_open = time.time()
            wb = excel.Workbooks.Open(self.file_path)
            load_time = time.time() - start_open
            
            # --- LOGIC AUDIT ---
            self.status.config(text="Mapping Data Density...")
            problems = []
            for sh in wb.Sheets:
                rows = sh.UsedRange.Rows.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                if rows > data + 2000:
                    problems.append(("Ghost Rows", f"'{sh.Name}' has {rows-data} bloated empty rows.", "GHOST"))

            self.pb['value'] = 75
            self.status.config(text="Benchmarking Formulas...")
            start_calc = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start_calc

            # --- RATING SYSTEM ---
            capability = "EXCELLENT"
            if ram.percent > 85 or load_time > 10: capability = "STRUGGLING"
            elif ram.percent > 60: capability = "MODERATE"

            self.audit_data = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "problems": problems,
                "calc_speed": f"{calc_time:.2f}s",
                "load_speed": f"{load_time:.2f}s",
                "cpu": f"{cpu_usage}%",
                "ram": f"{ram.percent}%",
                "disk": f"{disk.percent}%",
                "capability": capability
            }
            
            wb.Close(False)
            excel.Quit()
            self.pb['value'] = 100
            self.root.after(500, self.display_report)
        except Exception as e:
            messagebox.showerror("Scan Failed", f"Engine Error: {e}")
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- HARDWARE DASHBOARD ---
        tk.Label(self.container, text="System Hardware Status", font=("Segoe UI", 14, "bold"), bg=CLR_BG).pack(anchor="w", pady=(0,10))
        hw_frame = tk.Frame(self.container, bg=CLR_BG)
        hw_frame.pack(fill="x", pady=5)

        hw_metrics = [
            ("CPU Load", self.audit_data['cpu']),
            ("RAM Usage", self.audit_data['ram']),
            ("Disk Space", self.audit_data['disk']),
            ("System Rating", self.audit_data['capability'])
        ]

        for title, val in hw_metrics:
            card = tk.Frame(hw_frame, bg="white", highlightbackground="#E0E0E0", highlightthickness=1, padx=15, pady=10)
            card.pack(side="left", fill="both", expand=True, padx=4)
            tk.Label(card, text=title, font=("Segoe UI", 9), bg="white", fg="#605E5C").pack()
            tk.Label(card, text=val, font=("Segoe UI", 12, "bold"), bg="white", fg=CLR_BLUE if "Rating" not in title else CLR_EXCEL).pack()

        # --- FILE USAGE & PROBLEMS ---
        tk.Label(self.container, text=f"File Performance: {os.path.basename(self.file_path)}", font=("Segoe UI", 14, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20,10))
        
        perf_frame = tk.Frame(self.container, bg="white", padx=20, pady=15, highlightbackground="#E0E0E0", highlightthickness=1)
        perf_frame.pack(fill="x")
        tk.Label(perf_frame, text=f"📂 Size: {self.audit_data['size']:.2f} MB  |  ⏱️ Load Time: {self.audit_data['load_speed']}  |  🧮 Calc Time: {self.audit_data['calc_speed']}", bg="white", font=("Segoe UI", 10)).pack(side="left")

        # Problem List
        tk.Label(self.container, text="Detected Issues", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20, 5))
        if not self.audit_data['problems']:
            tk.Label(self.container, text="✅ No critical issues found in file structure.", fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w")
        else:
            for title, desc, tid in self.audit_data['problems']:
                f = tk.Frame(self.container, bg="white", pady=10, padx=20, highlightbackground="#EDEBE9", highlightthickness=1)
                f.pack(fill="x", pady=2)
                tk.Label(f, text=f"⚠️ {title}: {desc}", font=("Segoe UI", 10), bg="white").pack(side="left")
                tk.Button(f, text="Fix Only This", command=lambda t=tid: self.run_repair(t), bg=CLR_BLUE, fg="white", font=("Segoe UI", 8)).pack(side="right")

        # --- MASTER CONTROL ---
        btn_master = tk.Button(self.container, text="⚡ EXECUTE MASTER PERFORMANCE OVERHAUL (FIX ALL)", 
                               command=lambda: self.run_repair("ALL"), bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", pady=15)
        btn_master.pack(fill="x", pady=30)

    def run_repair(self, mode):
        # [Repair logic remains the same as v9.0, including time benchmarking for the final message]
        win = tk.Toplevel(self.root)
        win.title("Repairing")
        win.geometry("400x200")
        lbl = tk.Label(win, text="Starting repairs...", pady=40)
        lbl.pack()

        def execute():
            start_repair = time.time()
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

                wb.Save()
                wb.Close()
                excel.Quit()
                
                total_time = time.time() - start_repair
                lbl.config(text=f"✅ Repair Complete!\nTotal Process Time: {total_time:.2f}s", fg=CLR_EXCEL)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {e}", fg=CLR_ERR)

        threading.Thread(target=execute, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalHardwareSuite(root)
    root.mainloop()