import os, psutil, time, ctypes, threading, platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import pythoncom

# --- DESIGN SYSTEM ---
CLR_BG = "#F3F2F1"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_CARD = "#FFFFFF"

class MasterAuditSuiteV10:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Master Audit & Security Suite v10.0")
        self.root.geometry("1250x980")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=30, pady=20)
        
        self.initial_cleanup()
        self.show_home()

    def initial_cleanup(self):
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].upper() == "EXCEL.EXE": proc.kill()
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Master Intelligence", font=("Segoe UI", 32, "bold"), bg=CLR_BG).pack(pady=(100, 10))
        tk.Label(self.container, text="Full System Audit • Deep File Forensics • Security Intelligence", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 START DEEP BATCH SCAN", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 13, "bold"), relief="flat", padx=50, pady=22, cursor="hand2")
        btn.pack(pady=50)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = list(paths)
        self.start_engine()

    def start_engine(self):
        for widget in self.container.winfo_children(): widget.destroy()
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Booting Forensic Engine...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        threading.Thread(target=self.run_forensics, daemon=True).start()

    def run_forensics(self):
        self.batch_results = []
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3 
            
            for index, path in enumerate(self.file_paths):
                fname = os.path.basename(path)
                fsize = os.path.getsize(path) / (1024 * 1024)
                self.status.config(text=f"Analyzing {index+1}/{len(self.file_paths)}: {fname}")
                self.pb['value'] = ((index + 1) / len(self.file_paths)) * 100
                
                try:
                    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                    rows, cols, problems = 0, 0, []
                    formula_count = 0
                    
                    for sh in wb.Sheets:
                        rows += sh.UsedRange.Rows.Count
                        cols += sh.UsedRange.Columns.Count
                        # Detect volatile/heavy formulas or huge data
                        try: formula_count += sh.UsedRange.SpecialCells(win32com.client.constants.xlCellTypeFormulas).Count
                        except: pass
                        
                    # Logic for Affecting Factors
                    if wb.HasVBProject: problems.append("Virus Risk: VBA Macros")
                    if wb.LinkSources(1): problems.append(f"Performance: {len(wb.LinkSources(1))} Ext Links")
                    if formula_count > 5000: problems.append("Speed: Heavy Formula Load")
                    if fsize > 20: problems.append("Size: High Disk Footprint")

                    self.batch_results.append({
                        "name": fname, "size": f"{fsize:.2f} MB", "dims": f"{rows:,}x{cols:,}",
                        "formulas": f"{formula_count:,}", "problems": problems, "health": "Stable" if not problems else "Degraded"
                    })
                    wb.Close(False)
                except:
                    self.batch_results.append({"name": fname, "size": "N/A", "dims": "Locked", "formulas": "N/A", "problems": ["Encryption/Password"], "health": "Critical"})
            excel.Quit()
        finally:
            pythoncom.CoUninitialize()
            self.root.after(100, self.display_final_audit)

    def display_final_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # 1. SYSTEM INFORMATION (OS, CPU, RAM)
        sys_frame = tk.Frame(self.container, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#DDD")
        sys_frame.pack(fill="x", pady=(0, 20))
        tk.Label(sys_frame, text="1. SYSTEM ENVIRONMENT INFORMATION", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_BLUE).pack(anchor="w")
        sys_info = f"OS: {platform.system()} {platform.release()} | CPU: {psutil.cpu_percent()}% | RAM: {psutil.virtual_memory().percent}% | DISK: {psutil.disk_usage('/').percent}% Used"
        tk.Label(sys_frame, text=sys_info, font=("Segoe UI", 9), bg="white").pack(anchor="w")

        # 2. AUDIT PROCESS & METHODOLOGY
        proc_frame = tk.Frame(self.container, bg="white", padx=20, pady=10, highlightthickness=1, highlightbackground="#DDD")
        proc_frame.pack(fill="x", pady=(0, 20))
        tk.Label(proc_frame, text="2. AUDIT METHODOLOGY FOLLOWED", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_BLUE).pack(anchor="w")
        tk.Label(proc_frame, text="• Memory Flush -> COM Initialization -> Virus/Macro Scan -> Structural Dimension Audit -> Formula Density Check", 
                 font=("Segoe UI", 9), bg="white", fg="#666").pack(anchor="w")

        # 3. SCROLLABLE DETAILED RESULTS
        tk.Label(self.container, text="3. FILE-SPECIFIC ANALYSIS & HEALTH", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(0,5))
        
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=CLR_BG)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw", width=1180)
        canvas.configure(yscrollcommand=scroll.set)

        for res in self.batch_results:
            card = tk.Frame(frame, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=5)
            
            # File Health Indicator
            h_color = CLR_EXCEL if res['health'] == "Stable" else (CLR_ERR if res['health'] == "Critical" else "#FFB900")
            tk.Label(card, text=f"FILE: {res['name']}", font=("Segoe UI", 11, "bold"), bg="white").grid(row=0, column=0, sticky="w")
            tk.Label(card, text=f"HEALTH: {res['health']}", font=("Segoe UI", 9, "bold"), fg=h_color, bg="white").grid(row=0, column=1, sticky="e")
            
            # Data Stats (Rows, Cols, Disk)
            stats = f"Size: {res['size']} | Dim: {res['dims']} | Formulas: {res['formulas']}"
            tk.Label(card, text=stats, font=("Segoe UI", 9), bg="white", fg="#555").grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
            
            # Problems and Solutions
            if res['problems']:
                prob_text = "Problems Found: " + ", ".join(res['problems'])
                tk.Label(card, text=prob_text, font=("Segoe UI", 8, "bold"), bg="white", fg=CLR_ERR).grid(row=2, column=0, sticky="w")
                tk.Label(card, text="Recommended: Reset UsedRange, Remove VBA, or Break Links.", font=("Segoe UI", 8, "italic"), bg="white", fg="#777").grid(row=3, column=0, sticky="w")

        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # 5. FINAL NAVIGATION
        bottom_nav = tk.Frame(self.container, bg=CLR_BG)
        bottom_nav.pack(fill="x", pady=20)
        tk.Button(bottom_nav, text="➕ ADD MORE FILES", command=self.show_home, bg=CLR_BLUE, fg="white", font=("Segoe UI", 10, "bold"), padx=30, pady=10).pack(side="left")
        tk.Button(bottom_nav, text="⚡ APPLY MASTER FIXES TO BATCH", bg=CLR_EXCEL, fg="white", font=("Segoe UI", 10, "bold"), padx=30, pady=10).pack(side="right")

if __name__ == "__main__":
    root = tk.Tk()
    app = MasterAuditSuiteV10(root)
    root.mainloop()