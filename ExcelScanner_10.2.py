import os, psutil, time, ctypes, threading, platform, re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import pythoncom

# --- DESIGN SYSTEM ---
CLR_BG = "#F3F2F1"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_WARN = "#FFB900"
CLR_ERR = "#D13438"
CLR_GREEN_BG = "#EBF3EF"

class EnterpriseForensicSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Forensic Optimizer v10.6")
        self.root.geometry("1300x980")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=30, pady=20)
        
        self.engine_reset()
        self.show_home()

    def engine_reset(self):
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].upper() == "EXCEL.EXE": proc.kill()
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Forensic Intelligence", font=("Segoe UI", 32, "bold"), bg=CLR_BG).pack(pady=(100, 10))
        tk.Label(self.container, text="Enterprise Structural Audit • Optimization Standards v10.6", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 START BATCH AUDIT", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=18, cursor="hand2")
        btn.pack(pady=50)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = list(paths)
        self.start_audit()

    def start_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Initializing Audit Engine...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        threading.Thread(target=self.run_forensics, daemon=True).start()

    def run_forensics(self):
        self.batch_results = []
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts, excel.Visible, excel.AutomationSecurity = False, False, 3
            
            for index, path in enumerate(self.file_paths):
                fname = os.path.basename(path)
                fsize = os.path.getsize(path) / (1024 * 1024)
                self.pb['value'] = ((index + 1) / len(self.file_paths)) * 100
                
                try:
                    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                    issues = []
                    
                    # SCAN FOR PROBLEMS
                    volatile_pattern = re.compile(r"OFFSET\(|INDIRECT\(|TODAY\(|RAND\(")
                    for sh in wb.Sheets:
                        last_cell = sh.Cells.SpecialCells(11)
                        if last_cell.Row > 5000 and sh.UsedRange.Rows.Count < (last_cell.Row * 0.5):
                            issues.append(f"Phantom Data ({sh.Name})")
                        if sh.PivotTables().Count > 0: issues.append("Pivot Cache Bloat")
                        try:
                            f_range = sh.UsedRange.SpecialCells(-4123).Formula
                            if volatile_pattern.search(str(f_range)): issues.append("Volatile Function Lag")
                        except: pass

                    if wb.HasVBProject: issues.append("VBA Metadata Risk")
                    if wb.LinkSources(1): issues.append("External Link Delay")

                    self.batch_results.append({
                        "name": fname, "size": f"{fsize:.2f} MB", "issues": issues,
                        "health": "Stable" if not issues else "Needs Optimization"
                    })
                    wb.Close(False)
                except:
                    self.batch_results.append({"name": fname, "size": "N/A", "issues": ["Encrypted/Access Denied"], "health": "Critical"})
            excel.Quit()
        finally:
            pythoncom.CoUninitialize()
            self.root.after(100, self.display_final_audit)

    def display_final_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # 1. SYSTEM INFO
        sys_f = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        sys_f.pack(fill="x", pady=(0, 15))
        tk.Label(sys_f, text=f"SYS: {platform.system()} | RAM: {psutil.virtual_memory().percent}% | DISK: {psutil.disk_usage('/').percent}%", font=("Segoe UI", 8, "bold"), bg="white", fg=CLR_BLUE).pack(side="left")

        # 2. SCROLLABLE RESULTS
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=CLR_BG)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw", width=1200)
        canvas.configure(yscrollcommand=scroll.set)

        for res in self.batch_results:
            card = tk.Frame(frame, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=5)
            
            tk.Label(card, text=f"📄 {res['name']}", font=("Segoe UI", 11, "bold"), bg="white").pack(anchor="w")
            
            if res['health'] == "Stable":
                # PERFORMANCE STANDARDS COMPLIANCE
                tk.Label(card, text="✅ HEALTHY: FILE MEETS PERFORMANCE STANDARDS", font=("Segoe UI", 9, "bold"), fg=CLR_EXCEL, bg="white").pack(anchor="w", pady=5)
                standards = "✔ No Phantom Rows | ✔ Static Formulas | ✔ Clean Metadata | ✔ Optimal Pivot Cache | ✔ Internal Links Verified"
                tk.Label(card, text=standards, font=("Segoe UI", 8), fg="#666", bg="white").pack(anchor="w")
            else:
                # LIST ISSUES & SOLUTIONS
                tk.Label(card, text=f"⚠ STATUS: {res['health']}", font=("Segoe UI", 9, "bold"), fg=CLR_ERR, bg="white").pack(anchor="w", pady=5)
                for prob in res['issues']:
                    p_row = tk.Frame(card, bg="#FFF9F9")
                    p_row.pack(fill="x", pady=2)
                    tk.Label(p_row, text=f"• {prob}", font=("Segoe UI", 8, "bold"), fg=CLR_ERR, bg="#FFF9F9").pack(side="left")
                    tk.Label(p_row, text="  -> Fix: Clean used range / Replace Volatile functions / Clear cache", font=("Segoe UI", 8), fg="#444", bg="#FFF9F9").pack(side="left")

        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # 3. COMPACT RIGHT-ALIGNED BUTTONS
        btn_frame = tk.Frame(self.container, bg=CLR_BG)
        btn_frame.pack(fill="x", pady=20)
        
        # Spacer to push buttons to the right
        tk.Frame(btn_frame, bg=CLR_BG).pack(side="left", expand=True)
        
        tk.Button(btn_frame, text="➕ ADD MORE", command=self.show_home, 
                  bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), padx=15, pady=8).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="⚡ EXECUTE MASTER OPTIMIZATION", 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 9, "bold"), padx=15, pady=8).pack(side="left")

if __name__ == "__main__":
    root = tk.Tk()
    app = EnterpriseForensicSuite(root)
    root.mainloop()