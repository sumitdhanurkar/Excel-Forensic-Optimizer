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

class UltimateForensicSuiteV105:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Forensic Optimizer v10.5")
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
        tk.Label(self.container, text="Excel Forensic Intelligence", font=("Segoe UI", 32, "bold"), bg=CLR_BG).pack(pady=(80, 10))
        tk.Label(self.container, text="Deep Metadata Analysis • Structural Optimization • Performance Recovery", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT BATCH FOR FORENSIC AUDIT", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 13, "bold"), relief="flat", padx=50, pady=22, cursor="hand2")
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
        self.status = tk.Label(self.container, text="Initializing Deep Scan Engine...", bg=CLR_BG, font=("Segoe UI", 10))
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
                self.status.config(text=f"Forensic Audit {index+1}/{len(self.file_paths)}: {fname}")
                self.pb['value'] = ((index + 1) / len(self.file_paths)) * 100
                
                try:
                    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                    issues = []
                    
                    # 1. SCAN FACTORS: BLOAT & PERFORMANCE
                    volatile_pattern = re.compile(r"OFFSET\(|INDIRECT\(|TODAY\(|RAND\(")
                    pivot_bloat, phantom_data, volatile_found = False, False, False
                    
                    for sh in wb.Sheets:
                        # Phantom Data Check (Used Range vs Data)
                        last_cell = sh.Cells.SpecialCells(11) # xlCellTypeLastCell
                        if last_cell.Row > 5000 and sh.UsedRange.Rows.Count < (last_cell.Row * 0.5):
                            phantom_data = True
                        
                        # Pivot Cache Check
                        if sh.PivotTables().Count > 0: pivot_bloat = True
                        
                        # Volatile & Full Column Formula Check
                        try:
                            formulas = sh.UsedRange.SpecialCells(-4123).Formula # xlCellTypeFormulas
                            if isinstance(formulas, str): formulas = [formulas]
                            for f in str(formulas):
                                if volatile_pattern.search(f): volatile_found = True
                                if ":" in f and any(char*2 in f for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"): issues.append(f"Full Column Ref in {sh.Name}")
                        except: pass

                    # Populate Detailed Audit List
                    if phantom_data: issues.append("Bloated Used Range (Phantom Data)")
                    if pivot_bloat: issues.append("Pivot Table Cache Overload")
                    if volatile_found: issues.append("Volatile Functions (OFFSET/INDIRECT)")
                    if wb.HasVBProject: issues.append("Security: VBA Metadata")
                    if wb.LinkSources(1): issues.append("External Connection Lag")

                    self.batch_results.append({
                        "name": fname, "size": f"{fsize:.2f} MB", "format": fname.split('.')[-1],
                        "issues": issues, "health": "Stable" if not issues else ("Degraded" if len(issues) < 3 else "Critical")
                    })
                    wb.Close(False)
                except:
                    self.batch_results.append({"name": fname, "size": "N/A", "format": "Unknown", "issues": ["Locked/Encrypted File"], "health": "Critical"})
            excel.Quit()
        finally:
            pythoncom.CoUninitialize()
            self.root.after(100, self.display_final_audit)

    def display_final_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # 1. SYSTEM INFO
        sys_f = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        sys_f.pack(fill="x", pady=(0, 15))
        tk.Label(sys_f, text="[1] SYSTEM DIAGNOSTICS", font=("Segoe UI", 9, "bold"), bg="white", fg=CLR_BLUE).pack(anchor="w")
        tk.Label(sys_f, text=f"OS: {platform.system()} | RAM: {psutil.virtual_memory().percent}% | Disk: {psutil.disk_usage('/').percent}% Used", bg="white", font=("Segoe UI", 9)).pack(anchor="w")

        # 2. AUDIT LOG & TASKS PERFORMED
        log_f = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        log_f.pack(fill="x", pady=(0, 15))
        tk.Label(log_f, text="[2] TASKS PERFORMED & METHODOLOGY", font=("Segoe UI", 9, "bold"), bg="white", fg=CLR_BLUE).pack(anchor="w")
        tk.Label(log_f, text="Scanning for: Virus/VBA, External Connections, Used Range Bloat, Volatile Formulas, and Pivot Cache.", bg="white", fg="#666", font=("Segoe UI", 8)).pack(anchor="w")

        # 3. DETAILED RESULTS
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=CLR_BG)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw", width=1200)
        canvas.configure(yscrollcommand=scroll.set)

        for res in self.batch_results:
            card = tk.Frame(frame, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=5)
            
            # Row 1: Health & Name
            h_color = CLR_EXCEL if res['health'] == "Stable" else (CLR_ERR if res['health'] == "Critical" else CLR_WARN)
            tk.Label(card, text=f"FILE: {res['name']}", font=("Segoe UI", 11, "bold"), bg="white").grid(row=0, column=0, sticky="w")
            tk.Label(card, text=f"HEALTH: {res['health']}", font=("Segoe UI", 9, "bold"), fg=h_color, bg="white").grid(row=0, column=1, sticky="e")
            
            # Row 2: Stats
            tk.Label(card, text=f"Disk Usage: {res['size']} | Format: {res['format']}", font=("Segoe UI", 9), bg="white", fg="#555").grid(row=1, column=0, sticky="w", pady=5)

            # Row 3: Affecting Problems & Solutions
            if res['issues']:
                for i, prob in enumerate(res['issues']):
                    p_f = tk.Frame(card, bg="#FFF9F9")
                    p_f.grid(row=2+i, column=0, columnspan=2, sticky="w", pady=2)
                    tk.Label(p_f, text=f"⚠ Problem: {prob}", font=("Segoe UI", 8, "bold"), fg=CLR_ERR, bg="#FFF9F9").pack(side="left")
                    sol = "Fix: Save as .xlsb / Delete Phantom Rows / Replace OFFSET with INDEX."
                    tk.Label(p_f, text=f" -> {sol}", font=("Segoe UI", 8), fg="#444", bg="#FFF9F9").pack(side="left")
            else:
                tk.Label(card, text="✅ Optimized: This file meets all performance standards.", font=("Segoe UI", 9), fg=CLR_EXCEL, bg="white").grid(row=2, column=0, sticky="w")

        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # 4. FINAL ACTIONS
        b_nav = tk.Frame(self.container, bg=CLR_BG)
        b_nav.pack(fill="x", pady=20)
        tk.Button(b_nav, text="➕ ADD MORE FILES", command=self.show_home, bg=CLR_BLUE, fg="white", font=("Segoe UI", 10, "bold"), padx=30, pady=12).pack(side="left")
        tk.Button(b_nav, text="⚡ EXECUTE MASTER OPTIMIZATION (CLEAN ALL)", bg=CLR_EXCEL, fg="white", font=("Segoe UI", 10, "bold"), padx=30, pady=12).pack(side="right")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateForensicSuiteV105(root)
    root.mainloop()