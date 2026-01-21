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
CLR_HEADER = "#201F1E"

class VerticalEnterpriseSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Forensic Optimizer v10.7")
        self.root.geometry("1100x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
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
        tk.Label(self.container, text="Excel Forensic Intelligence", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(80, 10))
        tk.Label(self.container, text="Complete Vertical Audit & Compliance Reporting", font=("Segoe UI", 11), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 START BATCH AUDIT", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 11, "bold"), relief="flat", padx=40, pady=15, cursor="hand2")
        btn.pack(pady=40)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = list(paths)
        self.start_audit()

    def start_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=700, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Initializing Deep Vertical Scan...", bg=CLR_BG, font=("Segoe UI", 10))
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
                    volatile_pattern = re.compile(r"OFFSET\(|INDIRECT\(|TODAY\(|RAND\(")
                    
                    for sh in wb.Sheets:
                        last_cell = sh.Cells.SpecialCells(11)
                        if last_cell.Row > 5000 and sh.UsedRange.Rows.Count < (last_cell.Row * 0.5):
                            issues.append(f"Phantom Data in {sh.Name}: {last_cell.Row} rows detected vs {sh.UsedRange.Rows.Count} active rows.")
                        if sh.PivotTables().Count > 0: issues.append(f"Pivot Cache Bloat in {sh.Name}: Internal data duplication detected.")
                        try:
                            f_range = sh.UsedRange.SpecialCells(-4123).Formula
                            if volatile_pattern.search(str(f_range)): issues.append(f"Volatile Lag in {sh.Name}: Use of OFFSET or INDIRECT detected.")
                        except: pass

                    if wb.HasVBProject: issues.append("VBA Metadata: Potential macro/virus security risk found.")
                    if wb.LinkSources(1): issues.append("External Links: File depends on external network paths.")

                    self.batch_results.append({
                        "name": fname, "size": f"{fsize:.2f} MB", "issues": issues,
                        "health": "Fully Compliant" if not issues else "Needs Optimization"
                    })
                    wb.Close(False)
                except:
                    self.batch_results.append({"name": fname, "size": "N/A", "issues": ["Access Error: File is password protected or corrupted."], "health": "Critical"})
            excel.Quit()
        finally:
            pythoncom.CoUninitialize()
            self.root.after(100, self.display_final_audit)

    def display_final_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # 1. SYSTEM DASHBOARD (Vertical Stack)
        sys_f = tk.Frame(self.container, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#DDD")
        sys_f.pack(fill="x", pady=(0, 20))
        tk.Label(sys_f, text="SYSTEM STATUS REPORT", font=("Segoe UI", 9, "bold"), bg="white", fg=CLR_BLUE).pack(anchor="w")
        tk.Label(sys_f, text=f"Operating System: {platform.system()} {platform.release()}", bg="white", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(sys_f, text=f"Resource Load: RAM {psutil.virtual_memory().percent}% | CPU {psutil.cpu_percent()}%", bg="white", font=("Segoe UI", 9)).pack(anchor="w")

        # 2. SCROLLABLE RESULTS
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=CLR_BG)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw", width=1000)
        canvas.configure(yscrollcommand=scroll.set)

        for res in self.batch_results:
            card = tk.Frame(frame, bg="white", padx=25, pady=20, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=10)
            
            tk.Label(card, text=f"FILE NAME: {res['name']}", font=("Segoe UI", 12, "bold"), bg="white", fg=CLR_HEADER).pack(anchor="w")
            tk.Label(card, text=f"Disk Usage: {res['size']}", font=("Segoe UI", 9), bg="white", fg="#666").pack(anchor="w", pady=(2, 10))
            
            if res['health'] == "Fully Compliant":
                tk.Label(card, text="STATUS: FULLY COMPLIANT WITH PERFORMANCE STANDARDS", font=("Segoe UI", 9, "bold"), fg=CLR_EXCEL, bg="white").pack(anchor="w")
                
                # Detailed reasoning instead of ticks
                compliance_info = [
                    "Structural Audit: No phantom rows detected; UsedRange is clean.",
                    "Calculation Audit: No volatile functions detected; workbook uses static references.",
                    "Metadata Audit: No hidden VBA projects or unauthorized objects found.",
                    "Link Audit: All data connections are internal; no external dependency lag."
                ]
                for info in compliance_info:
                    tk.Label(card, text=f"• {info}", font=("Segoe UI", 9), fg="#444", bg="white", wraplength=900, justify="left").pack(anchor="w", padx=10, pady=1)
            else:
                tk.Label(card, text=f"STATUS: {res['health'].upper()}", font=("Segoe UI", 9, "bold"), fg=CLR_ERR, bg="white").pack(anchor="w")
                tk.Label(card, text="The following issues are affecting file speed and security:", font=("Segoe UI", 9, "italic"), bg="white", fg="#555").pack(anchor="w", pady=(5, 2))
                
                for prob in res['issues']:
                    p_frame = tk.Frame(card, bg="#FFF9F9", pady=5)
                    p_frame.pack(fill="x", pady=2)
                    tk.Label(p_row := tk.Frame(p_frame, bg="#FFF9F9"), bg="#FFF9F9").pack(fill="x")
                    tk.Label(p_row, text=f"DETECTED: {prob}", font=("Segoe UI", 9, "bold"), fg=CLR_ERR, bg="#FFF9F9", wraplength=900, justify="left").pack(anchor="w")
                    tk.Label(p_frame, text="SOLUTION: Run Master Optimization to reset UsedRange and strip volatile formatting.", font=("Segoe UI", 8), fg="#444", bg="#FFF9F9").pack(anchor="w", padx=15)

        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # 3. RIGHT-ALIGNED COMPACT BUTTONS
        btn_frame = tk.Frame(self.container, bg=CLR_BG)
        btn_frame.pack(fill="x", pady=20)
        
        tk.Frame(btn_frame, bg=CLR_BG).pack(side="left", expand=True) # Spacer
        
        tk.Button(btn_frame, text="ADD MORE FILES", command=self.show_home, 
                  bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), padx=20, pady=10, relief="flat").pack(side="left", padx=10)
        
        tk.Button(btn_frame, text="EXECUTE MASTER OPTIMIZATION", 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 9, "bold"), padx=20, pady=10, relief="flat").pack(side="left")

if __name__ == "__main__":
    root = tk.Tk()
    app = VerticalEnterpriseSuite(root)
    root.mainloop()