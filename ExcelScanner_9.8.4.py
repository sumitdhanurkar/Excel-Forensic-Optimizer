import os
import psutil
import time
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading

# --- THEME (v9.8 Style) ---
CLR_BG = "#F3F2F1"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_TXT = "#323130"

class AuditProfessionalV988:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Audit & Security v9.8.8")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
        self.pre_launch_cleanup()
        self.show_home()

    def pre_launch_cleanup(self):
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == "EXCEL.EXE": proc.kill()
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Audit & Security", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(60, 10))
        tk.Label(self.container, text="Disk Usage • Row/Col Audit • Security Factors", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT EXCEL FILES", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=18, cursor="hand2")
        btn.pack(pady=40)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = list(paths)
        self.start_batch_audit()

    def start_batch_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Analyzing Files...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        threading.Thread(target=self.perform_batch_scan, daemon=True).start()

    def perform_batch_scan(self):
        self.batch_results = []
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3 
            
            for index, path in enumerate(self.file_paths):
                fname = os.path.basename(path)
                fsize_mb = os.path.getsize(path) / (1024 * 1024)
                self.status.config(text=f"Auditing {index+1}/{len(self.file_paths)}: {fname}")
                self.pb['value'] = ((index + 1) / len(self.file_paths)) * 100
                
                try:
                    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                    
                    sheet_data = []
                    total_rows = 0
                    total_cols = 0
                    factors = []

                    for sh in wb.Sheets:
                        used = sh.UsedRange
                        r_count = used.Rows.Count
                        c_count = used.Columns.Count
                        total_rows += r_count
                        total_cols += c_count
                        
                        # Identify Bloat Factors
                        if r_count > 10000: factors.append(f"High Row Volume ({sh.Name})")
                        if c_count > 50: factors.append(f"Wide Column Range ({sh.Name})")
                        if sh.Shapes.Count > 10: factors.append(f"Excess Objects/Shapes ({sh.Name})")

                    if wb.HasVBProject: factors.append("VBA Project (Security Risk)")
                    if wb.LinkSources(1): factors.append("External Data Links")
                    
                    self.batch_results.append({
                        "name": fname,
                        "size": f"{fsize_mb:.2f} MB",
                        "dims": f"{total_rows} Rows / {total_cols} Cols",
                        "factors": factors if factors else ["Clean Structure"],
                        "status": "Verified"
                    })
                    wb.Close(False)
                except Exception as e:
                    self.batch_results.append({"name": fname, "size": "Error", "dims": "N/A", "factors": [str(e)[:30]], "status": "Failed"})
            
            excel.Quit()
            self.root.after(100, self.display_report)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- NAVIGATION ---
        nav = tk.Frame(self.container, bg=CLR_BG)
        nav.pack(fill="x", pady=(0, 10))
        tk.Label(nav, text="Detailed Audit History", font=("Segoe UI", 16, "bold"), bg=CLR_BG).pack(side="left")
        tk.Button(nav, text="➕ ADD MORE FILES", command=self.show_home, bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), padx=15).pack(side="right")

        # --- DISK & SYSTEM DASHBOARD ---
        disk = psutil.disk_usage('/')
        ram = psutil.virtual_memory().percent
        hw_frame = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        hw_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(hw_frame, text=f"DISK SPACE: {disk.percent}% Used", font=("Segoe UI", 9, "bold"), bg="white").pack(side="left", padx=20)
        tk.Label(hw_frame, text=f"SYSTEM RAM: {ram}%", font=("Segoe UI", 9, "bold"), bg="white", fg=CLR_BLUE if ram < 85 else CLR_ERR).pack(side="left")
        
        # --- RESULT CARDS (Enhanced with Rows/Cols/Factors) ---
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=CLR_BG)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=1100)
        canvas.configure(yscrollcommand=scrollbar.set)

        for res in self.batch_results:
            card = tk.Frame(scroll_frame, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=8, padx=10)
            
            # Row 1: File Name and Size
            r1 = tk.Frame(card, bg="white")
            r1.pack(fill="x")
            tk.Label(r1, text=f"📄 {res['name']}", font=("Segoe UI", 11, "bold"), bg="white").pack(side="left")
            tk.Label(r1, text=f"DISK SIZE: {res['size']}", font=("Segoe UI", 9, "bold"), bg="white", fg=CLR_BLUE).pack(side="right")

            # Row 2: Rows/Cols and Data Info
            r2 = tk.Frame(card, bg="white", pady=5)
            r2.pack(fill="x")
            tk.Label(r2, text=f"📊 DIMENSIONS: {res['dims']}", font=("Segoe UI", 9), bg="white", fg="#444").pack(side="left")
            
            # Row 3: Affecting Factors
            r3 = tk.Frame(card, bg="white", pady=2)
            r3.pack(fill="x")
            fact_str = " | ".join(res['factors'])
            tk.Label(r3, text=f"⚠️ AFFECTING FACTORS: {fact_str}", font=("Segoe UI", 8), bg="white", fg=CLR_ERR if "Clean" not in fact_str else CLR_EXCEL).pack(side="left")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Button(self.container, text="⚡ OPTIMIZE & CLEAN ENTIRE BATCH", 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15, command=lambda: messagebox.showinfo("Audit", "Deep Cleaning Started")).pack(fill="x", pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditProfessionalV988(root)
    root.mainloop()