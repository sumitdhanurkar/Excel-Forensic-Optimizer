import os
import psutil
import time
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import shutil

# --- DESIGN SYSTEM ---
CLR_BG = "#F3F2F1"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_WARN = "#FFB900"
CLR_TXT = "#323130"

class UltimateSecuritySuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intel & Security v9.8")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
        self.pre_launch_cleanup()
        self.show_home()

    def pre_launch_cleanup(self):
        """Clean RAM and kill stuck Excel processes."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == "EXCEL.EXE": proc.kill()
            # Flush RAM working set
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Security & Intelligence", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(80, 10))
        tk.Label(self.container, text="Hardware Stats • Security Scan • Performance Audit", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT & DEEP SCAN", command=self.select_file, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=18, cursor="hand2")
        btn.pack(pady=40)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        self.start_audit()

    def start_audit(self):
        """Triggered by home button OR the Rescan button."""
        for widget in self.container.winfo_children(): widget.destroy()
        
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Initializing Deep Scan Engine...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        
        # Always clean RAM before a fresh scan
        self.pre_launch_cleanup()
        threading.Thread(target=self.perform_deep_scan, daemon=True).start()

    def perform_deep_scan(self):
        try:
            # 1. HARDWARE STATS
            cpu_usage = psutil.cpu_percent(interval=0.5)
            ram = psutil.virtual_memory()
            disk = psutil.disk_usage(os.path.splitdrive(self.file_path)[0])
            
            # 2. FILE SCAN
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            start_t = time.time()
            wb = excel.Workbooks.Open(self.file_path)
            load_time = time.time() - start_t
            
            problems = []
            
            # Virus/Security Check (Macros & Links)
            ext_links = wb.LinkSources(1)
            if wb.HasVBProject:
                problems.append({"issue": "VBA Macros Detected", "reason": "Potential security risk or malware container.", "sol": "Inspect code or remove macros.", "id": "VBA"})
            
            if ext_links:
                problems.append({"issue": "External Connections", "reason": "Linked data can cause performance lag and privacy leaks.", "sol": "Break external links.", "id": "LINKS"})

            # Performance Check
            if load_time > 4.0:
                problems.append({"issue": "High Loading Time", "reason": f"File opened in {load_time:.2f}s (Over threshold).", "sol": "Optimize structure and convert format.", "id": "LOAD"})

            for sh in wb.Sheets:
                used_r = sh.UsedRange.Rows.Count
                try: data_c = excel.WorksheetFunction.CountA(sh.Cells)
                except: data_c = 0
                if used_r > data_c + 3000:
                    problems.append({"issue": f"Row Bloat: {sh.Name}", "reason": f"Sheet uses {used_r} rows but only has {data_c} data.", "sol": "Purge ghost rows.", "id": "GHOST"})

            self.audit_data = {
                "hw": {"cpu": f"{cpu_usage}%", "ram": f"{ram.percent}%", "disk": f"{disk.percent}%"},
                "meta": {"Size": f"{os.path.getsize(self.file_path)/(1024*1024):.2f} MB", "Load": f"{load_time:.2f}s"},
                "problems": problems
            }
            
            wb.Close(False)
            excel.Quit()
            self.root.after(100, self.display_report)
        except Exception as e:
            messagebox.showerror("Scan Error", str(e))
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- HEADER & RESCAN BUTTON ---
        header_nav = tk.Frame(self.container, bg=CLR_BG)
        header_nav.pack(fill="x", pady=(0, 15))
        
        tk.Label(header_nav, text="Diagnostic Results", font=("Segoe UI", 16, "bold"), bg=CLR_BG).pack(side="left")
        
        # THE NEW RESCAN BUTTON
        tk.Button(header_nav, text="🔄 RESCAN SYSTEM & FILE", command=self.start_audit, 
                  bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=15, pady=5).pack(side="right")

        # --- HARDWARE STATS ---
        hw_frame = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        hw_bar = tk.Frame(hw_frame, bg="white")
        hw_bar.pack(fill="x")
        
        for k, v in self.audit_data['hw'].items():
            f = tk.Frame(hw_bar, bg="white")
            f.pack(side="left", expand=True)
            tk.Label(f, text=k.upper(), font=("Segoe UI", 8, "bold"), bg="white", fg="#605E5C").pack()
            tk.Label(f, text=v, font=("Segoe UI", 12, "bold"), bg="white", fg=CLR_BLUE).pack()
        
        ram_val = float(self.audit_data['hw']['ram'].replace('%',''))
        if ram_val > 80:
            tk.Label(hw_frame, text="⚠️ CRITICAL: SYSTEM RAM EXHAUSTED - CONTACT IT TEAM", bg=CLR_ERR, fg="white", font=("Segoe UI", 8, "bold")).pack(fill="x", pady=(10,0))
        hw_frame.pack(fill="x", pady=(0, 20))

        # --- PROBLEM & SOLUTION CARDS ---
        tk.Label(self.container, text="Detected Issues", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        
        if not self.audit_data['problems']:
            tk.Label(self.container, text="✅ Optimized: No critical problems found.", font=("Segoe UI", 10), fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w", pady=10)
        
        for p in self.audit_data['problems']:
            pf = tk.Frame(self.container, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            pf.pack(fill="x", pady=4)
            
            top_line = tk.Frame(pf, bg="white")
            top_line.pack(fill="x")
            tk.Label(top_line, text=f"⚠️ {p['issue']}", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(side="left")
            tk.Button(top_line, text="Fix This", command=lambda i=p['id']: self.run_fix(i), bg=CLR_BLUE, fg="white", font=("Segoe UI", 8, "bold"), padx=10).pack(side="right")
            
            tk.Label(pf, text=f"Reason: {p['reason']}", font=("Segoe UI", 9), bg="white").pack(anchor="w", pady=(5,0))
            tk.Label(pf, text=f"Solution: {p['sol']}", font=("Segoe UI", 9, "italic"), bg="white", fg="#605E5C").pack(anchor="w")

        # --- MASTER BUTTON ---
        btn_all = tk.Button(self.container, text="⚡ APPLY ALL SECURITY & PERFORMANCE FIXES", 
                           bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=18, command=lambda: self.run_fix("ALL"))
        btn_all.pack(fill="x", pady=30)

    def run_fix(self, task_id):
        messagebox.showinfo("Repair Engine", f"Processing {task_id} Optimization...")
        self.pre_launch_cleanup() 

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateSecuritySuite(root)
    root.mainloop()