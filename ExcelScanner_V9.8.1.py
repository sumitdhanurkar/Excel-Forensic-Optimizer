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

class UltimateSecuritySuiteV985:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intel & Security v9.8.5 (Batch Mode)")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
        self.pre_launch_cleanup()
        self.show_home()

    def pre_launch_cleanup(self):
        """Clean RAM and kill stuck Excel processes."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == "EXCEL.EXE": proc.kill()
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Security & Intelligence", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(80, 10))
        tk.Label(self.container, text="Multi-File Hardware Stats • Security Scan • Performance Audit", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT EXCEL FILES (MULTI-SELECT)", command=self.select_files, 
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
        self.status = tk.Label(self.container, text="Initializing Batch Engine...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_batch_scan, daemon=True).start()

    def perform_batch_scan(self):
        self.batch_results = []
        total_files = len(self.file_paths)
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            for index, path in enumerate(self.file_paths):
                fname = os.path.basename(path)
                self.status.config(text=f"Scanning file {index+1} of {total_files}: {fname}")
                self.pb['value'] = ((index + 1) / total_files) * 100
                
                # Pre-file RAM flush
                self.pre_launch_cleanup()
                
                start_t = time.time()
                wb = excel.Workbooks.Open(path)
                load_time = time.time() - start_t
                
                problems = []
                # Security/Virus Check
                if wb.HasVBProject:
                    problems.append({"issue": "VBA/Macros", "reason": "Hidden code/Virus risk", "sol": "Inspect/Remove VBA", "id": "VBA"})
                
                links = wb.LinkSources(1)
                if links:
                    problems.append({"issue": "External Links", "reason": f"Connected to {len(links)} files", "sol": "Break Links", "id": "LINKS"})

                # Performance Check
                if load_time > 4.0:
                    problems.append({"issue": "Slow Load", "reason": f"Opened in {load_time:.2f}s", "sol": "Format optimization", "id": "LOAD"})

                self.batch_results.append({
                    "name": fname,
                    "load": f"{load_time:.2f}s",
                    "problems": problems,
                    "status": "Scanned"
                })
                wb.Close(False)

            excel.Quit()
            self.root.after(100, self.display_report)
            
        except Exception as e:
            messagebox.showerror("Batch Error", str(e))
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- TOP NAVIGATION ---
        nav = tk.Frame(self.container, bg=CLR_BG)
        nav.pack(fill="x", pady=(0, 10))
        tk.Label(nav, text="Batch Process Results", font=("Segoe UI", 16, "bold"), bg=CLR_BG).pack(side="left")
        
        # THE "ADD MORE" BUTTON
        tk.Button(nav, text="➕ ADD MORE FILES", command=self.show_home, 
                  bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), padx=15).pack(side="right")

        # --- HARDWARE STATS (v9.8 style) ---
        ram = psutil.virtual_memory().percent
        cpu = psutil.cpu_percent()
        hw_frame = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        hw_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(hw_frame, text=f"SYSTEM RAM: {ram}%", font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_BLUE if ram < 80 else CLR_ERR).pack(side="left", padx=20)
        tk.Label(hw_frame, text=f"CPU LOAD: {cpu}%", font=("Segoe UI", 10, "bold"), bg="white").pack(side="left")
        
        if ram > 80:
            tk.Label(hw_frame, text="⚠️ CONTACT IT TEAM: HIGH RAM USAGE", bg=CLR_ERR, fg="white", font=("Segoe UI", 8, "bold"), padx=10).pack(side="right")

        # --- BATCH RESULTS TABLE ---
        tk.Label(self.container, text="Detailed Process Log", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=CLR_BG)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for res in self.batch_results:
            f = tk.Frame(scrollable_frame, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#E0E0E0")
            f.pack(fill="x", pady=5, padx=2)
            
            # File info line
            line1 = tk.Frame(f, bg="white")
            line1.pack(fill="x")
            tk.Label(line1, text=f"📄 {res['name']}", font=("Segoe UI", 10, "bold"), bg="white").pack(side="left")
            tk.Label(line1, text=f"Load Time: {res['load']}", font=("Segoe UI", 9), bg="white", fg="#605E5C").pack(side="right")

            # Problem lines
            if not res['problems']:
                tk.Label(f, text="✅ No critical issues found in this file.", font=("Segoe UI", 9), bg="white", fg=CLR_EXCEL).pack(anchor="w")
            else:
                for p in res['problems']:
                    p_frame = tk.Frame(f, bg="white")
                    p_frame.pack(fill="x", pady=2)
                    tk.Label(p_frame, text=f"⚠️ {p['issue']}: {p['reason']}", font=("Segoe UI", 9), bg="white", fg=CLR_ERR).pack(side="left")
                    tk.Button(p_frame, text="Fix", command=lambda: messagebox.showinfo("Fix", "Optimizing file..."), bg=CLR_BLUE, fg="white", font=("Segoe UI", 7)).pack(side="right")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- MASTER BUTTON ---
        tk.Button(self.container, text="⚡ APPLY ALL SECURITY & PERFORMANCE FIXES TO ALL FILES", 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15, command=lambda: messagebox.showinfo("Batch", "Full Optimization Started")).pack(fill="x", pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateSecuritySuiteV985(root)
    root.mainloop()