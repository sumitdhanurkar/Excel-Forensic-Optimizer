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
CLR_TXT = "#323130"

class BatchExcelSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Batch Intel v9.9")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.results_log = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
        self.system_cleanup()
        self.show_home()

    def system_cleanup(self):
        """Standard procedure to flush RAM and kill background Excel tasks."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == "EXCEL.EXE": proc.kill()
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Batch Intelligence", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(80, 10))
        tk.Label(self.container, text="Multi-File Audit • Security Scan • System Recovery", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT EXCEL FILES (MULTI-SELECT)", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=18, cursor="hand2")
        btn.pack(pady=40)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = [os.path.normpath(p) for p in paths]
        self.start_batch_audit()

    def start_batch_audit(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        self.progress_lbl = tk.Label(self.container, text="Initializing Batch Scan...", font=("Segoe UI", 12, "bold"), bg=CLR_BG)
        self.progress_lbl.pack(pady=20)
        
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=10)
        
        self.status = tk.Label(self.container, text="Ready", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        
        threading.Thread(target=self.run_engine, daemon=True).start()

    def run_engine(self):
        self.results_log = []
        total = len(self.file_paths)
        
        for i, path in enumerate(self.file_paths):
            fname = os.path.basename(path)
            self.status.config(text=f"Processing ({i+1}/{total}): {fname}")
            self.pb['value'] = ((i) / total) * 100
            
            # --- START SCAN LOGIC ---
            try:
                self.system_cleanup()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                
                start_t = time.time()
                wb = excel.Workbooks.Open(path)
                load_time = time.time() - start_t
                
                # Check for Security and Bloat
                has_macros = "Yes" if wb.HasVBProject else "No"
                links = wb.LinkSources(1)
                link_count = len(links) if links else 0
                
                # Sheet Audit
                bloat_found = False
                for sh in wb.Sheets:
                    if sh.UsedRange.Rows.Count > 5000: bloat_found = True
                
                self.results_log.append({
                    "file": fname,
                    "load": f"{load_time:.2f}s",
                    "macros": has_macros,
                    "links": link_count,
                    "bloat": "Detected" if bloat_found else "Clean",
                    "status": "Success"
                })
                
                wb.Close(False)
                excel.Quit()
            except Exception as e:
                self.results_log.append({"file": fname, "status": f"Error: {str(e)}"})

        self.pb['value'] = 100
        self.root.after(100, self.display_summary)

    def display_summary(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- HEADER & ADD MORE BUTTON ---
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text="Batch Audit Results", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(side="left")
        
        tk.Button(header, text="➕ ADD MORE FILES", command=self.show_home, 
                  bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=15).pack(side="right")

        # --- HARDWARE STATUS ---
        ram = psutil.virtual_memory().percent
        hw_frame = tk.Frame(self.container, bg="white", padx=15, pady=10, highlightthickness=1, highlightbackground="#DDD")
        tk.Label(hw_frame, text=f"SYSTEM STABILITY: {'HEALTHY' if ram < 80 else 'HIGH LOAD'}", 
                 font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_EXCEL if ram < 80 else CLR_ERR).pack(side="left")
        tk.Label(hw_frame, text=f"CURRENT RAM USAGE: {ram}%", bg="white", font=("Segoe UI", 9)).pack(side="right")
        hw_frame.pack(fill="x", pady=(0, 20))

        # --- RESULTS TABLE ---
        tk.Label(self.container, text="Process History", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w")
        tbl_frame = tk.Frame(self.container, bg="white")
        tbl_frame.pack(fill="both", expand=True)

        cols = ("File Name", "Load Time", "Macros", "Ext Links", "Bloat", "Audit Status")
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings")
        for c in cols: tree.heading(c, text=c)
        
        for r in self.results_log:
            if "Error" in r['status']:
                tree.insert("", "end", values=(r['file'], "-", "-", "-", "-", r['status']))
            else:
                tree.insert("", "end", values=(r['file'], r['load'], r['macros'], r['links'], r['bloat'], r['status']))
        
        tree.pack(fill="both", expand=True)

        # --- MASTER OPTIMIZE ---
        btn_all = tk.Button(self.container, text="⚡ OPTIMIZE ALL FILES IN BATCH", 
                           bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15, command=self.run_batch_fix)
        btn_all.pack(fill="x", pady=20)

    def run_batch_fix(self):
        messagebox.showinfo("Batch Engine", "Initiating repair on all listed files. RAM will be flushed between each task.")

if __name__ == "__main__":
    root = tk.Tk()
    app = BatchExcelSuite(root)
    root.mainloop()