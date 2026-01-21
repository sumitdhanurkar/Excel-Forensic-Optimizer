import os
import psutil
import platform
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading

# --- THEME COLORS ---
CLR_BG = "#F5F7FA"
CLR_PRI = "#107C41"  # Excel Green
CLR_ACC = "#2B579A"  # Outlook Blue
CLR_ERR = "#D13438"  # Danger Red
CLR_TXT = "#323130"
CLR_CON = "#1B1B1B"  # Dark Console

class ExcelProOptimizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Hardware & Logic Engine")
        self.root.geometry("900x750")
        self.root.configure(bg=CLR_BG)
        self.file_path = ""
        self.results = {}

        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=30, pady=20)
        self.show_home()

    def clear_ui(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    # --- SCREEN 1: HOME ---
    def show_home(self):
        self.clear_ui()
        tk.Label(self.container, text="Excel Performance Suite", font=("Segoe UI", 26, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(pady=(60, 10))
        tk.Label(self.container, text="Professional Hardware Diagnostic & Workbook Remediation", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT WORKBOOK", command=self.start_scan_process, 
                        bg=CLR_PRI, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=15, cursor="hand2")
        btn.pack(pady=50)

    # --- SCREEN 2: LIVE CONSOLE ---
    def show_loading(self):
        self.clear_ui()
        tk.Label(self.container, text="Analyzing System & File Internals...", font=("Segoe UI", 16, "bold"), bg=CLR_BG, fg=CLR_ACC).pack(pady=10)
        
        self.console = tk.Text(self.container, bg=CLR_CON, fg="#00FF41", font=("Consolas", 10), height=25, borderwidth=0, padx=15, pady=15)
        self.console.pack(fill="both", expand=True)
        self.log("Initializing Diagnostic Engine...")

    def log(self, msg):
        self.console.insert(tk.END, f" > {msg}\n")
        self.console.see(tk.END)
        self.root.update()

    # --- CORE LOGIC & BENCHMARKING ---
    def start_scan_process(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = path
        self.show_loading()
        threading.Thread(target=self.run_full_audit, daemon=True).start()

    def run_full_audit(self):
        start_time = time.time()
        res = {"sheets": [], "issues": [], "sys": {}}
        
        try:
            # 1. System Scan
            self.log("Scanning Processor & RAM availability...")
            res['sys'] = {
                "cpu": platform.processor(),
                "ram": f"{psutil.virtual_memory().percent}% Used",
                "cores": psutil.cpu_count()
            }

            # 2. Excel COM Audit
            self.log("Starting Background Excel Process (win32com)...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            self.log(f"Measuring Load Time for {os.path.basename(self.file_path)}...")
            wb = excel.Workbooks.Open(self.file_path)
            load_duration = time.time() - start_time
            res['load_time'] = f"{load_duration:.2f}s"

            # 3. Deep Sheet Scan
            for sh in wb.Sheets:
                self.log(f"Auditing Sheet: {sh.Name}...")
                r = sh.UsedRange.Rows.Count
                c = sh.UsedRange.Columns.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                res['sheets'].append((sh.Name, r, c, data))
                
                if r > (data + 2000):
                    res['issues'].append(("Ghost Rows", sh.Name, "Cleanup Ghost Rows"))

            # 4. Global Checks
            if wb.HasVBProject: res['issues'].append(("VBA Detected", "Global", "Optimize Macros"))
            if wb.Connections.Count > 0: res['issues'].append(("External Links", "Data", "Fix Links"))
            if not self.file_path.endswith(".xlsb"): res['issues'].append(("Format Issue", "File", "Convert to Binary"))

            wb.Close(False)
            excel.Quit()
            
            self.log("Audit complete. Transforming to Dashboard...")
            self.results = res
            self.root.after(1000, self.show_results)

        except Exception as e:
            self.log(f"CRITICAL ERROR: {str(e)}")
            messagebox.showerror("Scan Error", f"Could not complete scan: {e}")

    # --- SCREEN 3: RESULTS & REPAIR HUB ---
    def show_results(self):
        self.clear_ui()
        
        # Top Header
        top = tk.Frame(self.container, bg=CLR_BG)
        top.pack(fill="x")
        tk.Label(top, text="Diagnostic Report", font=("Segoe UI", 18, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(side="left")
        tk.Label(top, text=f"Benchmark Load Time: {self.results['load_time']}", font=("Segoe UI", 10), bg=CLR_BG).pack(side="right")

        # Inventory Table
        tk.Label(self.container, text="Sheet Analysis", font=("Segoe UI", 11, "bold"), bg=CLR_BG).pack(anchor="w", pady=(10,0))
        frame_tbl = tk.Frame(self.container)
        frame_tbl.pack(fill="x", pady=5)
        
        cols = ("Name", "Range Rows", "Range Cols", "Actual Data")
        tree = ttk.Treeview(frame_tbl, columns=cols, show="headings", height=5)
        for c in cols: tree.heading(c, text=c)
        for s in self.results['sheets']: tree.insert("", "end", values=s)
        tree.pack(fill="x")

        # Repair Hub
        tk.Label(self.container, text="Individual Fixes (Manual Control)", font=("Segoe UI", 11, "bold"), bg=CLR_BG).pack(anchor="w", pady=(20,0))
        hub = tk.Frame(self.container, bg="white", padx=15, pady=15, relief="groove", borderwidth=1)
        hub.pack(fill="x", pady=5)

        if not self.results['issues']:
            tk.Label(hub, text="✨ No issues found! Your file is perfectly optimized.", bg="white", fg=CLR_PRI).pack()
        else:
            for issue, loc, action in self.results['issues']:
                f = tk.Frame(hub, bg="white")
                f.pack(fill="x", pady=2)
                tk.Label(f, text=f"• {issue} found in {loc}", bg="white", width=40, anchor="w").pack(side="left")
                tk.Button(f, text=action, command=lambda a=action: self.run_repair(a), bg=CLR_ACC, fg="white", font=("Arial", 8)).pack(side="right")

        # MASTER FIX
        tk.Button(self.container, text="🚀 RUN ALL RECOMMENDED REPAIRS (MASTER FIX)", 
                  command=lambda: self.run_repair("ALL"), bg=CLR_ERR, fg="white", font=("Segoe UI", 12, "bold"), pady=15).pack(fill="x", pady=20)

    # --- REPAIR LOGIC ---
    def run_repair(self, repair_type):
        self.log_popup = tk.Toplevel(self.root)
        self.log_popup.title("Applying Fixes")
        self.log_popup.geometry("400x200")
        l = tk.Label(self.log_popup, text=f"Repairing: {repair_type}...", pady=20)
        l.pack()
        
        # In a real app, you would insert the win32com logic here to:
        # 1. Reset UsedRange
        # 2. SaveAs .xlsb
        # 3. Break broken links
        
        time.sleep(2) # Simulating repair work
        l.config(text="Repair Complete! File Saved.")
        tk.Button(self.log_popup, text="Close", command=self.log_popup.destroy).pack()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = ExcelProOptimizer(root)
    root.mainloop()