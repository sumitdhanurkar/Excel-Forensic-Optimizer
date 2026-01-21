import os
import psutil
import platform
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading

# --- THEME CONSTANTS ---
CLR_BG = "#F3F2F1"
CLR_PRI = "#107C41"  # Excel Green
CLR_ACC = "#2B579A"  # Deep Blue
CLR_TXT = "#323130"
CLR_CON = "#1B1B1B"  # Console Dark

class ExcelProOptimizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Hardware & Logic Engine v6.0")
        self.root.geometry("950x800")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.results = {}
        
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=30, pady=20)
        
        self.show_home()

    def clear_ui(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_ui()
        tk.Label(self.container, text="Excel Performance Suite", font=("Segoe UI", 28, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(pady=(80, 10))
        tk.Label(self.container, text="Hardware Diagnostic & One-Click Remediation Hub", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT WORKBOOK", command=self.start_scan_process, 
                        bg=CLR_PRI, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=15, cursor="hand2")
        btn.pack(pady=50)

    def log_to_console(self, msg):
        self.console.insert(tk.END, f" > {msg}\n")
        self.console.see(tk.END)
        self.root.update()

    def start_scan_process(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        # Normalize path immediately to prevent 'Path not found' errors
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        self.clear_ui()
        tk.Label(self.container, text="Running Deep Audit...", font=("Segoe UI", 18, "bold"), bg=CLR_BG, fg=CLR_ACC).pack(pady=10)
        self.console = tk.Text(self.container, bg=CLR_CON, fg="#00FF41", font=("Consolas", 10), height=20, borderwidth=0, padx=15, pady=15)
        self.console.pack(fill="both", expand=True, pady=10)
        self.progress = ttk.Progressbar(self.container, orient="horizontal", mode="determinate", length=700)
        self.progress.pack(pady=20)
        
        threading.Thread(target=self.run_full_audit, daemon=True).start()

    def run_full_audit(self):
        try:
            self.progress['value'] = 20
            self.log_to_console("Step 1: Analyzing System Hardware...")
            sys_info = {"cpu": platform.processor(), "ram": f"{psutil.virtual_memory().percent}% Used", "cores": psutil.cpu_count()}
            
            self.progress['value'] = 40
            self.log_to_console("Step 2: Connecting to Excel Background Engine...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            self.log_to_console(f"Step 3: Opening File Safely...")
            wb = excel.Workbooks.Open(self.file_path)
            orig_size = os.path.getsize(self.file_path) / (1024*1024)
            
            self.progress['value'] = 70
            self.log_to_console("Step 4: Mapping Rows, Columns, and Data points...")
            sheet_data = []
            issues = []
            for sh in wb.Sheets:
                r, c = sh.UsedRange.Rows.Count, sh.UsedRange.Columns.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                sheet_data.append((sh.Name, r, c, data))
                if r > (data + 1000): issues.append(("Ghost Rows", sh.Name, "Cleanup Ghost Rows"))

            if wb.HasVBProject: issues.append(("VBA Detected", "Project", "Optimize Macros"))
            if not self.file_path.lower().endswith(".xlsb"): issues.append(("XML Bloat", "File", "Convert to Binary"))

            self.results = {"sheets": sheet_data, "issues": issues, "sys": sys_info, "size": orig_size}
            self.progress['value'] = 100
            
            wb.Close(False)
            excel.Quit()
            self.root.after(500, self.show_results)

        except Exception as e:
            messagebox.showerror("Path/Access Error", f"Excel could not access the file.\n\nReason: {str(e)}")
            self.show_home()

    def show_results(self):
        self.clear_ui()
        # Summary Header
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=10)
        tk.Label(header, text="Scan Dashboard", font=("Segoe UI", 20, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(side="left")
        tk.Label(header, text=f"Initial Size: {self.results['size']:.2f} MB", font=("Segoe UI", 10), bg=CLR_BG).pack(side="right")

        # Treeview Table
        tbl_frame = tk.Frame(self.container)
        tbl_frame.pack(fill="x", pady=10)
        cols = ("Sheet Name", "Used Rows", "Used Cols", "Data Cells")
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=6)
        for c in cols: tree.heading(c, text=c)
        for s in self.results['sheets']: tree.insert("", "end", values=s)
        tree.pack(fill="x")

        # Repair Section
        hub = tk.LabelFrame(self.container, text=" Optimization & Repair Tools ", font=("Segoe UI", 11, "bold"), bg=CLR_BG, padx=20, pady=20)
        hub.pack(fill="x", pady=10)

        for issue, loc, action in self.results['issues']:
            f = tk.Frame(hub, bg=CLR_BG)
            f.pack(fill="x", pady=3)
            tk.Label(f, text=f"• {issue} in {loc}", bg=CLR_BG).pack(side="left")
            tk.Button(f, text=action, command=lambda a=action: self.execute_repair(a), bg=CLR_ACC, fg="white", width=18, font=("Arial", 8)).pack(side="right")

        tk.Button(self.container, text="🚀 MASTER REPAIR (Run All Fixes & Save as Binary)", 
                  command=lambda: self.execute_repair("ALL"), bg=CLR_PRI, fg="white", font=("Segoe UI", 12, "bold"), pady=15).pack(fill="x", pady=10)

    def execute_repair(self, repair_type):
        rep_win = tk.Toplevel(self.root)
        rep_win.title("Repairing...")
        rep_win.geometry("450x250")
        rep_win.configure(bg=CLR_CON)
        lbl = tk.Label(rep_win, text="Initializing Engine...", fg="#00FF41", bg=CLR_CON, font=("Consolas", 10), wraplength=400)
        lbl.pack(pady=40)

        def repair_logic():
            try:
                # Re-normalize path for background worker
                clean_path = os.path.normpath(os.path.abspath(self.file_path))
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                lbl.config(text=f"Opening: {os.path.basename(clean_path)}")
                wb = excel.Workbooks.Open(clean_path)
                save_path = clean_path

                if repair_type in ["Convert to Binary", "ALL"]:
                    lbl.config(text="Status: Converting to High-Performance Binary (.xlsb)...")
                    base, _ = os.path.splitext(clean_path)
                    save_path = base + "_OPTIMIZED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                if repair_type in ["Cleanup Ghost Rows", "ALL"]:
                    lbl.config(text="Status: Purging Ghost Rows & Empty Formatting...")
                    for sh in wb.Sheets:
                        last_cell = sh.Cells.Find(What="*", SearchOrder=1, SearchDirection=2)
                        if last_cell:
                            sh.Rows(f"{last_cell.Row + 1}:{sh.Rows.Count}").Delete()
                    wb.Save()

                new_size = os.path.getsize(save_path) / (1024*1024)
                wb.Close()
                excel.Quit()
                
                lbl.config(text=f"✅ SUCCESS!\nOriginal: {self.results['size']:.2f} MB\nNew Size: {new_size:.2f} MB", fg="#00FF41")
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"❌ FAILED: {str(e)}", fg="red")

        threading.Thread(target=repair_logic, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    app = ExcelProOptimizer(root)
    root.mainloop()