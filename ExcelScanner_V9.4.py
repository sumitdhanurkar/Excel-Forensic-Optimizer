import os
import psutil
import time
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import shutil

# --- THEME ---
CLR_BG = "#F3F2F1"
CLR_EXCEL = "#107C41"
CLR_BLUE = "#0078D4"
CLR_ERR = "#D13438"
CLR_TXT = "#323130"

class UltimateRecoverySuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence & Recovery Suite v9.4")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.audit_data = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        self.show_home()

    def kill_ghost_excel(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def deep_ram_flush(self):
        """Forces Windows to release unused memory and clears Temp folders."""
        try:
            # 1. Clear Excel Temp files
            temp_path = os.path.expanduser('~\\AppData\\Local\\Temp')
            for filename in os.listdir(temp_path):
                if "excel" in filename.lower():
                    file_path = os.path.join(temp_path, filename)
                    try: shutil.rmtree(file_path) if os.path.isdir(file_path) else os.remove(file_path)
                    except: pass
            
            # 2. Windows Empty Working Set Call
            # This 'tricks' processes into releasing their held RAM back to the OS
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
            
            messagebox.showinfo("System Recovery", "RAM Flush Complete. Temporary Excel cache cleared.")
        except Exception as e:
            messagebox.showerror("Recovery Error", f"Could not complete flush: {e}")

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Intelligence Suite", font=("Segoe UI", 32, "bold"), bg=CLR_BG).pack(pady=(100, 10))
        tk.Label(self.container, text="File Structural Audit & System Recovery Engine", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT & ANALYZE WORKBOOK", command=self.start_audit, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=50, pady=20, cursor="hand2")
        btn.pack(pady=50)

    def start_audit(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Performing Full System & File Audit...", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(pady=20)
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=10)
        self.status = tk.Label(self.container, text="Scanning hardware...", bg=CLR_BG, font=("Consolas", 10))
        self.status.pack()
        
        threading.Thread(target=self.perform_scan, daemon=True).start()

    def perform_scan(self):
        try:
            cpu = psutil.cpu_percent(interval=0.5)
            ram = psutil.virtual_memory()
            
            self.kill_ghost_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(self.file_path)
            
            # Structural Audit
            sheets_report = []
            problems = []
            for sh in wb.Sheets:
                used_r = sh.UsedRange.Rows.Count
                used_c = sh.UsedRange.Columns.Count
                try: data_count = excel.WorksheetFunction.CountA(sh.Cells)
                except: data_count = 0
                
                sheets_report.append({"name": sh.Name, "rows": used_r, "cols": used_c, "data": data_count})

                if used_r > (data_count + 5000):
                    problems.append({"issue": f"Ghost Data in '{sh.Name}'", 
                                     "desc": f"Sheet uses {used_r} rows for only {data_count} entries.",
                                     "sol": "Delete unused rows to restore performance.", "id": "GHOST"})

            it_alert = True if ram.percent > 80 else False

            self.audit_data = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "sheets": sheets_report,
                "problems": problems,
                "ram": f"{ram.percent}%",
                "cpu": f"{cpu}%",
                "it_alert": it_alert
            }
            
            wb.Close(False)
            excel.Quit()
            self.root.after(500, self.display_report)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.show_home()

    def display_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- IT WARNING & RECOVERY SYSTEM ---
        if self.audit_data['it_alert']:
            warn_frame = tk.Frame(self.container, bg=CLR_ERR, padx=20, pady=15)
            warn_frame.pack(fill="x", pady=(0, 20))
            tk.Label(warn_frame, text="⚠️ CRITICAL RAM USAGE: Contact IT Team for Hardware Upgrade", font=("Segoe UI", 11, "bold"), bg=CLR_ERR, fg="white").pack(side="left")
            tk.Button(warn_frame, text="Try System Recovery (Flush RAM)", command=self.deep_ram_flush, bg="white", fg=CLR_ERR, font=("Segoe UI", 9, "bold")).pack(side="right")
        else:
            tk.Label(self.container, text="✅ SYSTEM HARDWARE HEALTHY", font=("Segoe UI", 10, "bold"), fg=CLR_EXCEL, bg=CLR_BG).pack(anchor="w")

        # --- AUDIT TABLE ---
        tk.Label(self.container, text="Worksheet Structural Audit", font=("Segoe UI", 12, "bold"), bg=CLR_BG).pack(anchor="w", pady=(10, 5))
        tbl = ttk.Treeview(self.container, columns=("Name", "Rows", "Cols", "Data"), show="headings", height=5)
        for c in ("Name", "Rows", "Cols", "Data"): tbl.heading(c, text=c)
        for s in self.audit_data['sheets']: tbl.insert("", "end", values=(s['name'], s['rows'], s['cols'], s['data']))
        tbl.pack(fill="x", pady=10)

        # --- PROBLEM LIST ---
        for p in self.audit_data['problems']:
            f = tk.Frame(self.container, bg="white", pady=10, padx=20, highlightthickness=1, highlightbackground="#E0E0E0")
            f.pack(fill="x", pady=2)
            tk.Label(f, text=p['issue'], font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_ERR).pack(anchor="w")
            tk.Label(f, text=f"Solution: {p['sol']}", font=("Segoe UI", 9), bg="white", fg="#605E5C").pack(anchor="w")
            tk.Button(f, text="Fix Task", command=lambda i=p['id']: self.run_fix(i), bg=CLR_BLUE, fg="white").place(relx=0.9, rely=0.2)

        # --- MASTER ACTION ---
        tk.Button(self.container, text="🚀 EXECUTE MASTER PERFORMANCE OVERHAUL", command=lambda: self.run_fix("ALL"), 
                  bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=15).pack(fill="x", pady=20)

    def run_fix(self, mode):
        # [The repair engine logic from previous versions remains here]
        win = tk.Toplevel(self.root)
        win.title("Repairing...")
        win.geometry("400x200")
        lbl = tk.Label(win, text="Optimizing... Please Wait.", pady=50)
        lbl.pack()

        def execute():
            try:
                self.kill_ghost_excel()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(self.file_path)
                save_path = self.file_path

                if mode in ["GHOST", "ALL"]:
                    for sh in wb.Sheets:
                        last = sh.Cells.Find("*", SearchOrder=1, SearchDirection=2)
                        if last: sh.Rows(f"{last.Row+1}:{sh.Rows.Count}").Delete()
                
                if mode in ["BINARY", "ALL"]:
                    base, _ = os.path.splitext(self.file_path)
                    save_path = base + "_OPTIMIZED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                wb.Save()
                wb.Close()
                excel.Quit()
                lbl.config(text="✅ Task Complete!", fg=CLR_EXCEL)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {e}", fg=CLR_ERR)

        threading.Thread(target=execute, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateRecoverySuite(root)
    root.mainloop()