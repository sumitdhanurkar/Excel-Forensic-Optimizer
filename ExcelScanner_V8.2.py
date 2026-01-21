import os
import psutil
import platform
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import zipfile

# --- THEME CONSTANTS ---
CLR_BG = "#F8F9FA"
CLR_PRI = "#107C41"  # Excel Green
CLR_ACC = "#2B579A"  # Deep Blue
CLR_TXT = "#323130"
CLR_CON = "#1B1B1B"

class FinalExcelOptimizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence Suite v8.2")
        self.root.geometry("1000x850")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.results = {}
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        self.show_home()

    def kill_excel_ghosts(self):
        """Clears stuck Excel processes to prevent 'Visible' property errors."""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try:
                    proc.kill()
                except:
                    pass

    def clear_ui(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_ui()
        tk.Label(self.container, text="Excel Intelligence Suite", font=("Segoe UI", 32, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(pady=(100, 10))
        tk.Label(self.container, text="Advanced Diagnostic & Binary Optimization Hub", font=("Segoe UI", 14), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 SELECT WORKBOOK", command=self.start_scan_process, 
                        bg=CLR_PRI, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=18, cursor="hand2")
        btn.pack(pady=50)

    def start_scan_process(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        
        self.clear_ui()
        tk.Label(self.container, text="Deep Audit in Progress...", font=("Segoe UI", 18, "bold"), bg=CLR_BG, fg=CLR_ACC).pack(pady=10)
        
        self.console = tk.Text(self.container, bg=CLR_CON, fg="#00FF41", font=("Consolas", 10), height=15, borderwidth=0, padx=15, pady=15)
        self.console.pack(fill="both", expand=True, pady=10)
        self.progress = ttk.Progressbar(self.container, orient="horizontal", mode="determinate", length=800)
        self.progress.pack(pady=20)
        
        threading.Thread(target=self.run_deep_audit, daemon=True).start()

    def log(self, msg):
        self.console.insert(tk.END, f" > {msg}\n")
        self.console.see(tk.END)
        self.root.update()

    def run_deep_audit(self):
        try:
            self.log("Pre-Check: Clearing stuck Excel processes...")
            self.kill_excel_ghosts() # Fix for the 'Visible' property error
            
            self.progress['value'] = 20
            self.log("Initializing Excel Engine...")
            excel = win32com.client.Dispatch("Excel.Application")
            
            # Critical sequence to avoid property errors:
            excel.DisplayAlerts = False
            excel.Visible = False 
            
            self.log(f"Opening {os.path.basename(self.file_path)}...")
            wb = excel.Workbooks.Open(self.file_path)
            
            # Logic Audit
            self.progress['value'] = 50
            start_calc = time.time()
            excel.CalculateFull()
            calc_time = time.time() - start_calc
            
            # Media Scan
            media_mb = 0
            if zipfile.is_zipfile(self.file_path):
                with zipfile.ZipFile(self.file_path, 'r') as z:
                    media_mb = sum(info.file_size for info in z.infolist() if 'media/' in info.filename) / (1024*1024)

            sheets = []
            for sh in wb.Sheets:
                r, c = sh.UsedRange.Rows.Count, sh.UsedRange.Columns.Count
                data = excel.WorksheetFunction.CountA(sh.Cells)
                sheets.append((sh.Name, r, c, data))

            self.results = {
                "size": os.path.getsize(self.file_path)/(1024*1024),
                "calc": f"{calc_time:.2f}s",
                "media": f"{media_mb:.1f} MB",
                "sheets": sheets,
                "vba": "Detected" if wb.HasVBProject else "None"
            }
            
            wb.Close(False)
            excel.Quit()
            self.root.after(500, self.show_results)
        except Exception as e:
            messagebox.showerror("Engine Error", f"System was unable to start Excel Engine.\nReason: {e}")
            self.show_home()

    def show_results(self):
        self.clear_ui()
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text="File Health Dashboard", font=("Segoe UI", 24, "bold"), bg=CLR_BG, fg=CLR_PRI).pack(side="left")
        
        # Stats Cards
        stats_frame = tk.Frame(self.container, bg=CLR_BG)
        stats_frame.pack(fill="x", pady=10)
        
        card_data = [
            ("Current Size", f"{self.results['size']:.2f} MB"),
            ("Calc Speed", self.results['calc']),
            ("Media Bloat", self.results['media']),
            ("VBA Status", self.results['vba'])
        ]
        
        for title, val in card_data:
            card = tk.Frame(stats_frame, bg="white", highlightbackground="#E0E0E0", highlightthickness=1, padx=20, pady=15)
            card.pack(side="left", fill="both", expand=True, padx=5)
            tk.Label(card, text=title, font=("Segoe UI", 10), bg="white", fg="#605E5C").pack()
            tk.Label(card, text=val, font=("Segoe UI", 14, "bold"), bg="white", fg=CLR_TXT).pack()

        # Action Hub
        hub = tk.LabelFrame(self.container, text=" Optimization Hub ", font=("Segoe UI", 12, "bold"), bg=CLR_BG, padx=20, pady=20)
        hub.pack(fill="x", pady=30)
        
        tk.Button(hub, text="🚀 FULL OPTIMIZATION (Binary + Cleanup)", command=lambda: self.execute_repair("ALL"), 
                  bg=CLR_PRI, fg="white", font=("Segoe UI", 11, "bold"), padx=30, pady=15, relief="flat").pack(fill="x")

    def execute_repair(self, r_type):
        rep_win = tk.Toplevel(self.root)
        rep_win.title("Repairing...")
        rep_win.geometry("400x250")
        rep_win.configure(bg="white")
        lbl = tk.Label(rep_win, text="Processing Intelligence Fixes...", font=("Segoe UI", 11), bg="white", pady=40)
        lbl.pack()
        
        def run():
            try:
                self.kill_excel_ghosts()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(self.file_path)
                
                # Binary Conversion
                base, _ = os.path.splitext(self.file_path)
                save_path = base + "_OPTIMIZED.xlsb"
                wb.SaveAs(save_path, FileFormat=50)

                # Ghost Row Purge
                for sh in wb.Sheets:
                    last_cell = sh.Cells.Find(What="*", SearchOrder=1, SearchDirection=2)
                    if last_cell:
                        sh.Rows(f"{last_cell.Row + 1}:{sh.Rows.Count}").Delete()
                
                wb.Save()
                new_size = os.path.getsize(save_path) / (1024*1024)
                wb.Close()
                excel.Quit()
                
                lbl.config(text=f"✅ Optimization Complete!\nNew Size: {new_size:.2f} MB", fg=CLR_PRI)
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {e}", fg="red")

        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalExcelOptimizer(root)
    root.mainloop()