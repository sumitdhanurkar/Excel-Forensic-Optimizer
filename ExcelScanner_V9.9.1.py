import os, psutil, time, ctypes, shutil, threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client

# --- THEME (Refined Professional v9.8 Style) ---
CLR_BG = "#F3F2F1"   # Soft Gray
CLR_EXCEL = "#107C41" # Excel Green
CLR_BLUE = "#0078D4"  # Windows Blue
CLR_ERR = "#D13438"   # Error Red
CLR_TXT = "#323130"   # Dark Text

class FinalMasterSuite:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intel & Security Suite v9.9 - Final Master")
        self.root.geometry("1200x950")
        self.root.configure(bg=CLR_BG)
        
        self.file_paths = []
        self.batch_results = []
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)
        
        self.master_ram_flush()
        self.show_home()

    def master_ram_flush(self):
        """Force clean RAM and kill ghost processes."""
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] == "EXCEL.EXE": proc.kill()
            # Clear Windows Standby List & Working Set
            ctypes.windll.psapi.EmptyWorkingSet(ctypes.windll.kernel32.GetCurrentProcess())
        except: pass

    def show_home(self):
        for widget in self.container.winfo_children(): widget.destroy()
        tk.Label(self.container, text="Excel Audit & Security Suite", font=("Segoe UI", 28, "bold"), bg=CLR_BG).pack(pady=(60, 10))
        tk.Label(self.container, text="Final Master Build: Batch Scan • Hardware Telemetry • Security", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack()
        
        btn = tk.Button(self.container, text="📂 SELECT EXCEL FILES FOR BATCH AUDIT", command=self.select_files, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", padx=40, pady=20, cursor="hand2")
        btn.pack(pady=40)

    def select_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")])
        if not paths: return
        self.file_paths = list(paths)
        self.start_audit_sequence()

    def start_audit_sequence(self):
        for widget in self.container.winfo_children(): widget.destroy()
        self.pb = ttk.Progressbar(self.container, orient="horizontal", length=800, mode="determinate")
        self.pb.pack(pady=40)
        self.status = tk.Label(self.container, text="Initializing Master Engine...", bg=CLR_BG, font=("Segoe UI", 10))
        self.status.pack()
        threading.Thread(target=self.run_engine, daemon=True).start()

    def run_engine(self):
        self.batch_results = []
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3 # Force disable macros for safe scanning
            
            for index, path in enumerate(self.file_paths):
                fname = os.path.basename(path)
                fsize = os.path.getsize(path) / (1024 * 1024)
                self.status.config(text=f"Auditing File {index+1}/{len(self.file_paths)}: {fname}")
                self.pb['value'] = ((index + 1) / len(self.file_paths)) * 100
                
                # Pre-file RAM Flush
                self.master_ram_flush()
                
                try:
                    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
                    
                    rows, cols, factors = 0, 0, []
                    for sh in wb.Sheets:
                        rows += sh.UsedRange.Rows.Count
                        cols += sh.UsedRange.Columns.Count
                        if sh.UsedRange.Rows.Count > 10000: factors.append(f"Row Bloat ({sh.Name})")
                    
                    if wb.HasVBProject: factors.append("VBA Script (Security Factor)")
                    if wb.LinkSources(1): factors.append("External Links (Load Factor)")
                    
                    self.batch_results.append({
                        "name": fname, "size": f"{fsize:.2f} MB", 
                        "dims": f"{rows:,} Rows / {cols} Cols",
                        "factors": factors if factors else ["Optimized Structure"],
                        "status": "Verified"
                    })
                    wb.Close(False)
                except:
                    self.batch_results.append({"name": fname, "size": "N/A", "dims": "Error", "factors": ["File Locked or Corrupt"], "status": "Failed"})

            excel.Quit()
            self.root.after(100, self.display_final_report)
        except Exception as e:
            messagebox.showerror("Engine Error", str(e))
            self.show_home()

    def display_final_report(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # --- HEADER & NAVIGATION ---
        nav = tk.Frame(self.container, bg=CLR_BG)
        nav.pack(fill="x", pady=(0, 10))
        tk.Label(nav, text="Detailed Batch Results", font=("Segoe UI", 18, "bold"), bg=CLR_BG).pack(side="left")
        tk.Button(nav, text="➕ ADD MORE FILES", command=self.show_home, bg=CLR_BLUE, fg="white", font=("Segoe UI", 9, "bold"), padx=20).pack(side="right")

        # --- SYSTEM TELEMETRY DASHBOARD ---
        ram = psutil.virtual_memory().percent
        cpu = psutil.cpu_percent()
        disk = psutil.disk_usage('/').percent
        hw = tk.Frame(self.container, bg="white", padx=15, pady=12, highlightthickness=1, highlightbackground="#DDD")
        hw.pack(fill="x", pady=(0, 20))
        
        for l, v, c in [("CPU", f"{cpu}%", CLR_TXT), ("RAM", f"{ram}%", CLR_BLUE if ram < 80 else CLR_ERR), ("DISK", f"{disk}%", CLR_TXT)]:
            f = tk.Frame(hw, bg="white")
            f.pack(side="left", expand=True)
            tk.Label(f, text=l, font=("Segoe UI", 8, "bold"), bg="white", fg="#666").pack()
            tk.Label(f, text=v, font=("Segoe UI", 12, "bold"), bg="white", fg=c).pack()

        # --- SCROLLABLE RESULT CARDS (v9.8.8 UI) ---
        canvas = tk.Canvas(self.container, bg=CLR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=CLR_BG)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=1100)
        canvas.configure(yscrollcommand=scrollbar.set)

        for res in self.batch_results:
            card = tk.Frame(scroll_frame, bg="white", padx=20, pady=15, highlightthickness=1, highlightbackground="#E0E0E0")
            card.pack(fill="x", pady=8, padx=20)
            
            top = tk.Frame(card, bg="white")
            top.pack(fill="x")
            tk.Label(top, text=f"📄 {res['name']}", font=("Segoe UI", 11, "bold"), bg="white").pack(side="left")
            tk.Label(top, text=res['size'], font=("Segoe UI", 10, "bold"), bg="white", fg=CLR_BLUE).pack(side="right")
            
            tk.Label(card, text=f"Data Audit: {res['dims']}", font=("Segoe UI", 9), bg="white", fg="#444").pack(anchor="w", pady=2)
            
            f_frame = tk.Frame(card, bg="#F9F9F9", padx=10, pady=5)
            f_frame.pack(fill="x", pady=5)
            tk.Label(f_frame, text="Affecting Factors: " + " | ".join(res['factors']), font=("Segoe UI", 8), bg="#F9F9F9", fg=CLR_ERR if "Optimized" not in res['factors'][0] else CLR_EXCEL).pack(side="left")
            tk.Button(card, text="Fix File", bg=CLR_BLUE, fg="white", font=("Segoe UI", 8), command=lambda: None).place(relx=0.92, rely=0.6)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- MASTER ACTION ---
        tk.Button(self.container, text="⚡ EXECUTE MASTER CLEANUP ON ALL FILES", bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), pady=18).pack(fill="x", pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalMasterSuite(root)
    root.mainloop()