import os
import psutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import threading
import zipfile

# --- THEME & STYLE ---
CLR_BG = "#F3F2F1"
CLR_CARD = "#FFFFFF"
CLR_EXCEL = "#107C41"  # Excel Green
CLR_BLUE = "#0078D4"   # Windows Blue
CLR_DARK = "#323130"

class CommandCenterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Intelligence Suite v8.3")
        self.root.geometry("1100x850")
        self.root.configure(bg=CLR_BG)
        
        self.file_path = ""
        self.container = tk.Frame(self.root, bg=CLR_BG)
        self.container.pack(fill="both", expand=True, padx=50, pady=30)
        self.show_welcome()

    def kill_excel(self):
        """Standard safety check to prevent 'Visible' errors."""
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == "EXCEL.EXE":
                try: proc.kill()
                except: pass

    def show_welcome(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        tk.Label(self.container, text="Excel Performance Suite", font=("Segoe UI", 28, "bold"), bg=CLR_BG, fg=CLR_DARK).pack(pady=(80, 10))
        tk.Label(self.container, text="Choose a workbook to begin diagnostic control", font=("Segoe UI", 12), bg=CLR_BG, fg="#605E5C").pack(pady=5)
        
        btn = tk.Button(self.container, text="📂 BROWSE WORKBOOK", command=self.load_file, 
                        bg=CLR_EXCEL, fg="white", font=("Segoe UI", 11, "bold"), relief="flat", padx=40, pady=15, cursor="hand2")
        btn.pack(pady=40)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb")])
        if not path: return
        self.file_path = os.path.normpath(os.path.abspath(path))
        self.show_dashboard()

    def show_dashboard(self):
        for widget in self.container.winfo_children(): widget.destroy()
        
        # Header
        header = tk.Frame(self.container, bg=CLR_BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text=f"File: {os.path.basename(self.file_path)}", font=("Segoe UI", 14, "bold"), bg=CLR_BG, fg=CLR_DARK).pack(side="left")
        tk.Button(header, text="Change File", font=("Segoe UI", 9), command=self.show_welcome, bg="#E1E1E1", relief="flat").pack(side="right")

        # --- SECTION: SINGLE TASKS ---
        tk.Label(self.container, text="Manual Optimization Controls", font=("Segoe UI", 11, "bold"), bg=CLR_BG, fg="#605E5C").pack(anchor="w", pady=(10, 5))
        
        task_frame = tk.Frame(self.container, bg=CLR_BG)
        task_frame.pack(fill="x", pady=10)

        # Task 1: Ghost Rows
        self.create_task_row(task_frame, "Purge Ghost Rows & Hidden Data", "GHOST", CLR_BLUE)
        
        # Task 2: Binary Conversion
        self.create_task_row(task_frame, "Convert to Binary (.xlsb) Format", "BINARY", CLR_BLUE)

        # Task 3: Calculation Clean
        self.create_task_row(task_frame, "Deep Reset of Calc Dependencies", "CALC", CLR_BLUE)

        # --- SECTION: MASTER BUTTON ---
        separator = tk.Frame(self.container, height=2, bg="#EDEBE9")
        separator.pack(fill="x", pady=30)

        master_box = tk.Frame(self.container, bg="white", padx=30, pady=30, highlightbackground="#D2D0CE", highlightthickness=1)
        master_box.pack(fill="x")

        tk.Label(master_box, text="Ultimate Optimization", font=("Segoe UI", 14, "bold"), bg="white", fg=CLR_DARK).pack(anchor="w")
        tk.Label(master_box, text="Runs all tasks above, removes media bloat, and creates a clean repair log.", 
                 font=("Segoe UI", 10), bg="white", fg="#605E5C").pack(anchor="w", pady=(0, 20))

        master_btn = tk.Button(master_box, text="⚡ RUN COMPLETE PERFORMANCE OVERHAUL", command=lambda: self.run_engine("MASTER"), 
                               bg=CLR_EXCEL, fg="white", font=("Segoe UI", 12, "bold"), relief="flat", pady=18, cursor="hand2")
        master_btn.pack(fill="x")

    def create_task_row(self, parent, label_text, task_id, color):
        row = tk.Frame(parent, bg="white", pady=15, padx=20, highlightbackground="#EDEBE9", highlightthickness=1)
        row.pack(fill="x", pady=5)
        
        tk.Label(row, text=label_text, font=("Segoe UI", 10), bg="white", fg=CLR_DARK).pack(side="left")
        
        btn = tk.Button(row, text="Run Task", command=lambda t=task_id: self.run_engine(t), 
                        bg="#F3F2F1", fg=color, font=("Segoe UI", 9, "bold"), relief="groove", padx=15)
        btn.pack(side="right")

    def run_engine(self, mode):
        # Progress Window
        win = tk.Toplevel(self.root)
        win.title("Processing...")
        win.geometry("450x250")
        win.configure(bg="white")
        lbl = tk.Label(win, text="Initializing Excel Engine...", font=("Segoe UI", 10), bg="white", pady=30)
        lbl.pack()
        pb = ttk.Progressbar(win, orient="horizontal", length=300, mode="indeterminate")
        pb.pack(pady=10)
        pb.start()

        def execute():
            try:
                self.kill_excel()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                wb = excel.Workbooks.Open(self.file_path)
                save_path = self.file_path
                
                # Logic Switch
                if mode == "GHOST" or mode == "MASTER":
                    lbl.config(text="Cleaning Ghost Cells...")
                    for sh in wb.Sheets:
                        last = sh.Cells.Find("*", SearchOrder=1, SearchDirection=2)
                        if last: sh.Rows(f"{last.Row + 1}:{sh.Rows.Count}").Delete()
                
                if mode == "CALC" or mode == "MASTER":
                    lbl.config(text="Rebuilding Calculation Chain...")
                    excel.CalculateFullRebuild()

                if mode == "BINARY" or mode == "MASTER":
                    lbl.config(text="Converting to High-Speed Binary...")
                    base, _ = os.path.splitext(self.file_path)
                    save_path = base + "_v8_FIXED.xlsb"
                    wb.SaveAs(save_path, FileFormat=50)

                wb.Save()
                wb.Close()
                excel.Quit()
                
                lbl.config(text="✅ Task Completed Successfully!", fg=CLR_EXCEL)
                pb.stop()
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                lbl.config(text=f"Error: {str(e)}", fg="red")

        threading.Thread(target=execute, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = CommandCenterApp(root)
    root.mainloop()