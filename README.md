# Excel-Forensic-Optimizer
The "Excel Forensic Pro" is a high-performance diagnostic utility designed to identify structural bloat, security risks, and performance bottlenecks within Microsoft Excel workbooks. It utilizes a background "COM (Component Object Model) Automation Engine" to perform deep-tissue scans without altering the original file data.<br><br>
These are all the version which is going to be developed having some bugs as well as error which we resolves with working with it...<br><br>
we are working on it and make it better day by day but at the point this is the final executable file which you can use it...<br><br>
If you want python file which you want to use it and develop according to your own need then you may use Excel Forensic Optimizer v10.8.py where you may find all the code....<br><br>
The final Executable file will be in the dist/Excel Forensic Optimizer v10.8.exe<br><br>
Feel Free to contact me for any query at : sumitdhanurkar@gmail.com<br><br>


To finalize our project, here is the professional **README.md** documentation. This is designed to be kept in the same folder as your software to explain the technical methodology, hardware requirements, and optimization logic to your IT department or future developers.

---

## 🛠 Excel Forensic Pro v10.8 - Technical Documentation

### 1. Executive Overview

The **Excel Forensic Pro** is a high-performance diagnostic utility designed to identify structural bloat, security risks, and performance bottlenecks within Microsoft Excel workbooks. It utilizes a background **COM (Component Object Model) Automation Engine** to perform deep-tissue scans without altering the original file data.

### 2. Core Diagnostic Logic

The software follows a 5-step forensic protocol for every file in the batch:

1. **Engine Reset:** Force-terminates any "zombie" `EXCEL.EXE` processes to ensure memory is clean.
2. **Authentication Layer:** Detects encryption and triggers an interactive password request if required.
3. **Metadata Audit:** Scans for hidden VBA projects, macro signatures, and external data connections.
4. **Structural Audit:** Compares the `UsedRange` (what Excel thinks is the data) against the actual filled cells to detect "Phantom Data" bloat.
5. **Formula Density Scan:** Identifies volatile functions (`OFFSET`, `INDIRECT`) that trigger constant CPU recalculation lag.

---

### 3. Hardware Telemetry & Performance Standards

The application monitors system health to ensure that the audit process does not exceed the physical limits of the host machine.

| Metric | Threshold | Impact on Excel |
| --- | --- | --- |
| **CPU Usage** | >80% | Calculation lag; "Not Responding" errors during file open. |
| **RAM Usage** | >85% | Excel may fail to initialize the COM interface. |
| **Disk Usage** | >95% | Temporary swap files cannot be created, leading to save-failure. |

**Compliance Standard:** A file is marked as "Fully Compliant" only if it passes all internal checks for zero phantom rows, no volatile formulas, and zero external network dependencies.

---

### 4. Interpretation of Affecting Factors

* **Phantom Data:** Occurs when formatting or deleted data remains in the XML background. **Solution:** Reset the UsedRange and save.
* **Volatile Lag:** Formulas like `=OFFSET()` force a full workbook recalculation on every edit. **Solution:** Replace with `=INDEX()`.
* **Pivot Cache Bloat:** Pivot tables saving source data internally. **Solution:** Uncheck "Save source data with file" in Pivot Options.
* **VBA Metadata:** Indicates the presence of macros. This is flagged for security review to prevent macro-based malware.

---

### 5. Deployment Instructions

To compile the source code into a standalone Enterprise Executable (`.exe`), use the following PyInstaller configuration:

```bash
python -m PyInstaller --noconsole --onefile --uac-admin --icon=app_icon.ico --name "Forensic_Pro_v10.8" main.py

```

**Note:** Ensure `win32com`, `psutil`, and `pythoncom` libraries are installed in the build environment.

---
