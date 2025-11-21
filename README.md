<h1 align="center">📊 OMK Attendance Processing Tool</h1>

<p align="center">
A complete automated attendance processing system built using <b>Python + Tkinter + Pandas</b>
</p>

---

## <p align="center">🚀 Status & Technology Badges</p>

<p align="center">

<img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" />
<img src="https://img.shields.io/badge/Tkinter-GUI-yellow?logo=windowsterminal" />
<img src="https://img.shields.io/badge/Pandas-Data%20Processing-green?logo=pandas" />
<img src="https://img.shields.io/badge/Platform-Windows-blue?logo=windows" />
<img src="https://img.shields.io/github/license/abhijeetraj22/OMK_ATTENDANCE_GUI" />
<img src="https://img.shields.io/github/stars/abhijeetraj22/OMK_ATTENDANCE_GUI?style=social" />
<img src="https://img.shields.io/github/issues/abhijeetraj22/OMK_ATTENDANCE_GUI" />

</p>

---

# 🎥 Live Demo (Animated GIF)

> Replace this GIF with your actual recorded screen:
> Save your GIF as: `docs/demo.gif`


---

# 🧩 Overview

The **OMK Attendance Processing Tool** provides an easy-to-use GUI workflow to automate attendance generation:

✔ Import Admin & Teacher attendance files  
✔ Auto-detect sheets & show as dropdown  
✔ Apply **grace-time rules**  
✔ Detect late arrival  
✔ Standardize time formats  
✔ Error handling  
✔ Export results to Excel or CSV  

---

# 🏗 UML Workflow Diagram

## 1️⃣ ASCII UML (text GitHub view)
            ┌────────────────────┐
            │  Start Application │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │  User Selects File │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │ Load Sheets + Show  │
            │      Dropdown       │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │ Configure Grace Time│
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │ Process Attendance  │
            │   (Pandas Logic)    │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │ Error Handling +    │
            │ Data Validation     │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │   Export Output     │
            │ (CSV / XLSX files)  │
            └─────────┬──────────┘
                      │
            ┌─────────▼──────────┐
            │      Complete       │
            └─────────────────────┘


---

# 🧰 Features

### 🔹 GUI (Tkinter)

* Browse Excel files
* Auto-detect & show sheet list
* Set grace time
* Run processing
* Save output

### 🔹 Attendance Processing

* Extract IN & OUT time
* Convert inconsistent formats
* Apply late-entry logic
* Identify incorrect or missing entries

### 🔹 Export Options

* `.csv` combined
* `.xlsx` multi-sheet

---

# 📁 Project Structure

```
OMK-Attendance/
│
├── attendance_tool.py          # Main script
├── /docs
│    ├── demo.gif               # Animation for README
│    ├── uml.png                # Optional diagram
│    └── logo.png               # App logo
├── requirements.txt
└── README.md
```

---

# 🛠 Installation

### 1️⃣ Clone the Repo

```bash
git clone https://github.com/abhijeetraj22/OMK_ATTENDANCE_GUI.git
cd OMK_ATTENDANCE_GUI
```

### 2️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

### 3️⃣ Run the Tool

```bash
python attendance_tool.py
```

---

# 🧪 Output Example

| Employee ID | Name         | In Time | Out Time |  Late  |   MoT   | Attend  |
| ----------- | ------------ | ------- | -------- |  ----  | ------- | ------- |
| 1024        | Rohan Sharma | 07:55   | 15:04    |   ADV  |   TRN   | Present |
| 1893        | Neha Verma   | 08:32   | 15:01    |   LATE |   SUP   | Present |

---

# 🧾 Grace Time Logic (Configurable)

```
Arrival < 07:40   → On Time  
07:40 – 08:00     → Adv  
>08:00            → Late
```

---

# 🤝 Contributing

1. Fork
2. Create branch
3. Commit
4. Open a PR

---

# 📜 License

MIT License

---

# ⭐ Support the Project

If you found this tool helpful:

👉 **Star ⭐ the repo**
👉 Share with team
👉 Follow on GitHub

---

## Want enhancements?

I can also generate:

✔ Installer (.exe)
✔ Packaging using PyInstaller
✔ Modern UI (CustomTkinter)
✔ Professional project logo PNG/SVG
✔ Flowchart PNG
✔ API integration
✔ Windows Task Scheduler automation

Just tell me!

>08:00            → Absent
--- # 🤝 Contributing 1. Fork 2. Create branch 3. Commit 4. Open a PR --- # 📜 License MIT License --- # ⭐ Support the Project If you found this tool helpful: 👉 **Star ⭐ the repo** 👉 Share with team 👉 Follow on GitHub ---
