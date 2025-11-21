"""
OMK Attendance Tool — Final fixed single module (robust ID matching)

- Normalizes IDs and uses set-based matching
- Detects teacher/admin ID columns (id / UPID / id-like)
- Copies Route / TRANS MODE columns when present
- Outputs required columns: ID, NAME, ROLE, FIRST_SCAN, LAST_SCAN, MoT, Attendance, LATE_STATUS
"""

import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
import warnings
from typing import List, Optional

warnings.filterwarnings("ignore", message="Cannot parse header or footer so it will be ignored")

# ----------------- Configuration / Regex -----------------
SALUTATIONS_PATTERN = re.compile(r'^\s*(Mr|Mrs|Ms|Dr|Miss)\.?\s+', flags=re.IGNORECASE)
DEFAULT_GRACE_START = "07:40:00"
DEFAULT_GRACE_END = "08:00:00"

# ----------------- Helper functions -----------------
def normalize_id_value(v) -> str:
    """Convert various ID formats to a trimmed string form for reliable matching."""
    if pd.isna(v):
        return ""
    s = str(v).strip()
    # drop trailing .0 for floats represented as strings like "123.0"
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    # remove accidental spaces and leading zeros? (we keep leading zeros)
    return s

def strip_salutation(name: Optional[str]) -> Optional[str]:
    if pd.isna(name):
        return name
    s = str(name).strip()
    return re.sub(SALUTATIONS_PATTERN, "", s).strip()

def clean_role_text(text: Optional[str]) -> Optional[str]:
    if pd.isna(text):
        return text
    t = str(text).strip()
    m = re.search(r"\((.*?)\)", t)
    if m:
        return m.group(1).strip()
    return t.replace("Default", "").strip()

def detect_id_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Return actual column name from df matching any candidate (case-insensitive); else None."""
    cols = df.columns.tolist()
    lower_map = {c.lower().replace(" ", ""): c for c in cols}
    for cand in candidates:
        key = cand.lower().replace(" ", "")
        if key in lower_map:
            return lower_map[key]
    return None

def classify_role(text: Optional[str]) -> Optional[str]:
    if pd.isna(text):
        return text
    t = clean_role_text(text).upper()

    mapping_checks = [
        ("HEADMISTRESS", "HEADMISTRESS"),
        ("VICE PRESIDENT", "VICE PRESIDENT"),
        ("SCHOOL COORDINATOR", "SCHOOL COORDINATOR"),
        ("VICE PRINCIPAL", "VICE PRINCIPAL"),
        ("PRINCIPAL", "PRINCIPAL"),
        ("TEACHER", "TEACHER"),
        (["ADMINISTRATOR", "FRONT OFFICE", "ACCOUNTANT", "ACCOUNT"], "ADMIN"),
        ("DIRECTOR", "DIRECTOR"),
        ("HELPER", "HELPER"),
        ("DRIVER", "DRIVER"),
        ("NURSE", "NURSE"),
        ("MAID", "MAID"),
        (["PERSONAL SECURITY", "GUARD"], "GUARD"),
        ("GROUND STAFF", "GROUND STAFF"),
        ("PEON", "PEON"),
        ("CLEANER", "CLEANER"),
        ("MAINTENANCE", "MAINTENANCE"),
        ("OTHER", "OTHER"),
    ]

    for key, val in mapping_checks:
        if isinstance(key, list):
            if any(k in t for k in key):
                return val
        else:
            if key in t:
                return val
    return t.strip()

def parse_time_safe(s: Optional[str], fallback: datetime.time) -> datetime.time:
    if pd.isna(s) or str(s).strip() in ("", "--"):
        return fallback
    s = str(s).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.datetime.strptime(s, fmt).time()
        except Exception:
            continue
    return fallback

def parse_scan_to_time_or_none(s: Optional[str]) -> Optional[datetime.time]:
    if pd.isna(s):
        return None
    s = str(s).strip()
    if s == "" or s == "--":
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.datetime.strptime(s, fmt).time()
        except Exception:
            continue
    return None

def compute_late_attendance_for_df_with_grace(df: pd.DataFrame, grace_start_str: str, grace_end_str: str) -> pd.DataFrame:
    gs = parse_time_safe(grace_start_str, datetime.time(7,40,0))
    ge = parse_time_safe(grace_end_str, datetime.time(8,0,0))
    if gs > ge:
        gs, ge = ge, gs
    out = df.copy()
    if "FIRST_SCAN" in out.columns:
        out["Attendance"] = out["FIRST_SCAN"].apply(lambda x: "AB" if pd.isna(x) or str(x).strip() in ("", "--") else "PR")
        def status(x):
            t = parse_scan_to_time_or_none(x)
            if t is None:
                return ""
            if t <= gs:
                return ""
            elif gs < t <= ge:
                return "GRACE"
            else:
                return "LATE"
        out["LATE_STATUS"] = out["FIRST_SCAN"].apply(status)
    else:
        out["Attendance"] = ""
        out["LATE_STATUS"] = ""
    return out

def format_scan_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in ("FIRST_SCAN", "LAST_SCAN"):
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%H:%M:%S")
            out[col] = out[col].fillna("--")
    return out

def read_ids_from_sheet(path: str, sheet_name: str) -> List[str]:
    """
    Read first ID-like column as a list of normalized strings.
    """
    xl = pd.ExcelFile(path)
    if sheet_name not in xl.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' not found in {path}. Available: {xl.sheet_names}")
    df = pd.read_excel(path, sheet_name=sheet_name, header=0)
    # find ID-like column
    cid = detect_id_column(df, ["id", "upid", "empid", "uid"])
    if cid:
        return [normalize_id_value(x) for x in df[cid].dropna().astype(str).tolist()]
    # fallback to first column
    return [normalize_id_value(x) for x in df.iloc[:,0].dropna().astype(str).tolist()]

# ----------------- Main Single-Module GUI -----------------
class OMKAttendanceSingleModule(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OMK Attendance Tool - Fixed ID Matching")
        self.geometry_center(1050, 640)
        self.resizable(False, False)

        # state
        self.raw_path: Optional[str] = None
        self.teacher_path: Optional[str] = None
        self.admin_path: Optional[str] = None
        self.processed_df: Optional[pd.DataFrame] = None

        self.teacher_sheets: List[str] = []
        self.admin_sheets: List[str] = []

        self.build_ui()

    def geometry_center(self, w, h):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def build_ui(self):
        padx = 8; pady = 6
        top = ttk.Frame(self, padding=10)
        top.pack(fill="both", expand=True)

        ttk.Label(top, text="OMK Attendance — Single Module (Fixed ID matching)", font=("Verdana", 14, "bold")).pack(pady=6)

        # Raw file browse
        row_raw = ttk.Frame(top); row_raw.pack(fill="x", pady=4)
        ttk.Label(row_raw, text="Raw Excel (attendance):", width=22).pack(side="left")
        self.raw_entry = ttk.Entry(row_raw, width=72); self.raw_entry.pack(side="left", padx=6)
        ttk.Button(row_raw, text="Browse RAW", command=self.browse_raw_file).pack(side="left", padx=4)
        ttk.Button(row_raw, text="Load Columns", command=self.load_raw_columns).pack(side="left", padx=4)

        # Name column selection
        row_name = ttk.Frame(top); row_name.pack(fill="x", pady=4)
        ttk.Label(row_name, text="NAME Column to Clean:", width=22).pack(side="left")
        self.name_col_var = tk.StringVar()
        self.name_col_cb = ttk.Combobox(row_name, textvariable=self.name_col_var, width=40, state="readonly")
        self.name_col_cb.pack(side="left", padx=6)

        # Teacher browse + sheet dropdown
        row_tr = ttk.Frame(top); row_tr.pack(fill="x", pady=4)
        ttk.Label(row_tr, text="Teacher Excel (MoT):", width=22).pack(side="left")
        self.teacher_entry = ttk.Entry(row_tr, width=62); self.teacher_entry.pack(side="left", padx=6)
        ttk.Button(row_tr, text="Browse Teacher", command=self.browse_teacher_file).pack(side="left", padx=4)
        ttk.Label(row_tr, text="Sheet:", width=6).pack(side="left", padx=6)
        self.teacher_sheet_cb = ttk.Combobox(row_tr, width=22, state="readonly"); self.teacher_sheet_cb.pack(side="left", padx=4)

        # Admin browse + sheet dropdown
        row_ad = ttk.Frame(top); row_ad.pack(fill="x", pady=4)
        ttk.Label(row_ad, text="Admin Excel (MoT):", width=22).pack(side="left")
        self.admin_entry = ttk.Entry(row_ad, width=62); self.admin_entry.pack(side="left", padx=6)
        ttk.Button(row_ad, text="Browse Admin", command=self.browse_admin_file).pack(side="left", padx=4)
        ttk.Label(row_ad, text="Sheet:", width=6).pack(side="left", padx=6)
        self.admin_sheet_cb = ttk.Combobox(row_ad, width=22, state="readonly"); self.admin_sheet_cb.pack(side="left", padx=4)

        # Grace times
        row_grace = ttk.Frame(top); row_grace.pack(fill="x", pady=8)
        ttk.Label(row_grace, text="Grace Start (ON-TIME until):", width=30).pack(side="left")
        self.grace_start_var = tk.StringVar(value=DEFAULT_GRACE_START)
        ttk.Entry(row_grace, textvariable=self.grace_start_var, width=12).pack(side="left", padx=6)
        ttk.Label(row_grace, text="Grace End (GRACE until):", width=22).pack(side="left")
        self.grace_end_var = tk.StringVar(value=DEFAULT_GRACE_END)
        ttk.Entry(row_grace, textvariable=self.grace_end_var, width=12).pack(side="left", padx=6)

        # Buttons row (Run Process and Save As)
        row_btns = ttk.Frame(top); row_btns.pack(fill="x", pady=10)
        ttk.Button(row_btns, text="RUN PROCESS", command=self.run_full_process, width=20).pack(side="left", padx=12)
        ttk.Button(row_btns, text="SAVE AS (XLSX / CSV)", command=self.save_as).pack(side="left", padx=12)
        ttk.Button(row_btns, text="SAVE FINAL (Auto-split & Export)", command=self.save_final_now).pack(side="left", padx=12)

        # Status / info area
        row_info = ttk.Frame(top); row_info.pack(fill="both", expand=True, pady=8)
        self.status_text = tk.Text(row_info, height=12, state="disabled", wrap="word")
        self.status_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Footer tips
        tips = ("Workflow: 1) Browse RAW and Load Columns  2) Select NAME column  3) Upload Teacher/Admin and choose sheet  "
                "4) Click RUN PROCESS  5) If successful, click SAVE AS to store output (XLSX/CSV).")
        ttk.Label(top, text=tips, foreground="gray").pack(fill="x", pady=4)

    # ---------------- UI Helpers ----------------
    def append_status(self, msg: str):
        self.status_text.configure(state="normal")
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_text.insert("end", f"[{ts}] {msg}\n")
        self.status_text.see("end")
        self.status_text.configure(state="disabled")

    # ---------------- File browse handlers ----------------
    def browse_raw_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if p:
            self.raw_path = p
            self.raw_entry.delete(0, tk.END); self.raw_entry.insert(0, p)
            self.append_status(f"Selected RAW file: {p}")

    def load_raw_columns(self):
        p = self.raw_entry.get().strip()
        if not p:
            messagebox.showerror("Error", "Select RAW Excel file first.")
            return
        try:
            df0 = pd.read_excel(p, nrows=0)
            cols = list(df0.columns)
            self.name_col_cb["values"] = cols
            if "NAME" in cols:
                self.name_col_var.set("NAME")
            elif cols:
                self.name_col_var.set(cols[0])
            self.append_status(f"Loaded columns from RAW: {cols}")
            messagebox.showinfo("Columns loaded", "Columns loaded into dropdown. Choose the NAME column to clean.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read columns:\n{e}")
            self.append_status(f"Error reading RAW columns: {e}")

    def browse_teacher_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if p:
            self.teacher_path = p
            self.teacher_entry.delete(0, tk.END); self.teacher_entry.insert(0, p)
            try:
                xl = pd.ExcelFile(p)
                self.teacher_sheets = xl.sheet_names
                self.teacher_sheet_cb["values"] = self.teacher_sheets
                if "SELF" in self.teacher_sheets:
                    self.teacher_sheet_cb.set("SELF")
                else:
                    self.teacher_sheet_cb.current(0)
                self.append_status(f"Teacher file loaded: {p} (sheets: {self.teacher_sheets})")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Teacher file:\n{e}")
                self.append_status(f"Error reading Teacher file: {e}")

    def browse_admin_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if p:
            self.admin_path = p
            self.admin_entry.delete(0, tk.END); self.admin_entry.insert(0, p)
            try:
                xl = pd.ExcelFile(p)
                self.admin_sheets = xl.sheet_names
                self.admin_sheet_cb["values"] = self.admin_sheets
                if "Sheet2" in self.admin_sheets:
                    self.admin_sheet_cb.set("Sheet2")
                else:
                    self.admin_sheet_cb.current(0)
                self.append_status(f"Admin file loaded: {p} (sheets: {self.admin_sheets})")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Admin file:\n{e}")
                self.append_status(f"Error reading Admin file: {e}")

    # ---------------- Core processing ----------------
    def run_full_process(self):
        raw = self.raw_entry.get().strip()
        if not raw:
            messagebox.showerror("Error", "Select RAW Excel file first.")
            return
        name_col = self.name_col_var.get().strip()
        if not name_col:
            messagebox.showerror("Error", "Select NAME column to clean (Load Columns first).")
            return

        try:
            df = pd.read_excel(raw)
            self.append_status(f"Read RAW file: {raw} ({len(df)} rows)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read RAW file:\n{e}")
            self.append_status(f"Failed to read RAW file: {e}")
            return

        # normalize column names spacing
        df.columns = [c.strip() for c in df.columns]

        # Ensure RAW has an ID-like column; detect and standardize to 'ID'
        id_col = detect_id_column(df, ["id", "ID", "upid", "empid", "uid", "staffid", "emp id", "up id"])
        if id_col is None:
            messagebox.showerror("Error", "RAW file must contain an ID-like column (ID/UPID/EMPID).")
            self.append_status("RAW missing ID-like column.")
            return
        if id_col != "ID":
            df.rename(columns={id_col: "ID"}, inplace=True)
        # normalize ID values
        df["ID"] = df["ID"].apply(normalize_id_value)

        # Clean name column
        if name_col not in df.columns:
            messagebox.showerror("Error", f"Selected NAME column '{name_col}' not found in RAW file.")
            self.append_status(f"NAME column '{name_col}' not in RAW columns.")
            return
        df[name_col] = df[name_col].apply(strip_salutation)
        df["NAME"] = df[name_col]
        self.append_status("Names cleaned (salutations removed).")

        # ROLE classification
        if "ROLE" not in df.columns:
            messagebox.showerror("Error", "RAW file must contain a 'ROLE' column for classification.")
            self.append_status("ROLE column missing in RAW.")
            return
        df["Category"] = df["ROLE"].apply(classify_role)
        self.append_status("Roles classified into Category.")

        # Format scans
        df = format_scan_columns(df)
        self.append_status("Formatted FIRST_SCAN and LAST_SCAN (HH:MM:SS).")

        # Remove placeholder & sample IDs
        try:
            df = df[df["NAME"] != "NEW REQ SST"]
        except Exception:
            pass
        sample_ids = [2312, 4990, 615, 4989, 4689, 5021]
        df = df[~df["ID"].isin([normalize_id_value(x) for x in sample_ids])]

        # Initialize MoT column (empty)
        if "MoT" not in df.columns:
            df["MoT"] = ""

        # --- Teacher MoT matching ---
        teacher_matches = 0
        if self.teacher_path and self.teacher_sheet_cb.get().strip():
            try:
                t_sheet = self.teacher_sheet_cb.get().strip()
                # read full sheet (prefer)
                tdf = pd.read_excel(self.teacher_path, sheet_name=t_sheet, header=0)
                # detect teacher id column; prefer 'id'
                t_id_col = detect_id_column(tdf, ["id", "ID"])
                if t_id_col is None:
                    # fallback to first column
                    t_id_col = tdf.columns[0]
                # normalize teacher id column name to 't__ID' to avoid confusion
                tdf["t__ID"] = tdf[t_id_col].astype(str).apply(normalize_id_value)
                teacher_id_set = set(tdf["t__ID"].dropna().unique().tolist())

                # target mask in RAW
                teacher_group_mask = df["Category"].isin(["PRINCIPAL", "VICE PRINCIPAL", "VICE PRESIDENT", "SCHOOL COORDINATOR", "HEADMISTRESS", "TEACHER"])
                # find matches
                matches_mask = teacher_group_mask & df["ID"].isin(teacher_id_set)
                teacher_matches = int(matches_mask.sum())
                if teacher_matches > 0:
                    df.loc[matches_mask, "MoT"] = "SUP"
                    # set remaining teacher-group rows with empty MoT to TRN
                    df.loc[teacher_group_mask & (df["MoT"].astype(str).str.strip() == ""), "MoT"] = "TRN"
                    # copy route column if present in teacher sheet: prefer 'Route No. / Self' or 'Route' or 'Route No'
                    for candidate in ["Route No. / Self", "Route No.", "Route", "Route No", "route"]:
                        if candidate in tdf.columns:
                            # map by t__ID
                            route_map = tdf.set_index("t__ID")[candidate].astype(str).to_dict()
                            df.loc[matches_mask, candidate] = df.loc[matches_mask, "ID"].map(route_map)
                            break
                else:
                    # mark teacher-group TRN if not already set
                    df.loc[teacher_group_mask & (df["MoT"].astype(str).str.strip() == ""), "MoT"] = "TRN"

                self.append_status(f"Teacher sheet '{t_sheet}' processed — teacher matches: {teacher_matches}, teacher set size: {len(teacher_id_set)}")
            except Exception as e:
                self.append_status(f"Teacher MoT error: {e}")

        # --- Admin MoT matching ---
        admin_matches = 0
        if self.admin_path and self.admin_sheet_cb.get().strip():
            try:
                a_sheet = self.admin_sheet_cb.get().strip()
                adf = pd.read_excel(self.admin_path, sheet_name=a_sheet, header=0)
                # detect admin id column; prefer 'upid'
                a_id_col = detect_id_column(adf, ["upid", "ID", "id"])
                if a_id_col is None:
                    a_id_col = adf.columns[0]
                adf["a__ID"] = adf[a_id_col].astype(str).apply(normalize_id_value)
                admin_id_set = set(adf["a__ID"].dropna().unique().tolist())

                admin_group_mask = df["Category"].isin(["ADMIN", "GROUND STAFF", "NURSE", "CLEANER", "MAINTENANCE", "MAID", "PEON", "OTHER"])
                matches_mask = admin_group_mask & df["ID"].isin(admin_id_set)
                admin_matches = int(matches_mask.sum())
                if admin_matches > 0:
                    df.loc[matches_mask, "MoT"] = "SUP"
                    df.loc[admin_group_mask & (df["MoT"].astype(str).str.strip() == ""), "MoT"] = "TRN"
                    # copy TRANS MODE if present
                    for candidate in ["TRANS MODE", "Trans Mode", "TRANS_MODE", "TRANS"]:
                        if candidate in adf.columns:
                            mode_map = adf.set_index("a__ID")[candidate].astype(str).to_dict()
                            df.loc[matches_mask, candidate] = df.loc[matches_mask, "ID"].map(mode_map)
                            break
                else:
                    df.loc[admin_group_mask & (df["MoT"].astype(str).str.strip() == ""), "MoT"] = "TRN"

                self.append_status(f"Admin sheet '{a_sheet}' processed — admin matches: {admin_matches}, admin set size: {len(admin_id_set)}")
            except Exception as e:
                self.append_status(f"Admin MoT error: {e}")

        # For Guard/Driver groups we keep MoT blank (if required). If you prefer TRN, change below.
        non_mot_groups = df["Category"].isin(["GUARD", "DRIVER", "HELPER"])
        df.loc[non_mot_groups, "MoT"] = df.loc[non_mot_groups, "MoT"].fillna("").astype(str)

        # Compute attendance & LATE_STATUS using grace times (24-hour)
        gs = self.grace_start_var.get().strip() or DEFAULT_GRACE_START
        ge = self.grace_end_var.get().strip() or DEFAULT_GRACE_END
        try:
            df = compute_late_attendance_for_df_with_grace(df, gs, ge)
            self.append_status(f"Computed Attendance & LATE_STATUS using grace {gs} -> {ge}.")
        except Exception as e:
            self.append_status(f"Compute attendance error: {e}")

        # Ensure required columns exist and in correct order
        for col in ["ID", "NAME", "ROLE", "FIRST_SCAN", "LAST_SCAN", "MoT", "Attendance", "LATE_STATUS"]:
            if col not in df.columns:
                df[col] = ""

        self.processed_df = df.copy()
        self.append_status(f"Processing finished. Teacher matches: {teacher_matches}, Admin matches: {admin_matches}. Stored in memory.")
        messagebox.showinfo("Processing Completed", "Processing Completed Successfully. Click SAVE AS to export (XLSX/CSV).")

    # ---------------- Save As (user chooses XLSX or CSV) ----------------
    def save_as(self):
        if self.processed_df is None:
            messagebox.showerror("Error", "No processed data in memory. Run PROCESS first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel Workbook", "*.xlsx"), ("CSV file", "*.csv")])
        if not p:
            return
        try:
            df = self.processed_df.copy()
            final_cols = [col for col in ["ID", "NAME", "ROLE", "FIRST_SCAN", "LAST_SCAN", "MoT", "Attendance", "LATE_STATUS"] if col in df.columns]
            df_out = df[final_cols].copy()
            if p.lower().endswith(".csv"):
                df_out.to_csv(p, index=False)
                messagebox.showinfo("Saved", f"Saved CSV:\n{p}")
                self.append_status(f"Saved CSV: {p}")
            else:
                group1 = df[df["Category"].isin(["PRINCIPAL", "VICE PRINCIPAL", "VICE PRESIDENT", "SCHOOL COORDINATOR", "HEADMISTRESS", "TEACHER"])].copy()
                group2 = df[df["Category"].isin(["ADMIN", "GROUND STAFF", "NURSE", "CLEANER", "MAINTENANCE", "MAID", "PEON", "OTHER"])].copy()
                group3 = df[df["Category"].isin(["GUARD"])].copy()
                group4 = df[df["Category"].isin(["DRIVER", "HELPER"])].copy()
                with pd.ExcelWriter(p, engine="openpyxl") as writer:
                    if not group1.empty:
                        group1[[c for c in final_cols if c in group1.columns]].to_excel(writer, sheet_name="Principal_Teacher", index=False)
                    if not group2.empty:
                        group2[[c for c in final_cols if c in group2.columns]].to_excel(writer, sheet_name="Admin_Staff", index=False)
                    if not group3.empty:
                        group3[[c for c in final_cols if c in group3.columns]].to_excel(writer, sheet_name="Guard", index=False)
                    if not group4.empty:
                        group4[[c for c in final_cols if c in group4.columns]].to_excel(writer, sheet_name="Driver_Helper", index=False)
                    df_out.to_excel(writer, sheet_name="Combined", index=False)
                messagebox.showinfo("Saved", f"Saved Excel (.xlsx):\n{p}")
                self.append_status(f"Saved XLSX: {p}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save file:\n{e}")
            self.append_status(f"Save error: {e}")

    # ---------------- Save Final (same as Save As but automatic grouping) ----------------
    def save_final_now(self):
        if self.processed_df is None:
            messagebox.showerror("Error", "No processed data in memory. Run PROCESS first.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")])
        if not p:
            return
        try:
            df = self.processed_df
            group1 = df[df["Category"].isin(["PRINCIPAL", "VICE PRINCIPAL", "VICE PRESIDENT", "SCHOOL COORDINATOR", "HEADMISTRESS", "TEACHER"])]
            group2 = df[df["Category"].isin(["ADMIN", "GROUND STAFF", "NURSE", "CLEANER", "MAINTENANCE", "MAID", "PEON", "OTHER"])]
            group3 = df[df["Category"].isin(["GUARD"])]
            group4 = df[df["Category"].isin(["DRIVER", "HELPER"])]

            # Sorting
            teacher_order = {"PRINCIPAL": 1, "VICE PRINCIPAL": 2,  "VICE PRESIDENT": 3 , "SCHOOL COORDINATOR": 4,"HEADMISTRESS": 5, "TEACHER": 6}
            group1 = group1.copy()
            group1["SortKey"] = group1["Category"].map(teacher_order)
            group1 = group1.sort_values(by=["SortKey", "NAME"]).drop(columns=["SortKey"])

            admin_order = {"ADMIN": 1, "NURSE": 2}
            group2 = group2.copy()
            group2["SortKey"] = group2["Category"].map(admin_order).fillna(4)
            group2 = group2.sort_values(by=["SortKey", "Category", "NAME"]).drop(columns=["SortKey"])

            group3 = group3.sort_values(by="NAME")

            dh_order = {"DRIVER": 1, "HELPER": 2}
            group4 = group4.copy()
            group4["SortKey"] = group4["Category"].map(dh_order)
            group4 = group4.sort_values(by=["SortKey", "NAME"]).drop(columns=["SortKey"])


            # Reorder columns
            final_cols = [col for col in ["ID", "NAME", "Category", "FIRST_SCAN", "LAST_SCAN", "MoT", "Attendance", "LATE_STATUS"] if col in df.columns]
            finalS_cols = [col for col in ["ID", "NAME", "Category", "FIRST_SCAN", "LAST_SCAN", "Attendance", "LATE_STATUS"] if col in df.columns]

            group1 = group1[final_cols]
            group2 = group2[final_cols]
            group3 = group3[finalS_cols]
            group4 = group4[finalS_cols]
            
            with pd.ExcelWriter(p, engine="openpyxl") as writer:
                if not group1.empty:
                    group1[[c for c in final_cols if c in group1.columns]].to_excel(writer, sheet_name="Principal_Teacher", index=False)
                if not group2.empty:
                    group2[[c for c in final_cols if c in group2.columns]].to_excel(writer, sheet_name="Admin_Staff", index=False)
                if not group3.empty:
                    group3[[c for c in final_cols if c in group3.columns]].to_excel(writer, sheet_name="Guard", index=False)
                if not group4.empty:
                    group4[[c for c in final_cols if c in group4.columns]].to_excel(writer, sheet_name="Driver_Helper", index=False)


            messagebox.showinfo("Saved", f"Final export saved:\n{p}")
            self.append_status(f"Final XLSX saved: {p}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save final file:\n{e}")
            self.append_status(f"Final save error: {e}")

# ----------------- Run App -----------------
if __name__ == "__main__":
    app = OMKAttendanceSingleModule()
    app.mainloop()
