import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
import warnings

# ignore openpyxl header/footer parse warning (non-fatal)
warnings.filterwarnings("ignore", message="Cannot parse header or footer so it will be ignored")

SALUTATIONS_PATTERN = re.compile(r'^\s*(Mr|Mrs|Ms|Dr|Miss)\.?\s+', flags=re.IGNORECASE)

def strip_salutation(name: str) -> str:
    if pd.isna(name):
        return name
    s = str(name).strip()
    s = re.sub(SALUTATIONS_PATTERN, '', s)
    return s.strip()

# --- Classification (same as your logic) ---
def clean_role_text(text: str) -> str:
    if pd.isna(text):
        return text
    t = str(text).strip()
    bracket_match = re.search(r"\((.*?)\)", t)
    if bracket_match:
        return bracket_match.group(1).strip()
    return t.replace("Default", "").strip()

def classify_role(text):
    if pd.isna(text):
        return text
    t = clean_role_text(str(text)).upper()

    if "HEADMISTRESS" in t:
        return "HEADMISTRESS"
    elif "VICE PRESIDENT" in t:
        return "VICE PRESIDENT"
    elif "SCHOOL COORDINATOR" in t:
        return "SCHOOL COORDINATOR"
    elif "VICE PRINCIPAL" in t:
        return "VICE PRINCIPAL"
    elif "PRINCIPAL" in t:
        return "PRINCIPAL"
    elif "TEACHER" in t:
        return "TEACHER"
    elif any(x in t for x in ["ADMINISTRATOR", "FRONT OFFICE", "ACCOUNTANT", "ACCOUNT"]):
        return "ADMIN"
    elif "DIRECTOR" in t:
        return "DIRECTOR"
    elif "HELPER" in t:
        return "HELPER"
    elif "DRIVER" in t:
        return "DRIVER"
    elif "NURSE" in t:
        return "NURSE"
    elif "MAID" in t:
        return "MAID"
    elif any(x in t for x in ["PERSONAL SECURITY", "GUARD"]):
        return "GUARD"
    elif "GROUND STAFF" in t:
        return "GROUND STAFF"
    elif "PEON" in t:
        return "PEON"
    elif "CLEANER" in t:
        return "CLEANER"
    elif "MAINTENANCE" in t:
        return "MAINTENANCE"
    elif "OTHER" in t:
        return "OTHER"
    else:
        return t.strip()

# --- Module 1 processing ---
def process_module1(input_file, output_file, name_col):
    role_col = "ROLE"
    df = pd.read_excel(input_file, header=0)

    if name_col not in df.columns:
        raise ValueError(f"Selected name column '{name_col}' not in Excel columns.")

    # strip salutations from selected name column and store back to NAME (or keep original NAME if exists)
    df[name_col] = df[name_col].apply(strip_salutation)

    # ensure we have a NAME column expected downstream; if user selected some other col, copy it to NAME
    if "NAME" not in df.columns:
        df["NAME"] = df[name_col]
    else:
        # prefer cleaned column in NAME
        df["NAME"] = df[name_col]

    # Apply role classification
    if role_col not in df.columns:
        raise ValueError(f"Expected column '{role_col}' not found in input file.")
    df["Category"] = df[role_col].apply(classify_role)

    # Format datetime columns
    for col in ["FIRST_SCAN", "LAST_SCAN"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%H:%M:%S")
            df[col] = df[col].fillna("--")

    # Remove unwanted row and IDs
    df = df[df["NAME"] != "NEW REQ SST"]
    ids = [2312, 4990, 615, 4989, 4689, 5021]
    if "ID" in df.columns:
        df = df[~df["ID"].isin(ids)]

    # --- Groups ---
    group1 = df[df["Category"].isin(["PRINCIPAL", "VICE PRINCIPAL", "VICE PRESIDENT" , "SCHOOL COORDINATOR", "HEADMISTRESS", "TEACHER"])]
    group2 = df[df["Category"].isin([
        "ADMIN", "GROUND STAFF", "NURSE",
        "CLEANER", "MAINTENANCE", "MAID", "PEON", "OTHER"
    ])]
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
    final_cols = [col for col in ["ID", "NAME", "Category", "FIRST_SCAN", "LAST_SCAN"] if col in df.columns]
    group1 = group1[final_cols]
    group2 = group2[final_cols]
    group3 = group3[final_cols]
    group4 = group4[final_cols]

    # Save
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        group1.to_excel(writer, sheet_name="Principal_Teacher", index=False)
        group2.to_excel(writer, sheet_name="Admin_Staff", index=False)
        group3.to_excel(writer, sheet_name="Guard", index=False)
        group4.to_excel(writer, sheet_name="Driver_Helper", index=False)

# --- Module 2 helpers ---
def read_ids_from_sheet(path, sheet_name):
    """Read ID list from a sheet. Prefer 'ID'/'UPID' column if present, else use first column."""
    xl = pd.ExcelFile(path)
    if sheet_name not in xl.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' not found in {path}. Available: {xl.sheet_names}")
    df = pd.read_excel(path, sheet_name=sheet_name, header=0)
    for c in df.columns:
        if str(c).strip().lower() in ("id", "upid", "uid", "empid", "up id"):
            return df[c].astype(str).str.strip().dropna().tolist()
    return df.iloc[:,0].astype(str).str.strip().dropna().tolist()

def apply_mot_by_id(df, id_lookup_list, id_col_name="ID"):
    """Set MoT column in df based on whether df[id_col_name] is in id_lookup_list"""
    lookup_set = set(str(x).strip() for x in id_lookup_list)
    if id_col_name in df.columns:
        df["MoT"] = df[id_col_name].astype(str).apply(lambda x: "SUP" if x.strip() in lookup_set else "TRN")
    else:
        df["MoT"] = df["NAME"].astype(str).apply(lambda x: "TRN")
    return df

def compute_late_attendance_for_df_with_grace(df, grace_start_str, grace_end_str):
    """
    Add Attendance and LATE_STATUS columns based on FIRST_SCAN and provided grace times.
      - ON-TIME: FIRST_SCAN <= grace_start
      - ADV: grace_start < FIRST_SCAN <= grace_end
      - LATE: FIRST_SCAN > grace_end
      - AB: FIRST_SCAN is '--' / missing
    """
    def parse_time_safe(s, fallback):
        try:
            if s is None:
                return fallback
            s = str(s).strip()
            if not s:
                return fallback
            for fmt in ("%H:%M:%S", "%H:%M"):
                try:
                    return datetime.datetime.strptime(s, fmt).time()
                except Exception:
                    continue
            return fallback
        except Exception:
            return fallback

    default_start = datetime.time(7, 40, 0)
    default_end = datetime.time(8, 0, 0)
    grace_start = parse_time_safe(grace_start_str, default_start)
    grace_end = parse_time_safe(grace_end_str, default_end)

    if grace_start > grace_end:
        grace_start, grace_end = grace_end, grace_start

    def parse_scan_or_none(x):
        if pd.isna(x):
            return None
        s = str(x).strip()
        if s == "--" or s == "":
            return None
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.datetime.strptime(s, fmt).time()
            except Exception:
                continue
        return None

    if "FIRST_SCAN" in df.columns:
        df["Attendance"] = df["FIRST_SCAN"].apply(lambda x: "AB" if (pd.isna(x) or str(x).strip() == "--") else "PR")
        def late_logic(t):
            if t is None:
                return " "
            if t <= grace_start:
                return " "
            elif grace_start < t <= grace_end:
                return "ADV"
            else:
                return "LATE"
        df["LATE_STATUS"] = df["FIRST_SCAN"].apply(lambda x: late_logic(parse_scan_or_none(x)))
    else:
        df["Attendance"] = ""
        df["LATE_STATUS"] = ""
    return df

# --- Full Tkinter App v1.4 (dropdowns) ---
class TwoModuleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OMK Attendance Tool v1.4 (Dropdowns & Grace-Time)")

        # initialize attributes
        self.m1_output_path = None
        self.teacher_mot_path = None
        self.admin_mot_path = None
        self.teacher_sheet_names = []
        self.admin_sheet_names = []

        # center window
        w, h = 1000, 640
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        x = (sw - w)//2; y = (sh - h)//2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.resizable(False, False)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.module1_frame = ttk.Frame(nb)
        self.module2_frame = ttk.Frame(nb)

        nb.add(self.module1_frame, text="Module 1 — Preprocess & Split")
        nb.add(self.module2_frame, text="Module 2 — MoT, Grace & Attendance")

        self.build_module1()
        self.build_module2()

    # ---------- Module 1 ----------
    def build_module1(self):
        f = self.module1_frame
        padx = 8; pady = 6

        ttk.Label(f, text="Module 1 — Select Input and Clean Names", font=("Verdana", 12, "bold")).pack(pady=12)

        row = ttk.Frame(f); row.pack(fill="x", padx=padx, pady=pady)
        ttk.Label(row, text="Input Excel:", width=16).pack(side="left")
        self.m1_input_entry = ttk.Entry(row, width=80)
        self.m1_input_entry.pack(side="left", padx=4)
        ttk.Button(row, text="Browse", command=self.m1_browse_input).pack(side="left", padx=4)

        row2 = ttk.Frame(f); row2.pack(fill="x", padx=padx, pady=pady)
        ttk.Label(row2, text="Output Excel (save as):", width=16).pack(side="left")
        self.m1_output_entry = ttk.Entry(row2, width=80)
        self.m1_output_entry.pack(side="left", padx=4)
        ttk.Button(row2, text="Save As", command=self.m1_save_as).pack(side="left", padx=4)

        row3 = ttk.Frame(f); row3.pack(fill="x", padx=padx, pady=pady)
        ttk.Label(row3, text="Select NAME column to clean:", width=22).pack(side="left")
        self.name_col_var = tk.StringVar()
        self.name_col_cb = ttk.Combobox(row3, textvariable=self.name_col_var, values=[], width=46, state="readonly")
        self.name_col_cb.pack(side="left", padx=4)
        ttk.Button(row3, text="Load Columns", command=self.populate_name_columns).pack(side="left", padx=4)

        ttk.Button(f, text="Run Module 1 (Clean & Split)", command=self.run_module1).pack(pady=12)
        ttk.Label(f, text="Creates: Principal_Teacher, Admin_Staff, Guard, Driver_Helper", foreground="blue").pack(pady=6)

    def m1_browse_input(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if p:
            self.m1_input_entry.delete(0, tk.END)
            self.m1_input_entry.insert(0, p)
            try:
                df = pd.read_excel(p, nrows=0)
                cols = list(df.columns)
                self.name_col_cb['values'] = cols
                if "NAME" in cols:
                    self.name_col_var.set("NAME")
                elif "Teacher's Name" in cols:
                    self.name_col_var.set("Teacher's Name")
                else:
                    self.name_col_var.set(cols[0] if cols else "")
            except Exception:
                pass

    def m1_save_as(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.m1_output_entry.delete(0, tk.END)
            self.m1_output_entry.insert(0, p)

    def populate_name_columns(self):
        p = self.m1_input_entry.get().strip()
        if not p:
            messagebox.showerror("Error", "Select input file first.")
            return
        try:
            df = pd.read_excel(p, nrows=0)
            cols = list(df.columns)
            self.name_col_cb['values'] = cols
            messagebox.showinfo("Columns loaded", "Choose the column that contains the names (to strip Mr/Ms/Dr/Miss).")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file columns:\n{e}")

    def run_module1(self):
        inp = self.m1_input_entry.get().strip()
        out = self.m1_output_entry.get().strip()
        namecol = self.name_col_var.get().strip()
        if not inp or not out:
            messagebox.showerror("Error", "Please select input and output for Module 1.")
            return
        if not namecol:
            messagebox.showerror("Error", "Please select the NAME column to clean.")
            return
        try:
            process_module1(inp, out, namecol)
            self.m1_output_path = out
            self.m1_output_entry.delete(0, tk.END); self.m1_output_entry.insert(0, out)
            messagebox.showinfo("Success", f"Module 1 finished and saved:\n{out}")
        except Exception as e:
            messagebox.showerror("Module 1 Error", str(e))

    # ---------- Module 2 ----------
    def build_module2(self):
        f = self.module2_frame
        padx = 8; pady = 6

        ttk.Label(f, text="Module 2 — MoT (Teacher/Admin), Grace time & Attendance", font=("Verdana", 12, "bold")).pack(pady=10)

        # processed file (Module1 output)
        row = ttk.Frame(f); row.pack(fill="x", padx=padx, pady=4)
        ttk.Label(row, text="Processed Module1 Excel:", width=20).pack(side="left")
        self.m2_input_entry = ttk.Entry(row, width=62)
        self.m2_input_entry.pack(side="left", padx=4)
        ttk.Button(row, text="Browse", command=self.m2_browse_processed).pack(side="left", padx=4)
        ttk.Button(row, text="Use Last Module1 Output", command=self.use_last_m1_output).pack(side="left", padx=4)

        # Teacher file upload + dropdown
        row_tr = ttk.Frame(f); row_tr.pack(fill="x", padx=padx, pady=4)
        ttk.Label(row_tr, text="Teacher Excel (any):", width=20).pack(side="left")
        self.tr_entry = ttk.Entry(row_tr, width=50); self.tr_entry.pack(side="left", padx=4)
        ttk.Button(row_tr, text="Upload Teacher File", command=self.upload_teacher_file_and_list_sheets).pack(side="left", padx=4)
        ttk.Label(row_tr, text="Select Teacher Sheet:", width=16).pack(side="left", padx=6)
        self.teacher_sheet_dropdown = ttk.Combobox(row_tr, width=28, state="readonly")
        self.teacher_sheet_dropdown.pack(side="left", padx=4)

        # Admin file upload + dropdown
        row_ad = ttk.Frame(f); row_ad.pack(fill="x", padx=padx, pady=4)
        ttk.Label(row_ad, text="Admin Excel (any):", width=20).pack(side="left")
        self.ad_entry = ttk.Entry(row_ad, width=50); self.ad_entry.pack(side="left", padx=4)
        ttk.Button(row_ad, text="Upload Admin File", command=self.upload_admin_file_and_list_sheets).pack(side="left", padx=4)
        ttk.Label(row_ad, text="Select Admin Sheet:", width=16).pack(side="left", padx=6)
        self.admin_sheet_dropdown = ttk.Combobox(row_ad, width=28, state="readonly")
        self.admin_sheet_dropdown.pack(side="left", padx=4)

        # Grace time inputs
        row_grace = ttk.Frame(f); row_grace.pack(fill="x", padx=padx, pady=8)
        ttk.Label(row_grace, text="Grace Start (ON-TIME until):", width=26).pack(side="left")
        self.grace_start_var = tk.StringVar(value="07:40:00")
        ttk.Entry(row_grace, textvariable=self.grace_start_var, width=12).pack(side="left", padx=6)
        ttk.Label(row_grace, text="Grace End (ADV until):", width=20).pack(side="left")
        self.grace_end_var = tk.StringVar(value="08:00:00")
        ttk.Entry(row_grace, textvariable=self.grace_end_var, width=12).pack(side="left", padx=6)

        # Buttons
        btn_row = ttk.Frame(f); btn_row.pack(pady=10)
        ttk.Button(btn_row, text="Apply Teacher MoT (Principal_Teacher)", command=self.run_apply_teacher_mot).pack(side="left", padx=10)
        ttk.Button(btn_row, text="Apply Admin MoT (Admin_Staff)", command=self.run_apply_admin_mot).pack(side="left", padx=10)
        ttk.Button(btn_row, text="Compute Late & Attendance (All sheets)", command=self.run_compute_late_attendance).pack(side="left", padx=10)

        # Save result
        row_save = ttk.Frame(f); row_save.pack(fill="x", padx=padx, pady=8)
        ttk.Label(row_save, text="Module2 Save As (optional):", width=20).pack(side="left")
        self.m2_output_entry = ttk.Entry(row_save, width=60); self.m2_output_entry.pack(side="left", padx=4)
        ttk.Button(row_save, text="Save As", command=self.m2_save_as).pack(side="left", padx=4)
        ttk.Button(f, text="Save Module2 Output Now", command=self.save_module2_output).pack(pady=8)

    # --- Module2 helpers and actions ---
    def m2_browse_processed(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.m2_input_entry.delete(0, tk.END)
            self.m2_input_entry.insert(0, p)

    def use_last_m1_output(self):
        if self.m1_output_path:
            self.m2_input_entry.delete(0, tk.END)
            self.m2_input_entry.insert(0, self.m1_output_path)
            messagebox.showinfo("Loaded", "Loaded last Module1 output for Module2.")
        else:
            messagebox.showerror("No Module1 Output", "Run Module1 first or select file manually.")

    def upload_teacher_file_and_list_sheets(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not p:
            return
        try:
            xl = pd.ExcelFile(p)
            self.teacher_sheet_names = xl.sheet_names
            self.teacher_sheet_dropdown['values'] = self.teacher_sheet_names
            if self.teacher_sheet_names:
                # auto-select SELF if present, else first sheet
                try:
                    idx = self.teacher_sheet_names.index("SELF")
                    self.teacher_sheet_dropdown.current(idx)
                except Exception:
                    self.teacher_sheet_dropdown.current(0)
            self.teacher_mot_path = p
            self.tr_entry.delete(0, tk.END); self.tr_entry.insert(0, p)
            messagebox.showinfo("Teacher file loaded", f"Found sheets: {self.teacher_sheet_names}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load teacher file:\n{e}")

    def upload_admin_file_and_list_sheets(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not p:
            return
        try:
            xl = pd.ExcelFile(p)
            self.admin_sheet_names = xl.sheet_names
            self.admin_sheet_dropdown['values'] = self.admin_sheet_names
            if self.admin_sheet_names:
                # auto-select 'Sheet2' if present else first
                try:
                    idx = self.admin_sheet_names.index("Sheet2")
                    self.admin_sheet_dropdown.current(idx)
                except Exception:
                    self.admin_sheet_dropdown.current(0)
            self.admin_mot_path = p
            self.ad_entry.delete(0, tk.END); self.ad_entry.insert(0, p)
            messagebox.showinfo("Admin file loaded", f"Found sheets: {self.admin_sheet_names}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load admin file:\n{e}")

    def run_apply_teacher_mot(self):
        proc = self.m2_input_entry.get().strip()
        if not proc:
            messagebox.showerror("Error", "Select processed Module1 Excel (Module1 output) first.")
            return
        if not getattr(self, "teacher_mot_path", None):
            messagebox.showerror("Error", "Upload Teacher Excel first.")
            return
        sheet = self.teacher_sheet_dropdown.get().strip()
        if not sheet:
            messagebox.showerror("Error", "Select Teacher sheet from dropdown.")
            return
        try:
            teacher_ids = read_ids_from_sheet(self.teacher_mot_path, sheet)
            book = pd.read_excel(proc, sheet_name=None)
            if "Principal_Teacher" not in book:
                messagebox.showerror("Error", "Principal_Teacher sheet not found in processed file.")
                return
            df_pt = book["Principal_Teacher"].copy()
            df_pt = apply_mot_by_id(df_pt, teacher_ids, id_col_name="ID")
            book["Principal_Teacher"] = df_pt
            outpath = self.m2_output_entry.get().strip() or proc
            with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
                for sh, df in book.items():
                    df.to_excel(writer, sheet_name=sh, index=False)
            self.m2_input_entry.delete(0, tk.END); self.m2_input_entry.insert(0, outpath)
            messagebox.showinfo("Done", f"Teacher MoT applied using sheet '{sheet}' and saved to:\n{outpath}")
        except Exception as e:
            messagebox.showerror("Teacher MoT Error", str(e))

    def run_apply_admin_mot(self):
        proc = self.m2_input_entry.get().strip()
        if not proc:
            messagebox.showerror("Error", "Select processed Module1 Excel (Module1 output) first.")
            return
        if not getattr(self, "admin_mot_path", None):
            messagebox.showerror("Error", "Upload Admin Excel first.")
            return
        sheet = self.admin_sheet_dropdown.get().strip()
        if not sheet:
            messagebox.showerror("Error", "Select Admin sheet from dropdown.")
            return
        try:
            admin_ids = read_ids_from_sheet(self.admin_mot_path, sheet)
            book = pd.read_excel(proc, sheet_name=None)
            if "Admin_Staff" not in book:
                messagebox.showerror("Error", "Admin_Staff sheet not found in processed file.")
                return
            df_ad = book["Admin_Staff"].copy()
            df_ad = apply_mot_by_id(df_ad, admin_ids, id_col_name="ID")
            book["Admin_Staff"] = df_ad
            outpath = self.m2_output_entry.get().strip() or proc
            with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
                for sh, df in book.items():
                    df.to_excel(writer, sheet_name=sh, index=False)
            self.m2_input_entry.delete(0, tk.END); self.m2_input_entry.insert(0, outpath)
            messagebox.showinfo("Done", f"Admin MoT applied using sheet '{sheet}' and saved to:\n{outpath}")
        except Exception as e:
            messagebox.showerror("Admin MoT Error", str(e))

    def run_compute_late_attendance(self):
        proc = self.m2_input_entry.get().strip()
        if not proc:
            messagebox.showerror("Error", "Select processed Module1 Excel (or Module2 intermediate) file first.")
            return
        grace_start = self.grace_start_var.get().strip()
        grace_end = self.grace_end_var.get().strip()
        try:
            book = pd.read_excel(proc, sheet_name=None)
            for sh, df in book.items():
                if "FIRST_SCAN" in df.columns:
                    df = compute_late_attendance_for_df_with_grace(df, grace_start, grace_end)
                    if "MoT" not in df.columns:
                        df["MoT"] = ""
                book[sh] = df
            outpath = self.m2_output_entry.get().strip() or proc
            with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
                for sh, df in book.items():
                    df.to_excel(writer, sheet_name=sh, index=False)
            self.m2_input_entry.delete(0, tk.END); self.m2_input_entry.insert(0, outpath)
            messagebox.showinfo("Done", f"Late & Attendance computed and saved to:\n{outpath}\nGrace Start: {grace_start} | Grace End: {grace_end}")
        except Exception as e:
            messagebox.showerror("Compute Error", str(e))

    def m2_save_as(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.m2_output_entry.delete(0, tk.END)
            self.m2_output_entry.insert(0, p)

    def save_module2_output(self):
        src = self.m2_input_entry.get().strip()
        out = self.m2_output_entry.get().strip()
        if not src:
            messagebox.showerror("Error", "Select input processed file first (Module1 output or Module2 interim).")
            return
        if not out:
            messagebox.showerror("Error", "Select save path for Module2 output.")
            return
        try:
            book = pd.read_excel(src, sheet_name=None)
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                for sh, df in book.items():
                    df.to_excel(writer, sheet_name=sh, index=False)
            messagebox.showinfo("Saved", f"Module2 output saved:\n{out}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

if __name__ == "__main__":
    app = TwoModuleApp()
    app.mainloop()
