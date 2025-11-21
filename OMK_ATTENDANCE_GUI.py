import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Clean text helper ---
def clean_role_text(text: str) -> str:
    if pd.isna(text):
        return text
    t = str(text).strip()
    bracket_match = re.search(r"\((.*?)\)", t)
    if bracket_match:
        return bracket_match.group(1).strip()
    return t.replace("Default", "").strip()

# --- Classification ---
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

def process_excel(input_file, output_file):
    role_col = "ROLE"
    df = pd.read_excel(input_file, header=0)

    # Apply role classification
    df["Category"] = df[role_col].apply(classify_role)

    # Format datetime columns
    for col in ["FIRST_SCAN", "LAST_SCAN"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%H:%M:%S")
            df[col] = df[col].fillna("--")

    # Remove unwanted row
    df = df[df["NAME"] != "NEW REQ SST"]
    #df = df[df["ID"] != "2312"]
    ids = [2312, 4990, 615, 4989, 4689, 5021]
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

    messagebox.showinfo("Success", f"Excel saved successfully!\n\n{output_file}")

# --- Tkinter App Class ---
class ExcelApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Task Tool")

        # --- Center Window ---
        window_width = 600
        window_height = 400
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.resizable(False, False)

        # --- Fonts ---
        self.label_font = ("Verdana", 14, "bold")
        self.entry_font = ("Verdana", 12)
        self.button_font = ("Verdana", 12, "bold")

        # Create main UI
        self.create_widgets()

        # Override close button (X)
        self.protocol("WM_DELETE_WINDOW", self.show_final_frame)

    def create_widgets(self):
        tk.Label(self, text="Select Input Excel File:", font=self.label_font).pack(pady=10)
        self.input_entry = tk.Entry(self, width=55, font=self.entry_font)
        self.input_entry.pack(pady=5)
        tk.Button(self, text="Browse", command=self.select_input,
                  font=self.button_font, bg="#258EFF", fg="white", width=12).pack(pady=5)

        tk.Label(self, text="Select Output Excel File:", font=self.label_font).pack(pady=10)
        self.output_entry = tk.Entry(self, width=55, font=self.entry_font)
        self.output_entry.pack(pady=5)
        tk.Button(self, text="Save As", command=self.select_output,
                  font=self.button_font, bg="#F825E7", fg="white", width=12).pack(pady=5)

        tk.Button(self, text="Run Process", command=self.run_process,
                  font=("Verdana", 12, "bold"), bg="green", fg="white", width=12).pack(pady=5)

        # Exit button
        tk.Button(self, text="Exit", command=self.show_final_frame,
                  font=("Verdana", 12, "bold"), bg="red", fg="white", width=12).pack(pady=5)

    def select_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)

    def select_output(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def run_process(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()
        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        try:
            process_excel(input_file, output_file)
        except Exception as e:
            messagebox.showerror("Error", f"Processing failed:\n{e}")

    def show_final_frame(self):
        for widget in self.winfo_children():
            widget.destroy()

        final_frame = tk.Frame(self)
        final_frame.pack(fill="both", expand=True)

        tk.Label(final_frame, text="🎉 Thank you for using this Excel Task Tool!",
                 font=("Verdana", 16, "bold"), fg="green").pack(pady=60)

        tk.Button(final_frame, text="Exit", command=self.destroy,
                  font=("Tahoma", 12), width=20, bd=0, bg="red", fg="white").pack(pady=10)


if __name__ == "__main__":
    app = ExcelApp()
    app.mainloop()
