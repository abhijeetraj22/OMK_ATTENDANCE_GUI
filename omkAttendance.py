import pandas as pd
import re

# Load Excel
input_file = "testing.xlsx"
output_file = "classified_output.xlsx"

df = pd.read_excel(input_file, header=0)

role_col = "ROLE"

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
    elif "VICE PRINCIPAL" in t:
        return "VICE PRINCIPAL"
    elif "PRINCIPAL" in t:
        return "PRINCIPAL"
    elif "TEACHER" in t:
        return "TEACHER"
    elif any(x in t for x in ["ADMINISTRATOR", "FRONT OFFICE", "ACCOUNTANT", "ACCOUNT"]):
        return "ADMINISTRATOR"
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

# Apply
df["Category"] = df[role_col].apply(classify_role)

# Format datetime columns
for col in df.columns:
    if pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = df[col].dt.strftime("%H:%M:%S")

# --- Groups ---
group1 = df[df["Category"].isin(["PRINCIPAL", "VICE PRINCIPAL", "HEADMISTRESS", "TEACHER"])]
# --- Groups ---
group2 = df[df["Category"].isin([
    "ADMINISTRATOR", "GROUND STAFF", "NURSE",
    "CLEANER", "MAINTENANCE", "MAID", "PEON", "OTHER"
])]

group3 = df[df["Category"].isin(["GUARD"])]
group4 = df[df["Category"].isin(["DRIVER", "HELPER"])]

# Sorting
teacher_order = {"PRINCIPAL": 1, "VICE PRINCIPAL": 2, "HEADMISTRESS": 3, "TEACHER": 4}
group1 = group1.copy()
group1["SortKey"] = group1["Category"].map(teacher_order)
group1 = group1.sort_values(by=["SortKey", "NAME"]).drop(columns=["SortKey"])

admin_order = {"ADMINISTRATOR": 1, "DIRECTOR": 2, "NURSE": 3}
group2 = group2.copy()
group2["SortKey"] = group2["Category"].map(admin_order).fillna(4)
group2 = group2.sort_values(by=["SortKey", "Category", "NAME"]).drop(columns=["SortKey"])

group3 = group3.sort_values(by="NAME")

dh_order = {"DRIVER": 1, "HELPER": 2}
group4 = group4.copy()
group4["SortKey"] = group4["Category"].map(dh_order)
group4 = group4.sort_values(by=["SortKey", "NAME"]).drop(columns=["SortKey"])

# Save
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    group1.to_excel(writer, sheet_name="Principal_Teacher", index=False)
    group2.to_excel(writer, sheet_name="Admin_Staff", index=False)
    group3.to_excel(writer, sheet_name="Guard", index=False)
    group4.to_excel(writer, sheet_name="Driver_Helper", index=False)

print("✅ Excel saved with Admin including DIRECTOR + OTHER roles")
