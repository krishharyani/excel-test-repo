from openpyxl import load_workbook

XLSX_PATH = "report.xlsx"
SHEET_NAME = "FR-Detail"   # change if needed

# Set these once based on your template
LABEL_COL = "A"            # Business Line column
NODEID_COL = "B"           # NodeID column
VALUE_COL = "M"            # Example: 2024 YTD Actual (change to your target)

def nbsp_level(x) -> int:
    if x is None:
        return 0
    s = str(x)
    i = 0
    while i < len(s) and ord(s[i]) == 160:  # NBSP
        i += 1
    return i

def clean_label(x) -> str:
    if x is None:
        return ""
    return str(x).replace("\u00a0", " ").strip()

def parse_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "" or s.lower() in {"n/a", "#n/a"}:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "")
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None

def main():
    wb = load_workbook(XLSX_PATH, data_only=True)
    ws = wb[SHEET_NAME]

    # Print a few rows so you confirm columns and levels
    for r in range(50, 80):
        label_raw = ws[f"{LABEL_COL}{r}"].value
        node_id = ws[f"{NODEID_COL}{r}"].value
        val = ws[f"{VALUE_COL}{r}"].value

        lvl = nbsp_level(label_raw)
        label = clean_label(label_raw)

        if label:
            print(r, "lvl=", lvl, "node=", node_id, "label=", label, "val=", val)

if __name__ == "__main__":
    main()