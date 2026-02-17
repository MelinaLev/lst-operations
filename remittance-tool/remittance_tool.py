from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

PIONEER = "Pioneer Natural Resources"
XTO = "XTO Energy"

def _norm_invoice(x) -> str:
    # normalize invoice numbers like 140248, " 140248 ", "140248.0"
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # handle excel numeric -> float text
    if s.endswith(".0"):
        s = s[:-2]
    # remove commas/spaces just in case
    s = s.replace(",", "").replace(" ", "")
    return s

def find_header_row(excel_file, required_cols: set[str], max_scan_rows: int = 25) -> int:
    """
    AXON has 1–6 metadata rows. This finds the real header row by scanning
    the first N rows and locating the row that contains all required column names.
    Returns 0-based row index to use as `header=` in pd.read_excel.
    """
    preview = pd.read_excel(excel_file, header=None, nrows=max_scan_rows)
    for i in range(len(preview)):
        row_vals = preview.iloc[i].astype(str).str.strip().tolist()
        row_set = set(row_vals)
        if required_cols.issubset(row_set):
            return i
    # fallback: assume header is row 7 (0-based 6)
    return 6

def read_axon(axon_file) -> pd.DataFrame:
    # Find header row (Invoice#, Name, Amount are the ones we care about)
    required = {"Invoice#", "Name", "Amount"}
    header_row = find_header_row(axon_file, required_cols=required)

    axon_file.seek(0)
    df = pd.read_excel(axon_file, header=header_row)

    # Keep only what we need
    df = df[["Invoice#", "Name", "Amount"]].copy()
    df["Invoice#"] = df["Invoice#"].apply(_norm_invoice)
    df["Name"] = df["Name"].astype(str).str.strip()

    # If duplicates exist, keep the first name; amount not used for totals, but keep for future checks
    df = df[df["Invoice#"] != ""]
    df = df.drop_duplicates(subset=["Invoice#"])
    return df

def read_remittance(rem_file) -> pd.DataFrame:
    df = pd.read_excel(rem_file)

    # Expected columns from your screenshot
    # A: Co Code, B: Document, C: Invoice Date, D: Reference, E: Net Amount
    needed = ["Co Code", "Document", "Invoice Date", "Reference", "Net Amount"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Remittance is missing columns: {missing}")

    df = df[needed].copy()
    df["Reference"] = df["Reference"].apply(_norm_invoice)
    df["Net Amount"] = pd.to_numeric(df["Net Amount"], errors="coerce").fillna(0.0)
    df = df[df["Reference"] != ""]
    return df

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/remittance")
async def remittance_compare(
    axon: UploadFile = File(...),
    remittance: UploadFile = File(...)
):
    axon_df = read_axon(axon.file)
    rem_df = read_remittance(remittance.file)

    # Map invoice -> customer from AXON
    inv_to_customer = dict(zip(axon_df["Invoice#"], axon_df["Name"]))

    # Attach customer name to each remittance row
    rem_df["Customer"] = rem_df["Reference"].map(inv_to_customer).fillna("NOT FOUND IN AXON")

    # Create the Pioneer/XTO breakout columns
    rem_df["PioneerNaturalResources"] = rem_df.apply(
        lambda r: r["Net Amount"] if r["Customer"] == PIONEER else "",
        axis=1
    )
    rem_df["XTO"] = rem_df.apply(
        lambda r: r["Net Amount"] if r["Customer"] == XTO else "",
        axis=1
    )

    # Totals
    pioneer_total = rem_df.loc[rem_df["Customer"] == PIONEER, "Net Amount"].sum()
    xto_total = rem_df.loc[rem_df["Customer"] == XTO, "Net Amount"].sum()

    # Output table (like your screenshot)
    out_cols = ["Co Code", "Document", "Invoice Date", "Reference", "Net Amount", "Customer", "PioneerNaturalResources", "XTO"]
    out_df = rem_df[out_cols].copy()

    # Add TOTAL row
    total_row = {
        "Co Code": "",
        "Document": "",
        "Invoice Date": "",
        "Reference": "",
        "Net Amount": "",
        "Customer": "TOTAL",
        "PioneerNaturalResources": float(pioneer_total),
        "XTO": float(xto_total),
    }
    out_df = pd.concat([out_df, pd.DataFrame([total_row])], ignore_index=True)

    not_found_df = rem_df[rem_df["Customer"] == "NOT FOUND IN AXON"][out_cols].copy()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Customer Breakdown")
        not_found_df.to_excel(writer, index=False, sheet_name="Not Found in AXON")

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=remittance_breakdown.xlsx"},
    )