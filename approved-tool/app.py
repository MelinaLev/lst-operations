from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

def _norm_ticket(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    # remove commas/spaces to normalize (keeps G-prefixed tickets intact)
    return s.replace(",", "").replace(" ", "")

def _split_axon_tickets(cell) -> list[str]:
    if pd.isna(cell):
        return []
    s = str(cell).strip()
    if not s:
        return []
    # you said AXON always separates by comma + space, but splitting on comma is safest
    parts = [p.strip() for p in s.split(",")]
    parts = [p for p in parts if p]
    return parts

def _read_any(file_obj) -> pd.DataFrame:
    # Try Excel first, fall back to CSV
    try:
        file_obj.seek(0)
        return pd.read_excel(file_obj)
    except Exception:
        file_obj.seek(0)
        return pd.read_csv(file_obj)

def _find_header_row_excel(file_obj, required_cols: set[str], max_scan_rows: int = 30) -> int:
    file_obj.seek(0)
    preview = pd.read_excel(file_obj, header=None, nrows=max_scan_rows)
    for i in range(len(preview)):
        row_vals = preview.iloc[i].astype(str).str.strip().tolist()
        if required_cols.issubset(set(row_vals)):
            return i
    return 0

def read_axon(file_obj) -> pd.DataFrame:
    required = {"Invoice#", "Tickets", "Date", "Name", "Balance Due"}

    # Excel with header scan (AXON sometimes has metadata rows)
    try:
        header_row = _find_header_row_excel(file_obj, required_cols=required)
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header_row)
    except Exception:
        # CSV: scan for header row
        file_obj.seek(0)
        preview = pd.read_csv(file_obj, header=None, nrows=50)
        header_row = 0
        for i in range(len(preview)):
            row_vals = preview.iloc[i].astype(str).str.strip().tolist()
            if required.issubset(set(row_vals)):
                header_row = i
                break
        file_obj.seek(0)
        df = pd.read_csv(file_obj, header=header_row)

    df.columns = df.columns.astype(str).str.strip()
    missing = [c for c in ["Invoice#", "Tickets", "Date", "Name", "Balance Due"] if c not in df.columns]
    if missing:
        raise ValueError(f"AXON missing columns: {missing}")

    df = df[["Invoice#", "Tickets", "Date", "Name", "Balance Due"]].copy()
    df["Invoice#"] = df["Invoice#"].astype(str).str.strip()
    df["Tickets"] = df["Tickets"].astype(str).fillna("").str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()

    # keep Balance Due numeric-friendly (but preserve display later)
    df["Balance Due"] = pd.to_numeric(df["Balance Due"], errors="coerce").fillna(0.0)

    return df

def read_openinvoice(file_obj) -> pd.DataFrame:
    df = _read_any(file_obj)
    df.columns = df.columns.astype(str).str.strip()

    if "Ticket" not in df.columns:
        raise ValueError('OpenInvoice missing column: "Ticket"')

    df = df.copy()
    df["Ticket"] = df["Ticket"].apply(_norm_ticket)
    df = df[df["Ticket"] != ""]
    return df

def compute_status(axon_ticket_list: list[str], openinvoice_ticket_set: set[str]) -> str:
    if not axon_ticket_list:
        return "Not Ready"

    normalized = [_norm_ticket(t) for t in axon_ticket_list if _norm_ticket(t)]
    if not normalized:
        return "Not Ready"

    matched = sum(1 for t in normalized if t in openinvoice_ticket_set)

    if matched == 0:
        return "Not Ready"
    if matched == len(normalized):
        return "Ready to Flip"
    return "Pending"

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/approved")
async def approved_invoice_status(
    axon: UploadFile = File(...),
    openinvoice: UploadFile = File(...)
):
    ax = read_axon(axon.file)
    oi = read_openinvoice(openinvoice.file)

    oi_set = set(oi["Ticket"].dropna().tolist())

    # status per invoice row
    statuses = []
    matched_counts = []
    total_counts = []

    for tickets_cell in ax["Tickets"].tolist():
        ticket_list = _split_axon_tickets(tickets_cell)
        normalized = [_norm_ticket(t) for t in ticket_list if _norm_ticket(t)]
        total = len(normalized)
        matched = sum(1 for t in normalized if t in oi_set)

        total_counts.append(total)
        matched_counts.append(matched)
        statuses.append(compute_status(ticket_list, oi_set))

    out = ax.copy()
    out["Status"] = statuses
    # optional helpers (super useful for “why is this pending?”)
    out["Matched Tickets"] = matched_counts
    out["Total Tickets"] = total_counts

    # order exactly how you asked (+ two helper columns at the end)
    out = out[["Invoice#", "Tickets", "Date", "Name", "Balance Due", "Status", "Matched Tickets", "Total Tickets"]]

    summary = pd.DataFrame([
        {"Metric": "Invoices (AXON rows)", "Value": len(out)},
        {"Metric": "Ready to Flip", "Value": int((out["Status"] == "Ready to Flip").sum())},
        {"Metric": "Pending", "Value": int((out["Status"] == "Pending").sum())},
        {"Metric": "Not Ready", "Value": int((out["Status"] == "Not Ready").sum())},
    ])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        out.to_excel(writer, index=False, sheet_name="Invoice Status")

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=approved_invoice_status.xlsx"},
    )