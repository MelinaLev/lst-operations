from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from io import BytesIO

app = FastAPI()
templates = Jinja2Templates(directory="templates")

def normalize_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\s+", "", regex=True)
        .str.replace("-", "", regex=False)
        .str.upper()
    )

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/compare")
async def compare(
    axon: UploadFile = File(...),
    approved: UploadFile = File(...)
):
    # Read files (supports .xlsx; add CSV support if you need it)
    axon_df = pd.read_excel(axon.file)
    appr_df = pd.read_excel(approved.file)

    # TODO: set these to your real column names
    AXON_KEY_COL = "TicketNumber"
    APPR_KEY_COL = "TicketNumber"

    if AXON_KEY_COL not in axon_df.columns:
        raise ValueError(f"AXON missing column: {AXON_KEY_COL}")
    if APPR_KEY_COL not in appr_df.columns:
        raise ValueError(f"Approved list missing column: {APPR_KEY_COL}")

    axon_df["_KEY"] = normalize_series(axon_df[AXON_KEY_COL])
    appr_df["_KEY"] = normalize_series(appr_df[APPR_KEY_COL])

    axon_keys = set(axon_df["_KEY"].dropna())
    appr_keys = set(appr_df["_KEY"].dropna())

    matched_keys = axon_keys & appr_keys
    missing_in_approved = axon_keys - appr_keys
    missing_in_axon = appr_keys - axon_keys

    matched_axon = axon_df[axon_df["_KEY"].isin(matched_keys)].copy()
    matched_appr = appr_df[appr_df["_KEY"].isin(matched_keys)].copy()

    miss_appr = axon_df[axon_df["_KEY"].isin(missing_in_approved)].copy()
    miss_axon = appr_df[appr_df["_KEY"].isin(missing_in_axon)].copy()

    summary = pd.DataFrame([
        {"Metric": "AXON rows", "Value": len(axon_df)},
        {"Metric": "Approved rows", "Value": len(appr_df)},
        {"Metric": "Matched unique keys", "Value": len(matched_keys)},
        {"Metric": "Missing in Approved", "Value": len(missing_in_approved)},
        {"Metric": "Missing in AXON", "Value": len(missing_in_axon)},
    ])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        matched_axon.to_excel(writer, index=False, sheet_name="Matched_AXON")
        matched_appr.to_excel(writer, index=False, sheet_name="Matched_Approved")
        miss_appr.to_excel(writer, index=False, sheet_name="Missing_in_Approved")
        miss_axon.to_excel(writer, index=False, sheet_name="Missing_in_AXON")

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=approved_compare_results.xlsx"},
    )