function log(el, msg) {
  el.textContent = msg;
}

function normalizeTicket(x) {
  if (x === null || x === undefined) return "";
  let s = String(x).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s.replaceAll(",", "").replaceAll(" ", "");
}

function splitAxonTickets(cell) {
  if (cell === null || cell === undefined) return [];
  const s = String(cell).trim();
  if (!s) return [];
  return s.split(",").map(p => p.trim()).filter(Boolean);
}

function normalizeInvoice(x) {
  if (x === null || x === undefined) return "";
  let s = String(x).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s.replaceAll(",", "").replaceAll(" ", "");
}

async function readTable(file) {
  const name = file.name.toLowerCase();
  const data = await file.arrayBuffer();

  if (name.endsWith(".csv")) {
    const text = new TextDecoder("utf-8").decode(data);
    const wb = XLSX.read(text, { type: "string" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  }

  const wb = XLSX.read(data, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function findHeaderRow(rows, required) {
  // When exported with metadata, headers might be shifted. This handles XLSX->json less well,
  // so we assume your CSV/XLSX first sheet is already tabular. If needed later, we can add
  // a raw sheet scan. For now: validate columns exist.
  return rows;
}

function downloadWorkbook(sheets, filename) {
  const wb = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }
  XLSX.writeFile(wb, filename);
}

/* ---------- Approved Tool (invoice-level status) ---------- */
async function runApproved(axonFile, oiFile, logEl) {
  log(logEl, "Reading files…");
  const axonRows = await readTable(axonFile);
  const oiRows = await readTable(oiFile);

  const requiredAxon = ["Invoice#", "Tickets", "Date", "Name", "Balance Due"];
  for (const c of requiredAxon) {
    if (!axonRows.length || !(c in axonRows[0])) throw new Error(`AXON missing column: ${c}`);
  }
  if (!oiRows.length || !("Ticket" in oiRows[0])) throw new Error('OpenInvoice missing column: Ticket');

  const oiSet = new Set(
    oiRows
      .map(r => normalizeTicket(r["Ticket"]))
      .filter(Boolean)
  );

  const out = axonRows.map(r => {
    const ticketList = splitAxonTickets(r["Tickets"]);
    const normalized = ticketList.map(normalizeTicket).filter(Boolean);
    const total = normalized.length;
    const matched = normalized.filter(t => oiSet.has(t)).length;

    let status = "Not Ready";
    if (total > 0 && matched === total) status = "Ready to Flip";
    else if (total > 0 && matched > 0) status = "Pending";

    return {
      "Invoice#": r["Invoice#"],
      "Tickets": r["Tickets"],
      "Date": r["Date"],
      "Name": r["Name"],
      "Balance Due": r["Balance Due"],
      "Status": status
    };
  });

  const summary = [
    { Metric: "Invoices (AXON rows)", Value: out.length },
    { Metric: "Ready to Flip", Value: out.filter(r => r.Status === "Ready to Flip").length },
    { Metric: "Pending", Value: out.filter(r => r.Status === "Pending").length },
    { Metric: "Not Ready", Value: out.filter(r => r.Status === "Not Ready").length },
  ];

  log(logEl, "Generating Excel…");
  downloadWorkbook(
    [["Summary", summary], ["Invoice Status", out]],
    "approved_invoice_status.xlsx"
  );
  log(logEl, "Done ✅ Download started.");
}

/* ---------- Remittance Tool (customer mapping + totals) ---------- */
async function runRemit(axonFile, remFile, logEl) {
  log(logEl, "Reading files…");
  const axonRows = await readTable(axonFile);
  const remRows = await readTable(remFile);

  const requiredAxon = ["Invoice#", "Name"];
  for (const c of requiredAxon) {
    if (!axonRows.length || !(c in axonRows[0])) throw new Error(`AXON missing column: ${c}`);
  }

  const requiredRem = ["Reference", "Net Amount"];
  // your remittance sometimes uses "Reference/Text"
  const hasReference = remRows.length && ("Reference" in remRows[0] || "Reference/Text" in remRows[0]);
  if (!hasReference) throw new Error('Remittance missing column: Reference (or Reference/Text)');
  if (!remRows.length || !("Net Amount" in remRows[0])) throw new Error('Remittance missing column: Net Amount');

  const invToCustomer = new Map(
    axonRows
      .map(r => [normalizeInvoice(r["Invoice#"]), String(r["Name"]).trim()])
      .filter(([k]) => k)
  );

  const PIONEER = "Pioneer Natural Resources";
  const XTO = "XTO Energy";

  const out = [];
  let pioneerTotal = 0;
  let xtoTotal = 0;

  for (const r of remRows) {
    const refRaw = ("Reference" in r) ? r["Reference"] : r["Reference/Text"];
    const inv = normalizeInvoice(refRaw);
    const amt = Number(String(r["Net Amount"]).replaceAll(",", "")) || 0;
    const customer = invToCustomer.get(inv) || "NOT FOUND IN AXON";

    const pioneer = customer === PIONEER ? amt : "";
    const xto = customer === XTO ? amt : "";

    if (customer === PIONEER) pioneerTotal += amt;
    if (customer === XTO) xtoTotal += amt;

    out.push({
      "Reference": inv,
      "Net Amount": amt,
      "Customer": customer,
      "PioneerNaturalResources": pioneer,
      "XTO": xto
    });
  }

  out.push({
    "Reference": "",
    "Net Amount": "",
    "Customer": "TOTAL",
    "PioneerNaturalResources": pioneerTotal,
    "XTO": xtoTotal
  });

  const notFound = out.filter(r => r.Customer === "NOT FOUND IN AXON");

  log(logEl, "Generating Excel…");
  downloadWorkbook(
    [["Customer Breakdown", out], ["Not Found in AXON", notFound]],
    "remittance_breakdown.xlsx"
  );
  log(logEl, "Done ✅ Download started.");
}

/* ---------- UI wiring ---------- */
const tabApproved = document.getElementById("tab-approved");
const tabRemit = document.getElementById("tab-remit");
const panelApproved = document.getElementById("panel-approved");
const panelRemit = document.getElementById("panel-remit");

tabApproved.onclick = () => {
  tabApproved.classList.add("active");
  tabRemit.classList.remove("active");
  panelApproved.style.display = "";
  panelRemit.style.display = "none";
};

tabRemit.onclick = () => {
  tabRemit.classList.add("active");
  tabApproved.classList.remove("active");
  panelRemit.style.display = "";
  panelApproved.style.display = "none";
};

document.getElementById("btn-approved").onclick = async () => {
  const axon = document.getElementById("approved-axon").files[0];
  const oi = document.getElementById("approved-oi").files[0];
  const logEl = document.getElementById("log-approved");
  try {
    if (!axon || !oi) throw new Error("Please select both files.");
    await runApproved(axon, oi, logEl);
  } catch (e) {
    log(logEl, `❌ ${e.message}`);
  }
};

document.getElementById("btn-remit").onclick = async () => {
  const axon = document.getElementById("remit-axon").files[0];
  const rem = document.getElementById("remit-rem").files[0];
  const logEl = document.getElementById("log-remit");
  try {
    if (!axon || !rem) throw new Error("Please select both files.");
    await runRemit(axon, rem, logEl);
  } catch (e) {
    log(logEl, `❌ ${e.message}`);
  }
};