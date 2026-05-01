import streamlit as st
import pdfplumber
import re
import json
import os
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="7501 Entry Tracker", layout="wide")
DATA_FILE = "entries.json"

# ── Data helpers ───────────────────────────────────────────────────────────────
def load_entries():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return []

def save_entries(entries):
    with open(DATA_FILE, "w") as f:
        json.dump(entries, f, indent=2)

# ── PDF extraction ─────────────────────────────────────────────────────────────
def extract_text_from_pdf(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except:
        return ""

def grab(pattern, text, default="", group=1):
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return m.group(group).strip() if m else default

COUNTRY_MAP = {
    "CN": "China", "MX": "Mexico", "BR": "Brazil", "DE": "Germany",
    "JP": "Japan", "KR": "South Korea", "IN": "India", "TW": "Taiwan",
    "CA": "Canada", "GB": "United Kingdom", "FR": "France", "IT": "Italy",
    "VN": "Vietnam", "TH": "Thailand", "MY": "Malaysia", "PH": "Philippines",
}

def parse_7501(text):
    # Entry number: 101-3526528-9
    entry_number = grab(r'\b(\d{3}-\d{7}-\d)\b', text)

    # Entry date: last date on the data line containing the entry number
    # Line: "101-3526528-9 01 ABI/P 05/05/26 457 8 3901 05/05/2026"
    entry_date = grab(r'\b\d{3}-\d{7}-\d\b[^\n]+(\d{2}/\d{2}/\d{4})', text)

    # Import date + country from carrier line
    # Line: "CMA CGM SWORDFISH (CMDU) 11 CN 05/04/2026"
    carrier_line = re.search(r'CMA CGM[^\n]*\b([A-Z]{2})\b\s+(\d{2}/\d{2}/\d{4})', text)
    if not carrier_line:
        carrier_line = re.search(r'(?:Mode Of Transport|Importing Carrier)[^\n]*\n[^\n]*\b([A-Z]{2})\b\s+(\d{2}/\d{2}/\d{4})', text)
    country_code = carrier_line.group(1) if carrier_line else ""
    import_date  = carrier_line.group(2) if carrier_line else ""
    country      = COUNTRY_MAP.get(country_code, country_code)

    # Broker
    broker = grab(r'(KUEHNE\s*\+\s*NAGEL[^\n,\.]*)', text).split('\n')[0].strip()
    if not broker:
        broker = grab(r'Broker/Filer Information[^\n]*\n([^\n]+)', text)

    # Broker file number: BUSxxxxxxx
    broker_number = grab(r'\b(BUS\d+)\b', text)

    # Invoice number
    invoice = grab(r'Invoice\s+Number\s+(\S+)', text)

    # Supplier: "MANUFACTURER SUPPLIER\nNINGBO LONGYUAN CO NINGBO LONGYUAN CO"
    sup_match = re.search(r'MANUFACTURER\s+SUPPLIER\s*\n([^\n]+)', text)
    if sup_match:
        # Name appears twice - just take the first occurrence
        raw = sup_match.group(1).strip()
        half = len(raw) // 2
        supplier = raw[:half].strip() if raw[:half].strip() == raw[half:].strip() else raw.split('  ')[0].strip()
    else:
        supplier = grab(r'SUPPLIER\s*\n([^\n]+)', text)

    # Part number: E1060047602A0 format
    part_match = re.search(r'\b(E\d{10}[A-Z0-9]\d)\b', text)
    part_number = part_match.group(1) if part_match else ""

    # Quantity: "3200 NO"
    quantity = grab(r'\b(\d{3,5})\s+NO\b', text)

    # Invoice value
    inv_value = grab(r'Invoice Value USD\s+([\d,]+\.?\d*)', text)

    # Total entered value
    total_value_str = grab(r'Total Entered Value \(Invoice\)\s+([\d,]+\.?\d*)', text)
    try:
        total_value_float = float(total_value_str.replace(',', ''))
    except:
        total_value_float = 0.0

    # Total duty: appears as "Total Other Fees 12670.50" (this is the duty subtotal)
    total_duty_str = grab(r'Total Other Fees\s+([\d,]+\.\d{2})', text)
    try:
        total_duty_float = float(total_duty_str.replace(',', ''))
    except:
        total_duty_float = 0.0

    tariff_lines = parse_tariff_lines(text, total_value_float)

    return {
        "entry_number":  entry_number,
        "entry_date":    entry_date,
        "import_date":   import_date,
        "broker":        broker,
        "broker_number": broker_number,
        "supplier":      supplier,
        "country":       country,
        "invoice":       invoice,
        "part_number":   part_number,
        "quantity":      quantity,
        "invoice_value": inv_value,
        "total_value":   total_value_float,
        "total_duty":    total_duty_float,
        "tariff_lines":  tariff_lines,
    }

def parse_tariff_lines(text, entered_value):
    """
    Real format from PDF (per line):
      9903.88.01 4473 KG 0 25% 6,335.25   -> keep
      9903.03.06 0.00 0 Free 0.00          -> skip (Free, duty=0)
      9903.82.09 0.00 0 25% 6,335.25       -> keep
      8481.90.9060 3985.00 KG 25,341 Free 0.00 -> skip (Free)
    """
    lines = []
    seen = set()

    for m in re.finditer(
        r'([0-9]{4}\.[0-9]{2}\.[0-9]{2,4})\s+.*?([0-9]+(?:\.[0-9]+)?)\s*%\s+([0-9,]+\.[0-9]{2})',
        text, re.ASCII
    ):
        hts  = m.group(1)
        rate = float(m.group(2))
        duty = float(m.group(3).replace(',', ''))

        if rate < 1.0 or duty == 0.0 or hts in seen:
            continue
        seen.add(hts)

        # Try to extract entered value from the specific HTS line
        ev = entered_value
        ev_match = re.search(
            re.escape(hts) + r'\s+([0-9,]+\.?[0-9]*)\s+\w+\s+([0-9,]+\.?[0-9]*)\s+[0-9]+\s*%',
            text, re.ASCII
        )
        if ev_match:
            try:
                candidate = float(ev_match.group(2).replace(',', ''))
                if candidate > 0:
                    ev = candidate
            except:
                pass

        lines.append({
            "hts":           hts,
            "entered_value": ev,
            "rate_pct":      rate,
            "duty_amount":   duty,
        })

    return lines

# ── Excel export ───────────────────────────────────────────────────────────────
def build_excel(entries):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "7501 Log"

    hfill = PatternFill("solid", start_color="1F4E79")
    hfont = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hal   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dfont = Font(name="Arial", size=10)
    cal   = Alignment(horizontal="center", vertical="center")

    headers = [
        "Entry #", "Entry Date", "Import Date", "Broker", "Broker #",
        "Supplier", "Country of Origin", "Invoice #", "Part Number",
        "Description", "Quantity", "Invoice Value", "Total Entered Value",
        "HTS Code", "Entered Value", "Rate (%)", "Duty (USD)",
        "Total Duty (USD)", "Date Logged"
    ]
    widths = [16, 12, 12, 20, 14, 22, 16, 18, 18, 28, 10, 14, 16, 14, 14, 10, 14, 14, 16]

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(1, c, h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = hal
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30

    row = 2
    for e in entries:
        tariff_lines = e.get("tariff_lines", [])
        if not tariff_lines:
            tariff_lines = [{"hts": "", "entered_value": "", "rate_pct": "", "duty_amount": ""}]

        for t in tariff_lines:
            values = [
                e.get("entry_number"), e.get("entry_date"), e.get("import_date"),
                e.get("broker"), e.get("broker_number"), e.get("supplier"),
                e.get("country"), e.get("invoice"), e.get("part_number"),
                e.get("description"), e.get("quantity"), e.get("invoice_value"),
                e.get("total_value"),
                t.get("hts"), t.get("entered_value"),
                t.get("rate_pct"), t.get("duty_amount"),
                e.get("total_duty"), e.get("date_logged")
            ]
            for c, v in enumerate(values, 1):
                cell = ws.cell(row, c, v)
                cell.font = dfont
                cell.alignment = cal
            row += 1

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Session state ──────────────────────────────────────────────────────────────
if "extracted" not in st.session_state:
    st.session_state.extracted = {}
if "to_delete" not in st.session_state:
    st.session_state.to_delete = []

# ── UI ─────────────────────────────────────────────────────────────────────────
st.title("📦 7501 Entry Tracker")
tab1, tab2 = st.tabs(["➕ New Entry", "📋 View All Entries"])

# ════════════════════════════════════════════════════════
# TAB 1 — NEW ENTRY
# ════════════════════════════════════════════════════════
with tab1:
    uploaded = st.file_uploader("Upload 7501 PDF", type="pdf")
    d = {}
    if uploaded:
        text = extract_text_from_pdf(uploaded)
        d = parse_7501(text)
        st.success("✅ PDF parsed — review and correct anything below before saving.")

    st.subheader("Entry Header")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        entry_number  = st.text_input("Entry #",            value=d.get("entry_number", ""))
        broker        = st.text_input("Broker",              value=d.get("broker", ""))
    with c2:
        entry_date    = st.text_input("Entry Date",          value=d.get("entry_date", ""))
        broker_number = st.text_input("Broker #",            value=d.get("broker_number", ""))
    with c3:
        import_date   = st.text_input("Import Date",         value=d.get("import_date", ""))
        supplier      = st.text_input("Supplier",            value=d.get("supplier", ""))
    with c4:
        country       = st.text_input("Country of Origin",   value=d.get("country", ""))
        invoice       = st.text_input("Invoice #",           value=d.get("invoice", ""))

    st.markdown("---")
    st.subheader("Merchandise")
    m1, m2, m3 = st.columns(3)
    with m1:
        part_number   = st.text_input("Part Number",         value=d.get("part_number", ""))
    with m2:
        quantity      = st.text_input("Quantity",            value=d.get("quantity", ""))
    with m3:
        invoice_value = st.text_input("Invoice Value (USD)", value=d.get("invoice_value", ""))

    st.markdown("---")
    st.subheader("Tariff Lines")
    st.caption("Only lines with a duty charge ≥ 1% — no Free lines, no MPF/HMF.")

    # Initialize tariff lines from extraction or blank
    if "tariff_lines" not in st.session_state or uploaded:
        st.session_state.tariff_lines = d.get("tariff_lines", [{}]) or [{}]

    th1, th2, th3, th4 = st.columns([2.5, 2, 1.5, 2])
    th1.markdown("**HTS Code**")
    th2.markdown("**Entered Value ($)**")
    th3.markdown("**Rate (%)**")
    th4.markdown("**Duty (USD)**")

    updated_tariffs = []
    for i, t in enumerate(st.session_state.tariff_lines):
        tc1, tc2, tc3, tc4 = st.columns([2.5, 2, 1.5, 2])
        hts = tc1.text_input("", value=t.get("hts", ""),              key=f"hts_{i}",  label_visibility="collapsed")
        ev  = tc2.number_input("", value=float(t.get("entered_value", 0) or 0), min_value=0.0, key=f"ev_{i}",   label_visibility="collapsed", format="%.2f")
        rt  = tc3.number_input("", value=float(t.get("rate_pct", 0) or 0),      min_value=0.0, max_value=100.0, key=f"rt_{i}", label_visibility="collapsed", format="%.2f")
        dy  = tc4.number_input("", value=float(t.get("duty_amount", 0) or 0),   min_value=0.0, key=f"dy_{i}",   label_visibility="collapsed", format="%.2f")
        updated_tariffs.append({"hts": hts, "entered_value": ev, "rate_pct": rt, "duty_amount": dy})

    st.session_state.tariff_lines = updated_tariffs

    ba, bb, _ = st.columns([1, 1, 5])
    if ba.button("➕ Add Tariff Line"):
        st.session_state.tariff_lines.append({})
        st.rerun()
    if bb.button("🗑 Remove Last") and len(st.session_state.tariff_lines) > 1:
        st.session_state.tariff_lines.pop()
        st.rerun()

    st.markdown("---")
    total_value  = d.get("total_value", 0.0)
    total_duty   = d.get("total_duty", 0.0)

    tv_input = st.number_input("Total Entered Value (USD)", value=float(total_value), min_value=0.0, format="%.2f")
    td_input = st.number_input("Total Duty (USD)",          value=float(total_duty),  min_value=0.0, format="%.2f")

    eff_rate = round((td_input / tv_input * 100) if tv_input else 0, 2)
    s1, s2, s3 = st.columns(3)
    s1.metric("Total Entered Value", f"${tv_input:,.2f}")
    s2.metric("Total Duty",          f"${td_input:,.2f}")
    s3.metric("Effective Duty Rate", f"{eff_rate:.2f}%")

    st.markdown("---")
    if st.button("✅ Save Entry", type="primary"):
        if not entry_number:
            st.error("Entry number is required.")
        else:
            entries = load_entries()
            # Check for duplicate
            existing = [e["entry_number"] for e in entries]
            if entry_number in existing:
                st.warning(f"Entry {entry_number} already exists in the log.")
            else:
                entries.append({
                    "date_logged":    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "entry_number":   entry_number,
                    "entry_date":     entry_date,
                    "import_date":    import_date,
                    "broker":         broker,
                    "broker_number":  broker_number,
                    "supplier":       supplier,
                    "country":        country,
                    "invoice":        invoice,
                    "part_number":    part_number,
                    "quantity":       quantity,
                    "invoice_value":  invoice_value,
                    "total_value":    tv_input,
                    "total_duty":     td_input,
                    "eff_rate":       eff_rate,
                    "tariff_lines":   [t for t in st.session_state.tariff_lines if t.get("hts")],
                    "filename":       uploaded.name if uploaded else "",
                })
                save_entries(entries)
                st.session_state.tariff_lines = [{}]
                st.success(f"✅ Entry {entry_number} saved!")
                st.rerun()

# ════════════════════════════════════════════════════════
# TAB 2 — VIEW ALL
# ════════════════════════════════════════════════════════
with tab2:
    entries = load_entries()

    if not entries:
        st.info("No entries logged yet.")
    else:
        # ── Summary metrics ────────────────────────────────────
        total_entries = len(entries)
        total_value   = sum(e.get("total_value", 0) for e in entries)
        total_duty    = sum(e.get("total_duty",  0) for e in entries)
        avg_rate      = round((total_duty / total_value * 100) if total_value else 0, 2)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Entries",       total_entries)
        m2.metric("Total Entered Value", f"${total_value:,.2f}")
        m3.metric("Total Duties Paid",   f"${total_duty:,.2f}")
        m4.metric("Avg Effective Rate",  f"{avg_rate:.2f}%")

        st.markdown("---")

        # ── Filters ────────────────────────────────────────────
        with st.expander("🔽 Filters", expanded=True):
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                search = st.text_input("🔍 Search", placeholder="Entry #, supplier, part #, HTS, country...")
            with fc2:
                all_brokers = sorted(set(e.get("broker", "") for e in entries if e.get("broker")))
                broker_filter = st.multiselect("Broker", all_brokers)
            with fc3:
                all_countries = sorted(set(e.get("country", "") for e in entries if e.get("country")))
                country_filter = st.multiselect("Country of Origin", all_countries)

            fc4, fc5, fc6 = st.columns(3)
            with fc4:
                all_suppliers = sorted(set(e.get("supplier", "") for e in entries if e.get("supplier")))
                supplier_filter = st.multiselect("Supplier", all_suppliers)
            with fc5:
                date_from = st.date_input("Entry Date From", value=None)
            with fc6:
                date_to   = st.date_input("Entry Date To",   value=None)

            fc7, fc8 = st.columns(2)
            with fc7:
                min_duty = st.number_input("Min Total Duty ($)", value=0.0, min_value=0.0, format="%.2f")
            with fc8:
                max_duty = st.number_input("Max Total Duty ($)", value=0.0, min_value=0.0, format="%.2f",
                                           help="Leave at 0 for no max")

        # ── Apply filters ──────────────────────────────────────
        filtered = entries

        if search:
            s = search.lower()
            filtered = [e for e in filtered if
                s in e.get("entry_number", "").lower() or
                s in e.get("supplier", "").lower() or
                s in e.get("country", "").lower() or
                s in e.get("part_number", "").lower() or
                s in e.get("description", "").lower() or
                s in e.get("broker", "").lower() or
                s in e.get("invoice", "").lower() or
                any(s in t.get("hts", "").lower() for t in e.get("tariff_lines", []))]

        if broker_filter:
            filtered = [e for e in filtered if e.get("broker") in broker_filter]

        if country_filter:
            filtered = [e for e in filtered if e.get("country") in country_filter]

        if supplier_filter:
            filtered = [e for e in filtered if e.get("supplier") in supplier_filter]

        if date_from:
            filtered = [e for e in filtered if e.get("entry_date", "") >= date_from.strftime("%m/%d/%Y")]

        if date_to:
            filtered = [e for e in filtered if e.get("entry_date", "") <= date_to.strftime("%m/%d/%Y")]

        if min_duty > 0:
            filtered = [e for e in filtered if e.get("total_duty", 0) >= min_duty]

        if max_duty > 0:
            filtered = [e for e in filtered if e.get("total_duty", 0) <= max_duty]

        st.markdown(f"Showing **{len(filtered)}** of **{total_entries}** entries")

        # ── Delete controls ────────────────────────────────────
        st.markdown("**Select entries to delete:**")
        to_delete = []
        for i, e in enumerate(reversed(filtered)):
            label = f"📄 {e['entry_number']}  |  {e.get('supplier', '')}  |  {e.get('country', '')}  |  {e.get('entry_date', '')}  |  ${e.get('total_value', 0):,.2f}  |  Duty: ${e.get('total_duty', 0):,.2f} ({e.get('eff_rate', 0):.1f}%)"

            col_check, col_expand = st.columns([0.05, 0.95])
            selected = col_check.checkbox("", key=f"del_{i}_{e['entry_number']}")
            if selected:
                to_delete.append(e["entry_number"])

            with col_expand.expander(label):
                h1, h2, h3, h4 = st.columns(4)
                h1.markdown(f"**Entry #:** {e.get('entry_number')}")
                h1.markdown(f"**Entry Date:** {e.get('entry_date')}")
                h1.markdown(f"**Import Date:** {e.get('import_date')}")
                h2.markdown(f"**Broker:** {e.get('broker')}")
                h2.markdown(f"**Broker #:** {e.get('broker_number')}")
                h2.markdown(f"**Invoice #:** {e.get('invoice')}")
                h3.markdown(f"**Supplier:** {e.get('supplier')}")
                h3.markdown(f"**Country:** {e.get('country')}")
                h3.markdown(f"**Part #:** {e.get('part_number')}")
                h4.markdown(f"**Quantity:** {e.get('quantity')}")
                h4.markdown(f"**Logged:** {e.get('date_logged')}")

                tariff_lines = e.get("tariff_lines", [])
                if tariff_lines:
                    st.markdown("**Tariff Lines:**")
                    tl1, tl2, tl3, tl4 = st.columns([2.5, 2, 1.5, 2])
                    tl1.markdown("**HTS Code**")
                    tl2.markdown("**Entered Value**")
                    tl3.markdown("**Rate**")
                    tl4.markdown("**Duty**")
                    for t in tariff_lines:
                        tl1, tl2, tl3, tl4 = st.columns([2.5, 2, 1.5, 2])
                        tl1.write(t.get("hts", ""))
                        tl2.write(f"${t.get('entered_value', 0):,.2f}")
                        tl3.write(f"{t.get('rate_pct', 0):.1f}%")
                        tl4.write(f"${t.get('duty_amount', 0):,.2f}")

                st.markdown("---")
                sv1, sv2, sv3 = st.columns(3)
                sv1.metric("Total Entered Value", f"${e.get('total_value', 0):,.2f}")
                sv2.metric("Total Duty",          f"${e.get('total_duty', 0):,.2f}")
                sv3.metric("Effective Rate",       f"{e.get('eff_rate', 0):.2f}%")

        # ── Delete button ──────────────────────────────────────
        if to_delete:
            st.warning(f"{len(to_delete)} entry/entries selected for deletion.")
            if st.button("🗑 Delete Selected Entries", type="primary"):
                updated = [e for e in entries if e["entry_number"] not in to_delete]
                save_entries(updated)
                st.success(f"Deleted {len(to_delete)} entry/entries.")
                st.rerun()

        # ── Export ─────────────────────────────────────────────
        st.markdown("---")
        if st.button("📥 Export All to Excel"):
            buf = build_excel(entries)
            st.download_button(
                label="⬇️ Download Excel",
                data=buf,
                file_name=f"7501_log_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
