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

def load_entries():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return []

def save_entries(entries):
    with open(DATA_FILE, "w") as f:
        json.dump(entries, f, indent=2)

def extract_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except:
        return ""

def extract_field(pattern, text, default=""):
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(1).strip() if m else default

def parse_7501(text):
    return {
        "entry_number":  extract_field(r'entry\s*(?:no|number|#)?[\s:.]*([A-Z0-9\-]{8,20})', text),
        "entry_date":    extract_field(r'entry\s*date[\s:.]*(\d{1,2}/\d{1,2}/\d{2,4})', text),
        "import_date":   extract_field(r'import\s*date[\s:.]*(\d{1,2}/\d{1,2}/\d{2,4})', text),
        "broker":        extract_field(r'broker[\s:]*([\w\s]+?)(?:\n|filer|#)', text),
        "broker_number": extract_field(r'broker\s*(?:no|number|#)[\s:.]*([0-9\-]+)', text),
        "supplier":      extract_field(r'(?:supplier|shipper|exporter)[\s:.]*([A-Za-z0-9\s,\.]+?)(?:\n)', text),
        "country":       extract_field(r'country\s*of\s*origin[\s:.]*([A-Za-z\s]+?)(?:\n)', text),
        "invoice":       extract_field(r'invoice\s*(?:no|number|#)?[\s:.]*([A-Z0-9\-]+)', text),
    }

def build_excel(entries):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "7501 Log"
    hfill = PatternFill("solid", start_color="1F4E79")
    hfont = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hal   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dfont = Font(name="Arial", size=10)
    headers = [
        "Entry #", "Entry Date", "Import Date", "Broker", "Supplier Name",
        "Country of Origin", "Broker #", "Invoice #", "Part Number", "DESC",
        "Quantity", "Price", "Total", "Tariff 1", "Tariff 2", "Tariff 3",
        "Rate 1 (%)", "Rate 2 (%)", "Rate 3 (%)",
        "Duty 1 (USD)", "Duty 2 (USD)", "Duty 3 (USD)",
        "Total Duty (USD)", "Total Duty %"
    ]
    widths = [14,12,12,10,16,16,12,14,18,24,10,10,12,14,14,14,10,10,10,12,12,12,14,12]
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(1, c, h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = hal
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30
    row = 2
    for e in entries:
        for item in e.get("line_items", []):
            total      = item.get("total", 0)
            r1         = item.get("rate1", 0) / 100
            r2         = item.get("rate2", 0) / 100
            r3         = item.get("rate3", 0) / 100
            duty1      = round(total * r1, 2)
            duty2      = round(total * r2, 2)
            duty3      = round(total * r3, 2)
            total_duty = round(duty1 + duty2 + duty3, 2)
            total_pct  = round((r1 + r2 + r3) * 100, 2)
            values = [
                e.get("entry_number"), e.get("entry_date"), e.get("import_date"),
                e.get("broker"), e.get("supplier"), e.get("country"),
                e.get("broker_number"), e.get("invoice"),
                item.get("part_number"), item.get("description"),
                item.get("quantity"), item.get("price"), total,
                item.get("tariff1"), item.get("tariff2"), item.get("tariff3"),
                item.get("rate1"), item.get("rate2"), item.get("rate3"),
                duty1, duty2, duty3, total_duty, total_pct
            ]
            for c, v in enumerate(values, 1):
                cell = ws.cell(row, c, v)
                cell.font = dfont
                cell.alignment = Alignment(horizontal="center", vertical="center")
            row += 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

if "line_items" not in st.session_state:
    st.session_state.line_items = [{}]
if "header" not in st.session_state:
    st.session_state.header = {}

st.title("📦 7501 Entry Tracker")

tab1, tab2 = st.tabs(["➕ New Entry", "📋 View All Entries"])

with tab1:
    uploaded = st.file_uploader("Upload 7501 PDF (optional)", type="pdf")
    extracted = {}
    if uploaded:
        text = extract_text(uploaded)
        extracted = parse_7501(text)
        st.success("✅ PDF parsed — review and correct fields below.")

    h = extracted or st.session_state.header
    st.subheader("Entry Header")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        entry_number  = st.text_input("Entry #",          value=h.get("entry_number", ""))
        broker        = st.text_input("Broker",            value=h.get("broker", ""))
    with c2:
        entry_date    = st.text_input("Entry Date",        value=h.get("entry_date", ""))
        supplier      = st.text_input("Supplier Name",     value=h.get("supplier", ""))
    with c3:
        import_date   = st.text_input("Import Date",       value=h.get("import_date", ""))
        country       = st.text_input("Country of Origin", value=h.get("country", ""))
    with c4:
        broker_number = st.text_input("Broker #",          value=h.get("broker_number", ""))
        invoice       = st.text_input("Invoice #",         value=h.get("invoice", ""))

    st.markdown("---")
    st.subheader("Line Items")
    st.caption("Add one row per part number.")

    cols = st.columns([2, 2.5, 1, 1, 1.5, 2, 2, 2, 1, 1, 1])
    for col, label in zip(cols, ["Part #", "Description", "Qty", "Price", "Total",
                                   "Tariff 1", "Tariff 2", "Tariff 3",
                                   "Rate 1%", "Rate 2%", "Rate 3%"]):
        col.markdown(f"**{label}**")

    updated_items = []
    for i, item in enumerate(st.session_state.line_items):
        cols = st.columns([2, 2.5, 1, 1, 1.5, 2, 2, 2, 1, 1, 1])
        part  = cols[0].text_input("", value=item.get("part_number",""),       key=f"part_{i}", label_visibility="collapsed")
        desc  = cols[1].text_input("", value=item.get("description",""),       key=f"desc_{i}", label_visibility="collapsed")
        qty   = cols[2].number_input("", value=float(item.get("quantity",0)),  key=f"qty_{i}",  label_visibility="collapsed", min_value=0.0)
        price = cols[3].number_input("", value=float(item.get("price",0)),     key=f"price_{i}",label_visibility="collapsed", min_value=0.0, format="%.2f")
        total = qty * price
        cols[4].markdown(f"<div style='padding-top:8px'>${total:,.2f}</div>", unsafe_allow_html=True)
        t1    = cols[5].text_input("", value=item.get("tariff1",""),           key=f"t1_{i}",   label_visibility="collapsed")
        t2    = cols[6].text_input("", value=item.get("tariff2",""),           key=f"t2_{i}",   label_visibility="collapsed")
        t3    = cols[7].text_input("", value=item.get("tariff3",""),           key=f"t3_{i}",   label_visibility="collapsed")
        r1    = cols[8].number_input("",  value=float(item.get("rate1",0)),    key=f"r1_{i}",   label_visibility="collapsed", min_value=0.0, max_value=100.0)
        r2    = cols[9].number_input("",  value=float(item.get("rate2",0)),    key=f"r2_{i}",   label_visibility="collapsed", min_value=0.0, max_value=100.0)
        r3    = cols[10].number_input("", value=float(item.get("rate3",0)),    key=f"r3_{i}",   label_visibility="collapsed", min_value=0.0, max_value=100.0)
        duty1 = round(total * r1 / 100, 2)
        duty2 = round(total * r2 / 100, 2)
        duty3 = round(total * r3 / 100, 2)
        td    = round(duty1 + duty2 + duty3, 2)
        tp    = round(r1 + r2 + r3, 2)
        updated_items.append({
            "part_number": part, "description": desc,
            "quantity": qty, "price": price, "total": total,
            "tariff1": t1, "tariff2": t2, "tariff3": t3,
            "rate1": r1, "rate2": r2, "rate3": r3,
            "duty1": duty1, "duty2": duty2, "duty3": duty3,
            "total_duty": td, "total_duty_pct": tp
        })

    st.session_state.line_items = updated_items

    ca, cb, _ = st.columns([1, 1, 4])
    if ca.button("➕ Add Line Item"):
        st.session_state.line_items.append({})
        st.rerun()
    if cb.button("🗑 Remove Last") and len(st.session_state.line_items) > 1:
        st.session_state.line_items.pop()
        st.rerun()

    all_total    = sum(i["total"]      for i in st.session_state.line_items)
    all_duty     = sum(i["total_duty"] for i in st.session_state.line_items)
    avg_duty_pct = round((all_duty / all_total * 100) if all_total else 0, 2)

    st.markdown("---")
    s1, s2, s3 = st.columns(3)
    s1.metric("Total Invoice Value", f"${all_total:,.2f}")
    s2.metric("Total Duties",        f"${all_duty:,.2f}")
    s3.metric("Effective Duty Rate", f"{avg_duty_pct:.2f}%")

    st.markdown("---")
    if st.button("✅ Save Entry", type="primary"):
        if not entry_number:
            st.error("Entry number is required.")
        else:
            entries = load_entries()
            entries.append({
                "date_logged":   datetime.now().strftime("%Y-%m-%d %H:%M"),
                "entry_number":  entry_number,
                "entry_date":    entry_date,
                "import_date":   import_date,
                "broker":        broker,
                "broker_number": broker_number,
                "supplier":      supplier,
                "country":       country,
                "invoice":       invoice,
                "line_items":    st.session_state.line_items,
                "total_value":   all_total,
                "total_duty":    all_duty,
                "duty_pct":      avg_duty_pct,
                "filename":      uploaded.name if uploaded else ""
            })
            save_entries(entries)
            st.session_state.line_items = [{}]
            st.session_state.header = {}
            st.success(f"✅ Entry {entry_number} saved with {len(updated_items)} line item(s)!")
            st.rerun()

with tab2:
    entries = load_entries()
    if not entries:
        st.info("No entries logged yet.")
    else:
        total_entries = len(entries)
        total_value   = sum(e.get("total_value", 0) for e in entries)
        total_duty    = sum(e.get("total_duty",  0) for e in entries)
        avg_duty_rate = round((total_duty / total_value * 100) if total_value else 0, 2)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Entries",       total_entries)
        m2.metric("Total Invoice Value", f"${total_value:,.2f}")
        m3.metric("Total Duties Paid",   f"${total_duty:,.2f}")
        m4.metric("Avg Duty Rate",       f"{avg_duty_rate:.2f}%")

        st.markdown("---")
        search = st.text_input("🔍 Search by entry #, supplier, part number, or country")
        filtered = entries
        if search:
            s = search.lower()
            filtered = [e for e in entries if
                s in e.get("entry_number","").lower() or
                s in e.get("supplier","").lower() or
                s in e.get("country","").lower() or
                any(s in i.get("part_number","").lower() or
                    s in i.get("description","").lower()
                    for i in e.get("line_items",[]))]

        st.markdown(f"Showing **{len(filtered)}** of **{total_entries}** entries")

        for e in reversed(filtered):
            label = f"📄 {e['entry_number']}  |  {e.get('supplier','')}  |  {e.get('entry_date','')}  |  ${e.get('total_value',0):,.2f}  |  Duty: ${e.get('total_duty',0):,.2f} ({e.get('duty_pct',0):.1f}%)"
            with st.expander(label):
                h1, h2, h3, h4 = st.columns(4)
                h1.markdown(f"**Entry #:** {e.get('entry_number')}")
                h1.markdown(f"**Entry Date:** {e.get('entry_date')}")
                h1.markdown(f"**Import Date:** {e.get('import_date')}")
                h2.markdown(f"**Broker:** {e.get('broker')}")
                h2.markdown(f"**Broker #:** {e.get('broker_number')}")
                h2.markdown(f"**Invoice #:** {e.get('invoice')}")
                h3.markdown(f"**Supplier:** {e.get('supplier')}")
                h3.markdown(f"**Country:** {e.get('country')}")
                h4.markdown(f"**File:** {e.get('filename','—')}")
                h4.markdown(f"**Logged:** {e.get('date_logged')}")

                st.markdown("**Line Items:**")
                li_cols = st.columns([2,2.5,1,1,1.5,2,2,2,1,1,1,1.5,1.5,1.5,1.5,1.5])
                for col, lbl in zip(li_cols, ["Part #","Desc","Qty","Price","Total",
                                               "Tariff 1","Tariff 2","Tariff 3",
                                               "R1%","R2%","R3%",
                                               "Duty 1","Duty 2","Duty 3","Total Duty","Duty %"]):
                    col.markdown(f"**{lbl}**")

                for item in e.get("line_items", []):
                    ic = st.columns([2,2.5,1,1,1.5,2,2,2,1,1,1,1.5,1.5,1.5,1.5,1.5])
                    ic[0].write(item.get("part_number",""))
                    ic[1].write(item.get("description",""))
                    ic[2].write(item.get("quantity",""))
                    ic[3].write(f"${item.get('price',0):.2f}")
                    ic[4].write(f"${item.get('total',0):,.2f}")
                    ic[5].write(item.get("tariff1",""))
                    ic[6].write(item.get("tariff2",""))
                    ic[7].write(item.get("tariff3",""))
                    ic[8].write(f"{item.get('rate1',0)}%")
                    ic[9].write(f"{item.get('rate2',0)}%")
                    ic[10].write(f"{item.get('rate3',0)}%")
                    ic[11].write(f"${item.get('duty1',0):,.2f}")
                    ic[12].write(f"${item.get('duty2',0):,.2f}")
                    ic[13].write(f"${item.get('duty3',0):,.2f}")
                    ic[14].write(f"${item.get('total_duty',0):,.2f}")
                    ic[15].write(f"{item.get('total_duty_pct',0):.1f}%")

        st.markdown("---")
        if st.button("📥 Export All to Excel"):
            buf = build_excel(entries)
            st.download_button(
                label="⬇️ Download Excel",
                data=buf,
                file_name=f"7501_log_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
