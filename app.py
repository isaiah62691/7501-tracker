import streamlit as st
import pdfplumber
import re
import json
import os
from datetime import datetime

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

def extract_entry_number(text):
    m = re.search(r'entry\s*(?:no|number|#)?[\s:.]*([0-9\-]{10,20})', text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def extract_hts_codes(text):
    codes = re.findall(r'\b(\d{4}\.\d{2}\.\d{4})\b', text)
    return ", ".join(dict.fromkeys(codes))

def extract_duties(text):
    amounts = re.findall(r'\$\s*([\d,]+\.\d{2})', text)
    return amounts[-1] if amounts else ""

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("7501 Entry Tracker")
st.markdown("Upload a CBP Form 7501 to auto-log it. Review and correct before saving.")

tab1, tab2 = st.tabs(["➕ New Entry", "📋 View All Entries"])

with tab1:
    uploaded = st.file_uploader("Upload 7501 PDF", type="pdf")

    if uploaded:
        text = extract_text(uploaded)

        st.subheader("Review & Correct Extracted Data")
        col1, col2 = st.columns(2)

        with col1:
            entry_number = st.text_input("Entry Number", value=extract_entry_number(text))
            hts_codes = st.text_input("HTS Codes", value=extract_hts_codes(text))

        with col2:
            duties = st.text_input("Total Duties ($)", value=extract_duties(text))
            notes = st.text_area("Notes (optional)", height=100)

        if st.button("✅ Save Entry", type="primary"):
            if not entry_number:
                st.error("Entry number is required before saving.")
            else:
                entries = load_entries()
                entries.append({
                    "date_logged": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "entry_number": entry_number,
                    "hts_codes": hts_codes,
                    "duties": duties,
                    "notes": notes,
                    "filename": uploaded.name
                })
                save_entries(entries)
                st.success(f"✅ Entry {entry_number} saved successfully!")

with tab2:
    entries = load_entries()

    if not entries:
        st.info("No entries logged yet. Upload a 7501 to get started.")
    else:
        st.markdown(f"**{len(entries)} total entries logged**")

        search = st.text_input("🔍 Search by entry number or HTS code")

        filtered = entries
        if search:
            filtered = [e for e in entries if
                       search.lower() in e["entry_number"].lower() or
                       search.lower() in e["hts_codes"].lower()]

        for e in reversed(filtered):
            with st.expander(f"📄 {e['entry_number']} — {e['date_logged']}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**File:** {e['filename']}")
                    st.markdown(f"**Entry Number:** {e['entry_number']}")
                    st.markdown(f"**HTS Codes:** {e['hts_codes']}")
                with c2:
                    st.markdown(f"**Duties:** ${e['duties']}")
                    st.markdown(f"**Notes:** {e['notes'] or '—'}")
