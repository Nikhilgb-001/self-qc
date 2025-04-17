import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re

def extract_fields_from_docx(file, fields):
    text = ""
    doc = Document(file)
    for p in doc.paragraphs:
        text += p.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            text += "\t".join(cell.text for cell in row.cells) + "\n"
    out = {}
    for f in fields:
        m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text)
        out[f] = m.group(1).strip() if m else ""
    return out

st.title("Self‑QC Automation")

docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

if docx_file and excel_file:
    df = pd.read_excel(excel_file, engine="openpyxl")
    cols = df.columns.tolist()
    st.write("🔍 Detected columns:", cols)

    # Let the user choose which column is which:
    field_col    = st.selectbox("Which column holds your field names?",    cols)
    status_col   = st.selectbox("Which column indicates active (TRUE/FALSE)?", cols)
    manual_col   = st.selectbox("Which column holds the manual values?",   cols)

    # Now proceed safely:
    fields = df[field_col].dropna().astype(str).tolist()
    extracted = extract_fields_from_docx(docx_file, fields)

    df["Auto Extracted Value"] = df[field_col].map(extracted)
    df["Match"] = df.apply(
        lambda r: (
            str(r[manual_col]).strip().lower()
            == str(r["Auto Extracted Value"]).strip().lower()
        ) if r[status_col] else None,
        axis=1,
    )

    st.dataframe(df)
    buf = BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        "Download QC Results",
        buf.getvalue(),
        "qc_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
