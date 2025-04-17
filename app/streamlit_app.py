import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

def extract_fields_from_docx(file, fields):
    text = ""
    doc = Document(file)
    for p in doc.paragraphs:
        text += p.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            text += "\t".join(cell.text for cell in row.cells) + "\n"
    results = {}
    for f in fields:
        m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text)
        results[f] = m.group(1).strip() if m else ""
    return results

st.title("Self‑QC Automation")

docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

if docx_file and excel_file:
    # 1) Load the sheet
    df = pd.read_excel(excel_file, engine="openpyxl")
    
    # 2) Show what columns we actually have
    st.write("🔍 Detected columns:", df.columns.tolist())
    
    # 3) Find the right column for "Field" (case‑insensitive match)
    field_col = None
    for col in df.columns:
        if col.strip().lower() == "field":
            field_col = col
            break
    
    if not field_col:
        st.error(
            "Could not find a column named ‘Field’.  \n"
            "Please rename your header to exactly “Field” (no extra spaces) or\n"
            "update the code to match your column name."
        )
        st.stop()
    
    # 4) Extract and compare
    fields = df[field_col].dropna().astype(str).tolist()
    extracted = extract_fields_from_docx(docx_file, fields)
    
    df["Auto Extracted Value"] = df[field_col].map(extracted)
    df["Match"] = df.apply(
        lambda r: (
            str(r.get("Manual Value", "")).strip().lower()
            == str(r["Auto Extracted Value"]).strip().lower()
        ) if r.get("Status") else None,
        axis=1
    )
    
    # 5) Show results and offer download
    st.dataframe(df)
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine="openpyxl")
    st.download_button(
        "Download QC Results",
        towrite.getvalue(),
        file_name="qc_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
