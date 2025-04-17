import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

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
        import re
        m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text)
        results[f] = m.group(1).strip() if m else ""
    return results

def highlight_mismatches(df):
    wb = load_workbook(filename=BytesIO())
    # (we’ll just let users download the raw dataframe for now)
    return df

st.title("Self‑QC Automation")

docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])
if docx_file and excel_file:
    df = pd.read_excel(excel_file, engine="openpyxl")
    fields = df["Field"].dropna().tolist()
    extracted = extract_fields_from_docx(docx_file, fields)
    df["Auto Extracted Value"] = df["Field"].map(extracted)
    df["Match"] = df.apply(
        lambda r: str(r["Manual Value"]).strip().lower() == str(r["Auto Extracted Value"]).strip().lower() 
                  if r["Status"] else None,
        axis=1
    )
    st.dataframe(df)
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine="openpyxl")
    st.download_button(
        "Download QC Results",
        towrite.getvalue(),
        file_name="qc_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
