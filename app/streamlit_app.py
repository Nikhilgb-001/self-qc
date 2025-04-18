# import streamlit as st
# import pandas as pd
# from docx import Document
# from io import BytesIO
# import re

# def extract_fields_from_docx(file, fields):
#     text = ""
#     doc = Document(file)
#     for p in doc.paragraphs:
#         text += p.text + "\n"
#     for table in doc.tables:
#         for row in table.rows:
#             text += "\t".join(cell.text for cell in row.cells) + "\n"
#     out = {}
#     for f in fields:
#         m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text)
#         out[f] = m.group(1).strip() if m else ""
#     return out

# st.title("Self‑QC Automation")

# docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
# excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

# if docx_file and excel_file:
#     df = pd.read_excel(excel_file, engine="openpyxl")
#     cols = df.columns.tolist()
#     st.write("🔍 Detected columns:", cols)

#     # Let the user choose which column is which:
#     field_col    = st.selectbox("Which column holds your field names?",    cols)
#     status_col   = st.selectbox("Which column indicates active (TRUE/FALSE)?", cols)
#     manual_col   = st.selectbox("Which column holds the manual values?",   cols)

#     # Now proceed safely:
#     fields = df[field_col].dropna().astype(str).tolist()
#     extracted = extract_fields_from_docx(docx_file, fields)

#     df["Auto Extracted Value"] = df[field_col].map(extracted)
#     df["Match"] = df.apply(
#         lambda r: (
#             str(r[manual_col]).strip().lower()
#             == str(r["Auto Extracted Value"]).strip().lower()
#         ) if r[status_col] else None,
#         axis=1,
#     )

#     st.dataframe(df)
#     buf = BytesIO()
#     df.to_excel(buf, index=False, engine="openpyxl")
#     st.download_button(
#         "Download QC Results",
#         buf.getvalue(),
#         "qc_results.xlsx",
#         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )
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

st.title("Self-QC Automation")

docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

if docx_file and excel_file:
    workbook = pd.ExcelFile(excel_file, engine="openpyxl")
    all_fields = set()

    # Extract fields from all sheets
    for sheet in workbook.sheet_names:
        df_temp = workbook.parse(sheet)
        all_fields.update(df_temp.iloc[:, 0].dropna().astype(str).tolist())

    fields = list(all_fields)
    extracted = extract_fields_from_docx(docx_file, fields)

    # Create final simplified DataFrame with two columns only
    result_df = pd.DataFrame({
        "Field Name": fields,
        "Value": [extracted.get(f, "") for f in fields]
    })

    # Sort alphabetically by Field Name for clarity
    result_df.sort_values(by="Field Name", inplace=True)

    st.subheader("✅ Extracted Results")
    st.dataframe(result_df, width=800)

    # Excel Download (proper formatting)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Extracted Data")

        # Adjust column widths
        ws = writer.sheets["Extracted Data"]
        for column_cells in ws.columns:
            max_length = max(len(str(cell.value)) for cell in column_cells) + 2
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = max_length

    st.download_button(
        "📥 Download QC Results",
        buf.getvalue(),
        "QC_Results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
