# # import streamlit as st
# # import pandas as pd
# # from docx import Document
# # from io import BytesIO
# # import re

# # def extract_fields_from_docx(file, fields):
# #     text = ""
# #     doc = Document(file)
# #     for p in doc.paragraphs:
# #         text += p.text + "\n"
# #     for table in doc.tables:
# #         for row in table.rows:
# #             text += "\t".join(cell.text for cell in row.cells) + "\n"
# #     out = {}
# #     for f in fields:
# #         m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text)
# #         out[f] = m.group(1).strip() if m else ""
# #     return out

# # st.title("Self‑QC Automation")

# # docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
# # excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

# # if docx_file and excel_file:
# #     df = pd.read_excel(excel_file, engine="openpyxl")
# #     cols = df.columns.tolist()
# #     st.write("🔍 Detected columns:", cols)

# #     # Let the user choose which column is which:
# #     field_col    = st.selectbox("Which column holds your field names?",    cols)
# #     status_col   = st.selectbox("Which column indicates active (TRUE/FALSE)?", cols)
# #     manual_col   = st.selectbox("Which column holds the manual values?",   cols)

# #     # Now proceed safely:
# #     fields = df[field_col].dropna().astype(str).tolist()
# #     extracted = extract_fields_from_docx(docx_file, fields)

# #     df["Auto Extracted Value"] = df[field_col].map(extracted)
# #     df["Match"] = df.apply(
# #         lambda r: (
# #             str(r[manual_col]).strip().lower()
# #             == str(r["Auto Extracted Value"]).strip().lower()
# #         ) if r[status_col] else None,
# #         axis=1,
# #     )

# #     st.dataframe(df)
# #     buf = BytesIO()
# #     df.to_excel(buf, index=False, engine="openpyxl")
# #     st.download_button(
# #         "Download QC Results",
# #         buf.getvalue(),
# #         "qc_results.xlsx",
# #         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
# #     )
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

# st.title("Self-QC Automation")

# docx_file = st.file_uploader("Upload Agreement (.docx)", type=["docx"])
# excel_file = st.file_uploader("Upload Checklist (.xlsx)", type=["xlsx"])

# if docx_file and excel_file:
#     workbook = pd.ExcelFile(excel_file, engine="openpyxl")
#     all_fields = set()

#     # Extract fields from all sheets
#     for sheet in workbook.sheet_names:
#         df_temp = workbook.parse(sheet)
#         all_fields.update(df_temp.iloc[:, 0].dropna().astype(str).tolist())

#     fields = list(all_fields)
#     extracted = extract_fields_from_docx(docx_file, fields)

#     # Create final simplified DataFrame with two columns only
#     result_df = pd.DataFrame({
#         "Field Name": fields,
#         "Value": [extracted.get(f, "") for f in fields]
#     })

#     # Sort alphabetically by Field Name for clarity
#     result_df.sort_values(by="Field Name", inplace=True)

#     st.subheader("✅ Extracted Results")
#     st.dataframe(result_df, width=800)

#     # Excel Download (proper formatting)
#     buf = BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as writer:
#         result_df.to_excel(writer, index=False, sheet_name="Extracted Data")

#         # Adjust column widths
#         ws = writer.sheets["Extracted Data"]
#         for column_cells in ws.columns:
#             max_length = max(len(str(cell.value)) for cell in column_cells) + 2
#             col_letter = column_cells[0].column_letter
#             ws.column_dimensions[col_letter].width = max_length

#     st.download_button(
#         "📥 Download QC Results",
#         buf.getvalue(),
#         "QC_Results.xlsx",
#         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )
import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# Set Streamlit page config
st.set_page_config(page_title="Self-QC Automation", layout="wide")

# Function to extract fields from docx
def extract_fields_from_docx(file, fields):
    text = ""
    doc = Document(file)
    for p in doc.paragraphs:
        text += p.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            text += "\t".join(cell.text.strip() for cell in row.cells) + "\n"
    extracted = {}
    for field in fields:
        pattern = re.escape(field) + r"\s*[:\-]\s*(.+)"
        match = re.search(pattern, text, re.IGNORECASE)
        extracted[field] = match.group(1).strip() if match else ""
    return extracted

# Title and UI enhancements
st.title("🔎 Self-QC Automation")
st.markdown("""
Upload your **Agreement (.docx)** and the **Excel Checklist (.xlsx)** to automatically extract and verify key data.
""")

# File uploaders in columns
col1, col2 = st.columns(2)

with col1:
    docx_file = st.file_uploader("📄 **Upload Agreement (.docx)**", type=["docx"])
with col2:
    excel_file = st.file_uploader("📑 **Upload Checklist (.xlsx)**", type=["xlsx"])

if docx_file and excel_file:
    excel = pd.ExcelFile(excel_file, engine="openpyxl")

    all_fields = set()

    # Process all sheets except 'Notification Email'
    for sheet in excel.sheet_names:
        if sheet.lower() != "notification email":
            df = excel.parse(sheet)
            if {"Field", "Status"}.issubset(df.columns):
                active_fields = df[df["Status"] == True]["Field"].dropna().astype(str).tolist()
                all_fields.update(active_fields)

    fields = sorted(all_fields)
    extracted_data = extract_fields_from_docx(docx_file, fields)

    # Create final simplified DataFrame
    final_df = pd.DataFrame({
        "Field Name": fields,
        "Extracted Value": [extracted_data.get(field, "") for field in fields]
    })

    # Display the DataFrame with improved styling
    st.subheader("📝 **Extracted QC Results**")
    st.dataframe(final_df, use_container_width=True)

    # Prepare well-formatted Excel download with highlighting for empty values
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="QC Results")

        # Access worksheet for styling
        worksheet = writer.sheets["QC Results"]

        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
        missing_fill = PatternFill(start_color="FFC7CE", fill_type="solid")

        # Style headers
        for col_num, column_title in enumerate(final_df.columns, start=1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill

        # Adjust column widths automatically and highlight missing values
        for idx, col in enumerate(final_df.columns, start=1):
            max_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 4
            worksheet.column_dimensions[get_column_letter(idx)].width = max_len

            # Highlight empty extracted values
            if col == "Extracted Value":
                for row_num, cell_value in enumerate(final_df[col], start=2):
                    if not cell_value:
                        worksheet.cell(row=row_num, column=idx).fill = missing_fill

    buf.seek(0)

    # Download button clearly placed
    st.download_button(
        label="📥 **Download QC Results Excel**",
        data=buf.getvalue(),
        file_name="QC_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Instructions clearly stated
    st.info("""
    📌 **Note**:
    - The final Excel clearly shows extracted values from the Word document.
    - Fields without extracted values (not found in the Word document) are **highlighted in red**.
    """)

else:
    st.warning("👆 Please upload both Agreement (.docx) and Checklist (.xlsx) files to proceed.")

