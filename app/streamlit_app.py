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
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Function to extract fields from Word Document
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
        m = re.search(re.escape(f) + r"\s*[:\-]\s*(.+)", text, re.IGNORECASE)
        out[f] = m.group(1).strip() if m else ""
    return out

# Streamlit UI
st.title("Self-QC Automation")

# File uploaders
docx_file = st.file_uploader("📄 Upload Agreement (.docx)", type=["docx"])
excel_file = st.file_uploader("📑 Upload Checklist (.xlsx)", type=["xlsx"])

if docx_file and excel_file:
    workbook = pd.ExcelFile(excel_file, engine="openpyxl")
    output_sheets = {}

    # Iterate through sheets except "Notification Email"
    for sheet_name in workbook.sheet_names:
        if sheet_name.lower() == "notification email":
            continue  # skip email sheet

        df = workbook.parse(sheet_name)

        # Check necessary columns exist
        required_cols = ["Field", "Status", "Manual Value"]
        if not all(col in df.columns for col in required_cols):
            st.error(f"Sheet '{sheet_name}' missing required columns.")
            continue

        # Get active fields
        active_df = df[df["Status"] == True].copy()
        fields = active_df["Field"].dropna().astype(str).tolist()

        # Extract data from Word
        extracted_values = extract_fields_from_docx(docx_file, fields)

        # Map extracted values back to DataFrame
        active_df["Auto Extracted Value"] = active_df["Field"].map(extracted_values)

        # Compare manual and auto-extracted values
        active_df["Match"] = active_df.apply(
            lambda row: str(row["Manual Value"]).strip().lower() == str(row["Auto Extracted Value"]).strip().lower(),
            axis=1,
        )

        # Store result for each sheet
        output_sheets[sheet_name] = active_df

        # Display results for each sheet
        st.subheader(f"Sheet: {sheet_name}")
        st.dataframe(active_df[["Field", "Manual Value", "Auto Extracted Value", "Match"]], width=900)

    # Prepare final Excel file with highlighting
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df_sheet in output_sheets.items():
            df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)

    buf.seek(0)
    wb = load_workbook(buf)

    # Highlight mismatches
    mismatch_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    for sheet_name in output_sheets:
        ws = wb[sheet_name]
        header = [cell.value for cell in ws[1]]

        match_col_idx = header.index("Match") + 1  # Excel is 1-indexed

        for row in range(2, ws.max_row + 1):
            match_cell = ws.cell(row=row, column=match_col_idx)
            if match_cell.value == False:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col_idx).fill = mismatch_fill

        # Adjust column widths for readability
        for col_cells in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col_cells) + 2
            ws.column_dimensions[col_cells[0].column_letter].width = max_length

    final_buf = BytesIO()
    wb.save(final_buf)
    final_buf.seek(0)

    # Final Download button
    st.download_button(
        "📥 Download QC Results (with Mismatches Highlighted)",
        final_buf.getvalue(),
        "QC_Results_Highlighted.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

