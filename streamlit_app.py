import io
from pathlib import Path
import re

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


PO_COLUMN = "Drop-Ship PO Nbr:"
INVENTORY_COLUMN = "Inventory ID"
VENDOR_COLUMN = "Vendor"
OUTPUT_COLUMNS = [VENDOR_COLUMN, PO_COLUMN, INVENTORY_COLUMN, "Row Number"]


def normalize_digits(value, width, field_name, row_number):
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError(f"Row {row_number}: {field_name} is blank.")

    text = str(value).strip()

    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]

    if not text.isdigit():
        raise ValueError(
            f"Row {row_number}: {field_name} must contain digits only; "
            f"received {value!r}."
        )

    if len(text) > width:
        raise ValueError(
            f"Row {row_number}: {field_name} has {len(text)} digits; "
            f"maximum allowed is {width}."
        )

    return text.zfill(width)


def clean_header(value):
    return str(value).replace("\ufeff", "").strip()


def find_header_row_from_excel(data):
    preview = pd.read_excel(
        io.BytesIO(data),
        header=None,
        dtype=str,
        keep_default_na=False,
        nrows=30,
    )
    for row_index in range(len(preview)):
        values = {clean_header(v) for v in preview.iloc[row_index].tolist()}
        if PO_COLUMN in values and INVENTORY_COLUMN in values:
            return row_index
    raise ValueError(
        "Could not find the true header row. The workbook must contain "
        f"'{PO_COLUMN}' and '{INVENTORY_COLUMN}'."
    )


def read_uploaded_file(uploaded_file):
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".csv":
        df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False)
    elif suffix in {".xlsx", ".xlsm"}:
        header_row = find_header_row_from_excel(data)
        df = pd.read_excel(
            io.BytesIO(data),
            header=header_row,
            dtype=str,
            keep_default_na=False,
        )
    else:
        raise ValueError("Please upload a CSV, XLSX, or XLSM file.")

    df.columns = [clean_header(c) for c in df.columns]
    return df


def transform(df, vendor_id=None):
    df = df.copy()
    df.columns = [clean_header(c) for c in df.columns]
    df = df.loc[~df.apply(lambda row: all(str(v).strip() == "" for v in row), axis=1)].reset_index(drop=True)

    missing = [c for c in (PO_COLUMN, INVENTORY_COLUMN) if c not in df.columns]
    if missing:
        raise ValueError("The uploaded file is missing required column(s): " + ", ".join(repr(c) for c in missing))

    original_columns = list(df.columns)
    if VENDOR_COLUMN not in df.columns:
        if not vendor_id or not vendor_id.strip():
            return None, "VENDOR_REQUIRED", None
        df[VENDOR_COLUMN] = vendor_id.strip()
    else:
        df[VENDOR_COLUMN] = df[VENDOR_COLUMN].astype(str).str.strip()
        if vendor_id and vendor_id.strip():
            df.loc[df[VENDOR_COLUMN] == "", VENDOR_COLUMN] = vendor_id.strip()

    if df[VENDOR_COLUMN].astype(str).str.strip().eq("").any():
        raise ValueError("The Vendor column contains blank values.")

    valid=[]; exceptions=[]
    for idx,row in df.iterrows():
        source_row=idx+2
        po=str(row[PO_COLUMN]).strip()
        inv=str(row[INVENTORY_COLUMN]).strip()
        if not po.isdigit() or len(po)>6:
            exceptions.append(row[original_columns].to_dict())
            continue
        if not inv.isdigit():
            raise ValueError(f"Inventory ID is invalid on source row {source_row}: {inv!r}. It must contain digits only.")
        if len(inv)>8:
            raise ValueError(f"Inventory ID is too long on source row {source_row}: {inv!r}.")
        valid.append({VENDOR_COLUMN:str(row[VENDOR_COLUMN]).strip(), PO_COLUMN:po.zfill(6), INVENTORY_COLUMN:inv.zfill(8)})

    result=pd.DataFrame(valid,columns=[VENDOR_COLUMN,PO_COLUMN,INVENTORY_COLUMN])
    result["Row Number"]=range(1,len(result)+1)
    result=result[OUTPUT_COLUMNS]
    exceptions=pd.DataFrame(exceptions,columns=original_columns)
    return result,"OK",exceptions


def make_excel(df):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Shipping Confirmations",
        )

    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb["Shipping Confirmations"]

    header_map = {
        cell.value: cell.column
        for cell in ws[1]
    }

    for field in (PO_COLUMN, INVENTORY_COLUMN):
        col = header_map[field]
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col).number_format = "@"

    ws.freeze_panes = "A2"

    widths = {
        VENDOR_COLUMN: 18,
        PO_COLUMN: 22,
        INVENTORY_COLUMN: 16,
        "Row Number": 14,
    }
    for cell in ws[1]:
        if cell.value in widths:
            ws.column_dimensions[cell.column_letter].width = widths[cell.value]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.set_page_config(
    page_title="Acumatica Drop-Ship Converter",
    page_icon="📦",
    layout="centered",
)

st.title("📦 Acumatica Drop-Ship Converter")
st.write(
    "Convert a vendor shipping-confirmation CSV or Excel file into "
    "an Acumatica-ready Excel file."
)

st.info(
    "Required fields: **Vendor**, **Drop-Ship PO Nbr:**, "
    "**Inventory ID**, and **Row Number**."
)

uploaded = st.file_uploader(
    "Upload vendor shipping confirmation",
    type=["csv", "xlsx", "xlsm"],
    help="CSV, XLSX, or XLSM files are supported.",
)

if uploaded:
    try:
        source = read_uploaded_file(uploaded)

        st.caption(
            f"Loaded **{len(source):,} row(s)** from `{uploaded.name}`."
        )

        vendor_missing = VENDOR_COLUMN not in source.columns
        vendor_id = None

        if vendor_missing:
            st.warning(
                "The uploaded file does not contain a `Vendor` column. "
                "Enter the Vendor ID below. It will be applied to every row."
            )
            vendor_id = st.text_input(
                "Vendor ID",
                placeholder="Enter Vendor ID",
            )
        else:
            blank_vendor = source[VENDOR_COLUMN].astype(str).str.strip().eq("").any()
            if blank_vendor:
                st.warning(
                    "The Vendor column contains blank rows. "
                    "Enter a Vendor ID below to fill them."
                )
                vendor_id = st.text_input(
                    "Vendor ID for blank Vendor rows",
                    placeholder="Enter Vendor ID",
                )

        if st.button(
            "Create Acumatica Excel File",
            type="primary",
            use_container_width=True,
        ):
            result, status, exceptions = transform(source, vendor_id)

            if status == "VENDOR_REQUIRED":
                st.error("Please enter the Vendor ID before creating the file.")
            else:
                output_bytes = make_excel(result)
                output_name = (
                    "Acumatica Create Purchase Receipt Import File "
                    + datetime.now().strftime("%Y-%m-%d")
                    + ".xlsx"
                )

                st.success(f"Ready! {len(result):,} row(s) are ready for Acumatica.")

                if exceptions is not None and not exceptions.empty:
                    st.warning(
                        f"{len(exceptions):,} row(s) were excluded because the "
                        "Drop-Ship PO Nbr: contains a non-numeric character "
                        "or is longer than 6 digits."
                    )
                    exception_bytes = make_excel(exceptions)
                    exception_name = Path(uploaded.name).stem + "_PO_Exceptions.xlsx"

                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            "⬇️ Download Acumatica Import File",
                            data=output_bytes,
                            file_name=output_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    with col2:
                        st.download_button(
                            "⬇️ Download PO Exception Rows",
                            data=exception_bytes,
                            file_name=exception_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    st.caption("The exception file contains the complete original rows for excluded records. They are not included in the Acumatica import file.")
                    st.dataframe(exceptions, use_container_width=True, hide_index=True)
                else:
                    st.download_button(
                        "⬇️ Download Acumatica Import File",
                        data=output_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    st.info("No PO exception rows were found.")

                st.subheader("Acumatica Import Rows")
                st.dataframe(result, use_container_width=True, hide_index=True)

    except Exception as exc:
        st.error(str(exc))
else:
    st.markdown(
        """
        **How it works**

        1. Upload the file returned by your vendor.
        2. If `Vendor` is missing, enter the Vendor ID.
        3. Click **Create Acumatica Excel File**.
        4. Review the rows shown on screen.
        5. Download the Excel file for your Acumatica import.

        PO numbers are padded to **6 digits** and Inventory IDs to
        **8 digits**. Both are stored as text so leading zeroes are preserved.
        """
    )
