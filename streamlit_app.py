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


def clean_header(value):
    if value is None:
        return ""
    return str(value).replace("\ufeff", "").strip()


def find_header_row_from_excel(data):
    """
    Find the actual column-header row instead of assuming row 1.

    This handles vendor workbooks that put a report title, date range,
    or other information above the real headers.
    """
    preview = pd.read_excel(
        io.BytesIO(data),
        header=None,
        dtype=str,
        keep_default_na=False,
        nrows=20,
    )

    required_markers = {PO_COLUMN, INVENTORY_COLUMN}

    for idx, row in preview.iterrows():
        values = {clean_header(v) for v in row.tolist()}
        if required_markers.issubset(values):
            return idx

    return None


def read_uploaded_file(uploaded_file):
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".csv":
        # CSV files normally have the headers on their first row.
        return pd.read_csv(
            io.BytesIO(data),
            dtype=str,
            keep_default_na=False,
        )

    if suffix in {".xlsx", ".xlsm"}:
        header_row = find_header_row_from_excel(data)

        if header_row is None:
            raise ValueError(
                "I could not find the actual header row. The workbook must "
                f"contain '{PO_COLUMN}' and '{INVENTORY_COLUMN}' as column names."
            )

        return pd.read_excel(
            io.BytesIO(data),
            header=header_row,
            dtype=str,
            keep_default_na=False,
        )

    raise ValueError("Please upload a CSV, XLSX, or XLSM file.")


def normalize_digits(value, width, field_name, row_number):
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError(f"Row {row_number}: {field_name} is blank.")

    text = str(value).strip()

    # Excel/pandas can represent an integer-looking value as 12345.0.
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


def transform(df, vendor_id=None):
    df = df.copy()
    df.columns = [clean_header(c) for c in df.columns]

    # Ignore completely blank rows.
    df = df.loc[
        ~df.apply(lambda row: all(str(v).strip() == "" for v in row), axis=1)
    ].reset_index(drop=True)

    missing = [c for c in (PO_COLUMN, INVENTORY_COLUMN) if c not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    original_columns = list(df.columns)

    if VENDOR_COLUMN not in df.columns:
        if not vendor_id:
            return None, "VENDOR_REQUIRED", None
        df[VENDOR_COLUMN] = vendor_id
    else:
        df[VENDOR_COLUMN] = df[VENDOR_COLUMN].astype(str).str.strip()
        if vendor_id:
            df.loc[df[VENDOR_COLUMN] == "", VENDOR_COLUMN] = vendor_id

    blank_vendor = df[VENDOR_COLUMN].astype(str).str.strip() == ""
    if blank_vendor.any():
        rows = (df.index[blank_vendor] + 2).tolist()
        raise ValueError(
            "Vendor is blank on source Excel row(s): " +
            ", ".join(map(str, rows))
        )

    valid_records = []
    exception_records = []

    for idx, row in df.iterrows():
        source_row_number = idx + 2
        po_raw = str(row[PO_COLUMN]).strip()
        inv_raw = str(row[INVENTORY_COLUMN]).strip()

        # PO values containing non-numeric characters, or more than 6 digits,
        # are excluded from the Acumatica import and copied in full to the
        # exception workbook.
        if not po_raw.isdigit() or len(po_raw) > 6:
            exception_records.append(row[original_columns].to_dict())
            continue

        if not inv_raw.isdigit():
            raise ValueError(
                f"Inventory ID is invalid on source Excel row {source_row_number}: "
                f"{inv_raw!r}. Inventory ID must contain digits only."
            )

        if len(inv_raw) > 8:
            raise ValueError(
                f"Inventory ID is too long on source Excel row {source_row_number}: "
                f"{inv_raw!r}."
            )

        valid_records.append({
            VENDOR_COLUMN: str(row[VENDOR_COLUMN]).strip(),
            PO_COLUMN: po_raw.zfill(6),
            INVENTORY_COLUMN: inv_raw.zfill(8),
        })

    output = pd.DataFrame(
        valid_records,
        columns=[VENDOR_COLUMN, PO_COLUMN, INVENTORY_COLUMN],
    )
    output["Row Number"] = range(1, len(output) + 1)
    output = output[OUTPUT_COLUMNS]

    exceptions = pd.DataFrame(exception_records, columns=original_columns)
    return output, "OK", exceptions

def make_excel(df):
    buffer = io.BytesIO()

    # Write identifier columns as strings.
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
        clean_header(cell.value): cell.column
        for cell in ws[1]
    }

    # Explicit Excel text format prevents leading zeroes from disappearing.
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
    "Required output fields: **Vendor**, **Drop-Ship PO Nbr:**, "
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
            f"Loaded **{len(source):,} non-empty row(s)** from `{uploaded.name}`. "
            "The converter automatically finds the actual header row in Excel files."
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
            blank_vendor = (
                source[VENDOR_COLUMN].astype(str).str.strip().eq("").any()
            )

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

                st.success(
                    f"Ready! {len(result):,} row(s) are ready for Acumatica."
                )

                if exceptions is not None and not exceptions.empty:
                    st.warning(
                        f"{len(exceptions):,} row(s) were excluded because the "
                        "Drop-Ship PO Nbr: contains a non-numeric character "
                        "or is longer than 6 digits."
                    )

                    exception_bytes = make_excel(exceptions)
                    exception_name = (
                        Path(uploaded.name).stem
                        + "_PO_Exceptions.xlsx"
                    )

                    col1, col2 = st.columns(2)

                    with col1:
                        st.download_button(
                            "⬇️ Download Acumatica Import File",
                            data=output_bytes,
                            file_name=output_name,
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            use_container_width=True,
                        )

                    with col2:
                        st.download_button(
                            "⬇️ Download PO Exception Rows",
                            data=exception_bytes,
                            file_name=exception_name,
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            use_container_width=True,
                        )

                    st.caption(
                        "This file contains the complete original rows for the "
                        "excluded records. They are not included in the "
                        "Acumatica import file."
                    )

                    st.dataframe(
                        exceptions,
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("No PO exception rows were found.")

                st.dataframe(
                    result,
                    use_container_width=True,
                    hide_index=True,
                )

                from datetime import datetime

                output_name = (
                    "Acumatica Create Purchase Receipt Import File "
                    + datetime.now().strftime("%Y-%m-%d")
                    + ".xlsx"
                )


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
