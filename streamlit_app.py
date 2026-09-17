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


def read_uploaded_file(uploaded_file):
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".csv":
        return pd.read_csv(
            io.BytesIO(data),
            dtype=str,
            keep_default_na=False,
        )

    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(
            io.BytesIO(data),
            dtype=str,
            keep_default_na=False,
        )

    raise ValueError("Please upload a CSV, XLSX, or XLSM file.")


def transform(df, vendor_id=None):
    missing = [
        c for c in (PO_COLUMN, INVENTORY_COLUMN)
        if c not in df.columns
    ]
    if missing:
        raise ValueError(
            "The uploaded file is missing required column(s): "
            + ", ".join(repr(c) for c in missing)
        )

    if VENDOR_COLUMN not in df.columns:
        if not vendor_id or not vendor_id.strip():
            return None, "VENDOR_REQUIRED"

        df[VENDOR_COLUMN] = vendor_id.strip()
    elif vendor_id and vendor_id.strip():
        df[VENDOR_COLUMN] = (
            df[VENDOR_COLUMN]
            .replace("", pd.NA)
            .fillna(vendor_id.strip())
        )

    if df[VENDOR_COLUMN].astype(str).str.strip().eq("").any():
        raise ValueError(
            "The Vendor column contains blank values. "
            "Enter a Vendor ID to fill the blanks."
        )

    result = pd.DataFrame()
    result[VENDOR_COLUMN] = df[VENDOR_COLUMN].astype(str).str.strip()
    result[PO_COLUMN] = [
        normalize_digits(v, 6, PO_COLUMN, i)
        for i, v in enumerate(df[PO_COLUMN], start=1)
    ]
    result[INVENTORY_COLUMN] = [
        normalize_digits(v, 8, INVENTORY_COLUMN, i)
        for i, v in enumerate(df[INVENTORY_COLUMN], start=1)
    ]
    result["Row Number"] = range(1, len(result) + 1)

    return result[OUTPUT_COLUMNS], None


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
            result, status = transform(source, vendor_id)

            if status == "VENDOR_REQUIRED":
                st.error("Please enter the Vendor ID before creating the file.")
            else:
                output_bytes = make_excel(result)

                st.success(
                    f"Ready! {len(result):,} row(s) were validated and formatted."
                )

                st.dataframe(
                    result,
                    use_container_width=True,
                    hide_index=True,
                )

                output_name = (
                    Path(uploaded.name).stem
                    + "_Acumatica_Ready.xlsx"
                )

                st.download_button(
                    "⬇️ Download Acumatica Excel File",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True,
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
