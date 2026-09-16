# Acumatica Drop-Ship Converter — Deployment

## The issue fixed in version 2

The FedCo workbook has a report title above the actual column headers.
The previous converter assumed Excel row 1 contained the headers, so it
saw names such as `Unnamed: 1` instead of `Inventory ID`.

The updated converter searches the first 20 Excel rows for the actual
`Drop-Ship PO Nbr:` and `Inventory ID` headers and uses that row.

It also ignores completely blank rows below the data while still
reporting partially populated rows as validation errors.

## Important validation behavior

A value such as `.` in `Drop-Ship PO Nbr:` is not a valid PO number.
The app will report the exact data row rather than silently changing it.

## Deploying

Put `streamlit_app.py` and `requirements.txt` in a GitHub repository and
deploy with Streamlit Community Cloud, or have IT host the same app in an
approved company-controlled environment.

Users only need a browser; they do not need Python installed.

Before using an external hosting service for vendor/business data, confirm
your company's security/privacy requirements.
