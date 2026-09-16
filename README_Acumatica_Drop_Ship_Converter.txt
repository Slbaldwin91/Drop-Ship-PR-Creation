Acumatica Drop-Ship Converter

1. Install dependencies:
   pip install pandas openpyxl

2. Run:
   python acumatica_drop_ship_converter.py "vendor_file.csv"

3. Or specify the Vendor ID if the source file has no Vendor column:
   python acumatica_drop_ship_converter.py "vendor_file.csv" --vendor "VENDOR01"

4. Optional output name:
   python acumatica_drop_ship_converter.py "vendor_file.csv" -o "Acumatica_Ready.xlsx"

The generated workbook contains:
Vendor
Drop-Ship PO Nbr:
Inventory ID
Row Number

PO numbers are exactly 6 digits and Inventory IDs exactly 8 digits.
Both identifier columns are stored as Excel text so leading zeroes remain visible.
