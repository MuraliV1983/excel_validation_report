import xlwings as xw
import re
import os
from datetime import datetime

# Load the input Excel file
wb = xw.Book('input_form.xls')
sheet = wb.sheets[0]
data = sheet.range("A1").expand().value

# Add "Processed" to headers
headers = data[0] + ["Processed"]
rows = data[1:]

# Basic email pattern
email_pattern = re.compile(r"[^@]+@[^@]+\.[^@]+")

# Prepare processed rows with validation
processed_rows = []
for row in rows:
    name, email, amount, status = row
    remarks = []

    # Validate amount
    if not isinstance(amount, (int, float)):
        try:
            amount = float(amount)
        except:
            remarks.append("Invalid Amount")

    # Validate email
    if not email_pattern.match(str(email)):
        remarks.append("Invalid Email")

    # Final status message (without emojis)
    if remarks:
        status_message = " and ".join(remarks)
    else:
        status_message = "Valid"

    processed_rows.append([name, email, amount, status, status_message])

# Create output folder
output_folder = "reports"
os.makedirs(output_folder, exist_ok=True)

# Create new workbook for output
new_wb = xw.Book()
report_name = "Report_Data"
output_sheet = new_wb.sheets[0]
output_sheet.name = report_name

# Write data
output_sheet.range("A1").value = [headers] + processed_rows

# Format header
header_range = output_sheet.range("A1:E1")
header_range.api.Font.Bold = True
header_range.api.Interior.ColorIndex = 15  # Medium Grey

# Autofit columns
output_sheet.autofit()

# Apply color formatting to 'Processed' column
last_row = len(processed_rows) + 1  # +1 for header
for i in range(2, last_row + 1):
    cell = output_sheet.range(f"E{i}")
    value = str(cell.value).lower()

    if value == "valid":
        cell.api.Interior.ColorIndex = 4  # Light Green
    else:
        cell.api.Interior.ColorIndex = 3  # Light Red

# Save with timestamped filename
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
output_filename = f"{report_name}_{timestamp}.xlsx"
output_path = os.path.join(output_folder, output_filename)
new_wb.save(output_path)

print(f"File saved: {output_path}")
