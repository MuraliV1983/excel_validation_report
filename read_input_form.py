import xlwings as xw
import csv
import os

# Load the workbook
wb = xw.Book('input_form.xls')
sheet = wb.sheets[0]

# Read the used range (assumes no gaps)
data = sheet.range("A1").expand().value  # returns list of rows

headers = data[0]
rows = data[1:]

# Create output folder if it doesn't exist
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

# Write to CSV
csv_path = os.path.join(output_folder, "output_data.csv")
with open(csv_path, mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(headers)
    writer.writerows(rows)

print(f"✅ Form data exported to: {csv_path}")
