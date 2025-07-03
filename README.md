# 🧾 Excel Validation Report with Python & xlwings

This Python project reads data from an Excel input form, validates specific fields, and generates a clean, color-coded Excel report automatically.

It’s ideal for scenarios where data needs to be validated before further processing — and gives a visual status update directly in Excel.

---

## ✅ Features

- 📥 Reads input from `input_form.xls`
- ✔️ Validates:
  - Email format
  - Amount field (must be numeric and > 0)
- 🧾 Adds a new column "Processed" with status
- 🎨 Color-coding:
  - ✅ **Green** for valid rows
  - ❌ **Red** for invalid rows
- 🕒 Output saved as a timestamped `.xlsx` file in `/output/`

---

## 🛠 Built With

- **Python 3.13**
- [xlwings](https://www.xlwings.org/) – for Excel-Python interaction
- `pywin32` – required for Excel COM interface on Windows

> **Note:** This project works on Windows only due to Excel COM dependency.

---

## 📂 Folder Structure
excel_validation_report
│
├── input_form.xls # Input Excel file
├── validate_and_save_report.py # Main script: read, validate, generate output
├── reports/ # Output folder for reports
├── .gitignore # Ignores output files and temp folders
└── README.md # Project overview and usage guide


---

## ▶️ How to Run

### 1. 🔧 Install Python and Dependencies

Make sure you have Python 3.10+ installed. Then install the required packages:
pip install xlwings pywin32
2. 🏃 Run the Script

In your terminal:
python validate_and_save_report.py
✅ A new report will be saved inside the /reports folder
📌 Filename will follow this format: Report_Data_YYYY-MM-DD_HH-MM.xlsx


💡 Notes
Excel must be installed on your system (as xlwings automates Excel directly)

Only validate_and_save_report.py is required for execution

📎 GitHub Repository
🔗 github.com/MuraliV1983/excel_validation_report

🤝 Contributions
Feel free to:
Fork the repo
Add improvements (CSV/Google Sheets input, web UI, logging)
Raise issues or suggestions

🔖 Author
V Muralidharan
Proudly built as part of my 🔁 #MuraliCodes series
🔗 LinkedIn: https://www.linkedin.com/in/dharanv/
