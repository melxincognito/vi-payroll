# VI Payroll Splitter

A simple web portal for splitting VI payroll Excel files by teacher. Upload your spreadsheet and instantly get a new file with each VI Teacher's students organized into their own tab.

---

## What It Does

1. You upload your payroll `.xlsx` or `.xls` file through the web portal
2. The app reads and validates all required columns
3. It generates a new Excel file (`VI_Teacher_Payroll.xlsx`) with:
   - **All Students** — the full original data with all 25 columns
   - **One tab per VI Teacher** — named after each teacher, containing only their students, sorted alphabetically by teacher name

---

## Required Columns

Your Excel file must contain the following columns (capitalization is flexible):

| Column | Column |
|---|---|
| Enrollment ID | Student First Name |
| VI Teacher | Student Last Name |
| School or District Name | Creation Date |
| Section | First Activity Date |
| Base Section Name | Last Activity Date |
| Tier | Start Date |
| Amount | Due Date |
| Quarter Pay | Days Active |
| Expected Progress | Completable Items |
| Actual Progress | Completed Items |
| Course Grade | Total Minutes |
| Messages Sent | Publisher |
| UserSpace | |

---

## Setup & Installation

### Prerequisites
- Python 3.9 or higher
- pip3

### Install Dependencies

```bash
pip3 install flask pandas openpyxl
```

### Run the App

```bash
cd /Users/melaniegonzalez/Desktop/programming/VI-Payroll
python3 app.py
```

Then open your browser and go to:

```
http://localhost:5000
```

---

## How to Use

1. Go to `http://localhost:5000`
2. Click **Browse** or drag and drop your Excel file onto the upload area
3. Click **Generate Teacher Tabs**
4. Your file (`VI_Teacher_Payroll.xlsx`) will download automatically

---

## Output File Details

- **All Students tab** — contains the full dataset with all 25 required columns
- **Teacher tabs** — one tab per unique VI Teacher, showing only that teacher's students
- Tabs are styled with a dark blue header row, alternating row colors, frozen header, and auto-sized columns
- Teacher tabs are sorted alphabetically
- Tab names longer than 31 characters (Excel's limit) are automatically trimmed

---

## Project Structure

```
VI-Payroll/
├── app.py              # Flask backend — handles file upload and Excel generation
├── templates/
│   └── index.html      # Web portal UI
└── README.md
```
