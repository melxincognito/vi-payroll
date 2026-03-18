# VI Payroll Splitter

A web portal for splitting VI payroll Excel files by teacher. Upload your spreadsheet and instantly get a new file with each VI Teacher's students organized into their own tab, complete with pay totals.

---

## What It Does

1. You upload your payroll `.xlsx` or `.xls` file through the web portal
2. A real-time progress bar tracks each step of processing by teacher name
3. It generates a new Excel file (`VI_Teacher_Payroll.xlsx`) with:
   - **All Students** — the full original data with all 25 columns
   - **One tab per VI Teacher** — named after each teacher, containing only their students, sorted alphabetically
   - **Pay totals row** at the bottom of each teacher tab showing the sum of Amount and Quarter Pay

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
4. Watch the progress bar as each teacher's tab is built
5. Your file (`VI_Teacher_Payroll.xlsx`) will download automatically when complete

---

## Output File Details

- **All Students tab** — contains the full dataset with all 25 required columns
- **Teacher tabs** — one tab per unique VI Teacher showing only their students
- **TOTAL row** — at the bottom of each teacher tab, showing the sum of Amount and Quarter Pay formatted as currency
- Tabs are styled with a dark blue header row, alternating row colors, frozen header, and auto-sized columns
- Teacher tabs are sorted alphabetically
- Tab names longer than 31 characters (Excel's limit) are automatically trimmed

---

## Progress Tracking

Processing runs in the background so the page stays responsive. The progress bar updates in real time with labeled steps:

| Stage | Progress |
|---|---|
| Uploading file | 0–5% |
| Reading file | 5–15% |
| Validating & organizing columns | 15–25% |
| Writing All Students tab | 25–30% |
| Processing each teacher (labeled by name) | 30–95% |
| Saving file | 95–100% |

When complete, the file downloads automatically and the form resets for the next upload.

---

## Deploying to Render

1. Push this project to a GitHub repository
2. Go to [render.com](https://render.com) and create a new **Web Service**
3. Connect your GitHub repo
4. Use these settings:
   - **Environment:** Python
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
5. Click **Deploy** — Render handles everything else automatically

> **Note:** The free tier on Render spins down after inactivity. The first request after a period of no use may be slow while it wakes up.

---

## Project Structure

```
VI-Payroll/
├── app.py                  # Flask backend — file processing, progress tracking, Excel generation
├── templates/
│   └── index.html          # Web portal UI with progress bar
├── static/
│   └── styles.css          # Stylesheet
├── requirements.txt        # Python dependencies
├── Procfile                # Render/Gunicorn start command
└── README.md
```
