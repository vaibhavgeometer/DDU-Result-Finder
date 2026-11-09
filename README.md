# DDU Result Finder

A Python-based toolkit to automate finding student Dates of Birth (DOB) and batch-downloading official semester results as merged PDFs from the Deen Dayal Upadhyaya Gorakhpur University (DDU) result portal.

---

## Table of Contents

- [Overview](#overview)
- [How It Works (Under the Hood)](#how-it-works-under-the-hood)
  - [1. DOB Finder (`dob-finder.py`)](#1-dob-finder-dob-finderpy)
  - [2. Result Saver (`result-saver.py`)](#2-result-saver-result-saverpy)
- [Prerequisites & Installation](#prerequisites--installation)
- [Step-by-Step Usage Guide](#step-by-step-usage-guide)
  - [Workflow A: When You Need to Find Student DOBs First](#workflow-a-when-you-need-to-find-student-dobs-first)
  - [Workflow B: When You Already Have an Excel Sheet with Roll Numbers & DOBs](#workflow-b-when-you-already-have-an-excel-sheet-with-roll-numbers--dobs)
- [Directory Structure](#directory-structure)
- [Troubleshooting & Tips](#troubleshooting--tips)
- [Disclaimer](#disclaimer)
- [AI Credits](#ai-credits)

---

## Overview

Accessing student results on the DDU portal requires both a **Roll Number** and the student's **Date of Birth (DOB)**. When processing results for an entire class or batch, retrieving individual DOBs and manually saving each student's scorecard is tedious and time-consuming.

This toolkit solves both problems:
1. **`dob-finder.py`**: Rapidly discovers the matching DOB for any student by testing high-probability birthdates using asynchronous HTTP requests.
2. **`result-saver.py`**: Reads any eligible Excel file containing roll numbers and DOBs, automates browser sessions to fetch result pages, renders first-page PDFs using Chrome DevTools Protocol, and merges them into a single, clean PDF per semester.

---

## How It Works (Under the Hood)

### 1. DOB Finder (`dob-finder.py`)

Finding a student's unknown DOB brute-force across thousands of calendar days would normally take minutes per student. `dob-finder.py` optimizes this through:

- **Asynchronous HTTP Concurrency (`aiohttp` + `asyncio`)**:
  Rather than making blocking sequential requests, the script pools asynchronous requests with a concurrency limit (`asyncio.Semaphore(50)`) and a `TCPConnector` cap. This allows testing dozens of dates per second without overwhelming local resources or the server.
- **ASP.NET Form & Token Handling**:
  The DDU result portal is built on ASP.NET Web Forms. For each lookup, `dob-finder.py` extracts the required dynamic form tokens (`__VIEWSTATE`, `__EVENTVALIDATION`, `__VIEWSTATEGENERATOR`), packages them with the target semester, roll number, and candidate DOB, and posts the payload.
- **HTTP 302 Success Detection**:
  When a Roll Number and DOB match a valid university record, the ASP.NET portal issues an **HTTP 302 redirect** to the results page. The script intercepts this redirect without downloading the heavy result HTML, confirming the correct DOB instantly.
- **Interleaved Priority Groups**:
  Students within a given college semester typically share a narrow age band. Dates are generated in tiered priority groups:
  1. *Primary cohort (highest probability)*: 2003 – 2008
  2. *Secondary cohort*: 1999 – 2002 & 2009 – 2010
  3. *Extended cohort*: 1996 – 1998 & 2011 – 2013  
  Within each group, dates are **interleaved across years** (e.g., testing January 1 across all priority years first) so that matches are discovered much earlier in the search cycle. You can also supply an optional **month filter** to speed up the process by ~12x if the birth month is known.
- **Resume & Dual Output Modes**:
  - The script checks if `Information/Saved DOBs/ddu_custom_results.xlsx` exists and automatically skips previously solved roll numbers when resuming.
  - You can choose to **save results to Excel** or **view them only in your terminal**.
  - A clean summary table is displayed in the terminal upon completion, with real-time discoveries logged cleanly via `tqdm.write` without disrupting progress bars.

---

### 2. Result Saver (`result-saver.py`)

Once you have roll numbers and DOBs, `result-saver.py` automates the entire download and compilation pipeline:

- **Dynamic Folder & Excel File Discovery**:
  On startup, the script verifies that the `Information/` directory and at least one source folder (`Input Info/` or `Saved DOBs/`) exist. It scans for `.xlsx` and `.xls` files, inspecting the headers of each file.
- **Flexible Column Recognition**:
  You do not need to rename your columns to an exact string. The script automatically detects and standardizes common column naming conventions:
  - **Roll Number**: `Roll Number`, `roll_no`, `Roll No`, `rollno`, `roll`, etc.
  - **Date of Birth**: `Date of Birth`, `dob`, `date_of_birth`, `birth date`, etc.
  If multiple eligible files are found, it presents an interactive selection menu.
- **Automated ChromeDriver Management**:
  Uses `webdriver-manager` to ensure the correct ChromeDriver matching your installed Google Chrome version is downloaded and configured before starting browser sessions, avoiding thread race conditions.
- **Parallel Multi-Semester Processing (`ThreadPoolExecutor`)**:
  When downloading results for multiple semesters (e.g., `1, 2, 3`), the script launches an independent, headless/kiosk-configured Chrome browser instance for each semester simultaneously, each with its own visual progress bar.
- **Resilient Navigation & Retry Handling**:
  For each student record, the script:
  1. Opens the result portal and selects the semester from the dropdown.
  2. Inputs the student's Roll Number and properly formatted DOB (`DD-MM-YYYY`).
  3. Submits the form and checks for "No Record Found" alerts or warnings.
  4. Includes up to 3 automatic retries per student, recovering and restarting the browser driver if connection drops occur.
- **Direct Chrome DevTools Protocol (CDP) PDF Generation**:
  Instead of relying on unstable OS print dialogs, the script sends a raw CDP command (`Page.printToPDF`) to Chrome to render the exact first page of the printable result into base64 PDF bytes.
- **Automatic PDF Merging & Cleanup**:
  All individual student PDFs for a semester are read in alphanumeric order using `pypdf`, combined into a single master document (`Information/Saved Results/<semester>.pdf`), and the temporary folder of individual PDFs is cleanly deleted.

---

## Prerequisites & Installation

### Requirements
- **Python 3.7+**
- **Google Chrome** browser installed on your machine

### 1. Clone or Download the Repository
```bash
git clone https://github.com/VaibhavPriyadarshi/DDU-Result-Finder.git
cd DDU-Result-Finder
```

### 2. Install Required Python Packages
```bash
pip install aiohttp pandas tqdm selenium webdriver-manager pypdf openpyxl requests
```

---

## Step-by-Step Usage Guide

### Workflow A: When You Need to Find Student DOBs First

If you have a list of roll numbers but do not know their dates of birth:

1. **Run the DOB finder**:
   ```bash
   python dob-finder.py
   ```
2. **Provide the requested inputs**:
   - **Roll Numbers**: Enter comma-separated roll numbers (e.g., `2515075160, 2515075161`).
   - **Semester**: Enter the semester number (e.g., `1`).
   - **Month (Optional)**: If you know the student's birth month (1–12), enter it to narrow the search; otherwise press <kbd>Enter</kbd> to search across the whole year.
   - **Output Mode**: Choose whether to save the results to an Excel file (`1` or `y`) or view them only in the terminal (`2` or `n`).
3. **Review Results**:
   - If saved to Excel, the output will be stored in `Information/Saved DOBs/ddu_custom_results.xlsx`.
   - A summary table will also be displayed directly in your terminal.

Next, proceed to [Workflow B](#workflow-b-when-you-already-have-an-excel-sheet-with-roll-numbers--dobs) to download the result PDFs using the generated sheet.

---

### Workflow B: When You Already Have an Excel Sheet with Roll Numbers & DOBs

1. **Place your Excel file**:
   - Place any `.xlsx` file containing roll numbers and DOBs into `Information/Input Info/` (or use the file created in `Information/Saved DOBs/` from Workflow A).
   - Ensure the sheet has a column for roll numbers (e.g., `Roll Number` or `roll_no`) and birth dates (e.g., `Date of Birth` or `dob`).
2. **Run the Result Saver**:
   ```bash
   python result-saver.py
   ```
3. **Select your file & semesters**:
   - The script will display all eligible Excel files detected in `Information/Input Info/` and `Information/Saved DOBs/`. Select the number corresponding to your desired file.
   - Enter the target semester(s) separated by commas (e.g., `1` or `1, 2, 3`).
4. **Collect Merged PDFs**:
   - The script will launch Chrome browser sessions, download the results, and merge them.
   - Find your final merged PDFs inside `Information/Saved Results/` named `1.pdf`, `2.pdf`, etc.

---

## Directory Structure

```text
DDU-Result-Finder/
├── dob-finder.py           # Asynchronous DOB discovery tool
├── result-saver.py         # Automated PDF result fetcher and merger
├── README.md               # Documentation
└── Information/            # Working data directory
    ├── Input Info/         # Place input student Excel sheets here (e.g., Math Group Student Info.xlsx)
    ├── Saved DOBs/         # Output location for dob-finder.py (ddu_custom_results.xlsx)
    └── Saved Results/      # Final merged semester PDFs (e.g., 1.pdf, 2.pdf)
```

---

## Troubleshooting & Tips

- **Website Down / 503 Service Unavailable**:
  The DDU result server occasionally experiences heavy traffic or maintenance downtimes. Both scripts include pre-flight checks that test if the website is responding before starting. If reported down, wait a short while and retry.
- **Excel File Permission Error (`PermissionError`)**:
  If you have an output Excel file (such as `ddu_custom_results.xlsx`) open in Microsoft Excel, Windows will lock the file and prevent the script from saving updates. Simply close the file in Excel and the script will automatically resume saving.
- **Non-US Date Formats**:
  Date parsing automatically handles both ISO dates (`YYYY-MM-DD`) and standard Indian/UK date formats (`DD-MM-YYYY`), preventing "Record Not Found" errors caused by swapped month/day fields.
- **Speeding up DOB Finder**:
  If you have partial information about students (e.g., birth month or approximate year), enter the specific month when prompted to cut the date search space by ~90%.

---

## Disclaimer

This project is intended strictly for educational, personal, and administrative automation purposes. Ensure you have appropriate authorization before querying and retrieving student records, and adhere to the terms of service of the host institution.

---

## AI Credits

This project was built and refined with the help of AI tools including ChatGPT, Claude, Gemini, and Google Antigravity IDE by [Vaibhav Priyadarshi](https://github.com/VaibhavPriyadarshi), with manual reviews and custom edits.
