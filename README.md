# DDU Result Finder

A collection of Python scripts to automate finding and downloading results from the DDU (Deen Dayal Upadhyaya Gorakhpur University) result portal.

## Features

- **DOB Finder (`dob-finder.py`)**: Automatically tests potential Dates of Birth for a given list of Roll Numbers to find the correct DOB needed to access a student's result. It uses asynchronous requests to quickly iterate through a configured range of possible dates and handles connectivity issues gracefully.
- **Result Saver (`result-saver.py`)**: Once the DOB is known, this automated Chrome-based script navigates the student result portal, inputs the credentials, fetches the results page, and saves them as PDF files. It handles concurrent semester downloads and merges individual PDFs automatically.

## Recent Improvements

- **Robust Error & Timeout Handling**: Scripts now correctly utilize strict page-load timeouts, protecting against infinite hangs when the university's response server stalls.
- **Improved Date Parsing:** Ensured Excel-fed dates safely abide by non-US formatting (`DD-MM-YYYY`), securing against incorrect DOB inputs resulting in "Record Not Found" misfires.
- **Resource Management & Safety**: Integrated strict `try...finally` teardowns securing system resources. Crashed Python tasks will now reliably terminate their child `chrome.exe` headless browser instances instead of permanently leaking background memory.
- **Empty PDF Protection**: Added safeguards restricting the app from generating corrupt, empty merged PDFs when web scraping fully fails for a designated semester.

## Prerequisites

- Python 3.7+
- A working Chrome browser installation (for `result-saver.py`)
- Install the required Python packages using `pip`:

  ```bash
  pip install aiohttp pandas tqdm selenium webdriver-manager pypdf openpyxl requests
  ```

## Step-by-Step Usage

### 1. Find Dates of Birth (Optional)

If you don't know the dates of birth for the students, run the `dob-finder.py` script:

  ```bash
  python dob-finder.py
  ```

You will be prompted to enter the Roll Numbers (comma-separated), Semester, optionally a specific month to speed up the search, and whether to save the results to an Excel file or view them directly in the terminal.
_Note: If saving to Excel is selected, results will be stored inside the `Information/Saved DOBs/ddu_custom_results.xlsx` file._

### 2. Prepare the Input File for Result Saver

Before automating PDF downloading, place your Excel files containing student Roll Numbers and DOBs in `Information/Input Info/` (or use the output saved in `Information/Saved DOBs/`).

- The script automatically detects eligible `.xlsx` / `.xls` files in either `Information/Input Info/` or `Information/Saved DOBs/`.
- Ensure the Excel file contains columns for Roll Numbers (e.g., `Roll Number`, `roll_no`) and Dates of Birth (e.g., `Date of Birth`, `dob`).

### 3. Save Results as PDF

Run the `result-saver.py` script to automate downloading results for the target students:

  ```bash
  python result-saver.py
  ```

1. The script verifies that the `Information/` directory and either `Input Info/` or `Saved DOBs/` exist.
2. It detects and lists all eligible Excel files, prompting you to choose which file to use.
3. Next, it prompts for the semester(s) to download (e.g., `1, 2, 3`).
4. It launches hidden Chrome instances simultaneously to fetch and download individual PDFs, and merges them into a single PDF per semester file located at `Information/Saved Results/1.pdf`, `2.pdf`, etc.

## Directory Structure Overview

To help you organize everything, place your files into the predefined standard `Information` folders as highlighted below:

```text
DDU-Result-Finder/
├── dob-finder.py
├── result-saver.py
└── Information/
    ├── Input Info/          <-- Place your "Math Group Student Info.xlsx" here
    ├── Saved DOBs/          <-- Extracted DOBs from dob-finder.py go here
    └── Saved Results/       <-- Result-saver.py outputs merged PDFs here
```

## Disclaimer

This codebase is specifically intended for educational purposes and internal automation use only. Make sure you have the proper authorization to process or retrieve the results using automated scripts and ensure you are not violating the terms of service of the target result portal.

## AI Credits

This whole project is vibecoded using AIs like ChatGPT, Claude, Gemini & Antigravity IDE by me (Vaibhav). I have also done minor edits wherever neccessary.
