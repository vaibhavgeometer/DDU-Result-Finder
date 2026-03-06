# DDU Result Finder

A collection of Python scripts to automate finding and saving results from the DDU (Deen Dayal Upadhyaya Gorakhpur University) result portal.

## Features

- **DOB Finder (`dob-finder.py`)**: Automatically tests potential Dates of Birth for a given list of Roll Numbers to find the correct DOB needed to access a student's result. It uses asynchronous requests to quickly iterate through a configured range of possible dates and handles connectivity issues gracefully.
- **Result Saver (`result-saver.py`)**: Once the DOB is known, this automated Chrome-based script navigates the student result portal, inputs the credentials, fetches the results page, and saves them as PDF files. It handles concurrent semester downloads and merges individual PDFs automatically.
- **Excel Maker (`excel-maker.py`)**: Parses the converted result `.docx` files to extract academic marks, SGPA/CGPA, and other relevant information. It compiles the gathered data into a structured Excel spreadsheet containing overall rank lists as well as individual subject-wise rankings.

## Prerequisites

- Python 3.7+
- A working Chrome browser installation (for `result-saver.py`)
- Install the required Python packages using `pip`:
  ```bash
  pip install aiohttp pandas tqdm selenium webdriver-manager pypdf python-docx openpyxl requests
  ```

## Step-by-Step Usage

### 1. Find Dates of Birth (Optional)
If you don't know the dates of birth for the students, run the `dob-finder.py` script:
```bash
python dob-finder.py
```
You will be prompted to enter the Roll Numbers (comma-separated), Semester, and optionally a specific month to speed up the search. 
*Note: The results from this script will be saved locally inside the `Information/Saved DOBs/ddu_custom_results.xlsx` file.*

### 2. Prepare the Input File for Result Saver
Before you can automate the PDF downloading, `result-saver.py` expects a specific file containing your students' Roll Numbers and DOBs. 
- Create or place an Excel file named `Math Group Student Info.xlsx` inside the `Information/Input Info/` directory.
- Ensure the Excel file contains columns named `Roll Number` and `Date of Birth` for all target students.

### 3. Save Results as PDF
Run the `result-saver.py` script to automate downloading results for the target students.
```bash
python result-saver.py
```
The script will ask which semesters to download (e.g., `1, 2, 3`). It will launch hidden Chrome instances simultaneously to fetch and download individual PDFs, and then merge them into a single PDF per semester file located at `Information/Saved Results/1.pdf`, `2.pdf`, etc.

### 4. Convert Merged PDFs to DOCX
Before parsing the data into Excel, you need to manually convert the merged result PDFs into Word (`.docx`) files. 
- A recommended platform that preserves the document formatting easily is [Smallpdf's PDF to Word Converter](https://smallpdf.com/pdf-to-word#r=convert-to-word). 
- Once converted, save these `.docx` files into the `Information/pdf2docx/` folder and name them exactly by their semester, such as `1.docx`, `2.docx`, or `3.docx`.

### 5. Combine to Excel
Run the `excel-maker.py` script to parse the recently saved `.docx` files.
```bash
python excel-maker.py
```
This script will transform the data into organized Excel files for every semester, neatly outputted to `Information/docx2xlsx/`. These files will contain two types of sheets:
- **Overall Results**: Ranked layout of all students based on their SGPA/CGPA.
- **Subject-Specific Sheets**: Individual layouts that rank students solely based on marks obtained in a given subject.

## Directory Structure Overview

To help you organize everything, place your files into the predefined standard `Information` folders as highlighted below:
```text
DDU-Result-Finder/
├── dob-finder.py
├── result-saver.py
├── excel-maker.py
└── Information/
    ├── Input Info/          <-- Place your "Math Group Student Info.xlsx" here
    ├── Saved DOBs/          <-- Extracted DOBs from dob-finder.py go here
    ├── Saved Results/       <-- Result-saver.py outputs merged PDFs here
    ├── pdf2docx/            <-- You must place your converted 1.docx, 2.docx files here
    └── docx2xlsx/           <-- Excel-maker.py outputs final structured Excel results here
```

## Disclaimer

This codebase is specifically intended for educational purposes and internal automation use only. Make sure you have the proper authorization to process or retrieve the results using automated scripts and ensure you are not violating the terms of service of the target result portal.
