# DDU Result Finder

A collection of Python scripts to automate finding and saving results from the DDU (Deen Dayal Upadhyaya Gorakhpur University) result portal.

## Features

- **DOB Finder (`dob-finder.py`)**: Automatically tests potential Dates of Birth for a given list of Roll Numbers to find the correct DOB needed to access a student's result. It uses asynchronous requests to quickly iterate through a configured range of possible dates and seamlessly handles retries and connections.
- **Result Saver (`result-saver.py`)**: Once the DOB is known, this script can automate the browser to navigate the student result portal, fetch the results, and save them as PDF files for offline viewing.
- **Excel Maker (`excel-maker.py`)**: Parses the converted result DOCX files to extract marks and other relevant information, compiling them into a structured Excel spreadsheet for easy viewing and analysis.

## Prerequisites

- Python 3.7+
- Necessary Python packages (you can install them using `pip`):
  ```bash
  pip install aiohttp pandas tqdm
  ```
  _Note: Depending on the specific execution methods of `result-saver.py` and `excel-maker.py`, additional packages like `selenium` or `PyPDF2`/`pdfplumber` might be required._

## Usage

### 1. Find Dates of Birth

Run the `dob-finder.py` script:

```bash
python dob-finder.py
```

You will be prompted to enter the Roll Numbers, Semester, and optionally a specific month to speed up the search. The results will be saved in an `Information/ddu_custom_results.xlsx` file.

### 2. Save Results as PDF

Run the `result-saver.py` script to automate downloading results tailored to the found DOBs.

```bash
python result-saver.py
```

### 3. Convert PDFs to DOCX

Before running the Excel Maker, you must convert the downloaded result PDFs into Word (.docx) files. We recommend using [Smallpdf's PDF to Word Converter](https://smallpdf.com/pdf-to-word#r=convert-to-word). Save these converted `.docx` files into the `Information/pdf2docx/` folder as `1.docx`, `2.docx`, etc., depending on the semester.

### 4. Combine to Excel

Run the `excel-maker.py` script to parse the `.docx` files and convert them into organized Excel spreadsheets containing overall result ranks and subject-wise ranks.

```bash
python excel-maker.py
```

## Directory Structure

- `Information/`: Default directory for saved output documents like generated Excel files and PDFs.
- `dob-finder.py`: Find date of births script.
- `result-saver.py`: PDF saving bot.
- `excel-maker.py`: DOCX to Excel parsing script.

## Disclaimer

This codebase is specifically intended for educational purposes and internal use only. Make sure you have the proper authorization to process or retrieve the results using automated scripts and ensure you are not violating the terms of service of the target result portal.
