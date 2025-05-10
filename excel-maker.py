import pdfplumber
import re
import openpyxl
from openpyxl.styles import PatternFill

# ---------- Excel Setup ----------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Results"

headers = [
    "Roll Number", "Student's Name", "Math/Bio", "SGPA", "Result",
    "Carry Over Paper", "MAT101F", "MAT102F", "PHY101F", "PHY102F",
    "CHE101F", "CHE102F", "PHED101F", "PHED102F", "BOT101F", "BOT102F",
    "ZOO101F", "ZOO102F", "AE1DDSP", "SE1PHI", "SE1BOT"
]
ws.append(headers)

# Styles
black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

# ---------- Helper Functions ----------
def safe_search(pattern, text, group=1, default=None, flags=0):
    match = re.search(pattern, text, flags)
    return match.group(group).strip() if match else default

def extract_obtained_marks(text, course_code):
    for line in text.splitlines():
        if course_code in line:
            # Split the line into parts and filter out empty strings
            parts = line.strip().split()
            if len(parts) >= 4:
                return parts[-4]  # 4th last column
    return None

# ---------- Load PDF ----------
pdf_path = "C:/Users/Vaibhav/Documents/Projects/ddu-result-code/1st Sem Result.pdf"  # Replace with your full PDF path
with pdfplumber.open(pdf_path) as pdf:
    full_text = ""
    for page in pdf.pages:
        full_text += page.extract_text() + "\n"

# ---------- Parse Students ----------
students_raw = re.split(r"Grade Sheet of Semester Examination.*?\n", full_text)[1:]

for student_text in students_raw:
    try:
        roll = safe_search(r"Roll No\s+(\d+)", student_text)
        name = safe_search(r"Name\s+([A-Z\s]+)", student_text)
        # --- SGPA ---
        sgpa_match = re.search(r"Semester Grade Point Average \(SGPA\)\s*:\s*([0-9]+\.[0-9]+)", student_text)
        sgpa = float(sgpa_match.group(1)) if sgpa_match else ""
        #--------------
        result = safe_search(r"Result\s*:\s*(PASSED|FAILED)", student_text)
        if not result:
            result = "FAILED"
        # --- Carry Over Paper ---
        carry = "-"
        for line in student_text.splitlines():
            if "Carry Over Paper" in line:
                parts = line.split(":", 1)
                if len(parts) > 1:
                    value = parts[1].strip()
                    if value:
                        carry = value
                break  # Only use the first match

        stream = "BIO" if "BOT101F" in student_text else "MATH"

        # Subjects to extract
        subject_codes = [
            "MAT101F", "MAT102F", "PHY101F", "PHY102F","CHE101F", "CHE102F",
            "PHED101F", "PHED102F", "BOT101F", "BOT102F", "ZOO101F", "ZOO102F",
            "AE1DDSP", "SE1PHI", "SE1BOT"
        ]

        marks = []
        subject_presence = []
        for code in subject_codes:
            mark = extract_obtained_marks(student_text, code)
            marks.append(mark if mark is not None else "")
            subject_presence.append(code in student_text)

        # Write data row
        row_data = [roll, name, stream, sgpa, result, carry] + marks
        ws.append(row_data)

        # Apply black fill for subjects not taken
        row_index = ws.max_row
        for idx, present in enumerate(subject_presence):
            if not present:
                cell = ws.cell(row=row_index, column=7 + idx)
                cell.fill = black_fill
                cell.value = ""

    except Exception as e:
        print(f"\nError processing student block:\n{student_text[:300]}...\nError: {e}")

# ---------- Save Excel ----------
output_file = "all_students_results.xlsx"
wb.save(output_file)
print(f"✅ Excel file saved as {output_file}")
