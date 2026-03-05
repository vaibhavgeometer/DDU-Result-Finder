import pdfplumber
import re
import openpyxl
from openpyxl.styles import PatternFill
import os

pdf_files = [
    "Information/merged_semester_1.pdf",
    "Information/merged_semester_2.pdf",
    "Information/merged_semester_3.pdf"
]

black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

wb = openpyxl.Workbook()
if wb.active:
    wb.remove(wb.active)

def safe_search(pattern, text, group=1, default=None, flags=0):
    match = re.search(pattern, text, flags)
    return match.group(group).strip() if match else default

def get_subjects_from_student(text):
    subjects = {}
    unresolved_subjects = []
    
    for line in text.splitlines():
        # Find all subject codes on this line, allowing them to be glued to lowercase text
        # e.g. PHED104FSports. `\b([A-Z]{1,6}\d+[A-Za-z0-9\-\*]*)\b` finds the whole block.
        codes_on_line = re.findall(r"\b([A-Z]{1,6}\d+[A-Za-z0-9\-\*]*)\b", line)
        for c in codes_on_line:
            # isolate the actual code part (uppercase letters, digits, and ending symbols, stopping before any letter that is followed by lowercase)
            # Find the split point: the index of the first lowercase letter
            idx = next((i for i, ch in enumerate(c) if ch.islower()), len(c))
            # If there was a lowercase letter, the preceding uppercase letter probably belongs to the word (e.g. "S" in "Sports").
            if idx < len(c) and idx > 0 and c[idx-1].isupper():
                idx -= 1
            code_only = c[:idx].upper()
            
            # Additional validation: the code must contain a digit to be real
            if re.search(r"\d", code_only):
                # Reject enrollment numbers (e.g. DDU5162...) or long digit strings like result2023
                if not re.search(r"\d{5,}", code_only) and not code_only.startswith("DDU") and "RESULT" not in code_only:
                    unresolved_subjects.append(code_only)
                
        # Find all 5-element mark blocks on this line (CIA, ESE, Total etc.)
        # A mark is either a digit, or '---' / '-'
        mark_blocks = re.findall(r"((?:(?:\d+|\-{1,3})\s+){4}(?:\d+|\-{1,3}))", line)
        
        for block in mark_blocks:
            parts = block.strip().split()
            if len(parts) >= 5:
                # The 5th element is the Total Marks
                obtained = parts[4]
                # Assign this mark block to the oldest unresolved subject
                if unresolved_subjects:
                    subj = unresolved_subjects.pop(0)
                    subjects[subj] = obtained
                    
    return subjects

for pdf_file in pdf_files:
    if not os.path.exists(pdf_file):
        print(f"File not found: {pdf_file}")
        continue
        
    print(f"Processing {pdf_file}...")
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    students_raw = re.split(r"Grade Sheet of Semester Examination.*?\n", full_text)[1:]
    if not students_raw:
        print(f"No student data found in {pdf_file}. Check if 'Grade Sheet' text matches.")
        continue
        
    all_subject_codes = set()
    student_data_list = []
    
    for student_text in students_raw:
        roll = safe_search(r"Roll No\s+(\d+)", student_text)
        
        name = safe_search(r"Name\s+([^\n]+)", student_text)
        if name:
            name = re.sub(r"\s+Student Type.*", "", name, flags=re.IGNORECASE).strip()
        
        sgpa_match = re.search(r"\(SGPA\)\s*:\s*([0-9]+(?:\.[0-9]+)?)", student_text)
        sgpa = float(sgpa_match.group(1)) if sgpa_match else ""
        
        cgpa_match = re.search(r"\(CGPA\)\s*:\s*([0-9]+(?:\.[0-9]+)?)", student_text)
        cgpa = float(cgpa_match.group(1)) if cgpa_match else ""
        
        result = safe_search(r"Result\s*:\s*([^\n]+)", student_text)
        if not result:
            result = "FAILED"
            
        carry = "-"
        for line in student_text.splitlines():
            if "Carry Over Paper" in line:
                parts = line.split(":", 1)
                if len(parts) > 1:
                    val = parts[1].strip()
                    if val:
                        carry = val.rstrip(",").strip()
                break
                
        subjects = get_subjects_from_student(student_text)
        
        # Determine Stream based on actual subjects to avoid false positives
        stream = "MATH"
        for code in subjects.keys():
            if code.startswith("BOT") or code.startswith("ZOO"):
                stream = "BIO"
                break
        
        all_subject_codes.update(subjects.keys())
        
        if roll:
            student_data_list.append({
                "roll": roll,
                "name": name,
                "stream": stream,
                "sgpa": sgpa,
                "cgpa": cgpa,
                "result": result,
                "carry": carry,
                "subjects": subjects
            })

    sorted_subject_codes = sorted(list(all_subject_codes))
    
    ws = wb.create_sheet(title=os.path.basename(pdf_file).replace('.pdf', ''))
    
    headers = ["Roll Number", "Student's Name", "Math/Bio", "SGPA", "CGPA", "Result", "Carry Over Paper"] + sorted_subject_codes
    ws.append(headers)
    
    for student in student_data_list:
        try:
            row_data = [
                student["roll"],
                student["name"],
                student["stream"],
                student["sgpa"],
                student["cgpa"],
                student["result"],
                student["carry"]
            ]
            
            subject_presence = []
            for code in sorted_subject_codes:
                mark = student["subjects"].get(code, "")
                row_data.append(mark)
                subject_presence.append(code in student["subjects"])
                
            ws.append(row_data)
            row_index = ws.max_row
            
            for idx, present in enumerate(subject_presence):
                if not present:
                    # 7 base headers + 1-based indexing = 8th column for 1st subject
                    cell = ws.cell(row=row_index, column=8 + idx)
                    cell.fill = black_fill
                    
        except Exception as e:
            print(f"Error writing student {student['roll']} to Excel: {e}")

output_file = os.path.join("Information", "all_results.xlsx")
wb.save(output_file)
print(f"✅ Consolidated Excel file saved as {output_file}\n")
