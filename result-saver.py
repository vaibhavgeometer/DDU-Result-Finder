import os
import sys
import time
import json
import shutil
import pandas as pd
import base64
import concurrent.futures
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from pypdf import PdfReader, PdfWriter
import requests
import re
import unicodedata

# Enable ANSI escape sequences on Windows console
os.system('')

class UI:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"
    CYAN    = "\033[96m"
    GREEN   = "\033[92m"
    YELLOW  = "\033[93m"
    RED     = "\033[91m"
    BLUE    = "\033[94m"
    WHITE   = "\033[97m"
    GRAY    = "\033[90m"

    @staticmethod
    def display_len(s):
        clean = re.sub(r'\033\[[0-9;]*m', '', str(s))
        return sum(2 if unicodedata.east_asian_width(c) in ('W', 'F') else 1 for c in clean)

    @staticmethod
    def pad_visible(s, target_len):
        extra = len(str(s)) - UI.display_len(s)
        return f"{str(s):<{target_len + extra}}"

    @staticmethod
    def banner(title, subtitle=""):
        width = 66
        inner = width - 2
        print(f"\n{UI.CYAN}{UI.BOLD}┌{'─' * inner}┐{UI.RESET}")
        title_str = f"  {title}"
        pad1 = inner - UI.display_len(title_str)
        print(f"{UI.CYAN}{UI.BOLD}│{UI.RESET}{UI.YELLOW}{UI.BOLD}{title_str}{' ' * max(0, pad1)}{UI.CYAN}{UI.BOLD}│{UI.RESET}")
        if subtitle:
            sub_str = f"  {subtitle}"
            pad2 = inner - UI.display_len(sub_str)
            print(f"{UI.CYAN}{UI.BOLD}│{UI.RESET}{UI.GRAY}{sub_str}{' ' * max(0, pad2)}{UI.CYAN}{UI.BOLD}│{UI.RESET}")
        print(f"{UI.CYAN}{UI.BOLD}└{'─' * inner}┘{UI.RESET}\n")

    @staticmethod
    def divider(width=66):
        print(f"{UI.GRAY}{'─' * width}{UI.RESET}")

    @staticmethod
    def card(title, items, width=66):
        print(f"{UI.CYAN}{UI.BOLD}📋 {title}{UI.RESET}")
        UI.divider(width)
        for label, val in items:
            lbl_str = f"  • {label:<18}: "
            print(f"{UI.GRAY}{lbl_str}{UI.WHITE}{val}{UI.RESET}")
        UI.divider(width)
        print()

    @staticmethod
    def success(msg):
        print(f"  {UI.GREEN}✔ {msg}{UI.RESET}")

    @staticmethod
    def error(msg):
        print(f"  {UI.RED}✖ {msg}{UI.RESET}")

    @staticmethod
    def warn(msg):
        print(f"  {UI.YELLOW}▲ {msg}{UI.RESET}")

    @staticmethod
    def info(msg):
        print(f"  {UI.BLUE}ℹ {msg}{UI.RESET}")

def find_eligible_excel_files(base_dirs):
    eligible = []
    roll_patterns = ["roll number", "roll no", "rollno", "roll_no", "roll_number", "roll"]
    dob_patterns = ["date of birth", "dob", "date_of_birth", "birth date", "birthdate"]
    
    for dir_path in base_dirs:
        if not os.path.exists(dir_path):
            continue
            
        for file_name in os.listdir(dir_path):
            if file_name.startswith("~$") or not file_name.lower().endswith((".xlsx", ".xls")):
                continue
                
            full_path = os.path.join(dir_path, file_name)
            try:
                sample_df = pd.read_excel(full_path, nrows=2)
                roll_col = None
                dob_col = None
                
                # Check exact or normalized matches first
                for col in sample_df.columns:
                    col_norm = str(col).strip().lower().replace("_", " ").replace("-", " ")
                    if not roll_col:
                        for pattern in roll_patterns:
                            if col_norm == pattern:
                                roll_col = col
                                break
                    if not dob_col:
                        for pattern in dob_patterns:
                            if col_norm == pattern:
                                dob_col = col
                                break
                
                # Fallback to substring matching if needed
                if not roll_col:
                    for col in sample_df.columns:
                        col_norm = str(col).strip().lower().replace("_", " ")
                        if "roll" in col_norm:
                            roll_col = col
                            break
                            
                if not dob_col:
                    for col in sample_df.columns:
                        col_norm = str(col).strip().lower().replace("_", " ")
                        if "dob" in col_norm or "birth" in col_norm:
                            dob_col = col
                            break
                            
                if roll_col and dob_col:
                    rel_dir = os.path.basename(dir_path)
                    eligible.append({
                        "file_name": file_name,
                        "display_name": f"[{rel_dir}] {file_name}",
                        "full_path": full_path,
                        "roll_col": roll_col,
                        "dob_col": dob_col,
                        "folder": rel_dir
                    })
            except Exception:
                continue
                
    return eligible

def get_manual_students():
    print(f"\n{UI.CYAN}{UI.BOLD}📝 Manual Student Entry{UI.RESET}")
    print(f"  {UI.GRAY}Enter single or comma-separated details, or add students one-by-one.{UI.RESET}")
    UI.divider()
    
    print(f"  {UI.WHITE}Entry style:{UI.RESET}")
    print(f"    {UI.WHITE}1.{UI.RESET} Quick Entry {UI.GRAY}(single student or comma-separated lists){UI.RESET} {UI.GRAY}[default]{UI.RESET}")
    print(f"    {UI.WHITE}2.{UI.RESET} Step-by-Step {UI.GRAY}(one student at a time){UI.RESET}")
    entry_choice = input(f"    {UI.YELLOW}Select style [1/2, default: 1]{UI.RESET}: ").strip()
    
    students = []
    
    if entry_choice == "2":
        count = 1
        while True:
            roll = input(f"\n  {UI.WHITE}Student #{count}{UI.RESET} Roll Number {UI.GRAY}(or press Enter to finish){UI.RESET}: ").strip()
            if not roll:
                if students:
                    break
                UI.error("You must enter at least one student.")
                continue
            while True:
                dob = input(f"  {UI.WHITE}Student #{count}{UI.RESET} Date of Birth {UI.GRAY}(DD-MM-YYYY or YYYY-MM-DD){UI.RESET}: ").strip()
                if dob:
                    break
                UI.error("Date of Birth cannot be empty.")
            students.append({'Roll Number': roll, 'Date of Birth': dob})
            count += 1
    else:
        while True:
            roll_input = input(f"\n  {UI.WHITE}{UI.BOLD}➤ Enter Roll Number(s){UI.RESET} {UI.GRAY}(comma-separated){UI.RESET}: ").strip()
            if roll_input:
                break
            UI.error("Roll number cannot be empty.")
            
        rolls = [r.strip() for r in roll_input.split(',') if r.strip()]
        
        while True:
            if len(rolls) > 1:
                prompt_msg = f"  {UI.WHITE}{UI.BOLD}➤ Enter Date(s) of Birth{UI.RESET} {UI.GRAY}(comma-separated, or single DOB for all){UI.RESET}: "
            else:
                prompt_msg = f"  {UI.WHITE}{UI.BOLD}➤ Enter Date of Birth{UI.RESET} {UI.GRAY}(DD-MM-YYYY or YYYY-MM-DD){UI.RESET}: "
            dob_input = input(prompt_msg).strip()
            if not dob_input:
                UI.error("Date of Birth cannot be empty.")
                continue
                
            if len(rolls) > 1:
                if ',' in dob_input:
                    dobs = [d.strip() for d in dob_input.split(',') if d.strip()]
                    if len(dobs) != len(rolls):
                        UI.error(f"Count mismatch: You entered {len(rolls)} roll numbers but {len(dobs)} DOBs.")
                        continue
                    for r, d in zip(rolls, dobs):
                        students.append({'Roll Number': r, 'Date of Birth': d})
                    break
                else:
                    for r in rolls:
                        students.append({'Roll Number': r, 'Date of Birth': dob_input})
                    break
            else:
                students.append({'Roll Number': rolls[0], 'Date of Birth': dob_input})
                break

    df = pd.DataFrame(students)
    return df

def process_semester(sem_input, df, driver_path, position):
    target_semester = "Semester " + sem_input
    
    # Create directory to save PDFs for the specific semester
    output_folder = os.path.join("Information", "Saved Results", f"Saved_PDFs_Sem_{sem_input}")
    chrome_junk_folder = os.path.join(output_folder, "chrome_junk")
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(chrome_junk_folder, exist_ok=True)
    
    # Setup Chrome for automatic PDF saving
    options = webdriver.ChromeOptions()
    options.add_argument("--log-level=3")  # Suppress browser logs from cluttering terminal

    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "printing.print_preview_sticky_settings.appState": json.dumps({
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "printPreviewSettings": {
                "pagesPerSheet": 1,
                "pages": "1"  # Save ONLY FIRST PAGE
            }
        })
    }

    prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(settings),
        'savefile.default_directory': os.path.abspath(chrome_junk_folder),
        'profile.default_content_settings.popups': 0,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True
    }

    options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing')
    options.add_argument("--start-maximized")

    # Initialize driver
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    driver.set_page_load_timeout(30)
    
    try:
        # URL
        url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"

        # Open the page once per semester
        driver.get(url)
        time.sleep(2)
        
        # Setup progress bar
        bar_format = f"  {UI.CYAN}{UI.BOLD}{{desc:<16}}{UI.RESET} {{percentage:3.0f}}%|{{bar:25}}| {{n_fmt}}/{{total_fmt}} [{UI.GRAY}{{elapsed}}<{{remaining}}{UI.RESET}]"
        pbar = tqdm(total=len(df), desc=f"Semester {sem_input}", position=position, leave=True, bar_format=bar_format)
        
        # Loop through students
        for index, row in df.iterrows():
            raw_roll = row['Roll Number']
            if pd.isna(raw_roll):
                pbar.update(1)
                continue
            if isinstance(raw_roll, float):
                roll_no = str(int(raw_roll)).zfill(6)
            else:
                roll_no = str(raw_roll).split('.')[0].strip().zfill(6)
            dob = row['Date of Birth']
        
            # Skip if no DOB
            if pd.isna(dob):
                pbar.update(1)
                continue
        
            # Format DOB
            if isinstance(dob, pd.Timestamp):
                dob_str = dob.strftime('%d-%m-%Y')
            else:
                try:
                    dob_obj = pd.to_datetime(dob, dayfirst=True)
                    dob_str = dob_obj.strftime('%d-%m-%Y')
                except:
                    dob_str = str(dob)
            
            for attempt in range(3):
                skip = False
                try:
                    # Always refresh the base page
                    driver.get(url)
                    time.sleep(2)
            
                    # Fill Semester
                    semester_dropdown = Select(driver.find_element(By.ID, "ddlsem"))
                    semester_dropdown.select_by_visible_text(target_semester)
            
                    # Fill Roll No and DOB
                    roll_input = driver.find_element(By.ID, "txtRollno")
                    roll_input.clear()
                    roll_input.send_keys(roll_no)
            
                    dob_input = driver.find_element(By.ID, "txtDob")
                    dob_input.clear()
                    dob_input.send_keys(dob_str)
            
                    # Click Search Result
                    search_button = driver.find_element(By.ID, "btnSearch")
                    search_button.click()
            
                    time.sleep(2)  # Wait for the result page to load
        
                    # Handle possible "Record Not Found" JS Alert
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        if "no record found" in alert_text.lower() or "not found" in alert_text.lower() or "invalid" in alert_text.lower():
                            alert.accept()
                            skip = True
                            break
                        alert.accept()
                    except:
                        pass
                        
                    # Handle "No Record found" text in page source (matches the red label)
                    if not skip and "no record found" in driver.page_source.lower():
                        skip = True
                        break
        
                    # Check if print button exists (Fail-safe, if the button is missing it implies no valid result loaded)
                    try:
                        print_button = driver.find_element(By.XPATH, "//input[@value='Print Result']")
                    except Exception as e:
                        raise Exception("Print button not found. Retrying in case of slow load...") from e
            
                    # Save main window handle
                    main_window = driver.current_window_handle
            
                    # Click Print Result
                    print_button.click()
                    time.sleep(2)
            
                    # Switch to the new popup window
                    for handle in driver.window_handles:
                        if handle != main_window:
                            driver.switch_to.window(handle)
                            break
            
                    # Trigger Save as PDF directly using Chrome DevTools Protocol
                    pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
                        "printBackground": True,
                        "preferCSSPageSize": True,
                        "pageRanges": "1"
                    })
            
                    new_filename = f"{roll_no}.pdf"
                    new_filepath = os.path.join(output_folder, new_filename)
            
                    with open(new_filepath, "wb") as f:
                        f.write(base64.b64decode(pdf_data['data']))
            
                    # Close the popup window
                    driver.close()
            
                    # Return to main window
                    driver.switch_to.window(main_window)
                    
                    # Success
                    break
            
                except Exception as e:
                    error_msg = str(e).lower()
                    if attempt < 2:
                        # If it's a critical webdriver crash/connection issue, restart the driver
                        if "connection" in error_msg or "actively refused" in error_msg or "forcibly closed" in error_msg or "max retries" in error_msg or "disconnected" in error_msg:
                            try:
                                driver.quit()
                            except:
                                pass
                            # Re-initialize driver
                            driver = webdriver.Chrome(service=Service(driver_path), options=options)
                            driver.set_page_load_timeout(30)
                        time.sleep(2)
                    else:
                        error_name = type(e).__name__
                        tqdm.write(f"  {UI.RED}✖ [{target_semester}] Failed Roll No: {roll_no} after 3 attempts ({error_name}){UI.RESET}")
                
            pbar.update(1)
                
        pbar.close()
        tqdm.write(f"  {UI.GREEN}✔ [{target_semester}] Finished downloading student scorecards.{UI.RESET}")

        # --- MERGING PROCESS ---
        tqdm.write(f"  {UI.CYAN}📑 [{target_semester}] Merging PDFs into single document...{UI.RESET}")
        output_pdf_path = os.path.join("Information", "Saved Results", f"{sem_input}.pdf")
        
        writer = PdfWriter()
        
        # Get all PDF files in the folder (sort alphabetically)
        pdf_files = [f for f in os.listdir(output_folder) if f.lower().endswith(".pdf")]
        pdf_files.sort()
        
        # Loop through each file and add the first page
        for pdf_file in pdf_files:
            pdf_path = os.path.join(output_folder, pdf_file)
            try:
                reader = PdfReader(pdf_path)
                if len(reader.pages) > 0:
                    writer.add_page(reader.pages[0])
            except Exception as e:
                tqdm.write(f"  {UI.RED}✖ [{target_semester}] Error processing {pdf_file}: {e}{UI.RESET}")
        
        # Write the combined PDF
        if len(writer.pages) > 0:
            with open(output_pdf_path, "wb") as out_file:
                writer.write(out_file)
            tqdm.write(f"  {UI.GREEN}✔ [{target_semester}] Merged PDF saved: {UI.BOLD}{output_pdf_path}{UI.RESET}")
        else:
            tqdm.write(f"  {UI.YELLOW}▲ [{target_semester}] No valid PDFs found to merge. Empty PDF was not created.{UI.RESET}")
        
        # --- CLEANUP PROCESS ---
        tqdm.write(f"  {UI.GRAY}🧹 [{target_semester}] Cleaning up intermediate files...{UI.RESET}")
        try:
            shutil.rmtree(output_folder)
            tqdm.write(f"  {UI.GREEN}✔ [{target_semester}] Intermediate folder cleaned.{UI.RESET}")
        except Exception as e:
            tqdm.write(f"  {UI.YELLOW}▲ [{target_semester}] Could not remove folder {output_folder}: {e}{UI.RESET}")

    finally:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    UI.banner(
        "📄 DDU RESULT FINDER  •  Result Saver",
        "Batch downloads & merges student scorecards into PDF"
    )

    info_dir = "Information"
    input_info_dir = os.path.join(info_dir, "Input Info")
    saved_dobs_dir = os.path.join(info_dir, "Saved DOBs")

    check_dirs = []
    if os.path.isdir(input_info_dir):
        check_dirs.append(input_info_dir)
    if os.path.isdir(saved_dobs_dir):
        check_dirs.append(saved_dobs_dir)

    eligible_files = find_eligible_excel_files(check_dirs) if check_dirs else []

    # Select Input Mode
    print(f"{UI.CYAN}{UI.BOLD}📋 Select Input Method{UI.RESET}")
    UI.divider()
    print(f"  {UI.WHITE}1.{UI.RESET} Load from Excel file {UI.GRAY}(scans Input Info & Saved DOBs){UI.RESET} {UI.GRAY}[default]{UI.RESET}")
    print(f"  {UI.WHITE}2.{UI.RESET} Enter manually {UI.GRAY}(Roll Numbers, DOBs & Semesters){UI.RESET}")
    UI.divider()
    
    mode_choice = input(f"  {UI.YELLOW}Choose option [1/2, default: 1]{UI.RESET}: ").strip()

    df = None
    source_label = ""

    if mode_choice == "2":
        # Manual Input Mode
        df = get_manual_students()
        source_label = f"Manual Input ({len(df)} records)"
        UI.success(f"Configured {UI.BOLD}{len(df)}{UI.RESET} student record(s) manually.")
        
        # Option to save manually entered students to Excel
        save_opt = input(f"\n  {UI.YELLOW}Save these students to Excel for future use? (y/N) [default: n]{UI.RESET}: ").strip().lower()
        if save_opt in ['y', 'yes']:
            os.makedirs(input_info_dir, exist_ok=True)
            manual_file_path = os.path.join(input_info_dir, "Manual_Students.xlsx")
            df.to_excel(manual_file_path, index=False)
            UI.success(f"Saved student sheet to: {UI.BOLD}{manual_file_path}{UI.RESET}")
    else:
        # Excel File Mode
        if not eligible_files:
            UI.warn("No eligible Excel files found in 'Information/Input Info' or 'Information/Saved DOBs'.")
            switch_manual = input(f"  {UI.YELLOW}Would you like to enter student details manually instead? (Y/n) [default: y]{UI.RESET}: ").strip().lower()
            if switch_manual in ['n', 'no']:
                print(f"\n  {UI.YELLOW}▲ Operation cancelled.{UI.RESET}\n")
                exit(0)
            df = get_manual_students()
            source_label = f"Manual Input ({len(df)} records)"
            UI.success(f"Configured {UI.BOLD}{len(df)}{UI.RESET} student record(s) manually.")
        else:
            # File Selection Menu
            print(f"\n{UI.CYAN}{UI.BOLD}📁 Available Student Data Sheets{UI.RESET}")
            UI.divider()
            for idx, file_info in enumerate(eligible_files, 1):
                tag = f"[{file_info['folder']}]"
                print(f"  {UI.WHITE}{UI.BOLD}[{idx}]{UI.RESET} {UI.CYAN}{tag:<14}{UI.RESET} {file_info['file_name']}")
            UI.divider()

            if len(eligible_files) == 1:
                choice_input = input(f"  {UI.YELLOW}Use this file? (Y/n) [default: y]{UI.RESET}: ").strip().lower()
                if choice_input in ['n', 'no']:
                    print(f"\n  {UI.YELLOW}▲ Operation cancelled by user.{UI.RESET}\n")
                    exit(0)
                selected_file = eligible_files[0]
            else:
                while True:
                    choice_input = input(f"  {UI.YELLOW}Select a file to use [1-{len(eligible_files)}, default: 1]{UI.RESET}: ").strip()
                    if not choice_input:
                        selected_file = eligible_files[0]
                        break
                    if choice_input.isdigit() and 1 <= int(choice_input) <= len(eligible_files):
                        selected_file = eligible_files[int(choice_input) - 1]
                        break
                    UI.error(f"Invalid selection. Please enter a number between 1 and {len(eligible_files)}.")

            df = pd.read_excel(selected_file['full_path'])
            df = df.rename(columns={
                selected_file['roll_col']: 'Roll Number',
                selected_file['dob_col']: 'Date of Birth'
            })
            df = df.dropna(subset=['Roll Number'])
            source_label = selected_file['display_name']
            UI.success(f"Loaded {UI.BOLD}{len(df)}{UI.RESET} student records from {UI.CYAN}{selected_file['display_name']}{UI.RESET}")

    # Semester Input
    while True:
        target_semesters_input = input(f"\n  {UI.WHITE}{UI.BOLD}➤ Enter semester(s){UI.RESET} {UI.GRAY}(comma-separated, e.g. 1 or 1, 2, 3){UI.RESET}: ").strip()
        if target_semesters_input:
            target_semester_list = [sem.strip() for sem in target_semesters_input.split(',') if sem.strip()]
            if target_semester_list:
                break
        UI.error("Please enter at least one semester number.")

    print()
    # Configuration Card
    job_items = [
        ("Source", source_label),
        ("Total Students", str(len(df))),
        ("Target Semesters", ", ".join(f"Semester {s}" for s in target_semester_list)),
        ("Destination", os.path.join(info_dir, "Saved Results"))
    ]
    UI.card("Download Configuration", job_items)

    # Check website status
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"
    print(f"  {UI.BLUE}🌐 University Portal Status :{UI.RESET} ", end="", flush=True)
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 503:
            print(f"{UI.RED}[OFFLINE]{UI.RESET}")
            UI.error("Portal is currently unavailable (503 Service Unavailable). Please try again later.")
            exit(1)
        elif response.status_code != 200:
            print(f"{UI.RED}[ERROR {response.status_code}]{UI.RESET}")
            UI.error(f"Website returned HTTP status {response.status_code}.")
            exit(1)
        print(f"{UI.GREEN}[ONLINE]{UI.RESET}")
    except Exception as e:
        print(f"{UI.RED}[ERROR]{UI.RESET}")
        UI.error(f"Error connecting to website: {e}")
        print(f"  {UI.GRAY}Please check your network connection and try again.{UI.RESET}\n")
        exit(1)

    # ChromeDriver Setup
    print(f"  {UI.BLUE}🔧 Chrome WebDriver         :{UI.RESET} ", end="", flush=True)
    try:
        driver_path = ChromeDriverManager().install()
        print(f"{UI.GREEN}[READY]{UI.RESET}")
    except Exception as e:
        print(f"{UI.RED}[FAILED]{UI.RESET}")
        UI.error(f"Could not initialize ChromeDriver: {e}")
        exit(1)

    # Ensure Saved Results directory exists
    os.makedirs(os.path.join(info_dir, "Saved Results"), exist_ok=True)

    print(f"\n  {UI.CYAN}🚀 Launching {len(target_semester_list)} browser session(s)...{UI.RESET}\n")
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(target_semester_list)) as executor:
        futures = {
            executor.submit(process_semester, sem, df, driver_path, idx): sem 
            for idx, sem in enumerate(target_semester_list)
        }
        
        for future in concurrent.futures.as_completed(futures):
            sem = futures[future]
            try:
                future.result()
            except Exception as exc:
                tqdm.write(f"  {UI.RED}✖ Task for Semester {sem} encountered an error: {exc}{UI.RESET}")

    # Completion Banner
    width = 66
    inner = width - 2
    print(f"\n{UI.GREEN}{UI.BOLD}┌{'─' * inner}┐{UI.RESET}")
    line1 = "  🎉 ALL SEMESTERS PROCESSED SUCCESSFULLY!"
    line2 = f"  Output saved in: {os.path.join(info_dir, 'Saved Results')}"
    pad1 = inner - UI.display_len(line1)
    pad2 = inner - UI.display_len(line2)
    print(f"{UI.GREEN}{UI.BOLD}│{UI.RESET}{UI.WHITE}{UI.BOLD}{line1}{' ' * max(0, pad1)}{UI.GREEN}{UI.BOLD}│{UI.RESET}")
    print(f"{UI.GREEN}{UI.BOLD}│{UI.RESET}{UI.GRAY}{line2}{' ' * max(0, pad2)}{UI.GREEN}{UI.BOLD}│{UI.RESET}")
    print(f"{UI.GREEN}{UI.BOLD}└{'─' * inner}┘{UI.RESET}\n")
