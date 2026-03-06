import os
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

def process_semester(sem_input, df, driver_path, position):
    target_semester = "Semester " + sem_input
    
    # Create directory to save PDFs for the specific semester
    output_folder = os.path.join("Information", f"Saved_PDFs_Sem_{sem_input}")
    os.makedirs(output_folder, exist_ok=True)
    
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
        'savefile.default_directory': os.path.abspath("Saved_PDFs"),
        'profile.default_content_settings.popups': 0,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True
    }

    options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing')
    options.add_argument("--start-maximized")

    # Initialize driver
    driver = webdriver.Chrome(service=Service(driver_path), options=options)

    # URL
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"

    # Open the page once per semester
    driver.get(url)
    time.sleep(2)
    
    # Setup progress bar
    pbar = tqdm(total=len(df), desc=f"Semester {sem_input}", position=position, leave=True)
    
    # Loop through students
    for index, row in df.iterrows():
        roll_no = str(row['Roll Number']).zfill(6)
        roll_suffix = roll_no[-3:]  # Last three digits
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
                dob_obj = pd.to_datetime(dob)
                dob_str = dob_obj.strftime('%d-%m-%Y')
            except:
                dob_str = dob
                
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
                except:
                    skip = True
                    break
        
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
                    "preferCSSPageSize": True
                })
        
                new_filename = f"{roll_suffix}.pdf"
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
                    time.sleep(2)
                else:
                    error_name = type(e).__name__
                    tqdm.write(f"❌ [{target_semester}] Failed Roll No: {roll_no} after 3 attempts ({error_name})")
            
        pbar.update(1)
            
    pbar.close()
    tqdm.write(f"🎉 [{target_semester}] Finished downloading PDFs!")

    # --- MERGING PROCESS ---
    tqdm.write(f"🔄 [{target_semester}] Starting merging process...")
    output_pdf_path = os.path.join("Information", f"{sem_input}.pdf")
    
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
            tqdm.write(f"❌ [{target_semester}] Error processing {pdf_file}: {e}")
    
    # Write the combined PDF
    with open(output_pdf_path, "wb") as out_file:
        writer.write(out_file)
    
    tqdm.write(f"✅ [{target_semester}] Merged PDF saved to {output_pdf_path}")
    
    # --- CLEANUP PROCESS ---
    tqdm.write(f"🧹 [{target_semester}] Cleaning up individual PDF files...")
    try:
        shutil.rmtree(output_folder)
        tqdm.write(f"✅ [{target_semester}] Removed intermediate folder.")
    except Exception as e:
        tqdm.write(f"⚠️ [{target_semester}] Could not remove folder {output_folder}: {e}")

    driver.quit()

if __name__ == "__main__":
    # Ask for semester input
    target_semesters_input = input("Enter the semesters separated by commas (e.g. 1, 2, 3): ").strip()
    if not target_semesters_input:
        print("❌ No semesters provided!")
        exit(1)
        
    target_semester_list = [sem.strip() for sem in target_semesters_input.split(',')]
    
    # Install driver once to avoid caching conflicts when starting multiple drivers concurrently
    print("🔄 Ensuring latest ChromeDriver is installed...")
    driver_path = ChromeDriverManager().install()

    # Read Excel
    file_path = "Information/Math Group Student Info.xlsx"
    if not os.path.exists(file_path):
        print(f"❌ Could not find {file_path}")
        exit(1)
        
    df = pd.read_excel(file_path)
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"

    # Check website status
    print("🔍 Checking website status...")
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 503:
            print("🛑 Website is DOWN: Service Unavailable (503)")
            print("Please try again later.")
            exit(1)
        elif response.status_code != 200:
            print(f"🛑 Website returned HTTP {response.status_code}")
            print("Please try again later.")
            exit(1)
        print("✅ Website is UP!")
    except Exception as e:
        print(f"🛑 Error connecting to website: {e}")
        print("Please check your internet connection and try again.")
        exit(1)

    # Use ThreadPoolExecutor to run tasks simultaneously
    print(f"🚀 Launching {len(target_semester_list)} simultaneous browser sessions...\n")
    with concurrent.futures.ThreadPoolExecutor(max_workers=len(target_semester_list)) as executor:
        # Submit tasks to threads with index for tqdm positioning
        futures = {
            executor.submit(process_semester, sem, df, driver_path, idx): sem 
            for idx, sem in enumerate(target_semester_list)
        }
        
        # Wait for them to finish and catch any exceptions
        for future in concurrent.futures.as_completed(futures):
            sem = futures[future]
            try:
                future.result()
            except Exception as exc:
                tqdm.write(f"❌ Task for Semester {sem} generated an exception: {exc}")

    print("\n🚀 All processes completed successfully!")
