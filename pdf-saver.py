import os
import time
import json
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Create directory to save PDFs
output_folder = "Saved_PDFs"
os.makedirs(output_folder, exist_ok=True)

# Setup Chrome for automatic PDF saving
options = webdriver.ChromeOptions()

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
    'savefile.default_directory': os.path.abspath(output_folder),
    'profile.default_content_settings.popups': 0,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True
}

options.add_experimental_option('prefs', prefs)
options.add_argument('--kiosk-printing')
options.add_argument("--start-maximized")

# Initialize driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# URL
url = "https://ddugorakhpur.com/result2023/searchresult_new.aspx"

# Read Excel
file_path = "1st Sem Result.xlsx"
df = pd.read_excel(file_path)

# Open the page once
driver.get(url)
time.sleep(2)

# Loop through students
for index, row in df.iterrows():
    roll_no = str(row['Roll Number']).zfill(6)
    roll_suffix = roll_no[-3:]  # Last three digits
    dob = row['Date of Birth']

    # Skip if no DOB
    if pd.isna(dob):
        print(f"⚠️ Skipping Roll No: {roll_no} due to missing DOB")
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

    try:
        # Always refresh the base page
        driver.get(url)
        time.sleep(2)

        # Fill Semester
        semester_dropdown = Select(driver.find_element(By.ID, "ddlsem"))
        semester_dropdown.select_by_visible_text("Semester 1")

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

        print(f"✅ Opened result for Roll No: {roll_no} | DOB: {dob_str}")
        time.sleep(2)  # Wait for the result page to load

        # Click Print Result
        print_button = driver.find_element(By.XPATH, "//input[@value='Print Result']")
        print_button.click()
        time.sleep(2)

        # Save main window handle
        main_window = driver.current_window_handle

        # Trigger Save as PDF
        driver.execute_script('window.print();')
        time.sleep(5)  # Let PDF print dialog (or new window) open

        # Close any popup windows
        all_windows = driver.window_handles
        for handle in all_windows:
            if handle != main_window:
                driver.switch_to.window(handle)
                print("🧹 Closing popup window...")
                driver.close()

        # Return to main window
        driver.switch_to.window(main_window)

        # Rename the latest downloaded PDF
        files = sorted(os.listdir(output_folder), key=lambda x: os.path.getmtime(os.path.join(output_folder, x)), reverse=True)
        latest_file = os.path.join(output_folder, files[0])

        new_filename = f"{roll_suffix}.pdf"
        new_filepath = os.path.join(output_folder, new_filename)

        shutil.move(latest_file, new_filepath)
        print(f"💾 Saved as {new_filename}")

    except Exception as e:
        print(f"❌ Error for Roll No: {roll_no}: {e}")

# Finished
print("🎉 All results saved as PDFs!")
input("🔵 Press ENTER to exit...")
driver.quit()
