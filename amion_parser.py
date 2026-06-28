#!/usr/bin/env python3
"""
Amion HTML Parser for Moonlighting Summaries

Parses Amion HTML exports (Residency Schedule + Contrast Coverage Schedule)
and generates trainee moonlighting summaries.

Usage: python amion_parser.py <folder_path>
Example: python amion_parser.py April
"""

import pandas as pd
from bs4 import BeautifulSoup
import os
import sys
import re
from datetime import datetime
from collections import defaultdict

# =============================================================================
# SHIFT MAPPINGS: Amion shift names → (display_name, hours)
# =============================================================================

# Residency Schedule shifts (shifts as columns, dates as rows)
RESIDENCY_SHIFT_MAPPINGS = {
    # ECEP shifts
    "ECEP Weekday 5p-11p": ("ECEP", 6),
    "ECEP Holiday-Weekend 5p-9p": ("ECEP", 5),

    # Supershift
    "Supershift Coverage 5p-9p": ("Bellevue Supershift", 4),

    # Weekday Scanner Coverage
    "Scanner Coverage - Gram Weekday 5p-9p": ("Gramercy weekday", 4),
    "Scanner Coverage - HCC Weekday 5p-7p": ("HCC weekday", 2),
    "Scanner Coverage - 41st St Weekday 7p-9p": ("41st St", 2),
    "Scanner Coverage - 53rd St Weekday 5p-9p": ("53rd St weekday", 4),
    "Scanner Coverage - CBI Weekday 5p-9p": ("CBI", 4),
    "Scanner Coverage - Bay Ridge Weekday 6p-9p": ("Bay Ridge", 4),  # includes travel
    "Scanner Coverage - CC Weekday 5p-8p": ("Cancer Center weekday", 3),

    # Weekend Scanner Coverage
    "Scanner Coverage - Gram Weekend 9a-5p": ("Gramercy weekend", 8),
    "Scanner Coverage - Gram Weekend 8a-4p": ("Gramercy weekend", 8),
    "Scanner Coverage - CBI Weekend 8a-4p": ("CBI weekend", 8),
    "Scanner Coverage - Bay Ridge Weekend 8a-4p": ("Bay Ridge", 8),
    "Scanner Coverage - Forest Hills Weekend 8a-4p": ("Forest Hills", 8),
    "Scanner Coverage - 53rd Weekend 8a-4p": ("53rd St weekend", 8),
    "Scanner Coverage - Garden City Weekend 8a-4p": ("Garden City", 8),
    "Scanner Coverage - Lake Success Weekend 8a-4p": ("Lake Success", 8),
    "Scanner Coverage - 32nd St Weekend 8a-4p": ("32nd St weekend", 8),
    "Scanner Coverage - 41st St Weekend 8a-4p": ("41st St weekend", 8),
    "Scanner Coverage - 41st St Weekend 4p-8p": ("41st St weekend PM", 4),
    "Scanner Coverage - Greenpoint Weekend 8a-4p": ("Greenpoint", 8),
    "Scanner Coverage - CC Weekend 8a-5:30p": ("Cancer Center weekend 8A-5:30PM", 9.5),
    "Scanner Coverage - CC Weekend 8a-4p": ("Cancer Center weekend 8A-4PM", 8),
    "Scanner Coverage - CC Weekend 8:30a-4p": ("Cancer Center weekend", 7.5),

    # IGNORED: Weekend and Holiday ED - BH/TH (not moonlighting)
    # IGNORED: Night Float Coverage - BH (not moonlighting)
}

# Contrast Coverage Schedule shifts (shifts as rows, dates as columns)
CONTRAST_SHIFT_MAPPINGS = {
    # Fellow early shifts
    "32nd Early (7a-8a)": ("7AM 32nd", 1),
    "CBI Early (7a-8a)": ("7AM CBI", 1),
    # IGNORED: 41st Early - attending only

    # Late shifts
    "32nd Late (5p-9p)": ("32nd Late", 4),
    "41st Late (5p-7p)": ("41st St late", 2),

    # Saturday shifts
    "32nd Sat (8a-4p)": ("32nd St weekend", 8),
    "41st Sat (8a-4p)": ("41st St weekend", 8),
    "CBI Sat (8a-4p)": ("CBI weekend", 8),
    "HCC Sat (9a-5p)": ("HCC weekend", 8),

    # Weekend coverage
    "53rd (8a-5p)": ("53rd St weekend", 9),
    "Gramercy (8a-5p)": ("Gramercy weekend", 9),
    "Midwood (8a-4p) 813 Quentin Rd (718) 832-1455": ("Midwood", 8),
    "Greenpoint (8a-4p) 74 Kent St (929) 455-3166": ("Greenpoint", 8),

    # FPO Weekend (for fellows)
    "FPO Weekend": ("FPO Weekend", 8),

    # IGNORED: MC Late shifts (MC Neuro, MSK, Body, Chest) - attending only

    # Winthrop
    "Winthrrop Chest Weekend 8a-11a": ("Winthrop Chest", 3),
    "Winthrop Peds Call Week": ("Winthrop Peds", 8),

    # LID
    "LID Sunday (Use Amion LID Set 24 Sunday Rules)": ("LID Sunday", 8),
    "LID Weekday C+ Coverage": ("LID Weekday", 4),

    # Premier Fellow - 8 hour shift but fellows only get paid for 1 hour
    "Premier Fellow": ("Premier", 1),
}


def load_trainees(base_path):
    """Load residents and fellows from Trainees.xlsx"""
    trainees_path = os.path.join(base_path, "Trainees.xlsx")
    df = pd.read_excel(trainees_path)

    residents = set(df['Residents'].dropna().str.strip().tolist())
    fellows = set(df['Fellows'].dropna().str.strip().tolist())

    # Create normalized name lookup (handles "Last, First" vs "First Last")
    all_trainees = residents | fellows

    return all_trainees, residents, fellows


# =============================================================================
# NAME ALIASES: Amion name → Display name (for people listed differently)
# =============================================================================
NAME_ALIASES = {
    "Young Joon Kwon": "Fred Kwon",
    "Lillian Chiu-Kennedy": "Lillian Chiu",
}


def normalize_name(name):
    """Normalize name to 'First Last' format"""
    if not name:
        return None

    name = name.strip()
    name = re.sub(r'\s+', ' ', name)  # Collapse whitespace
    name = name.replace('\xa0', ' ')  # Replace non-breaking space

    # Handle "Last, First" format
    if ',' in name:
        parts = name.split(',', 1)
        if len(parts) == 2:
            last = parts[0].strip()
            first = parts[1].strip()
            name = f"{first} {last}"

    # Apply name aliases
    if name in NAME_ALIASES:
        name = NAME_ALIASES[name]

    return name


def build_abbreviation_map(soup):
    """Build a mapping from abbreviations to full names from the HTML legend"""
    abbrev_map = {}

    # Look for the legend table that contains name mappings
    # Format: "Anna ChenACh" where "ACh" is the abbreviation
    # Or "Last, FirstAB" format

    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        for row in rows:
            cells = row.find_all(['td', 'th'])
            for cell in cells:
                cell_text = cell.get_text(strip=True)

                # Skip if too short or no letters
                if len(cell_text) < 4:
                    continue

                # Pattern: Full name followed by 2-3 uppercase letter abbreviation
                # The abbreviation is typically the last 2-3 chars, all uppercase or mixed case starting with uppercase
                # Examples: "Anna ChenACh", "Emmy HuEH", "Jacy LundbergJLu"

                # Try to find where the abbreviation starts
                # Look for pattern where we have "Name Name" followed by uppercase letters
                # The abbreviation starts where we see uppercase after a lowercase letter

                # First, strip any phone numbers at the end
                cell_text = re.sub(r'\d{3}[-.]?\d{3}[-.]?\d{4}.*$', '', cell_text)
                cell_text = re.sub(r'\d{3}-\d{3}-\d{4}.*$', '', cell_text)

                # Find the abbreviation: 2-3 chars at the end that are uppercase or start with uppercase
                # But we need to be careful about names like "Emmy Hu" where "Hu" looks like an abbrev
                match = re.match(r'^(.+?[a-z])\s*([A-Z][A-Za-z]?[a-z]?)$', cell_text)
                if match:
                    full_name = match.group(1).strip()
                    abbrev = match.group(2)
                    if len(abbrev) >= 2 and len(full_name) > 3 and (' ' in full_name or ',' in full_name):
                        normalized = normalize_name(full_name)
                        if normalized:
                            abbrev_map[abbrev] = normalized
                    continue

                # Also try: name ending with lowercase, then 2-4 uppercase chars
                match = re.match(r'^(.+[a-z])([A-Z]{2,4})$', cell_text)
                if match:
                    full_name = match.group(1).strip()
                    abbrev = match.group(2)
                    if len(full_name) > 3 and (' ' in full_name or ',' in full_name):
                        normalized = normalize_name(full_name)
                        if normalized:
                            abbrev_map[abbrev] = normalized
                    continue

                # Handle mixed case abbreviations like "ACh", "JKi", "MHa"
                # Pattern: name ending in lowercase, then Uppercase + lowercase + optional char
                match = re.match(r'^(.+[a-z])([A-Z][A-Za-z]{1,2})$', cell_text)
                if match:
                    full_name = match.group(1).strip()
                    abbrev = match.group(2)
                    if len(full_name) > 3 and (' ' in full_name or ',' in full_name):
                        normalized = normalize_name(full_name)
                        if normalized:
                            abbrev_map[abbrev] = normalized

    return abbrev_map


def parse_date_header(date_str, year=2026):
    """Parse date from header like 'Wed4-1' or 'Wed  4/1' to datetime"""
    date_str = date_str.strip()
    date_str = re.sub(r'\s+', '', date_str)  # Remove spaces

    # Pattern: DayM-D or DayM/D
    match = re.search(r'(\d{1,2})[-/](\d{1,2})', date_str)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        try:
            return datetime(year, month, day)
        except ValueError:
            return None
    return None


def parse_date_row(date_str, year=2026):
    """Parse date from row like 'Wed  4/1' to datetime"""
    return parse_date_header(date_str, year)


def parse_residency_schedule(html_path, trainees, year=2026):
    """
    Parse Residency Schedule HTML (shifts as columns, dates as rows)
    Returns list of (date, shift_display, person, hours) tuples
    """
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    results = []
    best_table = None
    best_table_data_rows = 0

    # Find the best schedule table (look for table with shift headers AND date rows)
    candidates = []
    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        if len(rows) < 10:
            continue

        # Check if first row has shift names
        first_row = rows[0]
        first_cells = first_row.find_all(['td', 'th'])
        first_texts = [c.get_text(strip=True) for c in first_cells]

        has_shift_headers = any('ECEP' in t or 'Scanner Coverage' in t for t in first_texts)
        if not has_shift_headers:
            continue

        # Count data rows (rows that start with a date)
        data_row_count = 0
        for row in rows[1:]:
            cells = row.find_all(['td', 'th'])
            if cells:
                first_text = cells[0].get_text(strip=True)
                if re.search(r'\d+[/-]\d+', first_text):  # Date pattern
                    data_row_count += 1

        # Check if first column header is clean (not mixed with other content)
        first_header = first_texts[0] if first_texts else ''
        is_clean = first_header == '' or first_header.startswith('ECEP')

        candidates.append((data_row_count, is_clean, len(rows), rows, first_texts))

    # Sort by: most data rows, then clean headers, then fewest total rows
    if candidates:
        candidates.sort(key=lambda x: (-x[0], -x[1], x[2]))
        best_table = (candidates[0][3], candidates[0][4])

    if not best_table:
        return results

    rows, shift_columns = best_table

    # Parse data rows (dates as first column)
    for row in rows[1:]:
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue

        cell_texts = [c.get_text(strip=True) for c in cells]
        if not cell_texts:
            continue

        # First cell should be the date
        date_str = cell_texts[0]
        date = parse_date_row(date_str, year)
        if not date:
            continue

        # Parse each shift column
        for col_idx, person_name in enumerate(cell_texts[1:], start=1):
            if col_idx >= len(shift_columns):
                break

            shift_name = shift_columns[col_idx]
            if not shift_name or shift_name == '-':
                continue

            # Normalize person name
            person_name = person_name.strip()
            if not person_name or person_name == '-':
                continue

            normalized_name = normalize_name(person_name)
            if not normalized_name:
                continue

            # Check if this person is a trainee
            if normalized_name not in trainees:
                continue

            # Map shift to display name and hours
            if shift_name in RESIDENCY_SHIFT_MAPPINGS:
                display_name, hours = RESIDENCY_SHIFT_MAPPINGS[shift_name]
                results.append((date, display_name, normalized_name, hours))

    return results


def parse_contrast_schedule(html_path, trainees, fellows, year=2026):
    """
    Parse Contrast Coverage Schedule HTML (shifts as rows, dates as columns)
    Returns list of (date, shift_display, person, hours) tuples

    fellows parameter is used to filter Premier Fellow shifts (only fellows get paid)
    """
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')

    results = []

    # Build abbreviation map from the legend
    abbrev_map = build_abbreviation_map(soup)

    # Find the main schedule table
    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        if len(rows) < 10:
            continue

        # Find the header row with dates
        date_columns = []
        header_row_idx = None

        for idx, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            cell_texts = [c.get_text(strip=True) for c in cells]

            # Check if this row contains date headers (e.g., "Wed4-1")
            if any(re.search(r'\d+-\d+', t) or re.search(r'\d+/\d+', t) for t in cell_texts[1:5]):
                header_row_idx = idx
                date_columns = cell_texts
                break

        if not date_columns or header_row_idx is None:
            continue

        # Parse dates from header
        dates = [parse_date_header(d, year) for d in date_columns]

        # Parse data rows (shifts as first column)
        for row in rows[header_row_idx + 1:]:
            cells = row.find_all(['td', 'th'])
            if not cells:
                continue

            cell_texts = [c.get_text(strip=True) for c in cells]
            if not cell_texts:
                continue

            # First cell should be the shift name
            shift_name = cell_texts[0]
            if not shift_name:
                continue

            # Normalize shift name (replace non-breaking spaces)
            shift_name = shift_name.replace('\xa0', ' ')

            # Skip non-shift rows
            if shift_name.startswith(('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', '20', '7:30')):
                continue
            if 'Doximity' in shift_name or 'GPT' in shift_name:
                continue

            # Map shift to display name and hours - try exact match first
            shift_info = CONTRAST_SHIFT_MAPPINGS.get(shift_name)

            # Try matching with key as prefix of shift_name (e.g., key="32nd Late" matches "32nd Late (5p-9p)")
            if not shift_info:
                for key in CONTRAST_SHIFT_MAPPINGS:
                    if shift_name.startswith(key):
                        shift_info = CONTRAST_SHIFT_MAPPINGS[key]
                        break

            if not shift_info:
                continue

            display_name, hours = shift_info

            # Determine day-of-week restrictions based on shift name
            # "Sat" shifts only count on Saturdays
            # "Sunday" shifts only on Sundays
            # "41st Late" only on weekdays (Mon-Fri)
            saturday_only = 'Sat' in shift_name and 'Sat' in shift_name.split('(')[0]
            sunday_only = 'Sunday' in shift_name
            weekday_only = '41st Late' in shift_name

            # Parse each date column
            for col_idx, abbrev in enumerate(cell_texts[1:], start=1):
                if col_idx >= len(dates) or not dates[col_idx]:
                    continue

                day_of_week = dates[col_idx].weekday()  # 0=Mon, 5=Sat, 6=Sun

                # Filter by day of week if needed
                if saturday_only and day_of_week != 5:  # 5 = Saturday
                    continue
                if sunday_only and day_of_week != 6:  # 6 = Sunday
                    continue
                if weekday_only and day_of_week >= 5:  # Skip Sat/Sun
                    continue

                abbrev = abbrev.strip()
                if not abbrev or abbrev == '-':
                    continue

                # Look up full name from abbreviation
                full_name = abbrev_map.get(abbrev)
                if not full_name:
                    # Try to find partial match
                    for key, value in abbrev_map.items():
                        if key.startswith(abbrev) or abbrev.startswith(key):
                            full_name = value
                            break

                if not full_name:
                    continue

                # Check if this person is a trainee
                if full_name not in trainees:
                    continue

                # Premier Fellow shifts only count for fellows (not residents)
                if 'Premier' in shift_name and full_name not in fellows:
                    continue

                results.append((dates[col_idx], display_name, full_name, hours))

        break  # Only process first matching table

    return results


def generate_summary(entries, month_name):
    """Generate text summary from entries"""
    # Group by person
    by_person = defaultdict(list)
    for date, shift, person, hours in entries:
        by_person[person].append((date, shift, hours))

    # Sort people alphabetically
    lines = []
    lines.append(f"{month_name} Moonlighting Summary")
    lines.append("=" * 50)
    lines.append("")

    for person in sorted(by_person.keys()):
        shifts = by_person[person]
        # Sort by date
        shifts.sort(key=lambda x: x[0])

        lines.append(person)
        total_hours = 0
        for date, shift, hours in shifts:
            date_str = date.strftime("%m-%d-%Y")
            lines.append(f"  {date_str} {shift:<25} {hours}")
            total_hours += hours
        lines.append(f"  {'Total:':<36} {total_hours}")
        lines.append("")

    return "\n".join(lines)


def main():
    if len(sys.argv) < 2:
        print("Usage: python amion_parser.py <folder_path>")
        print("Example: python amion_parser.py April")
        sys.exit(1)

    folder_path = sys.argv[1]
    base_path = os.path.dirname(os.path.abspath(__file__))

    # Handle relative or absolute path
    if not os.path.isabs(folder_path):
        folder_path = os.path.join(base_path, folder_path)

    if not os.path.isdir(folder_path):
        print(f"Error: Folder not found: {folder_path}")
        sys.exit(1)

    # Extract month name from folder
    month_name = os.path.basename(folder_path)

    # Find HTML files
    html_files = [f for f in os.listdir(folder_path) if f.endswith('.html')]

    residency_html = None
    contrast_html = None

    for f in html_files:
        if 'Residency' in f:
            residency_html = os.path.join(folder_path, f)
        elif 'Contrast' in f:
            contrast_html = os.path.join(folder_path, f)

    if not residency_html and not contrast_html:
        print("Error: No Amion HTML files found in folder")
        sys.exit(1)

    # Load trainees
    print("Loading trainees...")
    trainees, residents, fellows = load_trainees(base_path)
    print(f"  Found {len(residents)} residents and {len(fellows)} fellows")

    # Determine year from HTML filename
    year = 2026  # Default
    for f in html_files:
        match = re.search(r'(\d{4})', f)
        if match:
            year = int(match.group(1))
            break

    all_entries = []

    # Parse Residency Schedule
    if residency_html:
        print(f"Parsing Residency Schedule...")
        entries = parse_residency_schedule(residency_html, trainees, year)
        print(f"  Found {len(entries)} trainee shifts")
        all_entries.extend(entries)

    # Parse Contrast Coverage Schedule
    if contrast_html:
        print(f"Parsing Contrast Coverage Schedule...")
        entries = parse_contrast_schedule(contrast_html, trainees, fellows, year)
        print(f"  Found {len(entries)} trainee shifts")
        all_entries.extend(entries)

    # Remove duplicates (same date, shift, person)
    seen = set()
    unique_entries = []
    for entry in all_entries:
        key = (entry[0], entry[1], entry[2])
        if key not in seen:
            seen.add(key)
            unique_entries.append(entry)

    print(f"Total unique shifts: {len(unique_entries)}")

    # Generate summary
    summary = generate_summary(unique_entries, month_name)

    # Write output
    output_path = os.path.join(folder_path, f"{month_name}_Trainee_Summary.txt")
    with open(output_path, 'w') as f:
        f.write(summary)

    print(f"Summary written to: {output_path}")

    # Also print to console
    print("\n" + "=" * 50)
    print(summary)


if __name__ == "__main__":
    main()
