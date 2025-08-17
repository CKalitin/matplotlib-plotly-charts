import csv
from datetime import datetime, timedelta
from collections import defaultdict

# Helper to parse duration string (e.g. '1:02:47:52' or '0:00:15:22')
def parse_duration(duration):
    duration = duration.strip()
    if duration == '0' or duration == '' or duration == '0:00:00:00':
        return timedelta(0)
    parts = duration.split(':')
    try:
        if len(parts) == 4:
            days, hours, minutes, seconds = map(int, parts)
            return timedelta(days=days, hours=hours, minutes=minutes, seconds=seconds)
        elif len(parts) == 3:
            hours, minutes, seconds = map(int, parts)
            return timedelta(hours=hours, minutes=minutes, seconds=seconds)
        elif len(parts) == 2:
            minutes, seconds = map(int, parts)
            return timedelta(minutes=minutes, seconds=seconds)
        else:
            return timedelta(0)
    except ValueError:
        # If any part is not an integer (e.g. '13?'), treat as zero duration
        return timedelta(0)

# Helper to parse date string (e.g. '1961 Jan 31 1654:51' or '1961 Apr 12 0607:00')
def parse_date(date_str):
    date_str = date_str.strip()
    for fmt in ["%Y %b %d %H%M:%S", "%Y %b %d %H%M", "%Y %b %d %H%M:%S", "%Y %b %d %H%M", "%Y %b %d", "%Y-%m-%d"]:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    # If date is just year-month-day
    if '-' in date_str:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            pass
    return None



# Read missions.tsv and process
missions = defaultdict(list)
with open("gantt_orbital_crew_missions/missions.tsv", encoding="utf-8") as f:
    # Find the header line (first line starting with # but not '# Updated')
    lines = f.readlines()
    header = None
    data_lines = []
    for line in lines:
        if line.startswith('#HSFID'):
            header = line.lstrip('#').strip().split('\t')
        elif not line.startswith('#') and line.strip():
            data_lines.append(line.strip())
    if not header:
        raise ValueError("Header line not found in TSV file.")
    reader = csv.DictReader(data_lines, fieldnames=header, delimiter='\t')
    for row in reader:
        duration = parse_duration(row['Dur'])
        if duration < timedelta(hours=1):
            continue  # Cull missions less than an hour
        project = row['Project'].strip()
        ship = row['Ship'].strip()
        # Translate project names
        if project == 'SZ':
            project = 'Shenzhou'
        elif project == 'STS':
            project = 'Space Shuttle'
        elif project == 'Soyuz/MirCorp':
            project = 'Soyuz'
        # Also translate ship names for clarity
        if ship.startswith('SZ'):
            ship = ship.replace('SZ', 'Shenzhou', 1)
        if ship.startswith('STS'):
            ship = ship.replace('STS', 'Space Shuttle', 1)
        ldate = parse_date(row['LDate'])
        edate_raw = row['EDate'].strip()
        if edate_raw.endswith('?'):
            edate_raw = edate_raw[:-1].strip()
        edate = parse_date(edate_raw)
        # Format dates as YYYY-MM-DD
        ldate_str = ldate.strftime('%Y-%m-%d') if ldate else '-'
        edate_str = edate.strftime('%Y-%m-%d') if edate else '-'
        missions[project].append((ship, ldate_str, edate_str))

# Sort projects alphabetically
sorted_projects = sorted(missions.keys())

# Write to missions_data.txt
with open("gantt_orbital_crew_missions/missions_data.txt", "w", encoding="utf-8") as out:
    for project in sorted_projects:
        out.write(f"{project}\n")
        for ship, ldate, edate in missions[project]:
            out.write(f"{ship}, {ldate}, {edate}\n")
        out.write("\n")

print("Translation complete. Output written to gantt_orbital_crew_missions/missions_data.txt.")
