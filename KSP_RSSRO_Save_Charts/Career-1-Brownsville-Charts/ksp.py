from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
import warnings
try:
    from tqdm import tqdm
except ImportError:
    # Fallback if tqdm is not available
    def tqdm(iterable, **kwargs):
        return iterable

# Configure matplotlib to prevent windows from opening
plt.ioff()  # Turn OFF interactive mode to prevent window popups
plt.switch_backend('Agg')  # Use non-interactive backend
warnings.filterwarnings('ignore', message='.*figure.*opened.*', category=RuntimeWarning)

# Global chart counter
chart_counter = {'current': 0, 'total': 0}

def update_progress(description="Creating chart"):
    """Update progress counter and display progress"""
    chart_counter['current'] += 1
    current = min(chart_counter['current'], chart_counter['total'])  # Prevent overflow
    progress = f"[{current}/{chart_counter['total']}] {description}"
    # Use carriage return to overwrite the same line
    print(f"\r{progress:<100}", end='', flush=True)

def set_total_charts(total):
    """Set the total number of charts to be created"""
    chart_counter['total'] = total
    chart_counter['current'] = 0

# ==================== HELPER FUNCTIONS ====================

def get_outcome_order():
    """Define the proper order for mission outcomes (best to worst)"""
    return ['Great Success', 'Success', 'Partial Success', 'In Progress', 'Partial Failure', 'Failure', 'Catastrophic Failure']

def sort_outcome_columns(data, outcome_order=None):
    """Sort dataframe columns by outcome order"""
    if outcome_order is None:
        outcome_order = get_outcome_order()
    
    # Get columns that exist in the data, in the proper order
    columns_to_plot = [col for col in outcome_order if col in data.columns]
    
    # Add any remaining columns that weren't in our order (shouldn't happen, but safety)
    remaining_columns = [col for col in data.columns if col not in columns_to_plot]
    columns_to_plot.extend(remaining_columns)
    
    return data[columns_to_plot]

def extract_program_name(mission_payload):
    """Extract program name from Mission/Payload field"""
    if pd.isna(mission_payload) or not str(mission_payload).strip():
        return 'Other'
    
    mission_str = str(mission_payload).strip()
    
    # Special case: First navigational satellite (1958 Oct 10, Goliath 1A)
    # This should be classified as Selenium 1
    if mission_str == '' or pd.isna(mission_payload):
        return 'Other'  # Will be handled by fill_mission_payload_column for the specific case
    
    if mission_str.startswith('ATP Flight'):
        return 'ATP Flight'
    elif mission_str.startswith('ATP Spaceflight'):
        return 'ATP Spaceflight'
    elif mission_str.startswith('Humboldt Orbital'):
        return 'Humboldt Orbital'
    elif mission_str.startswith('Humboldt Suborbital'):
        return 'Humboldt Suborbital'
    elif mission_str.startswith('Ricochet'):
        return 'Ricochet'
    elif mission_str.startswith('Heinlein'):
        return 'Heinlein'
    elif mission_str.startswith('Colussus'):
        return 'Colussus'
    elif mission_str.startswith('Goliath'):
        return 'Goliath'
    elif mission_str.startswith('Polaris'):
        return 'Polaris'
    elif mission_str.startswith('Selenium'):
        return 'Selenium'
    elif mission_str.startswith('Amundsen'):
        return 'Amundsen'
    elif mission_str.startswith('Belgica') or mission_str.startswith('CRS'):
        return 'Belgica'
    elif mission_str.startswith('Cassiope'):
        return 'Cassiope'
    elif mission_str.startswith('Perplex'):
        return 'Perplex'
    else:
        return mission_str.split()[0] if mission_str.split() else 'Other'

def calculate_flight_time_hours(row):
    """Calculate flight time in hours based on mission type and notes"""
    # If no crew, return 0
    if pd.isna(row.get("Crew")) or not str(row.get("Crew")).strip():
        return 0
    
    launch_vehicle = str(row.get("Launch Vehicle", "")).strip()
    notes = str(row.get("Notes", "")).strip().lower()
    mission_payload = str(row.get("Mission / Payload", "")).strip()
    
    # ATP flights - assume 1 hour each
    if launch_vehicle.startswith("ATP"):
        return 1.0
    
    # Belgica missions - parse from notes
    if "belgica" in mission_payload.lower():
        # Look for explicit flight time in notes
        if "3:08" in notes:  # Belgica 1
            return 3.13  # 3 hours 8 minutes
        elif "2 orbits" in notes:  # Belgica 2
            return 3.0  # Assume ~3 hours for 2 orbits
        elif "24hrs" in notes or "24 hrs" in notes:  # Belgica 3
            return 24.0
        elif "7 day" in notes or "7 days" in notes:  # Belgica 6
            return 168.0  # 7 days = 168 hours
        else:
            # Default Belgica missions (single orbit LEO)
            return 1.5
    
    # Polaris missions - parse from notes
    if "polaris" in mission_payload.lower():
        # Look for day/hour patterns in notes
        if "2 days 18 hours" in notes:  # Polaris 1
            return 66.0  # 2*24 + 18 = 66 hours
        elif "6 days 20 hours" in notes:  # Polaris 3
            return 164.0  # 6*24 + 20 = 164 hours
        elif "14 day" in notes:  # Polaris 5
            return 336.0  # 14*24 = 336 hours
        elif "first eva" in notes or "rendezvous" in notes:  # Polaris 4
            return 24.0  # Assume ~1 day for LEO EVA mission
        else:
            # Other Polaris missions (lunar flyby/orbit)
            return 168.0  # Default 7 days for lunar missions
    
    # Other crewed missions - estimate based on mission type
    if "lunar" in mission_payload.lower():
        return 168.0  # 7 days for lunar missions
    elif "orbital" in mission_payload.lower():
        return 24.0  # 1 day for orbital missions
    else:
        return 3.0  # Default for other crewed missions

def create_pie_chart(data, title, filename, output_dir, colors=None, figsize=(8, 8)):
    """Generic pie chart creation function"""
    if len(data) == 0:
        return
        
    update_progress(f"Creating pie chart: {title}")
    plt.figure(figsize=figsize)
    
    try:
        if colors:
            if callable(colors):
                chart_colors = [colors(item) for item in data.index]
            else:
                chart_colors = colors
        else:
            chart_colors = None
            
        plt.pie(data, labels=data.index, autopct='%1.1f%%', startangle=140, colors=chart_colors)
        plt.title(title)
        plt.savefig(output_dir / filename)
    finally:
        plt.close()

def format_year_axis(years, skip_every=2):
    """Format year axis for time series charts"""
    if len(years) > 10:
        year_positions = list(range(0, len(years), skip_every))
        year_labels = [str(int(years[i])) for i in year_positions]
        plt.xticks(year_positions, year_labels, rotation=0)
    else:
        plt.xticks(rotation=0)

def format_month_axis(dates):
    """Format month axis for IRL time series charts"""
    month_positions = []
    month_labels = []
    
    for i, date in enumerate(dates):
        if date.day == 1:
            month_positions.append(i)
            month_labels.append(date.strftime('%Y-%m'))
    
    plt.xticks(month_positions, month_labels, rotation=45)

def create_time_series_chart(data, title, filename, output_dir, colors=None, figsize=(12, 6), 
                           width=0.9, stacked=False, time_type='ksp'):
    """Generic time series chart creation function"""
    if len(data) == 0:
        return
        
    update_progress(f"Creating time series: {title}")
    plt.figure(figsize=figsize)
    
    try:
        if stacked:
            ax = data.plot(kind='bar', stacked=True, figsize=figsize, width=width, color=colors)
        else:
            ax = data.plot(kind='bar', figsize=figsize, width=width, color=colors)
        
        plt.ylabel("Number of Missions")
        plt.title(title)
        
        # Format time axis
        if time_type == 'ksp':
            format_year_axis(data.index)
        elif time_type == 'irl':
            format_month_axis(data.index)
        
        plt.tight_layout()
        
        # Handle legend for stacked charts with many series
        if stacked and len(data.columns) > 5:
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.savefig(output_dir / filename, bbox_inches='tight')
        else:
            plt.savefig(output_dir / filename)
    finally:
        plt.close()

def prepare_irl_date_data(df):
    """Prepare IRL date data for time series analysis"""
    df["IRL Date"] = pd.to_datetime(df["IRL Date"], format='mixed')
    min_date = df["IRL Date"].min().date()
    max_date = df["IRL Date"].max().date()
    date_range = pd.date_range(start=min_date, end=max_date, freq='D')
    df["IRL Day"] = df["IRL Date"].dt.date
    return df, date_range

# ==================== MAIN FUNCTIONS ====================

def fill_mission_payload_column(df):
    """Fill in missing Mission/Payload values based on vehicle patterns"""
    df = df.copy()
    
    # Initialize Mission/Payload column if it doesn't exist or has NaN
    if 'Mission / Payload' not in df.columns:
        df['Mission / Payload'] = ''
    df['Mission / Payload'] = df['Mission / Payload'].fillna('')
    
    # Track counters for different program types
    humboldt_orbital_counter = 1
    humboldt_suborbital_counter = 1
    atp_flight_counter = 1
    atp_spaceflight_counter = 1
    
    for idx, row in df.iterrows():
        if pd.isna(row['Mission / Payload']) or row['Mission / Payload'].strip() == '':
            vehicle = row['Launch Vehicle']
            contract_mission = str(row.get('Contract / Mission', '')).strip()
            
            # Special case: First navigational satellite (1958 Oct 10, Goliath 1A)
            if (pd.notna(vehicle) and vehicle == 'Goliath 1A' and 
                'First Navigational Satellite' in contract_mission):
                df.at[idx, 'Mission / Payload'] = "Selenium 1"
            
            # Humboldt-4X -> "Humboldt Orbital Program" with incrementing numbers
            elif pd.notna(vehicle) and 'Humboldt-4' in vehicle:
                df.at[idx, 'Mission / Payload'] = f"Humboldt Orbital {humboldt_orbital_counter}"
                humboldt_orbital_counter += 1
            
            # Other Humboldt vehicles -> "Humboldt Suborbital Program"
            elif pd.notna(vehicle) and vehicle.startswith('Humboldt') and 'Humboldt-4' not in vehicle:
                df.at[idx, 'Mission / Payload'] = f"Humboldt Suborbital {humboldt_suborbital_counter}"
                humboldt_suborbital_counter += 1
            
            # ATP flights
            elif pd.notna(vehicle) and vehicle.startswith('ATP-'):
                # Extract series number (first digit after ATP-)
                match = re.search(r'ATP-(\d)', vehicle)
                if match:
                    series = match.group(1)
                    if series in ['2', '4']:  # ATP 2XX or 4XX -> ATP Spaceflight
                        df.at[idx, 'Mission / Payload'] = f"ATP Spaceflight {atp_spaceflight_counter}"
                        atp_spaceflight_counter += 1
                    else:  # Other ATP -> ATP Flight
                        df.at[idx, 'Mission / Payload'] = f"ATP Flight {atp_flight_counter}"
                        atp_flight_counter += 1
    
    return df

def create_output_directories():
    """Create organized directory structure for charts"""
    base_dir = Path("mission_charts")
    base_dir.mkdir(exist_ok=True)
    
    # Create subdirectories
    subdirs = [
        "pie_charts",
        "ksp_time_series", 
        "irl_time_series",
        "custom_charts",
        "program_analysis"
    ]
    
    for subdir in subdirs:
        (base_dir / subdir).mkdir(exist_ok=True)
    
    return base_dir

def get_color_schemes():
    """Define all color schemes used across charts"""
    
    # Mission outcome colors
    outcome_colors = {
        'Great Success': "#008100",  # Dark Green
        'Success': "#2AA72A",  # Forest Green  
        'Partial Success': "#6AEB6A",  # Light Green
        'Partial Failure': "#FF7300",  # Orange
        'Failure': "#E61010",  # Red Orange
        'Catastrophic Failure': '#8B0000',  # Dark Red
        'In Progress': "#FFFF00"  # Yellow
    }
    
    # Launch vehicle series colors
    series_colors = {
        'ATP-1XX': '#FFD700',  # Gold (Plane)
        'ATP-2XX': "#FFB13C",  # Orange (Plane)
        'ATP-3XX': "#FF7809",  # Dark Orange (Plane)
        'ATP-4XX': "#F8561B",  # Coral (Plane)
        'Goliath 1': '#90EE90',  # Light Green
        'Goliath 2': '#32CD32',  # Lime Green
        'Goliath 3': '#228B22',  # Forest Green
        'Goliath 4': '#006400',  # Dark Green
        'Colussus 1': "#F83333",  # Light Red
        'Colussus 2': '#8B0000',  # Dark Red
        'Humboldt': '#4682B4',  # Steel Blue
        'Unknown': '#708090'   # Slate Gray
    }
    
    # Program colors based on mission nature
    program_colors = {
        'ATP Flight': '#FFD700',        
        'ATP Spaceflight': '#FF8C00',   
        'Humboldt Suborbital': '#CA0707',
        'Humboldt Orbital': "#FF4A09",  
        'Heinlein': "#B9D6EC",          
        'Polaris': "#28C0FC",           
        'Ricochet': "#0022B9",          
        'Cassiope': "#05AD05",          
        'Belgica': "#0F4FFF",           
        'Selenium': "#63E463",          
        'Perplex': "#AC39FF",           
        'Amundsen': '#4B0082',          
        'Other': "#333435"              
    }
    
    # Success/failure colors
    success_colors = ['#2AA72A', '#FF6B35']  # Green for Successful, Red for Unsuccessful
    
    # Crewed vs uncrewed colors
    crewed_colors = ['#87CEEB', '#FF6B35']  # Sky Blue for Uncrewed, Orange-Red for Crewed
    
    # Vehicle type colors
    vehicle_colors = ['#FFD700', '#8B4513']  # Gold for Plane, Saddle Brown for Rocket
    
    return outcome_colors, series_colors, success_colors, crewed_colors, vehicle_colors, program_colors

def create_pie_charts(df, output_dir, outcome_colors, series_colors):
    """Create all pie charts"""
    pie_dir = output_dir / "pie_charts"
    
    # Pie chart: Mission Outcomes
    outcome_counts = df["Outcome"].value_counts()
    create_pie_chart(outcome_counts, "Mission Outcomes Distribution", "pie_mission_outcomes.png", 
                    pie_dir, lambda x: outcome_colors.get(x, '#888888'))

    # Pie chart: Crewed vs Uncrewed
    crew_counts = df["Crewed"].value_counts().rename({True: "Crewed", False: "Uncrewed"})
    create_pie_chart(crew_counts, "Crewed vs Uncrewed Missions", "pie_crewed_vs_uncrewed.png",
                    pie_dir, ['#87CEEB', '#FF6B35'])

    # Pie chart: Plane vs Rocket
    vehicle_counts = df["Vehicle Type"].value_counts()
    create_pie_chart(vehicle_counts, "Plane vs Rocket Missions", "pie_plane_vs_rocket.png",
                    pie_dir, ['#FFD700', '#8B4513'])

    # Pie chart: Launch Vehicle Series
    series_counts = df["Launch Vehicle Series"].value_counts()
    create_pie_chart(series_counts, "Launch Vehicle Series Distribution", "pie_launch_vehicle_series.png",
                    pie_dir, lambda x: series_colors.get(x, '#888888'), figsize=(10, 10))

def create_ksp_time_series(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors):
    """Create KSP date time series charts"""
    ksp_dir = output_dir / "ksp_time_series"
    
    # Mission outcomes over time (stacked bar chart)
    outcome_by_year = df.groupby(["Year", "Outcome"]).size().unstack(fill_value=0)
    
    # Sort columns from best to worst outcome
    outcome_by_year = sort_outcome_columns(outcome_by_year)
    
    colors = [outcome_colors.get(col, '#888888') for col in outcome_by_year.columns]
    create_time_series_chart(outcome_by_year, "Mission Outcomes Over Time (KSP Date)", 
                           "ksp_mission_outcomes_vs_time.png", ksp_dir, colors, stacked=True)

    # Total flights per year
    flight_counts = df.groupby("Year").size()
    create_time_series_chart(flight_counts, "Total Flights Over Time (KSP Date)",
                           "ksp_total_flights_vs_time.png", ksp_dir, '#4682B4')

    # Crewed vs Uncrewed over time
    crewed_by_year = df.groupby(["Year", "Crewed"]).size().unstack(fill_value=0)
    crewed_by_year.rename(columns={True: "Crewed", False: "Uncrewed"}, inplace=True)
    create_time_series_chart(crewed_by_year, "Crewed vs Uncrewed Missions Over Time (KSP Date)",
                           "ksp_crewed_vs_uncrewed_vs_time.png", ksp_dir, crewed_colors, stacked=True)

    # Plane vs Rocket over time
    vehicle_by_year = df.groupby(["Year", "Vehicle Type"]).size().unstack(fill_value=0)
    create_time_series_chart(vehicle_by_year, "Plane vs Rocket Missions Over Time (KSP Date)",
                           "ksp_plane_vs_rocket_vs_time.png", ksp_dir, vehicle_colors, stacked=True)

    # Launch Vehicle Series over time
    series_by_year = df.groupby(["Year", "Launch Vehicle Series"]).size().unstack(fill_value=0)
    colors = [series_colors.get(col, '#888888') for col in series_by_year.columns]
    create_time_series_chart(series_by_year, "Launch Vehicle Series Over Time (KSP Date)",
                           "ksp_launch_vehicle_series_vs_time.png", ksp_dir, colors, 
                           figsize=(14, 8), stacked=True)

def create_irl_time_series(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors):
    """Create IRL date time series charts"""
    irl_dir = output_dir / "irl_time_series"
    
    # Prepare IRL date data
    df, date_range = prepare_irl_date_data(df)
    
    # Total flights per IRL day
    irl_flight_counts = df.groupby("IRL Day").size()
    irl_flight_counts = irl_flight_counts.reindex(date_range.date, fill_value=0)
    create_time_series_chart(irl_flight_counts, "Total Flights Over Time (IRL Date - Daily)",
                           "irl_total_flights_vs_time.png", irl_dir, '#4682B4', 
                           figsize=(16, 6), width=1.0, time_type='irl')
    
    # Mission outcomes over IRL time
    outcome_by_irl_day = df.groupby(["IRL Day", "Outcome"]).size().unstack(fill_value=0)
    outcome_by_irl_day = outcome_by_irl_day.reindex(date_range.date, fill_value=0)
    
    outcome_by_irl_day = sort_outcome_columns(outcome_by_irl_day)
    
    colors = [outcome_colors.get(col, '#888888') for col in outcome_by_irl_day.columns]
    create_time_series_chart(outcome_by_irl_day, "Mission Outcomes Over Time (IRL Date - Daily)",
                           "irl_mission_outcomes_vs_time.png", irl_dir, colors,
                           figsize=(16, 6), width=1.0, stacked=True, time_type='irl')
    
    # Crewed vs Uncrewed over IRL time
    crewed_by_irl_day = df.groupby(["IRL Day", "Crewed"]).size().unstack(fill_value=0)
    crewed_by_irl_day = crewed_by_irl_day.reindex(date_range.date, fill_value=0)
    crewed_by_irl_day.rename(columns={True: "Crewed", False: "Uncrewed"}, inplace=True)
    create_time_series_chart(crewed_by_irl_day, "Crewed vs Uncrewed Missions Over Time (IRL Date - Daily)",
                           "irl_crewed_vs_uncrewed_vs_time.png", irl_dir, crewed_colors,
                           figsize=(16, 6), width=1.0, stacked=True, time_type='irl')
    
    # Plane vs Rocket over IRL time
    vehicle_by_irl_day = df.groupby(["IRL Day", "Vehicle Type"]).size().unstack(fill_value=0)
    vehicle_by_irl_day = vehicle_by_irl_day.reindex(date_range.date, fill_value=0)
    create_time_series_chart(vehicle_by_irl_day, "Plane vs Rocket Missions Over Time (IRL Date - Daily)",
                           "irl_plane_vs_rocket_vs_time.png", irl_dir, vehicle_colors,
                           figsize=(16, 6), width=1.0, stacked=True, time_type='irl')
    
    # Launch Vehicle Series over IRL time
    series_by_irl_day = df.groupby(["IRL Day", "Launch Vehicle Series"]).size().unstack(fill_value=0)
    series_by_irl_day = series_by_irl_day.reindex(date_range.date, fill_value=0)
    colors = [series_colors.get(col, '#888888') for col in series_by_irl_day.columns]
    create_time_series_chart(series_by_irl_day, "Launch Vehicle Series Over Time (IRL Date - Daily)",
                           "irl_launch_vehicle_series_vs_time.png", irl_dir, colors,
                           figsize=(16, 8), width=1.0, stacked=True, time_type='irl')

def create_custom_charts(df, output_dir, series_colors, success_colors, outcome_colors, program_colors):
    """Create custom analysis charts"""
    custom_dir = output_dir / "custom_charts"
    
    # Extract program names from Mission/Payload
    df['Program'] = df['Mission / Payload'].apply(extract_program_name)
    
    # Success/Failure by Launch Vehicle Series
    df["Success Category"] = "Unsuccessful"
    success_outcomes = ['Great Success', 'Success', 'Partial Success']
    df.loc[df["Outcome"].isin(success_outcomes), "Success Category"] = "Successful"
    
    success_by_series = df.groupby(["Launch Vehicle Series", "Success Category"]).size().unstack(fill_value=0)
    
    update_progress("Creating success/failure by vehicle series chart")
    try:
        ax = success_by_series.plot(kind='bar', stacked=True, figsize=(14, 8), width=0.8, color=success_colors)
        plt.ylabel("Number of Missions")
        plt.title("Successful vs Unsuccessful Flights by Launch Vehicle Series")
        plt.xlabel("Launch Vehicle Series")
        plt.xticks(rotation=45, ha='right')
        plt.legend(title="Mission Category")
        plt.tight_layout()
        plt.savefig(custom_dir / "vehicle_series_flights_by_success.png", bbox_inches='tight')
    finally:
        plt.close()

    # Chart: Outcome by Vehicle Series (NEW)
    outcome_by_series = df.groupby(["Launch Vehicle Series", "Outcome"]).size().unstack(fill_value=0)
    outcome_by_series = sort_outcome_columns(outcome_by_series)
    
    update_progress("Creating detailed outcomes by vehicle series chart")
    try:
        # Get colors for the columns that exist in the data
        series_outcome_colors = [outcome_colors.get(col, '#888888') for col in outcome_by_series.columns]
        
        ax = outcome_by_series.plot(kind='bar', stacked=True, figsize=(14, 10), width=0.8, color=series_outcome_colors)
        plt.ylabel("Number of Missions")
        plt.title("Detailed Mission Outcomes by Launch Vehicle Series")
        plt.xlabel("Launch Vehicle Series")
        plt.xticks(rotation=45, ha='right')
        plt.legend(title="Mission Outcome", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(custom_dir / "vehicle_series_flights_by_outcome.png", bbox_inches='tight')
    finally:
        plt.close()

    # Chart: Outcome by Program (NEW)
    outcome_by_program = df.groupby(["Program", "Outcome"]).size().unstack(fill_value=0)
    outcome_by_program = sort_outcome_columns(outcome_by_program)
    
    update_progress("Creating detailed outcomes by program chart")
    try:
        # Get colors for the columns that exist in the data
        program_outcome_colors = [outcome_colors.get(col, '#888888') for col in outcome_by_program.columns]
        
        ax = outcome_by_program.plot(kind='bar', stacked=True, figsize=(14, 10), width=0.8, color=program_outcome_colors)
        plt.ylabel("Number of Missions")
        plt.title("Detailed Mission Outcomes by Program")
        plt.xlabel("Program")
        plt.xticks(rotation=45, ha='right')
        plt.legend(title="Mission Outcome", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(custom_dir / "program_flights_by_outcome.png", bbox_inches='tight')
    finally:
        plt.close()

    # Chart: Program by Vehicle Series (NEW)
    program_by_series = df.groupby(["Program", "Launch Vehicle Series"]).size().unstack(fill_value=0)
    
    update_progress("Creating program flights by vehicle series chart")
    try:
        # Get colors for the columns that exist in the data
        program_series_colors = [series_colors.get(col, '#888888') for col in program_by_series.columns]
        
        ax = program_by_series.plot(kind='bar', stacked=True, figsize=(14, 10), width=0.8, color=program_series_colors)
        plt.ylabel("Number of Missions")
        plt.title("Launch Vehicle Series by Program")
        plt.xlabel("Program")
        plt.xticks(rotation=45, ha='right')
        plt.legend(title="Launch Vehicle Series", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(custom_dir / "program_flights_by_vehicle_series.png", bbox_inches='tight')
    finally:
        plt.close()

    # Chart: Vehicle Series by Program (NEW)
    series_by_program = df.groupby(["Launch Vehicle Series", "Program"]).size().unstack(fill_value=0)
    
    update_progress("Creating vehicle series flights by program chart")
    try:
        # Get colors for the columns that exist in the data
        series_program_colors = [program_colors.get(col, '#888888') for col in series_by_program.columns]
        
        ax = series_by_program.plot(kind='bar', stacked=True, figsize=(14, 10), width=0.8, color=series_program_colors)
        plt.ylabel("Number of Missions")
        plt.title("Programs by Launch Vehicle Series")
        plt.xlabel("Launch Vehicle Series")
        plt.xticks(rotation=45, ha='right')
        plt.legend(title="Program", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        plt.savefig(custom_dir / "vehicle_series_flights_by_program.png", bbox_inches='tight')
    finally:
        plt.close()

    # Pilots by Launch Vehicle Series AND Success/Failure
    pilot_vehicle_data = []
    
    for _, row in df.iterrows():
        if pd.notna(row["Crew"]) and row["Crew"].strip():
            crew_members = [name.strip() for name in row["Crew"].split(',')]
            for pilot in crew_members:
                if pilot:
                    pilot_vehicle_data.append({
                        'Pilot': pilot,
                        'Launch Vehicle Series': row["Launch Vehicle Series"],
                        'Success Category': row["Success Category"],
                        'KSP Date': row["KSP Date"]
                    })
    
    pilot_df = pd.DataFrame(pilot_vehicle_data)
    
    if not pilot_df.empty:
        # Find first flight date for each pilot and sort
        pilot_first_flight = pilot_df.groupby('Pilot')['KSP Date'].min().sort_values()
        pilot_order = pilot_first_flight.index.tolist()
        
        # Chart: Pilots by Vehicle Series
        pilot_by_series = pilot_df.groupby(["Pilot", "Launch Vehicle Series"]).size().unstack(fill_value=0)
        pilot_by_series = pilot_by_series.reindex(pilot_order)
        
        vehicle_colors = [series_colors.get(col, '#888888') for col in pilot_by_series.columns]
        
        update_progress("Creating pilots by vehicle series chart")
        try:
            ax = pilot_by_series.plot(kind='bar', stacked=True, figsize=(16, 10), width=0.8, color=vehicle_colors)
            plt.ylabel("Number of Flights")
            plt.title("Launch Vehicle Series Flown by Each Pilot (Sorted by First Flight Date)")
            plt.xlabel("Pilot")
            plt.xticks(rotation=45, ha='right')
            plt.legend(title="Launch Vehicle Series", bbox_to_anchor=(1.05, 1), loc='upper left', ncol=1)
            plt.tight_layout()
            plt.savefig(custom_dir / "pilot_flights_by_vehicle_series.png", bbox_inches='tight')
        finally:
            plt.close()
        
        # Chart: Pilots by Success/Failure
        pilot_by_success = pilot_df.groupby(["Pilot", "Success Category"]).size().unstack(fill_value=0)
        pilot_by_success = pilot_by_success.reindex(pilot_order)
        
        update_progress("Creating pilots by success/failure chart")
        try:
            ax = pilot_by_success.plot(kind='bar', stacked=True, figsize=(16, 8), width=0.8, color=success_colors)
            plt.ylabel("Number of Flights")
            plt.title("Successful vs Unsuccessful Flights by Pilot (Sorted by First Flight Date)")
            plt.xlabel("Pilot")
            plt.xticks(rotation=45, ha='right')
            plt.legend(title="Mission Category")
            plt.tight_layout()
            plt.savefig(custom_dir / "pilot_flights_by_success.png", bbox_inches='tight')
        finally:
            plt.close()
        
        # Chart: Pilots by Detailed Outcome Types
        # Create a mapping of pilots to their detailed outcomes
        pilot_outcome_data = []
        for _, row in df.iterrows():
            if pd.notna(row["Crew"]) and row["Crew"].strip() and pd.notna(row["Outcome"]):
                crew_members = [name.strip() for name in row["Crew"].split(',')]
                for pilot in crew_members:
                    if pilot:
                        pilot_outcome_data.append({
                            'Pilot': pilot,
                            'Outcome': row["Outcome"],
                            'KSP Date': row["KSP Date"]
                        })
        
        pilot_outcome_df = pd.DataFrame(pilot_outcome_data)
        
        if not pilot_outcome_df.empty:
            # Find first flight date for each pilot and sort
            pilot_first_flight_outcomes = pilot_outcome_df.groupby('Pilot')['KSP Date'].min().sort_values()
            pilot_order_outcomes = pilot_first_flight_outcomes.index.tolist()
            
            pilot_by_outcome = pilot_outcome_df.groupby(["Pilot", "Outcome"]).size().unstack(fill_value=0)
            pilot_by_outcome = pilot_by_outcome.reindex(pilot_order_outcomes)
            pilot_by_outcome = sort_outcome_columns(pilot_by_outcome)
            
            # Get colors for the columns that exist in the data using the main outcome_colors
            detail_colors = [outcome_colors.get(col, '#888888') for col in pilot_by_outcome.columns]
            
            update_progress("Creating pilots by detailed outcome chart")
            try:
                ax = pilot_by_outcome.plot(kind='bar', stacked=True, figsize=(16, 10), width=0.8, color=detail_colors)
                plt.ylabel("Number of Flights")
                plt.title("Detailed Mission Outcomes by Pilot (Sorted by First Flight Date)")
                plt.xlabel("Pilot")
                plt.xticks(rotation=45, ha='right')
                plt.legend(title="Mission Outcome", bbox_to_anchor=(1.05, 1), loc='upper left')
                plt.tight_layout()
                plt.savefig(custom_dir / "pilot_flights_by_outcome.png", bbox_inches='tight')
            finally:
                plt.close()

    # Chart: Pilot Flight Time Analysis (NEW)
    # Calculate flight times for all missions
    df['Flight_Time_Hours'] = df.apply(calculate_flight_time_hours, axis=1)
    
    # Create pilot flight time data
    pilot_time_data = []
    for _, row in df.iterrows():
        if pd.notna(row["Crew"]) and row["Crew"].strip() and row['Flight_Time_Hours'] > 0:
            crew_members = [name.strip() for name in row["Crew"].split(',')]
            for pilot in crew_members:
                if pilot:
                    pilot_time_data.append({
                        'Pilot': pilot,
                        'Flight_Time_Hours': row['Flight_Time_Hours'],
                        'KSP Date': row["KSP Date"],
                        'Mission': row.get('Mission / Payload', 'Unknown')
                    })
    
    pilot_time_df = pd.DataFrame(pilot_time_data)
    
    if not pilot_time_df.empty:
        # Calculate total flight time per pilot
        pilot_total_time = pilot_time_df.groupby('Pilot')['Flight_Time_Hours'].sum()
        
        # Find first flight date for each pilot for ordering (consistent with other charts)
        pilot_first_flight = pilot_time_df.groupby('Pilot')['KSP Date'].min().sort_values()
        pilot_order = pilot_first_flight.index.tolist()
        
        # Reorder pilot total time by first flight date
        pilot_total_time = pilot_total_time.reindex(pilot_order)
        
        update_progress("Creating pilot total flight time chart")
        try:
            ax = pilot_total_time.plot(kind='bar', figsize=(16, 8), width=0.8, color='#4682B4')
            plt.ylabel("Total Flight Time (Hours)")
            plt.title("Total Flight Time by Pilot (Sorted by First Flight Date)")
            plt.xlabel("Pilot")
            plt.xticks(rotation=45, ha='right')
            
            # Add value labels on top of bars
            for i, v in enumerate(pilot_total_time.values):
                if v >= 24:
                    label = f"{v:.0f}h\n({v/24:.1f}d)"
                else:
                    label = f"{v:.1f}h"
                plt.text(i, v + max(pilot_total_time.values) * 0.01, label, 
                        ha='center', va='bottom', fontsize=9)
            
            plt.tight_layout()
            plt.savefig(custom_dir / "pilot_flight_time_total.png", bbox_inches='tight')
        finally:
            plt.close()

        # Chart: Pilot Flight Time by Program (NEW)
        # Create pilot flight time by program breakdown
        pilot_time_program_data = []
        for _, row in df.iterrows():
            if pd.notna(row["Crew"]) and row["Crew"].strip() and row['Flight_Time_Hours'] > 0:
                crew_members = [name.strip() for name in row["Crew"].split(',')]
                for pilot in crew_members:
                    if pilot:
                        pilot_time_program_data.append({
                            'Pilot': pilot,
                            'Program': row['Program'],
                            'Flight_Time_Hours': row['Flight_Time_Hours'],
                            'KSP Date': row["KSP Date"]
                        })
        
        pilot_time_program_df = pd.DataFrame(pilot_time_program_data)
        
        if not pilot_time_program_df.empty:
            # Group by pilot and program, sum flight time
            pilot_time_by_program = pilot_time_program_df.groupby(['Pilot', 'Program'])['Flight_Time_Hours'].sum().unstack(fill_value=0)
            pilot_time_by_program = pilot_time_by_program.reindex(pilot_order)
            
            # Get program colors for the chart
            program_time_colors = [program_colors.get(col, '#888888') for col in pilot_time_by_program.columns]
            
            update_progress("Creating pilot flight time by program chart")
            try:
                ax = pilot_time_by_program.plot(kind='bar', stacked=True, figsize=(16, 10), width=0.8, color=program_time_colors)
                plt.ylabel("Flight Time (Hours)")
                plt.title("Flight Time by Pilot and Program (Sorted by First Flight Date)")
                plt.xlabel("Pilot")
                plt.xticks(rotation=45, ha='right')
                plt.legend(title="Program", bbox_to_anchor=(1.05, 1), loc='upper left')
                
                # Add value labels on top of bars (total flight time per pilot)
                pilot_totals = pilot_time_by_program.sum(axis=1)
                for i, total_time in enumerate(pilot_totals.values):
                    if total_time > 0:
                        if total_time >= 24:
                            label = f"{total_time:.0f}h\n({total_time/24:.1f}d)"
                        else:
                            label = f"{total_time:.1f}h"
                        plt.text(i, total_time + max(pilot_totals.values) * 0.01, label, 
                                ha='center', va='bottom', fontsize=9)
                
                plt.tight_layout()
                plt.savefig(custom_dir / "pilot_flight_time_by_program.png", bbox_inches='tight')
            finally:
                plt.close()

        # Chart: Pilot Flight Time by Vehicle Series (NEW)
        # Create pilot flight time by vehicle series breakdown
        pilot_time_vehicle_data = []
        for _, row in df.iterrows():
            if pd.notna(row["Crew"]) and row["Crew"].strip() and row['Flight_Time_Hours'] > 0:
                crew_members = [name.strip() for name in row["Crew"].split(',')]
                for pilot in crew_members:
                    if pilot:
                        pilot_time_vehicle_data.append({
                            'Pilot': pilot,
                            'Launch Vehicle Series': row['Launch Vehicle Series'],
                            'Flight_Time_Hours': row['Flight_Time_Hours'],
                            'KSP Date': row["KSP Date"]
                        })
        
        pilot_time_vehicle_df = pd.DataFrame(pilot_time_vehicle_data)
        
        if not pilot_time_vehicle_df.empty:
            # Group by pilot and vehicle series, sum flight time
            pilot_time_by_vehicle = pilot_time_vehicle_df.groupby(['Pilot', 'Launch Vehicle Series'])['Flight_Time_Hours'].sum().unstack(fill_value=0)
            pilot_time_by_vehicle = pilot_time_by_vehicle.reindex(pilot_order)
            
            # Get vehicle series colors for the chart
            vehicle_time_colors = [series_colors.get(col, '#888888') for col in pilot_time_by_vehicle.columns]
            
            update_progress("Creating pilot flight time by vehicle series chart")
            try:
                ax = pilot_time_by_vehicle.plot(kind='bar', stacked=True, figsize=(16, 10), width=0.8, color=vehicle_time_colors)
                plt.ylabel("Flight Time (Hours)")
                plt.title("Flight Time by Pilot and Vehicle Series (Sorted by First Flight Date)")
                plt.xlabel("Pilot")
                plt.xticks(rotation=45, ha='right')
                plt.legend(title="Launch Vehicle Series", bbox_to_anchor=(1.05, 1), loc='upper left')
                
                # Add value labels on top of bars (total flight time per pilot)
                pilot_totals = pilot_time_by_vehicle.sum(axis=1)
                for i, total_time in enumerate(pilot_totals.values):
                    if total_time > 0:
                        if total_time >= 24:
                            label = f"{total_time:.0f}h\n({total_time/24:.1f}d)"
                        else:
                            label = f"{total_time:.1f}h"
                        plt.text(i, total_time + max(pilot_totals.values) * 0.01, label, 
                                ha='center', va='bottom', fontsize=9)
                
                plt.tight_layout()
                plt.savefig(custom_dir / "pilot_flight_time_by_vehicle_series.png", bbox_inches='tight')
            finally:
                plt.close()

def create_program_analysis(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors):
    """Create program-specific analysis charts"""
    program_dir = output_dir / "program_analysis"
    
    # Extract program names from Mission/Payload
    df['Program'] = df['Mission / Payload'].apply(extract_program_name)
    
    # Get unique programs
    programs = df['Program'].unique()
    programs = [p for p in programs if p != 'Other' and p.strip()]
    
    for program in programs:
        program_df = df[df['Program'] == program].copy()
        
        if len(program_df) == 0:
            continue
            
        # Periodic cleanup to prevent memory buildup
        if len(plt.get_fignums()) > 10:
            plt.close('all')
            
        # Create program-specific directory
        safe_program_name = re.sub(r'[^\w\s-]', '', program).strip().replace(' ', '_')
        program_subdir = program_dir / safe_program_name
        program_subdir.mkdir(exist_ok=True)
        
        # Create various charts for this program
        _create_program_charts(program_df, program, safe_program_name, program_subdir, 
                              outcome_colors, series_colors, crewed_colors)

def _create_program_charts(program_df, program, safe_program_name, program_subdir, 
                          outcome_colors, series_colors, crewed_colors):
    """Helper function to create charts for a specific program"""
    
    # 1. Mission Outcomes Pie Chart
    if len(program_df["Outcome"].value_counts()) > 1:
        outcome_counts = program_df["Outcome"].value_counts()
        create_pie_chart(outcome_counts, f"{program} - Mission Outcomes", 
                        f"{safe_program_name}_outcomes_pie.png", program_subdir,
                        lambda x: outcome_colors.get(x, '#888888'))
    
    # 2. Vehicle Series Pie Chart
    if len(program_df["Launch Vehicle Series"].value_counts()) > 1:
        series_counts = program_df["Launch Vehicle Series"].value_counts()
        create_pie_chart(series_counts, f"{program} - Vehicle Series",
                        f"{safe_program_name}_vehicles_pie.png", program_subdir,
                        lambda x: series_colors.get(x, '#888888'))
    
    # 3. Crewed vs Uncrewed Pie Chart
    if len(program_df["Crewed"].value_counts()) > 1:
        crew_counts = program_df["Crewed"].value_counts().rename({True: "Crewed", False: "Uncrewed"})
        create_pie_chart(crew_counts, f"{program} - Crewed vs Uncrewed",
                        f"{safe_program_name}_crewed_pie.png", program_subdir, crewed_colors)
    
    # 4. Mission Outcomes Bar Chart
    outcome_counts = program_df["Outcome"].value_counts()
    if len(outcome_counts) > 0:
        update_progress(f"Creating {program} outcomes bar chart")
        try:
            outcome_colors_bar = [outcome_colors.get(outcome, '#888888') for outcome in outcome_counts.index]
            ax = outcome_counts.plot(kind='bar', figsize=(10, 6), color=outcome_colors_bar)
            plt.ylabel("Number of Missions")
            plt.title(f"{program} - Mission Outcomes")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(program_subdir / f"{safe_program_name}_outcomes_bar.png")
        finally:
            plt.close()
    
    # 5. Vehicle Series Bar Chart
    series_counts = program_df["Launch Vehicle Series"].value_counts()
    if len(series_counts) > 0:
        update_progress(f"Creating {program} vehicle series bar chart")
        try:
            series_colors_bar = [series_colors.get(series, '#888888') for series in series_counts.index]
            ax = series_counts.plot(kind='bar', figsize=(10, 6), color=series_colors_bar)
            plt.ylabel("Number of Missions")
            plt.title(f"{program} - Vehicle Series")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.savefig(program_subdir / f"{safe_program_name}_vehicles_bar.png")
        finally:
            plt.close()
    
    # 6. Program Time Series Chart (KSP Date)
    if len(program_df) > 1:
        program_df_copy = program_df.copy()
        
        # Create complete year range for the program
        min_year = int(program_df_copy["Year"].min())
        max_year = int(program_df_copy["Year"].max())
        all_years = list(range(min_year, max_year + 1))
        
        program_by_year = program_df_copy.groupby("Year").size()
        program_by_year = program_by_year.reindex(all_years, fill_value=0)
        
        if len(program_by_year) > 0:
            create_time_series_chart(program_by_year, f"{program} - Missions Over Time (KSP Date)",
                                   f"{safe_program_name}_time_series.png", program_subdir, '#4682B4')
    
    # 7. Program Outcomes vs Time
    if len(program_df) > 1 and len(program_df["Year"].unique()) > 1 and len(program_df["Outcome"].unique()) > 1:
        program_df_copy = program_df.copy()
        
        min_year = int(program_df_copy["Year"].min())
        max_year = int(program_df_copy["Year"].max())
        all_years = list(range(min_year, max_year + 1))
        
        outcome_by_year = program_df_copy.groupby(["Year", "Outcome"]).size().unstack(fill_value=0)
        outcome_by_year = outcome_by_year.reindex(all_years, fill_value=0)
        
        outcome_by_year = sort_outcome_columns(outcome_by_year)
        
        colors = [outcome_colors.get(col, '#888888') for col in outcome_by_year.columns]
        create_time_series_chart(outcome_by_year, f"{program} - Mission Outcomes Over Time (KSP Date)",
                               f"{safe_program_name}_outcomes_vs_time.png", program_subdir, colors, stacked=True)
    
    # 8. Program Vehicle Series vs Time
    if len(program_df) > 1 and len(program_df["Year"].unique()) > 1 and len(program_df["Launch Vehicle Series"].unique()) > 1:
        program_df_copy = program_df.copy()
        
        min_year = int(program_df_copy["Year"].min())
        max_year = int(program_df_copy["Year"].max())
        all_years = list(range(min_year, max_year + 1))
        
        series_by_year = program_df_copy.groupby(["Year", "Launch Vehicle Series"]).size().unstack(fill_value=0)
        series_by_year = series_by_year.reindex(all_years, fill_value=0)
        
        colors = [series_colors.get(col, '#888888') for col in series_by_year.columns]
        create_time_series_chart(series_by_year, f"{program} - Launch Vehicle Series Over Time (KSP Date)",
                               f"{safe_program_name}_vehicles_vs_time.png", program_subdir, colors,
                               figsize=(14, 8), stacked=True)

def create_program_breakdowns(df, output_dir, outcome_colors, series_colors, success_colors, crewed_colors, vehicle_colors, program_colors):
    """Create program-specific versions of all chart types"""
    
    # Extract program names
    df['Program'] = df['Mission / Payload'].apply(extract_program_name)
    
    # Use existing directories
    pie_dir = output_dir / "pie_charts"
    ksp_dir = output_dir / "ksp_time_series"
    irl_dir = output_dir / "irl_time_series"
    custom_dir = output_dir / "custom_charts"
    
    # Filter programs to show (more than 1 mission, not 'Other')
    program_counts = df['Program'].value_counts()
    programs_to_show = [p for p in program_counts.index if p != 'Other' and program_counts[p] > 1]
    
    # 1. PIE CHARTS BY PROGRAM
    if programs_to_show:
        program_counts_filtered = program_counts[programs_to_show]
        create_pie_chart(program_counts_filtered, "Mission Distribution by Program",
                        "pie_programs_distribution.png", pie_dir,
                        lambda x: program_colors.get(x, '#888888'), figsize=(12, 12))
    
    # 2. KSP TIME SERIES BY PROGRAM
    program_by_year = df.groupby(["Year", "Program"]).size().unstack(fill_value=0)
    programs_cols = [col for col in program_by_year.columns if col in programs_to_show]
    if programs_cols:
        program_by_year_filtered = program_by_year[programs_cols]
        colors = [program_colors.get(col, '#888888') for col in programs_cols]
        create_time_series_chart(program_by_year_filtered, "Program Activity Over Time (KSP Date)",
                               "ksp_programs_vs_time.png", ksp_dir, colors, 
                               figsize=(16, 8), stacked=True)
    
    # 3. IRL TIME SERIES BY PROGRAM
    df, date_range = prepare_irl_date_data(df)
    program_by_irl_day = df.groupby(["IRL Day", "Program"]).size().unstack(fill_value=0)
    program_by_irl_day = program_by_irl_day.reindex(date_range.date, fill_value=0)
    
    programs_cols = [col for col in program_by_irl_day.columns if col in programs_to_show]
    if programs_cols:
        program_by_irl_day_filtered = program_by_irl_day[programs_cols]
        colors = [program_colors.get(col, '#888888') for col in programs_cols]
        create_time_series_chart(program_by_irl_day_filtered, "Program Activity Over Time (IRL Date - Daily)",
                               "irl_programs_vs_time.png", irl_dir, colors,
                               figsize=(18, 8), width=1.0, stacked=True, time_type='irl')
    
    # 4. CUSTOM CHARTS BY PROGRAM
    df["Success Category"] = "Unsuccessful"
    success_outcomes = ['Great Success', 'Success', 'Partial Success']
    df.loc[df["Outcome"].isin(success_outcomes), "Success Category"] = "Successful"
    
    # Success/Failure by Program
    success_by_program = df.groupby(["Program", "Success Category"]).size().unstack(fill_value=0)
    programs_rows = [p for p in success_by_program.index if p in programs_to_show]
    if programs_rows:
        success_by_program_filtered = success_by_program.loc[programs_rows]
        
        update_progress("Creating success/failure by program chart")
        try:
            ax = success_by_program_filtered.plot(kind='bar', stacked=True, figsize=(14, 8), width=0.8, color=success_colors)
            plt.ylabel("Number of Missions")
            plt.title("Successful vs Unsuccessful Flights by Program")
            plt.xlabel("Program")
            plt.xticks(rotation=45, ha='right')
            plt.legend(title="Mission Category")
            plt.tight_layout()
            plt.savefig(custom_dir / "program_flights_by_success.png", bbox_inches='tight')
        finally:
            plt.close()
    
    # Crewed vs Uncrewed by Program
    crewed_by_program = df.groupby(["Program", "Crewed"]).size().unstack(fill_value=0)
    crewed_by_program.rename(columns={True: "Crewed", False: "Uncrewed"}, inplace=True)
    programs_rows = [p for p in crewed_by_program.index if p in programs_to_show]
    if programs_rows:
        crewed_by_program_filtered = crewed_by_program.loc[programs_rows]
        
        update_progress("Creating crewed/uncrewed by program chart")
        try:
            ax = crewed_by_program_filtered.plot(kind='bar', stacked=True, figsize=(14, 8), width=0.8, color=crewed_colors)
            plt.ylabel("Number of Missions")
            plt.title("Crewed vs Uncrewed Missions by Program")
            plt.xlabel("Program")
            plt.xticks(rotation=45, ha='right')
            plt.legend(title="Mission Type")
            plt.tight_layout()
            plt.savefig(custom_dir / "program_flights_by_crewed.png", bbox_inches='tight')
        finally:
            plt.close()
    
    # Pilots by Program
    pilot_program_data = []
    
    for _, row in df.iterrows():
        if pd.notna(row["Crew"]) and row["Crew"].strip():
            crew_members = [name.strip() for name in row["Crew"].split(',')]
            for pilot in crew_members:
                if pilot:
                    pilot_program_data.append({
                        'Pilot': pilot,
                        'Program': row["Program"],
                        'KSP Date': row["KSP Date"]
                    })
    
    pilot_program_df = pd.DataFrame(pilot_program_data)
    
    if not pilot_program_df.empty:
        # Find first flight date for each pilot and sort
        pilot_first_flight = pilot_program_df.groupby('Pilot')['KSP Date'].min().sort_values()
        pilot_order = pilot_first_flight.index.tolist()
        
        # Pilots by Program chart
        pilot_by_program = pilot_program_df.groupby(["Pilot", "Program"]).size().unstack(fill_value=0)
        pilot_by_program = pilot_by_program.reindex(pilot_order)
        
        # Only show programs that appear in programs_to_show
        program_cols = [col for col in pilot_by_program.columns if col in programs_to_show]
        if program_cols:
            pilot_by_program_filtered = pilot_by_program[program_cols]
            program_colors_list = [program_colors.get(col, '#888888') for col in program_cols]
            
            update_progress("Creating pilots by program chart")
            try:
                ax = pilot_by_program_filtered.plot(kind='bar', stacked=True, figsize=(16, 10), width=0.8, color=program_colors_list)
                plt.ylabel("Number of Flights")
                plt.title("Programs Flown by Each Pilot (Sorted by First Flight Date)")
                plt.xlabel("Pilot")
                plt.xticks(rotation=45, ha='right')
                plt.legend(title="Program", bbox_to_anchor=(1.05, 1), loc='upper left', ncol=1)
                plt.tight_layout()
                plt.savefig(custom_dir / "program_flights_by_pilot.png", bbox_inches='tight')
            finally:
                plt.close()

def calculate_total_charts(df):
    """Calculate the total number of charts that will be created"""
    total = 0
    
    # Basic pie charts: 4
    total += 4
    
    # KSP time series: 5 
    total += 5
    
    # IRL time series: 5
    total += 5
    
    # Custom charts: 10 (success by series, outcome by series, outcome by program, program by series, series by program, pilots by series, pilots by success, pilots by detailed outcome, pilot flight time total, pilot flight time by program, pilot flight time by vehicle series)
    total += 10
    
    # Program breakdown charts: 6 (programs pie, ksp time, irl time, success by program, crewed by program, pilots by program)
    total += 6
    
    # Program analysis charts - calculate actual count
    df['Program'] = df['Mission / Payload'].apply(extract_program_name)
    programs = df['Program'].unique()
    programs = [p for p in programs if p != 'Other' and p.strip()]
    
    # For each program, count actual charts that will be created
    for program in programs:
        program_df = df[df['Program'] == program].copy()
        if len(program_df) == 0:
            continue
            
        program_charts = 0
        
        # Pie charts (up to 3) - use actual column names
        if 'Outcome' in df.columns and len(program_df["Outcome"].unique()) > 1:
            program_charts += 1  # outcomes pie
        if 'Launch Vehicle Series' in df.columns and len(program_df["Launch Vehicle Series"].unique()) > 1:
            program_charts += 1  # vehicle series pie  
        if 'Crewed' in df.columns and len(program_df["Crewed"].unique()) > 1:
            program_charts += 1  # crewed pie
            
        # Bar charts (2)
        if 'Outcome' in df.columns and len(program_df["Outcome"].unique()) > 0:
            program_charts += 1  # outcomes bar
        if 'Launch Vehicle Series' in df.columns and len(program_df["Launch Vehicle Series"].unique()) > 0:
            program_charts += 1  # vehicle series bar
            
        # Time series (up to 3)
        if len(program_df) > 1:
            program_charts += 1  # basic time series
            if 'Year' in df.columns and 'Outcome' in df.columns:
                if len(program_df["Year"].unique()) > 1 and len(program_df["Outcome"].unique()) > 1:
                    program_charts += 1  # outcomes vs time
            if 'Year' in df.columns and 'Launch Vehicle Series' in df.columns:
                if len(program_df["Year"].unique()) > 1 and len(program_df["Launch Vehicle Series"].unique()) > 1:
                    program_charts += 1  # vehicles vs time
                
        total += program_charts
    
    return total

def main(csv_path):
    # Clear any existing figures to prevent memory warnings
    plt.close('all')
    
    print("Loading and filtering data...")
    
    # Load and prepare data
    df = pd.read_csv(csv_path)
    
    # Filter out launches without KSP Date
    initial_count = len(df)
    df = df.dropna(subset=['KSP Date'])
    df = df[df['KSP Date'].str.strip() != '']  # Remove empty strings too
    filtered_count = len(df)
    
    if initial_count != filtered_count:
        print(f"Filtered out {initial_count - filtered_count} launches without KSP Date")
        print(f"Processing {filtered_count} launches with valid KSP Date")
    
    # Fill missing Mission/Payload values
    df = fill_mission_payload_column(df)
    
    # Clean and prepare data
    df["KSP Date"] = pd.to_datetime(df["KSP Date"], format='mixed', errors='coerce')
    
    # Remove any rows where KSP Date couldn't be parsed
    df = df.dropna(subset=['KSP Date'])
    final_count = len(df)
    
    if filtered_count != final_count:
        print(f"Removed {filtered_count - final_count} additional launches with unparseable KSP Date")
        print(f"Final dataset: {final_count} launches")
    
    df["Outcome"] = df["Result"].fillna("In Progress")
    df["Crewed"] = df["Crew"].notna() & df["Crew"].str.strip().ne("")
    df["Year"] = df["KSP Date"].dt.year
    
    # Create Vehicle Type column
    df["Vehicle Type"] = "Rocket"
    df.loc[df["Launch Vehicle"].str.contains("ATP", na=False), "Vehicle Type"] = "Plane"
    
    # Create Launch Vehicle Series column
    df["Launch Vehicle Series"] = df["Launch Vehicle"].fillna("Unknown")
    
    # For ATP vehicles: ATP-XYY -> ATP-XXX
    atp_mask = df["Launch Vehicle Series"].str.contains("ATP-", na=False)
    df.loc[atp_mask, "Launch Vehicle Series"] = df.loc[atp_mask, "Launch Vehicle Series"].str.replace(r'ATP-(\d)\d+.*', r'ATP-\1XX', regex=True)
    
    # For Colussus 2 vehicles: remove everything after "Colussus 2"
    colussus_mask = df["Launch Vehicle Series"].str.contains("Colussus 2", na=False)
    df.loc[colussus_mask, "Launch Vehicle Series"] = "Colussus 2"
    
    # For other vehicles: remove trailing letters and -number patterns
    other_mask = ~atp_mask & ~colussus_mask & (df["Launch Vehicle Series"] != "Unknown")
    df.loc[other_mask, "Launch Vehicle Series"] = df.loc[other_mask, "Launch Vehicle Series"].str.replace(r'[A-Z]+$', '', regex=True)
    df.loc[other_mask, "Launch Vehicle Series"] = df.loc[other_mask, "Launch Vehicle Series"].str.replace(r'-\d+$', '', regex=True)
    
    # NOW calculate total charts after all data processing
    total_charts = calculate_total_charts(df.copy())
    set_total_charts(total_charts)
    print(f"\nGenerating {total_charts} charts...")
    print()  # Empty line for progress bar
    
    # Create output directories
    output_dir = create_output_directories()
    
    # Get color schemes
    outcome_colors, series_colors, success_colors, crewed_colors, vehicle_colors, program_colors = get_color_schemes()
    
    # Create all chart types
    create_pie_charts(df, output_dir, outcome_colors, series_colors)
    create_ksp_time_series(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors)
    create_irl_time_series(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors)
    create_custom_charts(df, output_dir, series_colors, success_colors, outcome_colors, program_colors)
    create_program_analysis(df, output_dir, outcome_colors, series_colors, crewed_colors, vehicle_colors)
    
    # Create program-specific versions of all chart types
    create_program_breakdowns(df, output_dir, outcome_colors, series_colors, success_colors, crewed_colors, vehicle_colors, program_colors)
    
    # Complete progress
    print(f"\n\nAll charts generated successfully in {output_dir}")
    print(f"Total missions processed: {len(df)}")
    print(f"Programs identified: {df['Program'].value_counts().to_dict()}")
    
    # Final cleanup
    plt.close('all')

if __name__ == "__main__":
    main("Career-1-Brownsville - Sheet1.csv")
