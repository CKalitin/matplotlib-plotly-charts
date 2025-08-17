import plotly.express as px
import plotly.io as pio
import pandas as pd
from html2image import Html2Image

def generate(file_name="station-modules-gantt/station_modules_gantt_chart",
             max_name_line_len=20,
             label_font_size=32,
             lead_name_font_size=20,
             title_font_size=48,
             dimensions=(3840,2160)
             ):
        
    team_color_map = {
        "Almaz": "#970c10",
        "Skylab": "#1155cc",
        "TKS": "#ff8000",
        "Russian Docking Module": "#fad711",
        "ISS US": "#1155cc",
        "ISS ESA": "#00e1ff",
        "ISS JAXA": "#00e1ff",
        "ISS Bigelow": "#00e1ff",
        "Tiangong": "#eb1700",
    }
    
    team_order = [
        "Almaz",
        "Skylab",
        "TKS",
        "Russian Docking Module",
        "ISS US",
        "ISS ESA",
        "ISS JAXA",
        "ISS Bigelow",
        "Tiangong",
    ]

    # Need color map per person because Plotly
    display_color_map = {}

    def parse_file():
        data = []
        current_team = None
        with open("station-modules-gantt/station_data.txt") as f:
            for line in f:
                line = line.strip()
                line = line.replace(", -", ", 2025-07-11") # Replace ", -" with current date
                if not line:
                    continue
                tokens = line.split()
                if len(tokens) >= 3 and tokens[-2].count("-") == 2 and tokens[-1].count("-") == 2:
                    start, end = tokens[-2], tokens[-1]
                    name = " ".join(tokens[:-2]).replace(",", "")
                    data.append({
                        "Team": current_team,
                        "Lead": name,
                        "Start": start,
                        "End": end
                    })
                    display_color_map[name] = team_color_map.get(current_team, "#000000")  # Default to black if not found
                else:
                    current_team = line
        return data

    def truncate_name(name, max_len=max_name_line_len):
        name_split = name.split("\n")
        out = ""
        for ns in name_split:
            out += ns if len(ns) <= max_len else ns[:max_len - 3] + "...\n"
        return out.strip()

    def linebreak_name(name):
        return name.replace(" ", "\n ")

    data = parse_file()

    df = pd.DataFrame(data)
    df["Start"] = pd.to_datetime(df["Start"].str.replace(",", ""))
    df["End"] = pd.to_datetime(df["End"].str.replace(",", ""))
    df["Offset"] = 0

    df["Lead"] = df["Lead"].apply(linebreak_name)
    df["Lead"] = df["Lead"].apply(truncate_name)

    offset_tracker = {}  # Dict of {team: list of (end_time, offset)}
    for idx, row in df.iterrows():
        team = row["Team"]
        start = row["Start"]
        end = row["End"]
        
        if team not in offset_tracker:
            offset_tracker[team] = []

        taken_offsets = set()

        # Find which offsets are occupied at this bar's start time
        for prev_end, offset in offset_tracker[team]:
            if prev_end > start:
                taken_offsets.add(offset)
        
        # Assign lowest available offset
        offset = 0
        while offset in taken_offsets:
            offset += 1

        df.loc[idx, "Offset"] = offset
        offset_tracker[team].append((end, offset))

    # Initialize 'Y' column as float to avoid future warning
    df["Y"] = 0.0

    # Set custom y for each team, no gap between same team (eg. BMS 0 and 1) gap between different teams

    team_y_mapping = {}
    current_y_position = 0

    # Calculate Y positions based on team order and max offsets
    for team in team_order:
        team_rows = df[df["Team"] == team]
        max_offset = team_rows["Offset"].max()
        if pd.isna(max_offset):
            max_offset = -1 # Treat as if no offsets, so current_y_position increments by 0 + 1 + 0.5 below

        # Store the base Y position for each offset within this team
        for offset_val in range(max_offset + 1):
            team_y_mapping[(team, offset_val)] = current_y_position + offset_val

        current_y_position += (max_offset + 1) + 0.0 # Add 0.5 gap between teams, nvm doesn't work bc y labels

    # Create label column to display on Y-axis
    df["Y_Label"] = df.apply(lambda row: f"{row['Team']} ({int(row['Offset'])})", axis=1)

    # Loop through Y_Label and remove all containing (1)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(1)" not in x else " " * len(x))
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(2)" not in x else " " * len(x) * 2)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(3)" not in x else " " * len(x) * 3)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(4)" not in x else " " * len(x) * 4)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(5)" not in x else " " * len(x) * 5)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x if "(6)" not in x else " " * len(x) * 6)
    df["Y_Label"] = df["Y_Label"].apply(lambda x: x.replace("(0)", ""))

    # Assign Y values to DataFrame rows directly
    for idx, row in df.iterrows():
        df.loc[idx, "Y"] = team_y_mapping[(row["Team"], row["Offset"])]

    df["ColorGroup"] = df["Team"]

    # Build ordered Y-axis list respecting team_order
    y_order = []
    for team in team_order:
        offsets = sorted(df[df["Team"] == team]["Offset"].unique())
        for offset in offsets:
            y_order.append(f"{team} ({offset})")
    y_order.reverse()

    # Plot
    fig = px.timeline(df,
                    x_start="Start",
                    x_end="End",
                    y="Y_Label",
                    color="ColorGroup",
                    text="Lead",
                    category_orders={"Y": y_order},
                    color_discrete_map=team_color_map
                    )

    fig.update_yaxes(autorange="reversed")
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        title = {
            "text": "<b>Space Station Modules Timeline</b>",
            "font": {"size": title_font_size, "color": "#000000"},
            "x": 0.5,
            "y": 0.99,
            "xanchor": "center",
            "yanchor": "top",
        },
        showlegend=False,
        bargap=0,
        yaxis_title="",
        font_size=label_font_size,
        margin=dict(t=120), # Higher number = more space
    )
    fig.update_traces(textfont_size=lead_name_font_size)  # Change 16 to your desired font size

    # Top Left add "Christopher Kalitin 2025"
    fig.add_annotation(
        text="Christopher Kalitin 2025",
        xref="paper", yref="paper",
        x=-0.15, y=1.05,
        showarrow=False,
        font=dict(size=label_font_size, color="#202020"),
        align="left"
    )

    pio.write_html(fig, f"{file_name}.html", auto_open=False)

    # pio.write_image is broken and I'm not fixing it this is the downside of using LLMs

    hti = Html2Image()
    print(f"{file_name.split("/")[-1]}.png")
    hti.screenshot(html_file=f"{file_name}.html", save_as=f"{file_name.split("/")[-1]}.png", size=dimensions)

generate("station-modules-gantt/station_modules_gantt_chart_4k",
         max_name_line_len=20,
         label_font_size=32,
         lead_name_font_size=20,
         title_font_size=48,
         dimensions=(3840, 2160)
)

generate("station-modules-gantt/station_modules_gantt_chart_1080p",
         max_name_line_len=20,
         label_font_size=16,
         lead_name_font_size=9,
         title_font_size=24,
         dimensions=(1920, 1080),
)
