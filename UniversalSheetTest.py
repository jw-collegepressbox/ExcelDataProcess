import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from collections import defaultdict
import re
from openpyxl import load_workbook
import io
import zipfile
import requests
import toml
import time # Added for demonstration
import streamlit as st
import toml, os


st.set_page_config(layout="wide")
st.title("🏈 Football XML Stats Parser")
st.markdown("Enter your Dropbox ZIP URL, upload the Excel template, and generate reports for all teams.")

# Load the Dropbox token from secrets_folder/secrets.toml
# Get the absolute path to the current script directory
base_dir = os.path.dirname(__file__)

# Construct the relative path to the secrets file
secrets_path = os.path.join(base_dir, "secrets_folder", "secrets.toml")

# Load the file
secrets = toml.load(secrets_path)
DROPBOX_ACCESS_TOKEN = secrets["DROPBOX_ACCESS_TOKEN"]

# --- TEAM MAPPING (Canonical Names for Schedule/Selection) ---
TEAM_NAME_MAP = {
    "airforce": "Air Force", "akron": "Akron", "alabama": "Alabama", "appstate": "App State",
    "arizona": "Arizona", "asu": "Arizona State", "arkansas": "Arkansas", "arkstate": "Arkansas State", "astate": "Arkansas State",
    "army": "Army West Point", "auburn": "Auburn", "ballstate": "Ball State", "baylor": "Baylor",
    "boisestate": "Boise State", "bc": "Boston College", "bgsu": "Bowling Green", "buffalo": "Buffalo",
    "byu": "BYU", "cal": "California", "cmu": "Central Michigan", "charlotte": "Charlotte",
    "cincinnati": "Cincinnati", "clemson": "Clemson", "coastal": "Coastal Carolina", "colorado": "Colorado",
    "csu": "Colorado State", "delaware": "Delaware", "duke": "Duke", "ecu": "East Carolina",
    "emu": "Eastern Michigan", "fiu": "FIU", "florida": "Florida", "fau": "Florida Atlantic",
    "fsu": "Florida State", "fresnostate": "Fresno State", "georgia": "Georgia", "gasouthern": "Georgia Southern",
    "gastate": "Georgia State", "gatech": "Georgia Tech", "hawaii": "Hawaii", "houston": "Houston",
    "illinois": "Illinois", "indiana": "Indiana", "iowa": "Iowa", "iowastate": "Iowa State",
    "jmu": "James Madison", "jaxstate": "Jacksonville State", "kansas": "Kansas", "kstate": "Kansas State",
    "kennesaw": "Kennesaw State", "kentstate": "Kent State", "kentucky": "Kentucky", "liberty": "Liberty",
    "louisiana": "Louisiana", "latech": "Louisiana Tech", "louisville": "Louisville", "lsu": "LSU",
    "marshall": "Marshall", "maryland": "Maryland", "umass": "Massachusetts", "memphis": "Memphis",
    "miami": "Miami (FL)", "muohio": "Miami (OH)", "michigan": "Michigan", "michstate": "Michigan State",
    "mtsu": "Middle Tenn", "minnesota": "Minnesota", "mstate": "Mississippi State", "missouri": "Missouri",
    "mostate": "Missouri State", "navy": "Navy", "ncsu": "NC State", "nebraska": "Nebraska",
    "nevada": "Nevada", "unm": "New Mexico", "nmsu": "New Mexico State", "niu": "NIU",
    "unc": "North Carolina", "unt": "North Texas", "northwestern": "Northwestern", "notredame": "Notre Dame",
    "ohio": "Ohio", "ohiostate": "Ohio State", "oklahoma": "Oklahoma", "ostate": "Oklahoma State",
    "odu": "Old Dominion", "olemiss": "Ole Miss", "oregon": "Oregon", "oregonstate": "Oregon State",
    "pennstate": "Penn State", "pitt": "Pitt", "purdue": "Purdue", "rice": "Rice",
    "rutgers": "Rutgers", "shsu": "Sam Houston", "sdsu": "San Diego State", "sjsu": "San Jose State",
    "smu": "SMU", "usa": "South Alabama", "sc": "South Carolina", "usf": "South Florida",
    "usm": "Southern Miss", "stanford": "Stanford", "syracuse": "Syracuse", "tcu": "TCU",
    "temple": "Temple", "tennessee": "Tennessee", "texas": "Texas", "tamu": "Texas A&M",
    "texasstate": "Texas State", "texastech": "Texas Tech", "toledo": "Toledo", "troy": "Troy",
    "tulane": "Tulane", "tulsa": "Tulsa", "uab": "UAB", "ucf": "UCF",
    "ucla": "UCLA", "uconn": "UConn", "ulm": "ULM", "unlv": "UNLV",
    "usc": "USC", "utah": "Utah", "utahstate": "Utah State", "utep": "UTEP",
    "utsa": "UTSA", "vanderbilt": "Vanderbilt", "virginia": "Virginia", "vt": "Virginia Tech",
    "wfu": "Wake Forest", "washington": "Washington", "washstate": "Washington State", "wvu": "West Virginia",
    "wmu": "Western Michigan", "wisconsin": "Wisconsin", "wku": "WKU", "wyoming": "Wyoming"
}

# --- XML NAME ALIASES (Fixes for Inconsistent XML Names) ---
XML_NAME_ALIASES = {
    "appalachian state": "App State", 
    "central mich": "Central Michigan",
    "eastern mich": "Eastern Michigan",
    "western mich": "Western Michigan",
    "fla atlantic": "Florida Atlantic",
    "florida atlantic": "Florida Atlantic",
    "ga southern": "Georgia Southern",
    "Georgia  Southern": "Georgia Southern",
    "ga. southern": "Georgia Southern",
    "ull": "Louisiana",
    "northern illinois": "NIU",
    "#6/#5 notre dame": "Notre Dame",
    "nm state": "New Mexico State",
    "new mexico st": "New Mexico State",
    "pittsburgh": "Pitt",
    "south fla": "South Florida",
    "#18/#23 south florida": "South Florida",
    "southern california": "USC",
    "western ky": "WKU",
    "western kentucky": "WKU",
    "miami": "Miami (FL)",
    "#5 miami (fl)": "Miami (FL)",
    "#12 miami (fl)": "Miami (FL)",
    "#2/#2 miami (fl)": "Miami (FL)",
    "#4/#6 miami (fl)": "Miami (FL)",
    "#5/#6 miami (fl)": "Miami (FL)",
    "#5/#7 miami (fl)": "Miami (FL)",
    "#9/#9 miami (fl)": "Miami (FL)",
    "#10/#10 miami (fl)": "Miami (FL)",
    "hawai'i": "Hawaii",
    "jax state": "Jacksonville State"
}

# --- HELPER FUNCTIONS (Adjusted to take template_data instead of template_file) ---

def normalize_team_name(name):
    """
    Normalizes team names from XML, aggressively handling inconsistent spacing and aliases.
    The fix is here: re.sub(r'\s+', ' ', name).strip()
    """
    if not name:
        return ""
        
    # 1. AGGRESSIVE WHITESPACE CLEANUP: Replace all sequences of whitespace 
    # (spaces, tabs, non-breaking spaces like the one in your file) with a single space, then strip.
    name_cleaned = re.sub(r'\s+', ' ', name).strip()
    
    # 2. Apply existing cleanup rules on the now clean name
    name_cleaned = name_cleaned.replace('.', '')
    name_cleaned = re.sub(r'^#\d+\s+', '', name_cleaned)
    name_cleaned = re.sub(r'\bSt\b', 'State', name_cleaned)
    name_cleaned = name_cleaned.strip()
    
    # 3. CRITICAL FIX: Alias Check using the cleaned, lowercase name
    name_lower = name_cleaned.lower()
    
    if name_lower in XML_NAME_ALIASES:
        return XML_NAME_ALIASES[name_lower]
    
    # 4. Default Return
    return name_cleaned

def build_cumulative_df(core_stats, master_order, category):
    data = []
    # Simplified keys for cleaner code, matching the required output columns
    stat_keys = {
        'rush': ['Rush Att','Rush Yds','Rush TD'],
        'pass': ['Pass Comp','Pass Att','Pass Yds','Pass TD','Pass INT'],
        'recv': ['Rec','Rec Yds','Rec TD']
    }
    
    for p in master_order[category]:
        # Default stats for players not found in core_stats (shouldn't happen, but safe)
        s = core_stats.get(p, {k: 0 for k in stat_keys['rush'] + stat_keys['pass'] + stat_keys['recv']})

        if category == 'rush':
            row = {'Player': p, 'Att': s['Rush Att'], 'Yds': s['Rush Yds'], 'TD': s['Rush TD']}
        elif category == 'pass':
            row = {'Player': p, 'Comp': s['Pass Comp'], 'Att': s['Pass Att'], 'Yds': s['Pass Yds'], 'TD': s['Pass TD'], 'INT': s['Pass INT']}
        elif category == 'recv':
            row = {'Player': p, 'Rec': s['Rec'], 'Yds': s['Rec Yds'], 'TD': s['Rec TD']}
        data.append(row)
    return pd.DataFrame(data)

def fill_excel_df(ws, df, start_row, columns):
    for i, row in df.iterrows():
        excel_row = start_row + i
        for col_letter, col_name in columns.items():
            ws[f"{col_letter}{excel_row}"] = row[col_name]

def set_cell_value_safe(ws, row, col_letter, value):
    cell_coord = f"{col_letter}{row}"
    # Check if the cell is part of a merged area but is not the top-left cell
    for merged in ws.merged_cells.ranges:
        if cell_coord in merged:
            if cell_coord != merged.start_cell.coordinate:
                return # Skip writing to merged cells that aren't the primary cell
    ws[cell_coord].value = value

def build_game_df(core_game_stats, master_list, category):
    data = []
    stat_keys = {
        'rush': ['Rush Att', 'Rush Yds', 'Rush TD'],
        'pass': ['Pass Comp', 'Pass Att', 'Pass Yds', 'Pass TD', 'Pass INT', 'Sacks'],
        'recv': ['Rec', 'Rec Yds', 'Rec TD']
    }

    for p in master_list:
        # Get stats for the current player in the current game
        stats = core_game_stats.get(p, {k: 0 for k in stat_keys['rush'] + stat_keys['pass'] + stat_keys['recv']})
        
        if category == 'rush':
            row = {'Player': p, 'Att': stats['Rush Att'], 'Yds': stats['Rush Yds'], 'TD': stats['Rush TD']}
        elif category == 'pass':
            row = {'Player': p, 'Comp': stats['Pass Comp'], 'Att': stats['Pass Att'],
                   'Yds': stats['Pass Yds'], 'TD': stats['Pass TD'], 'INT': stats['Pass INT']}
        elif category == 'recv':
            row = {'Player': p, 'Rec': stats['Rec'], 'Yds': stats['Rec Yds'], 'TD': stats['Rec TD']}
        data.append(row)
    return pd.DataFrame(data)

# --- CORE PROCESSING FUNCTION (Updated to take template_data) ---
def process_team(core_team_name, team_stats, game_stats, parsed_games, template_data, keep_vba, progress_bar, progress_index, total_teams):
    
    # Update progress bar
    progress_percent = (progress_index / total_teams)
    progress_bar.progress(progress_percent, text=f"Processing {core_team_name}... ({progress_index}/{total_teams})")
    
    core_stats = team_stats[core_team_name]

    # --- Step 3b: Get schedule games ---
    schedule_games = []
    
    def get_full_name_from_file_segment(segment):
        if segment in TEAM_NAME_MAP:
            return TEAM_NAME_MAP[segment]
        return segment.title().replace("St", "State") 

    for game_date, team1, team2, file_stream, file_name in parsed_games:
        full_name1 = get_full_name_from_file_segment(team1)
        full_name2 = get_full_name_from_file_segment(team2)
        
        if full_name1 == core_team_name:
            is_home = False
            opponent_file_segment = team2 
        elif full_name2 == core_team_name:
            is_home = True
            opponent_file_segment = team1
        else:
            continue
            
        game_name = file_name.replace(".xml", "")

        schedule_games.append((game_date, game_name, opponent_file_segment, is_home))

    if not schedule_games:
        return None 

    # --- Step 4: Build master order (Logic remains the same) ---
    master_order = {'rush': [], 'pass': [], 'recv': []}

    for game_idx, (game_date, game_name, opponent, is_home) in enumerate(schedule_games, start=1):
        core_game_stats = game_stats[game_name][core_team_name] 

        rush_players = [p for p, s in core_game_stats.items() if s['Rush Att'] > 0 or s['Rush Yds'] != 0 or s['Rush TD'] != 0]
        pass_players = [p for p, s in core_game_stats.items() if (s['Pass Att'] > 0 or s['Pass Yds'] != 0 or s['Pass TD'] != 0 or s['Pass Comp'] > 0 or s['Pass INT'] > 0 or s['Sacks'] > 0)]
        recv_players = [p for p, s in core_game_stats.items() if s['Rec'] > 0 or s['Rec Yds'] != 0 or s['Rec TD'] != 0]

        if game_idx == 1:
            rush_players.sort(key=lambda p: core_game_stats[p]['Rush Yds'], reverse=True)
            pass_players.sort(key=lambda p: core_game_stats[p]['Pass Yds'], reverse=True)
            recv_players.sort(key=lambda p: core_game_stats[p]['Rec Yds'], reverse=True)

        for category, players in [('rush', rush_players), ('pass', pass_players), ('recv', recv_players)]:
            for p in players:
                if p not in master_order[category]:
                    master_order[category].append(p)

    # --- Step 5: Build cumulative DataFrames (Logic remains the same) ---
    rush_df = build_cumulative_df(core_stats, master_order, 'rush')
    pass_df = build_cumulative_df(core_stats, master_order, 'pass')
    recv_df = build_cumulative_df(core_stats, master_order, 'recv')
    
    # --- Step 6: Load Excel template from in-memory data and Fill cumulative stats ---
    
    # CRITICAL FIX: Load from in-memory template_data buffer, not the original st.file_uploader object
    template_buffer = io.BytesIO(template_data) 
    
    try:
        wb = load_workbook(template_buffer, keep_vba=keep_vba)
        ws = wb["BOX SCORES"] 
    except KeyError:
        st.error(f"Failed to find the worksheet named 'BOX SCORES' in the template for {core_team_name}. Skipping.")
        return None
    except Exception as e:
        st.error(f"Error loading Excel template for {core_team_name}: {e}. Skipping.")
        return None

    # Fill cumulative stats
    fill_excel_df(ws, rush_df, 3, {'B':'Player','C':'Att','D':'Yds','F':'TD'})
    fill_excel_df(ws, pass_df, 3, {'G':'Player','H':'Comp','I':'Att','K':'Yds','M':'TD','O':'INT'})
    fill_excel_df(ws, recv_df, 3, {'R':'Player','S':'Rec','T':'Yds','V':'TD'})

    # --- Step 7: Fill game stats ---
    for game_idx, (game_date, game_name, opponent, is_home) in enumerate(schedule_games, start=1):
        core_game_stats = game_stats[game_name][core_team_name]
        
        # Build game DFs using the master order
        rush_df_game = build_game_df(core_game_stats, master_order['rush'], 'rush')
        pass_df_game = build_game_df(core_game_stats, master_order['pass'], 'pass')
        recv_df_game = build_game_df(core_game_stats, master_order['recv'], 'recv')

        # Fill game blocks
        base_row = 32 + (game_idx - 1) * 29
        fill_excel_df(ws, rush_df_game, base_row, {'C': 'Att', 'D': 'Yds', 'F': 'TD'})
        fill_excel_df(ws, pass_df_game, base_row, {'H': 'Comp', 'I': 'Att', 'K': 'Yds', 'M': 'TD', 'O': 'INT'})
        fill_excel_df(ws, recv_df_game, base_row, {'S': 'Rec', 'T': 'Yds', 'V': 'TD'})

        # Fill player names (only once per game block)
        for i, p_name in enumerate(master_order['rush']):
             set_cell_value_safe(ws, base_row + i, 'B', p_name)
        
        for i, p_name in enumerate(master_order['pass']):
            set_cell_value_safe(ws, base_row + i, 'G', p_name)

        for i, p_name in enumerate(master_order['recv']):
            set_cell_value_safe(ws, base_row + i, 'R', p_name)


    # --- Step 8: Fill Schedule Columns (Z, AA, AB) ---
    # --- Step 8: Fill Schedule Columns (Z, AA, AB) ---
    for i, (game_date, game_name, opponent, is_home) in enumerate(schedule_games[:17], start=2):
        ws[f"Z{i}"] = game_date.strftime("%m/%d")

        is_playoff = (i - 1) > 12
        if is_playoff and not is_home:
            ws[f"AA{i}"] = "†"
        elif not is_home:
            ws[f"AA{i}"] = "at"
        else:
            ws[f"AA{i}"] = ""

        # ✅ SAFER TEAM NAME FIX
        opponent_cleaned = normalize_team_name(opponent)
        opponent_key = opponent.lower().replace(" ", "")

        # Try to map abbreviation to full name
        opponent_full = TEAM_NAME_MAP.get(opponent_key, opponent.title())

        # Write to Excel
        ws[f"AB{i}"] = opponent_full

        
    # --- Step 9: Save and Return file data ---
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    file_name = core_team_name + " Sheet.xlsm" if keep_vba else core_team_name + "_filled_stats.xlsx"
    
    return output, file_name 
# END of process_team function

# ----------------------------------------------------------------------
# --- MAIN APPLICATION LOGIC ---
# ----------------------------------------------------------------------

dropbox_zip_url = st.text_input(
    "Enter the Direct ZIP Download Link for the XML files (must end with &dl=1):",
    "" 
)

template_file = st.file_uploader("Upload your Excel template (.xlsx or .xlsm) **AFTER** entering the URL.", type=["xlsx", "xlsm"])

# Only start the process if the button is clicked, the URL is provided, AND the template is uploaded
if st.button("Generate Reports") and dropbox_zip_url and template_file:
    
    # ------------------ FILE DOWNLOAD & INITIAL XML PARSING ------------------
    
    try:
        # Download the ZIP file content
        with st.spinner(f"Downloading files from {dropbox_zip_url}..."):
            response = requests.get(dropbox_zip_url, stream=True, timeout=60)
            response.raise_for_status() 
            zip_buffer = io.BytesIO(response.content)

        all_xml_data = {}  
        with st.spinner("Extracting and preparing XML files..."):
            with zipfile.ZipFile(zip_buffer, 'r') as z:
                for file_info in z.infolist():
                    file_name = file_info.filename
                    if file_name.endswith('.xml') and not file_info.is_dir() and not file_name.startswith('__MACOSX'):
                        content = z.read(file_name)
                        all_xml_data[file_name] = io.BytesIO(content)
    
    except requests.exceptions.RequestException as e:
        st.error(f"Error downloading ZIP file. Please ensure the link is correct and ends with `&dl=1`. Error: {e}")
        st.stop()
    except zipfile.BadZipFile:
        st.error("The downloaded file is not a valid ZIP archive. Ensure the URL ends with `&dl=1`.")
        st.stop()
    except Exception as e:
        st.error(f"An unexpected error occurred during file extraction: {e}")
        st.stop()


    parsed_games = []
    
    for file_name_with_path, file_stream in all_xml_data.items():
        filename_base_with_ext = file_name_with_path.split('/')[-1].split('\\')[-1]
        filename_base = filename_base_with_ext.replace(".xml", "")
        
        match = re.match(r"(\d{6})gamebook_(.+)_(.+)", filename_base)
        
        if match:
            date_str, team1, team2 = match.groups()
            game_date = pd.to_datetime(date_str, format="%y%m%d")
            parsed_games.append((game_date, team1.lower(), team2.lower(), file_stream, filename_base_with_ext))
        else:
            st.warning(f"Filename format not recognized: {file_name_with_path}. Skipping.")

    if not parsed_games:
        st.error("No valid XML files found in the ZIP archive matching the required naming convention.")
        st.stop()
        
    parsed_games.sort(key=lambda x: x[0]) 

    # XML Parsing Structures
    team_stats = defaultdict(lambda: defaultdict(lambda: {
        'Rush Att': 0, 'Rush Yds': 0, 'Rush TD': 0, 'Rec': 0, 'Rec Yds': 0, 'Rec TD': 0,
        'Pass Comp': 0, 'Pass Att': 0, 'Pass Yds': 0, 'Pass TD': 0, 'Pass INT': 0, 'Sacks': 0
    }))
    team_names_set = set()
    game_stats = {}


    # XML PARSING LOOP
    with st.spinner("Parsing XML data and aggregating stats..."):
        for game_date, team1, team2, file_stream, file_name in parsed_games: 
            file_stream.seek(0)
            
            # Using try/except block here is safer for potentially malformed XML files
            try:
                tree = ET.parse(file_stream) 
                root = tree.getroot()
            except ET.ParseError:
                st.error(f"Error parsing XML file: {file_name}. Skipping.")
                continue

            game_name = file_name.replace(".xml", "")
            game_stats[game_name] = {}

            for team in root.findall('.//team'):
                team_name = normalize_team_name(team.attrib.get('name', ''))
                team_names_set.add(team_name)
                players_dict = {}

                for player in team.findall('./player'):
                    name = player.attrib.get('name', '')
                    opos = player.attrib.get('opos', None)
                    if opos in (None, 'None') and name != "TEAM":
                        continue
                    if ',' in name:
                        last, first = name.split(',', 1)
                        name = f"{first.strip()} {last.strip()}"

                    stats = {k: 0 for k in team_stats[team_name][name].keys()} # Initialize all stats to 0

                    if (r := player.find('rush')) is not None:
                        stats['Rush Att'] = int(r.attrib.get('att', 0))
                        stats['Rush Yds'] = int(r.attrib.get('yds', 0))
                        stats['Rush TD'] = int(r.attrib.get('td', 0))

                    if (rcv := player.find('rcv')) is not None:
                        stats['Rec'] = int(rcv.attrib.get('no', 0))
                        stats['Rec Yds'] = int(rcv.attrib.get('yds', 0))
                        stats['Rec TD'] = int(rcv.attrib.get('td', 0))

                    if (p := player.find('pass')) is not None:
                        stats['Pass Comp'] = int(p.attrib.get('comp', 0))
                        stats['Pass Att'] = int(p.attrib.get('att', 0))
                        stats['Pass Yds'] = int(p.attrib.get('yds', 0))
                        stats['Pass TD'] = int(p.attrib.get('td', 0))
                        stats['Pass INT'] = int(p.attrib.get('int', 0))
                        stats['Sacks'] = int(p.attrib.get('sacks', 0))
                    else:
                        stats['Sacks'] = 0
                    
                    for k in stats:
                        team_stats[team_name][name][k] += stats[k]
                    players_dict[name] = stats

                game_stats[game_name][team_name] = players_dict

    # ----------------------------------------------------------------------
    # --- TEMPLATE PROCESSING SETUP (The fix for the freeze!) ---
    # ----------------------------------------------------------------------
    
    # 1. Filter teams to only those in the TEAM_NAME_MAP
    valid_teams = sorted([team for team in team_names_set if team in TEAM_NAME_MAP.values()])
    total_teams = len(valid_teams)

    if not valid_teams:
        st.error("No teams found in the XML files that match the accepted team list for report generation.")
        st.stop()
        
    # 2. Load template data ONCE into a byte buffer
    keep_vba = template_file.name.endswith(".xlsm")
    # Read the full content into memory. This is the key to efficiency.
    template_data = template_file.getvalue() 
    
    st.header(f"📊 Starting Automated Report Generation for {total_teams} Teams...")
    progress_bar = st.progress(0, text="Initializing...")

    results = []
    
    # 3. Loop through every valid team and generate a report
    for idx, team in enumerate(valid_teams):
        # The process_team function now receives the raw template_data bytes.
        result = process_team(team, team_stats, game_stats, parsed_games, template_data, keep_vba, progress_bar, idx + 1, total_teams)
        
        if result:
            results.append(result)
        else:
            st.warning(f"Could not generate report for {team}.")

    # Set progress to 100% when finished
    progress_bar.progress(1.0, text="✅ All reports generated!")
    
    # 4. ZIP ARCHIVE CREATION 
    # ------------------ ZIP ARCHIVE CREATION ------------------
    st.header("📦 Packaging All Sheets into One File")

    if results:
        zip_buffer = io.BytesIO()
        
        with st.spinner("Creating ZIP file..."):
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for file_data_io, file_name in results:
                    zip_file.writestr(file_name, file_data_io.getvalue())
        
        zip_buffer.seek(0)

        # --- Local Download ---
        st.download_button(
            label = "📥 Download All Reports as ZIP",
            data = zip_buffer,
            file_name = "All_Team_Sheets.zip",
            mime = "application/zip"
        )

        # --- Dropbox Upload ---
        import dropbox

        dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)

        dropbox_folder = "/FootballReports"  # Folder path in Dropbox
        file_path = f"{dropbox_folder}/All_Team_Sheets.zip"

        with st.spinner("Uploading ZIP to Dropbox..."):
            try:
                zip_buffer.seek(0)  # Reset pointer to start
                dbx.files_upload(zip_buffer.read(), file_path, mode=dropbox.files.WriteMode.overwrite)
                st.success(f"✅ All reports uploaded to Dropbox at {file_path}!")
            except Exception as e:
                st.error(f"Failed to upload ZIP to Dropbox: {e}")

    else:
        st.error("No reports were successfully processed. Check the warnings above for potential errors (e.g., missing 'BOX SCORES' sheet).")
