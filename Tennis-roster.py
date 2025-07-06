import pandas as pd

# Configuration
team_name = "Hadfield"
player_list  = ["Soheil", "Carla", "Jesse", "Georgia", "Jen", "Joe","Zac"]
unavailable_dates = {
    "Carla": ["2025-09-24"],
    "Jen": ["2025-10-01","2025-10-08","2025-10-15"]
}

file_path = r"C:\Users\Soheil\Desktop\projects\NSNTA-Timetable\Mixed Fixture.xlsx"


df = pd.read_excel(file_path,sheet_name='B Res 3')

df['Dates'] = df['Dates'].dt.strftime('%Y-%m-%d')

player_stats = {p:{'games':0, 'home_games':0} for p in player_list}

def is_available(player,date):
    return date not in unavailable_dates.get(player,[])

def available_players(date):
    return [p for p in player_list if is_available(p,date)]

schedule = []
for date in df['Dates']:
    available = available_players(date)
    
    # Sort available players by least games played
    available.sort(key=lambda p: player_stats[p]['games'])
    
    while len(available) >= 4:
        group = available[:4]
        available = available[4:]
        
        total_home_games = sum(player_stats[p]['home_games'] for p in group)
        total_games = sum(player_stats[p]['games'] for p in group)
        total_away_games = total_games - total_home_games
        
        # Decide location for whole group
        if total_home_games <= total_away_games:
            location = 'Home'
            for p in group:
                player_stats[p]['games'] += 1
                player_stats[p]['home_games'] += 1
        else:
            location = 'Away'
            for p in group:
                player_stats[p]['games'] += 1
        
        schedule.append({
            'Date': date,
            'Players': group,
            'Location': location
        })

    rows = []
    for game in schedule:
        for p in game['Players']:
            rows.append({
                'Date': game['Date'],
                'Player': p,
                'Location': game['Location']
            })

df_schedule = pd.DataFrame(rows)

# Save to Excel
output_file = r"C:\Users\Soheil\Desktop\projects\NSNTA-Timetable\Generated_Tennis_Schedule.xlsx"
df_schedule.to_excel(output_file, index=False) 

print(f"Schedule generated and saved to {output_file}")