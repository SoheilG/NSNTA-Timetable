import pandas as pd

# Configuration
team_name = "Hadfield"
player_list = ["Soheil", "Carla", "Jesse", "Georgia", "Jen", "Joe", "Zac"]

genders = {
    "Soheil": "boy",
    "Carla": "girl",
    "Jesse": "boy",
    "Georgia": "girl",
    "Jen": "girl",
    "Joe": "boy",
    "Zac": "boy"
}

unavailable_dates = {
    "Carla": ["2025-09-24"],
    "Jen": ["2025-09-10"]
}

file_path = r"C:\Users\Soheil\Desktop\projects\NSNTA-Timetable\Mixed Fixture.xlsx"
df = pd.read_excel(file_path, sheet_name='B Res 3')
df['Dates'] = df['Dates'].dt.strftime('%Y-%m-%d')

# Determine if Hadfield is home or away each week
df['Location'] = df['Location'].apply(lambda x: 'Home' if x == team_name else 'Away')

player_stats = {p: {'games': 0, 'home_games': 0} for p in player_list}

def is_available(player, date):
    return date not in unavailable_dates.get(player, [])

def available_players_by_gender(date, gender):
    return [p for p in player_list if genders[p] == gender and is_available(p, date)]

schedule = []

for i, row in df.iterrows():
    date = row['Dates']
    location = row['Location']
    
    available_girls = available_players_by_gender(date, "girl")
    available_boys = available_players_by_gender(date, "boy")

    available_girls.sort(key=lambda p: (player_stats[p]['games'], player_stats[p]['home_games']))
    available_boys.sort(key=lambda p: (player_stats[p]['games'], player_stats[p]['home_games']))

    if len(available_girls) >= 2 and len(available_boys) >= 2:
        group = available_girls[:2] + available_boys[:2]

        for p in group:
            player_stats[p]['games'] += 1
            if location == 'Home':
                player_stats[p]['home_games'] += 1

        schedule.append({
            'Date': date,
            'Players': group,
            'Location': location
        })

# Flatten schedule
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
