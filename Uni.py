import pandas as pd
import folium

# Load your data
file_path = r'C:\Users\shaun\Downloads\Final_Masters_Programs_Shortlist.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Define the correct coordinates mapping
full_location_coordinates = {
    'Hamburg': (53.550341, 10.000654),
    'Stuttgart': (48.781974, 9.177942),
    'Bonn': (50.733346, 7.100659),
    'Berlin': (52.512407, 13.326416),
    'Freiburg': (47.993784, 7.84856),
    'Aachen': (50.778217, 6.060549),
    'Munich': (48.148434, 11.567867),
    'Dresden': (51.028430, 13.728965),
    'Kleve': (51.788453, 6.135309),
    'Siegen': (50.909843, 8.024191),
    'Muenster': (51.960664, 7.626134),
    'Trier': (49.749992, 6.637143),
    'Cologne': (50.936191, 6.957857),
    'Mannheim': (49.482729, 8.463343),
    'Augsburg': (48.366512, 10.894446),
    'Tuebingen': (48.523616, 9.054358),
    'Chemnitz': (50.832260, 12.929600),
    'Leipzig': (51.339695, 12.373075)
}

# Replace incorrect location names
df['Location'] = df['Location'].replace({
    'Hmburg': 'Hamburg',
    'BERLIN': 'Berlin',
    'Seigen': 'Siegen',
    'Köln': 'Cologne',
    'Münster': 'Muenster'
})

# Create a dataframe with unique university locations and their coordinates
locations = df[['University', 'Location']].drop_duplicates()
locations['Coordinates'] = locations['Location'].map(full_location_coordinates)

# Initialize the map centered on Germany
germany_center = [51.1657, 10.4515]
m = folium.Map(location=germany_center, zoom_start=6)

# Add markers for each university location with course and link information
for i, row in locations.dropna(subset=['Coordinates']).iterrows():
    course_info = df[df['University'] == row['University']]
    if not course_info.empty:
        courses_links = '<br>'.join([f"{course['Course']}: <a href='{course['Link']}' target='_blank'>Link</a>" for _, course in course_info.iterrows()])
        popup_text = f"{row['University']} ({row['Location']})<br>{courses_links}"
    else:
        popup_text = f"{row['University']} ({row['Location']})"
    
    folium.Marker(
        location=row['Coordinates'],
        popup=folium.Popup(popup_text, max_width=300),
        tooltip=row['University']
    ).add_to(m)

# Save the map to an HTML file
map_file_path = 'masters_programs_map_with_courses.html'
m.save(map_file_path)

print(f"Map has been saved to {map_file_path}")
