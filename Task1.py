# Importing necessary libraries
import folium
import pandas as pd

# Loading outlet information from an Excel file
outlet_info = pd.read_excel('Worksheet in Assessment question (006).xlsx',
                            sheet_name='outlet info')

# Removing the rows where latitude or longitude is missing
# 'dropna' is used to ensure that only complete entries with both latitude and longitude are included
outlet_info = outlet_info.dropna(subset=['latitude', 'longitude'])

# Calculating the center of the map based on the average latitude and longitude
# 'mean()' calculates the average value of latitude and longitude to center the map
map_center = [outlet_info['latitude'].mean(), outlet_info['longitude'].mean()]

# Creating a folium map object centered at the calculated location with a zoom level of 7
# 'location' sets the center point of the map and 'zoom_start' controls the initial zoom level
map = folium.Map(location=map_center, zoom_start=7)

# Plot each outlet on the map
for _, outlet in outlet_info.iterrows():
    folium.Marker(
        location=[outlet['latitude'], outlet['longitude']],
        popup=f"Outlet: {outlet['outlet']}\nRepresentative: {outlet['representative']}"
        # 'popup' provides information about each outlet
    ).add_to(map)   # 'add_to(map)' adds the marker to the map

# Save the map as an HTML file
map.save('outlets_map.html')
