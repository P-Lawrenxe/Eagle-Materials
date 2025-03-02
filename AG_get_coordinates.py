import pandas as pd
from geopy.geocoders import Nominatim
from time import sleep
from tqdm import tqdm
import os

# Print current working directory to verify file location
print(f"Current working directory: {os.getcwd()}")

# List files in directory to confirm the Excel file exists
print("Files in directory:")
for file in os.listdir():
    print(f"  - {file}")

# Use the correct file name with underscore as shown in Explorer
file_path = "RP_Lanes.xlsx"  
try:
    df = pd.read_excel(file_path)
    print(f"Successfully loaded Excel file. Columns: {df.columns.tolist()}")
    print(f"First few rows:\n{df.head()}")
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit(1)

# Initialize the geocoder
geolocator = Nominatim(user_agent="geo_lookup", timeout=15)

# Function to fetch latitude and longitude
def get_lat_lon(city):
    if not isinstance(city, str) or pd.isna(city):
        print(f"Skipping invalid city: {city}")
        return pd.Series([None, None])
    
    try:
        print(f"Geocoding: {city}")
        sleep(1)  # Add delay to avoid hitting rate limits
        location = geolocator.geocode(city)
        if location:
            print(f"Found coordinates for {city}: {location.latitude}, {location.longitude}")
            return pd.Series([location.latitude, location.longitude])
        else:
            print(f"No coordinates found for {city}")
            return pd.Series([None, None])
    except Exception as e:
        print(f"Error geocoding {city}: {e}")
        return pd.Series([None, None])

# Use the correct column name "Ship to City" instead of "Ship from city"
city_column = 'Ship from city'

# Apply function to all cities with a progress bar
print(f"Starting geocoding process for column '{city_column}'...")
tqdm.pandas(desc="Geocoding cities")
df['Lat'] = None
df['Lon'] = None
coordinates = df[city_column].progress_apply(get_lat_lon)
df['Lat'] = coordinates.iloc[:, 0]
df['Lon'] = coordinates.iloc[:, 1]

# Save the results back to Excel
output_file = "RP_cities_with_coordinates.xlsx"
df.to_excel(output_file, index=False)
print(f"✅ Done! The updated file is saved as {output_file}")