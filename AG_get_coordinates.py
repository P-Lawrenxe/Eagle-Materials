import pandas as pd
from geopy.geocoders import Nominatim
from time import sleep
from tqdm import tqdm
import os

# Print current working directory
print(f"Current working directory: {os.getcwd()}")

# List files in directory
print("Files in directory:")
for file in os.listdir():
    print(f"  - {file}")

# Load the Excel file - specifically the "Experiment" sheet
file_path = "23-24_Duke_Volume.xlsx"  
try:
    # Specify the sheet name "Experiment"
    df = pd.read_excel(file_path, sheet_name="Experiment")
    print(f"Successfully loaded 'Experiment' sheet. Columns: {df.columns.tolist()}")
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

# Check if "City, State" column exists
if 'City, State' not in df.columns:
    print(f"ERROR: 'City, State' column not found in the Experiment sheet. Available columns: {df.columns.tolist()}")
    exit(1)

print(f"Starting geocoding process for 'City, State' column...")
tqdm.pandas(desc="Geocoding cities")
df['Lat'] = None
df['Lon'] = None
coordinates = df['City, State'].progress_apply(get_lat_lon)
df['Lat'] = coordinates.iloc[:, 0]
df['Lon'] = coordinates.iloc[:, 1]

# Save the results back to Excel, but only update the "Experiment" sheet
# First read all sheets
with pd.ExcelFile(file_path) as xls:
    all_sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

# Replace the "Experiment" sheet with our updated dataframe
all_sheets["Experiment"] = df

# Write all sheets back to a new Excel file
output_file = "23-24_Duke_Volume_with_coordinates.xlsx"
with pd.ExcelWriter(output_file) as writer:
    for sheet_name, sheet_df in all_sheets.items():
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✅ Done! The updated file is saved as {output_file}")