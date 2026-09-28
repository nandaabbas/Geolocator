# GeoLocator

A Streamlit app for geocoding industrial facilities (e.g. smelters and refineries) from an Excel list, then reviewing each result on a satellite map before exporting.

## What it does

1. **Geocodes by facility name** using the Google Maps Geocoding API. It searches the facility name (with city and country as context) and can fall back to a secondary name column such as a parent company or group. City-only and region-only matches are rejected, so you don't get false pins in town centres.
2. **Processes in batches**, so large lists can be run, paused and downloaded part-way.
3. **Satellite review.** Step through each result on a satellite map with the original row data alongside. You can move the pin, paste corrected coordinates, and add notes.
4. **Exports to Excel** with all your original columns kept and the geocoding columns (lat/lon, confidence, match source, notes) added on the right.

## Requirements

- Python 3.9+
- A **Google Maps API key** with the Geocoding API enabled:
  1. Go to [console.cloud.google.com](https://console.cloud.google.com)
  2. Create a project, search for **Geocoding API**, and click **Enable**
  3. Go to **Credentials**, then **+ Create Credentials**, then **API Key**
  4. Add a billing account (Google's free monthly credit covers typical use)

The key is entered in the app each time and only held in memory. It is never saved to disk.

## Run it

**Mac / Linux, easiest:**
```bash
./start_mac_linux.sh
```

**Or manually (any OS):**
```bash
pip3 install -r requirements.txt
streamlit run app.py
```

The app opens in your browser at http://localhost:8501. Press `Ctrl + C` in the terminal to stop it.

## Input file

An `.xlsx` file with, at minimum, these columns (the names can be changed in the app's *Column settings*):

| Column | Default name |
|---|---|
| Facility name | `Facility Name` |
| City | `City` |
| Country | `Country` |
| Secondary name (optional) | e.g. `Group Name` |
