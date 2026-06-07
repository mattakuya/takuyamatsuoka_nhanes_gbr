import os
from pathlib import Path
import urllib.request
import time
from urllib.error import HTTPError, URLError

# ==========================================
# Configuration
# ==========================================
# Mapping of cycle years to their respective NHANES file name suffixes (2007-2018)
CYCLES = {
    "2007-2008": "E",
    "2009-2010": "F",
    "2011-2012": "G",
    "2013-2014": "H",
    "2015-2016": "I",
    "2017-2018": "J",
}

# List of required datasets based on the analysis design
# Add or remove datasets as necessary.
DATASETS = [
    "DEMO",    # Demographics
    "DPQ",     # Depression
    "BIOPRO",  # Standard Biochemistry
    "BMX",     # Body Measures
    "SMQ",     # Smoking
    "ALQ",     # Alcohol
    "PAQ",     # Physical Activity
    "MCQ",     # Medical Conditions
    "DIQ",     # Diabetes
    "RXQ_RX",  # Prescription Medications
    "CBC",     # Complete Blood Count with 5-part Differential
    "CRP",     # C-Reactive Protein (older cycles)
    "HSCRP",   # High-Sensitivity CRP (newer cycles)
]

# Output directory path
ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = str(ROOT / "data" / "raw")

# Base URL for CDC NHANES data files
# Example: https://wwwn.cdc.gov/Nchs/Data/Nhanes/Public/2007/DataFiles/DEMO_E.xpt
BASE_URL = "https://wwwn.cdc.gov/Nchs/Data/Nhanes/Public/{start_year}/DataFiles/{filename}"

def download_nhanes_data():
    # Ensure the output directory exists
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"Created output directory: {OUTPUT_DIR}")

    total_downloads = len(CYCLES) * len(DATASETS)
    current_download = 0

    for cycle_years, suffix in CYCLES.items():
        cycle_dir = os.path.join(OUTPUT_DIR, cycle_years)
        if not os.path.exists(cycle_dir):
            os.makedirs(cycle_dir)
            
        start_year = cycle_years.split('-')[0]
        for dataset in DATASETS:
            current_download += 1
            
            filename = f"{dataset}_{suffix}.XPT"
            url_filename = f"{dataset}_{suffix}.xpt"
                
            url = BASE_URL.format(start_year=start_year, filename=url_filename)
            filepath = os.path.join(cycle_dir, filename)

            # Skip if the file already exists and is not an HTML error page (e.g., 404 page)
            if os.path.exists(filepath):
                with open(filepath, 'rb') as f:
                    header = f.read(15)
                    # Overwrite if not a valid SAS Transport file (e.g., is HTML)
                    if b"<!DOCTYPE html>" not in header and len(header) > 10:
                        print(f"[{current_download}/{total_downloads}] Skip (valid file already exists): {filepath}")
                        continue

            print(f"[{current_download}/{total_downloads}] Downloading: {filename} ...")
            
            try:
                # Set User-Agent to avoid potential blocks
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as response, open(filepath, 'wb') as out_file:
                    data = response.read()
                    out_file.write(data)
                print(f"  -> Saved successfully: {filepath}")
                
                # Polite delay to avoid overloading server
                time.sleep(1)
                
            except HTTPError as e:
                print(f"  -> Error (HTTP {e.code}): {filename} may not exist in this cycle or has a different name.")
            except URLError as e:
                print(f"  -> Network Error: {e.reason}")
            except Exception as e:
                print(f"  -> Unexpected Error: {e}")

if __name__ == "__main__":
    print("Starting NHANES batch data download...")
    download_nhanes_data()
    print("\nDownload process completed.")
