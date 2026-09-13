import os
import sys
import zipfile
import pandas as pd
import warnings
from datetime import datetime
from ftplib import FTP_TLS
from os import makedirs
from os.path import join, exists
from pathlib import Path
import re

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import env_int, env_path_str, env_str

warnings.simplefilter("ignore", UserWarning)

# ----------------------------------------------------------------------
# FTPS Configuration - shares the same FTP_HOST/FTP_USERNAME/FTP_PASSWORD
# used by reports/subscriber_reports.py (same server), but a different
# port/protocol (explicit-TLS FTP on 21, not the SFTP-on-22 used there).
# ----------------------------------------------------------------------
FTPS_HOST = env_str("FTP_HOST")
FTPS_PORT = env_int("CELL_INFO_FTP_PORT", 21)
FTPS_USER = env_str("FTP_USERNAME")
FTPS_PASS = env_str("FTP_PASSWORD")

if not FTPS_HOST or not FTPS_USER or not FTPS_PASS:
    raise RuntimeError("FTP_HOST, FTP_USERNAME, and FTP_PASSWORD must be configured in .env or environment variables.")

# Remote directories
REMOTE_DIR_MAIN = env_str("CELL_INFO_REMOTE_DIR_MAIN", "/ftproot/New")  # 2G,3G,4G files
REMOTE_DIR_EPT = env_str("CELL_INFO_REMOTE_DIR_EPT", "/ftproot/RF DATA")  # EPT file

# Exact prefix for EPT
EPT_PREFIX = 'Libyana MS EPT_'

# Base output folder (data lives outside the project folder, like every
# other report here - see README's "never commit generated reports" note)
base_folder = env_path_str("CELL_INFO_OUTPUT_DIR", os.path.join(
    os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data"), "Output", "Cell_Info"
))
os.makedirs(base_folder, exist_ok=True)

# Month-specific subfolder
current_year_month = datetime.now().strftime('%Y-%m')
month_folder = os.path.join(base_folder, current_year_month)
os.makedirs(month_folder, exist_ok=True)

# After download, we set folder_path to the month folder for processing
folder_path = month_folder

# Required patterns for checking after download (used for raw files only)
REQUIRED_PATTERNS = ['2G', '3G', '4G', 'EPT']


# ----------------------------------------------------------------------
# Helper functions for FTPS
# ----------------------------------------------------------------------
def get_timestamp_from_filename(fname):
    """Extract the last numeric sequence of >=14 digits from a filename."""
    numbers = re.findall(r'\d{14,}', fname)
    if numbers:
        ts_str = numbers[-1][:14]  # first 14 digits (YYYYMMDDHHMMSS)
        try:
            return datetime.strptime(ts_str, "%Y%m%d%H%M%S")
        except:
            return None
    return None


def get_date_from_ept_filename(fname):
    """Extract date from EPT filename like 'Libyana MS EPT_v16032026-Whole Network.xlsx'."""
    match = re.search(r'v(\d{8})', fname)
    if match:
        try:
            return datetime.strptime(match.group(1), "%d%m%Y")
        except:
            return None
    return None


def get_latest_tech_file_for_month(ftp, tech, target_month):
    """Return the newest .zip file (by timestamp) for the given tech in the given month."""
    all_files = ftp.nlst()
    matching = []
    for f in all_files:
        if 'Libyana Cell info' in f and tech in f and f.endswith('.zip'):
            ts = get_timestamp_from_filename(f)
            if ts and ts.strftime('%Y-%m') == target_month:
                matching.append((f, ts))
    if not matching:
        return None
    matching.sort(key=lambda x: x[1], reverse=True)
    return matching[0][0]


def get_latest_ept_file(ftp):
    """Return the newest EPT file (by date in filename)."""
    all_files = ftp.nlst()
    matching = []
    for f in all_files:
        if f.startswith(EPT_PREFIX) and f.endswith('.xlsx'):
            date_obj = get_date_from_ept_filename(f)
            if date_obj:
                matching.append((f, date_obj))
    if not matching:
        return None
    matching.sort(key=lambda x: x[1], reverse=True)
    return matching[0][0]


def download_file(ftp, remote_file, local_path):
    """Download a file if it doesn't exist locally."""
    if exists(local_path):
        print(f"  {remote_file} already exists, skipping.")
        return False
    print(f"Downloading {remote_file} ...")
    with open(local_path, 'wb') as f:
        ftp.retrbinary(f'RETR {remote_file}', f.write)
    return True


def download_from_ftps():
    """Connect to FTPS and download the latest files for the current month."""
    try:
        ftp = FTP_TLS()
        ftp.connect(FTPS_HOST, FTPS_PORT)
        ftp.login(FTPS_USER, FTPS_PASS)
        ftp.prot_p()

        # --- 2G, 3G, 4G files in REMOTE_DIR_MAIN ---
        ftp.cwd(REMOTE_DIR_MAIN)
        for tech in ['2G', '3G', '4G']:
            latest = get_latest_tech_file_for_month(ftp, tech, current_year_month)
            if latest:
                local_file = join(month_folder, latest)
                download_file(ftp, latest, local_file)
            else:
                print(f"No {tech} file for {current_year_month} found.")

        # --- EPT file in REMOTE_DIR_EPT ---
        ftp.cwd(REMOTE_DIR_EPT)
        latest_ept = get_latest_ept_file(ftp)
        if latest_ept:
            local_ept = join(month_folder, latest_ept)
            download_file(ftp, latest_ept, local_ept)
        else:
            print("No EPT file found.")

        ftp.quit()
        print("FTPS download completed.")
    except Exception as e:
        print(f"FTPS error: {e}. Proceeding with existing local files.")


def extract_files():
    """Extract all .zip files in the month folder."""
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".zip"):
            zip_path = os.path.join(folder_path, file_name)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(folder_path)
            print(f"Extracted: {file_name}")


# ----------------------------------------------------------------------
# New functions to get raw files (avoid picking processed outputs)
# ----------------------------------------------------------------------
def get_raw_cell_file(tech, folder):
    """Return path to raw 2G/3G/4G Excel file extracted from FTPS."""
    for f in os.listdir(folder):
        if f.startswith('Libyana Cell info') and tech in f and f.endswith(('.xls', '.xlsx')):
            return os.path.join(folder, f)
    return None


def get_raw_ept_file(folder):
    """Return path to raw EPT Excel file."""
    for f in os.listdir(folder):
        if f.startswith('Libyana MS EPT_') and f.endswith('.xlsx'):
            return os.path.join(folder, f)
    return None


def check_required_files():
    """Verify that all required raw files exist. Exit if any missing."""
    missing = []
    for pattern in ['2G', '3G', '4G']:
        if not get_raw_cell_file(pattern, folder_path):
            missing.append(pattern)
    if not get_raw_ept_file(folder_path):
        missing.append('EPT')
    if missing:
        print(f"ERROR: Required raw file(s) missing for: {missing}")
        print("Please check FTPS connection and remote paths.")
        exit(1)


# ----------------------------------------------------------------------
# Main execution: download, extract, then process
# ----------------------------------------------------------------------
print("=== Downloading raw data from FTPS for month:", current_year_month, "===")
download_from_ftps()

print("=== Extracting zip files in", folder_path, "===")
extract_files()

# After extraction, check that all required raw files are present
check_required_files()

# ----------------------------------------------------------------------
# Column definitions for output
# ----------------------------------------------------------------------
columns = [
    "site_name", "Cell_name", "Node B Id", "CI_DEC", "CI_HEX", "LAC_DEC", "LAC_HEX",
    "E_State", "A_State", "2/3G Voice traffic_(Erl)or4G data traffic (TB)", "Latitude",
    "longitude", "Antenna Azimuth", "GCI/eGCI", "band", "Area Name",
    "Network Type", "ECI"
]


def p():
    print("**" * 8)


p()
print("Starting processing of 2G, 3G, 4G cell info...")
p()


# ----------------------------------------------------------------------
# 2G Processing
# ----------------------------------------------------------------------
def cell_info_2G():
    p()
    df = pd.DataFrame(columns=columns)
    file_path_2G_raw = get_raw_cell_file('2G', folder_path)
    if not file_path_2G_raw:
        raise FileNotFoundError("No raw 2G file found.")
    df_raw_2G = pd.read_excel(file_path_2G_raw, engine="openpyxl")
    print("\nRaw 2G file columns:", df_raw_2G.columns.tolist())

    # Map raw columns to our output format (adjust names as needed)
    df["site_name"] = df_raw_2G["Site Name"]
    df["Cell_name"] = df_raw_2G["Cell Name"]
    # Manual correction for specific cells
    df.loc[df["Cell_name"] == "DBRYQ007-1", "Cell_name"] = "DCSBRYQ007-1"
    df.loc[df["Cell_name"] == "DBRYQ007-2", "Cell_name"] = "DCSBRYQ007-2"
    df.loc[df["Cell_name"] == "DBRYQ007-3", "Cell_name"] = "DCSBRYQ007-3"

    df["Node B Id"] = df_raw_2G["Cell CI"].astype(str).str[:-1]
    df["CI_DEC"] = df_raw_2G["Cell CI"]
    df["LAC_DEC"] = df_raw_2G["Cell LAC"]
    df["band"] = df_raw_2G["DL frequency"]
    df["2/3G Voice traffic_(Erl)or4G data traffic (TB)"] = df_raw_2G["K3014:Traffic Volume on TCH(Erl)"]
    df["Area Name"] = "East"
    df["Network Type"] = "2G"

    # Merge with EPT data
    ept_path = get_raw_ept_file(folder_path)
    if not ept_path:
        raise FileNotFoundError("No EPT file found for 2G merging.")
    df_ept = pd.read_excel(ept_path, engine="openpyxl")
    df_ept = df_ept.rename(columns={
        "Longitude": "longitude_ept",
        "Latitude": "Latitude_ept",
        "Azimuth": "Azimuth_ept"
    })

    df = df.merge(
        df_ept[["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"]],
        left_on="Cell_name",
        right_on="Cell Name",
        how="left"
    )

    df["Latitude"] = df["Latitude_ept"]
    df["longitude"] = df["longitude_ept"]
    df["Antenna Azimuth"] = df["Azimuth_ept"]

    # Convert to HEX
    df["CI_HEX"] = df["CI_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)
    df["LAC_HEX"] = df["LAC_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)

    # Ensure integer columns for concatenation
    df["LAC_DEC"] = df["LAC_DEC"].fillna(0).astype(int)
    df["CI_DEC"] = df["CI_DEC"].fillna(0).astype(int)
    df["ECI"] = df["ECI"].fillna(0).astype(int)

    df["GCI/eGCI"] = (
            "60600_"
            + df["LAC_DEC"].astype(str)
            + "_"
            + df.apply(lambda row: str(row["ECI"]) if row["Network Type"] == "4G" else str(row["CI_DEC"]),
                       axis=1).astype(str)
    )

    df = df.drop(columns=["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"])
    output_path = os.path.join(folder_path, "2G_cell_info.xlsx")
    df.to_excel(output_path, index=False)
    print(f"2G processing completed. Output saved to {output_path}")
    return df


df_2G = cell_info_2G()
p()


# ----------------------------------------------------------------------
# 3G Processing
# ----------------------------------------------------------------------
def cell_info_3G():
    p()
    df = pd.DataFrame(columns=columns)
    file_path_3G_raw = get_raw_cell_file('3G', folder_path)
    if not file_path_3G_raw:
        raise FileNotFoundError("No raw 3G file found.")
    df_raw_3G = pd.read_excel(file_path_3G_raw, engine="openpyxl")
    print("\nRaw 3G file columns:", df_raw_3G.columns.tolist())

    df["site_name"] = df_raw_3G["NODEBNAME"].astype(str).str[1:]
    df["Cell_name"] = df_raw_3G["Cell Name"]
    df["Node B Id"] = df_raw_3G["NodeB ID"]
    df["CI_DEC"] = df_raw_3G["Cell ID"]
    df["LAC_DEC"] = df_raw_3G["Location Area Code"]

    band_mapping = {"BAND8": "UMTS900", "BAND1": "UMTS2100"}
    df["band"] = df_raw_3G["Band Indicator"].map(band_mapping)
    df["2/3G Voice traffic_(Erl)or4G data traffic (TB)"] = df_raw_3G["CS Traffic(Erl)"]
    df["Area Name"] = "East"
    df["Network Type"] = "3G"

    df["CI_HEX"] = df["CI_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)
    df["LAC_HEX"] = df["LAC_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)

    df["LAC_DEC"] = df["LAC_DEC"].fillna(0).astype(int)
    df["CI_DEC"] = df["CI_DEC"].fillna(0).astype(int)
    df["ECI"] = df["ECI"].fillna(0).astype(int)

    df["GCI/eGCI"] = (
            "60600_"
            + df["LAC_DEC"].astype(str)
            + "_"
            + df.apply(lambda row: str(row["ECI"]) if row["Network Type"] == "4G" else str(row["CI_DEC"]),
                       axis=1).astype(str)
    )

    ept_path = get_raw_ept_file(folder_path)
    if not ept_path:
        raise FileNotFoundError("No EPT file found for 3G merging.")
    # 3G uses sheet_name="UMTS"
    df_ept = pd.read_excel(ept_path, sheet_name="UMTS", engine="openpyxl")
    df_ept = df_ept.rename(columns={
        "Longitude": "longitude_ept",
        "Latitude": "Latitude_ept",
        "Azimuth": "Azimuth_ept"
    })
    df = df.merge(
        df_ept[["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"]],
        left_on="Cell_name",
        right_on="Cell Name",
        how="left"
    )

    df["Latitude"] = df["Latitude_ept"]
    df["longitude"] = df["longitude_ept"]
    df["Antenna Azimuth"] = df["Azimuth_ept"]
    df = df.drop(columns=["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"])

    output_path = os.path.join(folder_path, "3G_cell_info.xlsx")
    df.to_excel(output_path, index=False)
    print(f"3G processing completed. Output saved to {output_path}")
    return df


df_3G = cell_info_3G()
p()


# ----------------------------------------------------------------------
# 4G Processing
# ----------------------------------------------------------------------
def cell_info_4G():
    p()
    df = pd.DataFrame(columns=columns)
    file_path_4G_raw = get_raw_cell_file('4G', folder_path)
    if not file_path_4G_raw:
        raise FileNotFoundError("No raw 4G file found.")
    df_raw_4G = pd.read_excel(file_path_4G_raw, engine="openpyxl")
    print("\nRaw 4G file columns:", df_raw_4G.columns.tolist())

    LTE_BAND_MAP = {1: "LTE2100", 3: "LTE1800", 8: "LTE900", 28: "LTE700"}
    df["site_name"] = df_raw_4G["eNodeB Function Name"].astype(str).str[1:]
    df["Cell_name"] = df_raw_4G["Cell Name"]
    df["Node B Id"] = df_raw_4G["eNodeB identity"]
    df["CI_DEC"] = df_raw_4G["Cell ID"]
    df["LAC_DEC"] = df_raw_4G["Tracking area code"]
    df["band"] = df_raw_4G[" Frequency band"].map(LTE_BAND_MAP)
    df["2/3G Voice traffic_(Erl)or4G data traffic (TB)"] = (
                                                                   df_raw_4G["Downlink Traffic Volume(GB)"] + df_raw_4G[
                                                               "Uplink Traffic Volume (GB)"]
                                                           ) / 1024

    df["CI_HEX"] = df["CI_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)
    df["LAC_HEX"] = df["LAC_DEC"].apply(lambda x: format(int(x), "X") if pd.notnull(x) else None)
    df["ECI"] = (256 * df["Node B Id"]) + df["CI_DEC"]
    df["Area Name"] = "East"
    df["Network Type"] = "4G"

    df["LAC_DEC"] = df["LAC_DEC"].fillna(0).astype(int)
    df["CI_DEC"] = df["CI_DEC"].fillna(0).astype(int)
    df["ECI"] = df["ECI"].fillna(0).astype(int)

    df["GCI/eGCI"] = (
            "60600_"
            + df["LAC_DEC"].astype(str)
            + "_"
            + df.apply(lambda row: str(row["ECI"]) if row["Network Type"] == "4G" else str(row["CI_DEC"]),
                       axis=1).astype(str)
    )

    ept_path = get_raw_ept_file(folder_path)
    if not ept_path:
        raise FileNotFoundError("No EPT file found for 4G merging.")
    # 4G uses sheet_name="LTE"
    df_ept = pd.read_excel(ept_path, sheet_name="LTE", engine="openpyxl")
    df_ept.loc[df_ept["Site Name"].str.contains("NWH", na=False), "Site Name"] = "NWH Site"
    print("NWH Site updated:", df_ept[df_ept["Site Name"] == "NWH Site"]["Site Name"].count())

    df_ept = df_ept.rename(columns={
        "Longitude": "longitude_ept",
        "Latitude": "Latitude_ept",
        "Azimuth": "Azimuth_ept"
    })

    df = df.merge(
        df_ept[["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"]],
        left_on="Cell_name",
        right_on="Cell Name",
        how="left"
    )
    df["Latitude"] = df["Latitude_ept"]
    df["longitude"] = df["longitude_ept"]
    df["Antenna Azimuth"] = df["Azimuth_ept"]
    df = df.drop(columns=["Cell Name", "longitude_ept", "Latitude_ept", "Azimuth_ept"])

    output_path = os.path.join(folder_path, "4G_cell_info.xlsx")
    df.to_excel(output_path, index=False)
    print(f"4G processing completed. Output saved to {output_path}")
    return df


df_4G = cell_info_4G()
print("\ndf_4G sample:")
print(df_4G.head())
df_4G.to_excel(os.path.join(folder_path, "df_4G_nwh.xlsx"), index=False)

# ----------------------------------------------------------------------
# Merge all RATs
# ----------------------------------------------------------------------
merged_df = pd.concat([df_2G, df_3G, df_4G], ignore_index=True)
merged_df = merged_df[merged_df["2/3G Voice traffic_(Erl)or4G data traffic (TB)"] != 0]
merged_df = merged_df[~merged_df["Cell_name"].str.contains("Trail", na=False)]

# Special handling for stadium site
merged_df.loc[merged_df["Cell_name"].str.contains("stadium", case=False, na=False), "site_name"] = "BGZStadium"
merged_df.rename(columns={"longitude": "Longitude"}, inplace=True)
merged_df.loc[merged_df["site_name"] == "BGZStadium", ["Latitude", "Longitude"]] = [32.10098124, 20.07180587]

output_merged = os.path.join(folder_path, f"merged_df_{datetime.now().strftime('%Y-%m')}.xlsx")
merged_df.to_excel(output_merged, index=False)
print(f"\nAll processing completed successfully. Merged file saved to {output_merged}")
