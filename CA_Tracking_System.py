"""
Corporate Action Tracking System
Processes daily CA files and creates Excel workbook with upcoming, recent, and archived CAs.
"""

import pandas as pd
import os
import glob
import re
from datetime import datetime, timedelta, date
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# Configuration
INPUT_FOLDER = r"F:\Trade Support\Corporate Actions\CA check\CA Raw file"
OUTPUT_FILE = r"F:\Trade Support\Corporate Actions\CA check\CA Raw file\CA_Tracking.xlsx"

# Date ranges
TODAY = date.today()
NEXT_15_DAYS = TODAY + timedelta(days=15)
LAST_7_DAYS = TODAY - timedelta(days=7)

# Columns to extract (by name from CSV)
REQUIRED_COLUMNS = [
    "Security ID",
    "Security Name",
    "Event Type",
    "Response Status(ELIG)",
    "Client",
    "Reference ID",
    "Action Class",
    "ISIN",
    "Client Deadline Date",
    "Early Deadline Date"
]

# Filter criteria
EXCLUDED_EVENT_TYPES = [
    "OPTIONAL DIVIDEND",
    "CASH DISTRIBUTIONS",
    "DIVIDEND REINVESTMENT               "  # Note: includes trailing spaces
]


def print_progress(message):
    """Print progress message to console"""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")


def show_message(title, message, is_error=False):
    """Show pop-up message box"""
    if is_error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)


def find_latest_file(folder_path):
    """Find the latest file in the folder (CSV, XLSX, or XLS)"""
    print_progress(f"Searching for files in: {folder_path}")
    
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"Input folder not found: {folder_path}")
    
    # Look for CSV, XLSX, and XLS files
    csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
    xlsx_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    xls_files = glob.glob(os.path.join(folder_path, "*.xls"))
    
    all_files = csv_files + xlsx_files + xls_files
    
    if not all_files:
        raise FileNotFoundError(f"No CSV or Excel files found in: {folder_path}")
    
    # Get the most recently modified file
    latest_file = max(all_files, key=os.path.getmtime)
    print_progress(f"Found latest file: {os.path.basename(latest_file)}")
    return latest_file


def load_data(file_path):
    """Load data from CSV or Excel file"""
    print_progress(f"Loading data from: {os.path.basename(file_path)}")
    
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.csv':
        # Read CSV file
        df = pd.read_csv(file_path)
    elif file_ext == '.xlsx':
        # Use openpyxl engine (already imported)
        df = pd.read_excel(file_path, engine='openpyxl')
    elif file_ext == '.xls':
        # Try to read .xls file without specifying engine (pandas may handle it)
        # Note: This may require xlrd, but we'll try without it first
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            raise ValueError(f"Cannot read .xls file without xlrd. Please convert {os.path.basename(file_path)} to .xlsx or .csv format. Error: {str(e)}")
    else:
        raise ValueError(f"Unsupported file format: {file_ext}. Expected .csv, .xls, or .xlsx")
    
    print_progress(f"Loaded {len(df)} rows")
    return df


def extract_columns(df):
    """Extract required columns from dataframe"""
    print_progress("Extracting required columns...")
    
    # Check which columns exist
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")
    
    # Extract required columns
    columns_to_extract = REQUIRED_COLUMNS.copy()
    
    # Also include Comments if it exists in the input data
    if 'Comments' in df.columns:
        columns_to_extract.append('Comments')
    
    output = df[columns_to_extract].copy()
    
    # Remove rows where all key columns are empty
    output = output.dropna(subset=["Security ID", "Security Name"], how='all')
    
    return output


def parse_dates(df):
    """Parse date columns"""
    print_progress("Parsing date columns...")
    
    # Parse Client Deadline Date
    if 'Client Deadline Date' in df.columns:
        # Handle date format: "29 Dec 2050 03:30:00 PM EST" -> take first 11 chars
        df['Client Deadline Date'] = df['Client Deadline Date'].astype(str)
        df['Client Deadline Date'] = pd.to_datetime(
            df['Client Deadline Date'].str[:11], 
            format='%d %b %Y', 
            errors='coerce'
        )
    
    # Parse Early Deadline Date
    if 'Early Deadline Date' in df.columns:
        df['Early Deadline Date'] = df['Early Deadline Date'].astype(str)
        df['Early Deadline Date'] = pd.to_datetime(
            df['Early Deadline Date'].str[:11], 
            format='%d %b %Y', 
            errors='coerce'
        )
    
    return df


def apply_filters(df):
    """Apply all filter criteria"""
    print_progress("Applying filters...")
    
    initial_count = len(df)
    
    # Filter 1: Response Status = "RESPONSE REQUIRED"
    df = df[df['Response Status(ELIG)'].str.contains("RESPONSE REQUIRED", na=False)]
    
    # Filter 2: Client = "CIF"
    df = df[df['Client'].str.contains("CIF", na=False)]
    
    # Filter 3: Client not equal to "nil" or empty
    df = df[df['Client'].notna()]
    df = df[df['Client'] != "nil"]
    df = df[df['Client'] != ""]
    
    # Filter 4: Event Type NOT IN excluded list
    df = df[~df['Event Type'].isin(EXCLUDED_EVENT_TYPES)]
    
    # Filter 5: Action Class not equal to "Mandatory"
    df = df[df['Action Class'] != "Mandatory"]
    
    filtered_count = len(df)
    print_progress(f"Filtered from {initial_count} to {filtered_count} rows")
    
    return df


def determine_deadline_date(row):
    """
    Determine which deadline date to use and return (date, type)
    Logic:
    - If both dates are in next 15 days: use earlier date
    - If Early Deadline is in next 15 days: use Early Deadline
    - If Client Deadline is in next 15 days: use Client Deadline
    - If Early Deadline is past but Client Deadline is future: use Client Deadline
    - If Client Deadline is in last 7 days: use Client Deadline
    """
    client_deadline = row['Client Deadline Date']
    early_deadline = row['Early Deadline Date']
    
    # Convert to date if datetime
    if pd.notna(client_deadline):
        if isinstance(client_deadline, pd.Timestamp):
            client_deadline = client_deadline.date()
        elif isinstance(client_deadline, datetime):
            client_deadline = client_deadline.date()
    
    if pd.notna(early_deadline):
        if isinstance(early_deadline, pd.Timestamp):
            early_deadline = early_deadline.date()
        elif isinstance(early_deadline, datetime):
            early_deadline = early_deadline.date()
    
    # Check if dates are in range
    client_in_next_15 = (pd.notna(client_deadline) and 
                        TODAY <= client_deadline <= NEXT_15_DAYS)
    early_in_next_15 = (pd.notna(early_deadline) and 
                       TODAY <= early_deadline <= NEXT_15_DAYS)
    
    client_in_last_7 = (pd.notna(client_deadline) and 
                       LAST_7_DAYS <= client_deadline < TODAY)
    early_in_last_7 = (pd.notna(early_deadline) and 
                      LAST_7_DAYS <= early_deadline < TODAY)
    
    # Priority: Use earlier date if both are in next 15 days
    if client_in_next_15 and early_in_next_15:
        if early_deadline <= client_deadline:
            return early_deadline, "Early"
        else:
            return client_deadline, "Client"
    
    # Use Early Deadline if in next 15 days
    if early_in_next_15:
        return early_deadline, "Early"
    
    # Use Client Deadline if in next 15 days
    if client_in_next_15:
        return client_deadline, "Client"
    
    # Check for last 7 days (Tab 2)
    if early_in_last_7:
        return early_deadline, "Early"
    
    if client_in_last_7:
        return client_deadline, "Client"
    
    # If Early Deadline is past but Client Deadline is future (beyond 15 days)
    if (pd.notna(early_deadline) and early_deadline < TODAY and 
        pd.notna(client_deadline) and client_deadline > NEXT_15_DAYS):
        return client_deadline, "Client"
    
    # No valid date in range
    return None, None


def generate_tabs_from_archive(archive_df, previous_tab1_df=None):
    """Generate Tab 1 and Tab 2 from archive by applying filters and date criteria"""
    print_progress("Generating tabs from archive...")
    
    if archive_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    
    # First, apply filters to archive (same filters as before)
    filtered_df = apply_filters_to_archive(archive_df)
    
    # Create a dictionary of ALL previous Tab 1 CAs by Reference ID (with their comments)
    # We include ALL CAs from previous tab, then match by Reference ID when generating new tab
    previous_cas_dict = {}
    if previous_tab1_df is not None and not previous_tab1_df.empty:
        if 'Reference ID' in previous_tab1_df.columns and 'Comments' in previous_tab1_df.columns:
            for idx, row in previous_tab1_df.iterrows():
                ref_id = str(row.get('Reference ID', '')).strip()
                comment = row.get('Comments', '')
                
                # Include ALL CAs from previous tab (don't filter by date here)
                if pd.notna(ref_id) and ref_id != '':
                    comment_str = str(comment).strip() if pd.notna(comment) and str(comment).strip() != '' else ''
                    previous_cas_dict[ref_id] = comment_str
            
            comments_with_data = sum(1 for v in previous_cas_dict.values() if v != '')
            print_progress(f"Loaded {len(previous_cas_dict)} CAs from previous Next 15 Days tab")
            print_progress(f"  - {comments_with_data} of these have comments to transfer")
    else:
        print_progress("No previous Next 15 Days tab found - starting fresh")
    
    tab1_data = []
    tab2_data = []
    
    for idx, row in filtered_df.iterrows():
        deadline_date = row.get('Deadline Date', None)
        
        if pd.isna(deadline_date) or deadline_date == "":
            continue
        
        # Convert to date if needed
        if isinstance(deadline_date, str):
            try:
                deadline_date = pd.to_datetime(deadline_date).date()
            except:
                continue
        elif isinstance(deadline_date, pd.Timestamp):
            deadline_date = deadline_date.date()
        elif isinstance(deadline_date, datetime):
            deadline_date = deadline_date.date()
        
        row_data = row.copy()
        
        # Tab 1: Next 15 days
        if TODAY <= deadline_date <= NEXT_15_DAYS:
            ref_id = str(row_data.get('Reference ID', '')).strip()
            
            # Comments only exist in the previous Excel file (they're cleared in new data)
            # So we copy directly from the previous "Next 15 Days" tab
            if ref_id in previous_cas_dict:
                # Previous Next 15 Days tab had this CA - copy its comment
                previous_comment = previous_cas_dict[ref_id]
                row_data['Comments'] = previous_comment if previous_comment else ""
            else:
                # CA not in previous tab - no comment to transfer
                row_data['Comments'] = ""
            
            tab1_data.append(row_data)
        # Tab 2: Last 7 days (auto-remove after 7 days)
        elif LAST_7_DAYS <= deadline_date < TODAY:
            tab2_data.append(row_data)
    
    tab1_df = pd.DataFrame(tab1_data) if tab1_data else pd.DataFrame()
    tab2_df = pd.DataFrame(tab2_data) if tab2_data else pd.DataFrame()
    
    # Count how many comments were transferred
    if not tab1_df.empty:
        comments_transferred = 0
        matched_ref_ids = 0
        unmatched_ref_ids = []
        for idx, row in tab1_df.iterrows():
            ref_id = str(row.get('Reference ID', '')).strip()
            comment = str(row.get('Comments', ''))
            if ref_id in previous_cas_dict:
                matched_ref_ids += 1
                if comment.strip() != '':
                    comments_transferred += 1
            else:
                unmatched_ref_ids.append(ref_id)
        
        print_progress(f"Matched {matched_ref_ids} CAs by Reference ID from previous tab")
        if comments_transferred > 0:
            print_progress(f"Transferred {comments_transferred} comments from previous Next 15 Days tab")
        else:
            print_progress("No comments were transferred (no comments in previous tab for matched CAs)")
        if unmatched_ref_ids and len(unmatched_ref_ids) <= 5:
            print_progress(f"Unmatched Reference IDs (sample): {unmatched_ref_ids[:5]}")
    
    # Sort by deadline date
    if not tab1_df.empty and 'Deadline Date' in tab1_df.columns:
        tab1_df = tab1_df.sort_values(by='Deadline Date')
    if not tab2_df.empty and 'Deadline Date' in tab2_df.columns:
        tab2_df = tab2_df.sort_values(by='Deadline Date')
    
    print_progress(f"Tab 1 (Next 15 Days): {len(tab1_df)} CAs")
    print_progress(f"Tab 2 (Last 7 Days): {len(tab2_df)} CAs")
    
    return tab1_df, tab2_df


def apply_filters_to_archive(df):
    """Apply filter criteria to archive data (for Tab 1 and Tab 2)"""
    # Filter 1: Response Status = "RESPONSE REQUIRED"
    df = df[df['Response Status(ELIG)'].str.contains("RESPONSE REQUIRED", na=False)]
    
    # Filter 2: Client = "CIF"
    df = df[df['Client'].str.contains("CIF", na=False)]
    
    # Filter 3: Client not equal to "nil" or empty
    df = df[df['Client'].notna()]
    df = df[df['Client'] != "nil"]
    df = df[df['Client'] != ""]
    
    # Filter 4: Event Type NOT IN excluded list
    df = df[~df['Event Type'].isin(EXCLUDED_EVENT_TYPES)]
    
    # Filter 5: Action Class not equal to "Mandatory"
    df = df[df['Action Class'] != "Mandatory"]
    
    return df


def prepare_output_columns(df):
    """Prepare columns for output"""
    output_cols = [
        "Security ID",
        "Security Name",
        "Event Type",
        "Response Status(ELIG)",
        "Client",
        "Reference ID",
        "Action Class",
        "ISIN",
        "Deadline Date",
        "Deadline Type",
        "Comments"
    ]
    
    # Ensure all columns exist
    for col in output_cols:
        if col not in df.columns:
            df[col] = ""
    
    # Select and reorder columns
    return df[output_cols]


def find_most_recent_excel(folder_path, exclude_today=True):
    """Find the most recent Excel file in the folder with CA_Tracking_YYYYMMDD_HHMMSS pattern"""
    if not os.path.exists(folder_path):
        return None
    
    # Look for CA_Tracking files with date pattern: CA_Tracking_YYYYMMDD_HHMMSS.xlsx
    # This pattern matches files like: CA_Tracking_20241217_125653.xlsx
    xlsx_files = glob.glob(os.path.join(folder_path, "CA_Tracking_*.xlsx"))
    
    if not xlsx_files:
        # Fallback: also check for CA_Tracking.xlsx (main file without date)
        main_file = os.path.join(folder_path, "CA_Tracking.xlsx")
        if os.path.exists(main_file):
            return main_file
        return None
    
    # Filter to only files with date pattern (YYYYMMDD_HHMMSS) - exclude backup files
    # Pattern: CA_Tracking_ followed by 8 digits, underscore, 6 digits
    date_pattern_files = []
    today_str = TODAY.strftime('%Y%m%d')
    
    for file in xlsx_files:
        filename = os.path.basename(file)
        # Match pattern: CA_Tracking_YYYYMMDD_HHMMSS.xlsx
        if re.match(r'CA_Tracking_\d{8}_\d{6}\.xlsx$', filename):
            # Extract date from filename
            date_match = re.search(r'CA_Tracking_(\d{8})_\d{6}\.xlsx$', filename)
            if date_match:
                file_date = date_match.group(1)
                # Exclude today's files if requested
                if exclude_today and file_date == today_str:
                    continue
            date_pattern_files.append(file)
        # Also include CA_Tracking.xlsx if it exists (and not excluding today)
        elif filename == "CA_Tracking.xlsx" and not exclude_today:
            date_pattern_files.append(file)
    
    if not date_pattern_files:
        return None
    
    # Get the most recently modified file (this will be the most recent run, excluding today)
    most_recent = max(date_pattern_files, key=os.path.getmtime)
    return most_recent


def load_existing_excel(file_path):
    """Load existing Excel file from the most recent run (CA_Tracking_YYYYMMDD_HHMMSS.xlsx)"""
    output_folder = os.path.dirname(file_path)
    
    # Find the most recent file with date pattern (excluding today's files)
    most_recent_file = find_most_recent_excel(output_folder, exclude_today=True)
    
    if not most_recent_file:
        print_progress("No previous CA_Tracking file found - starting fresh")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    file_to_load = most_recent_file
    print_progress(f"Loading previous run's file: {os.path.basename(file_to_load)}")
    
    if not os.path.exists(file_to_load):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        # Load archive tab (needed for merging)
        archive_df = pd.read_excel(file_to_load, sheet_name="Archive", engine='openpyxl')
        print_progress(f"Loaded {len(archive_df)} CAs from archive")
    except Exception as e:
        print_progress(f"Could not load Archive tab: {str(e)}")
        archive_df = pd.DataFrame()
    
    try:
        # Load Tab 1 (Next 15 Days) to get comments - THIS IS WHERE COMMENTS COME FROM
        print_progress(f"*** Loading 'Next 15 Days' sheet (Tab 1) for comment transfer ***")
        tab1_df = pd.read_excel(file_to_load, sheet_name="Next 15 Days", engine='openpyxl')
        print_progress(f"✓ Successfully loaded {len(tab1_df)} CAs from previous Next 15 Days tab (Tab 1) - THIS IS FOR COMMENTS")
        # Ensure Comments column exists
        if 'Comments' not in tab1_df.columns:
            tab1_df['Comments'] = ""
            print_progress("Warning: Comments column not found, created empty column")
        # Ensure Reference ID column exists for matching
        if 'Reference ID' not in tab1_df.columns:
            print_progress("ERROR: Reference ID column not found in previous Next 15 Days tab - cannot match comments!")
        else:
            # Debug: count how many have comments
            comments_count = 0
            if 'Comments' in tab1_df.columns:
                for idx, row in tab1_df.iterrows():
                    comment = row.get('Comments', '')
                    if pd.notna(comment) and str(comment).strip() != '':
                        comments_count += 1
            print_progress(f"✓ Found {comments_count} CAs with comments in previous Next 15 Days tab (ready to transfer)")
    except Exception as e:
        print_progress(f"*** ERROR: Could not load Next 15 Days tab: {str(e)} ***")
        print_progress("*** This means comments CANNOT be transferred! ***")
        tab1_df = pd.DataFrame()
    
    try:
        # Load Tab 2 (Last 7 Days) - this is NOT used for comments, only for other data
        tab2_df = pd.read_excel(file_to_load, sheet_name="Last 7 Days", engine='openpyxl')
        print_progress(f"Loaded {len(tab2_df)} CAs from Tab 2 (Last 7 Days) - NOT used for comment transfer")
    except Exception as e:
        print_progress(f"Could not load Last 7 Days tab: {str(e)}")
        tab2_df = pd.DataFrame()
    
    return archive_df, tab1_df, tab2_df


def data_changed(new_row, existing_row, exclude_cols=['Comments', 'Deadline Date', 'Deadline Type']):
    """Check if data has changed (excluding specified columns)"""
    for col in new_row.index:
        if col in exclude_cols:
            continue
        new_val = new_row[col]
        existing_val = existing_row.get(col, None)
        
        # Handle NaN comparisons
        if pd.isna(new_val) and pd.isna(existing_val):
            continue
        if pd.isna(new_val) or pd.isna(existing_val):
            return True
        if str(new_val) != str(existing_val):
            return True
    return False


def merge_with_archive(new_df, existing_archive_df, previous_tab1_df=None):
    """Merge new data with archive, preserving comments and only updating if data changed"""
    print_progress("Merging with archive...")
    
    # Create a dictionary of previous Tab 1 comments by Reference ID (for fallback)
    previous_tab1_comments = {}
    if previous_tab1_df is not None and not previous_tab1_df.empty:
        if 'Reference ID' in previous_tab1_df.columns and 'Comments' in previous_tab1_df.columns:
            for idx, row in previous_tab1_df.iterrows():
                ref_id = str(row.get('Reference ID', ''))
                comment = row.get('Comments', '')
                if pd.notna(ref_id) and ref_id != '' and pd.notna(comment) and str(comment).strip() != '':
                    previous_tab1_comments[ref_id] = str(comment).strip()
            print_progress(f"Loaded {len(previous_tab1_comments)} comments from previous Next 15 Days tab for archive merge")
    
    if existing_archive_df.empty:
        # First time - all new data goes to archive
        new_df = new_df.copy()
        # Handle Comments: use from input data if provided, otherwise set to empty
        if 'Comments' not in new_df.columns:
            new_df['Comments'] = ""
        else:
            # Ensure Comments column exists and clean up values
            new_df['Comments'] = new_df['Comments'].apply(
                lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != '' else ""
            )
        # Add deadline date and type columns
        deadline_dates = []
        deadline_types = []
        for idx, row in new_df.iterrows():
            deadline_date, deadline_type = determine_deadline_date(row)
            deadline_dates.append(deadline_date if deadline_date else "")
            deadline_types.append(deadline_type if deadline_type else "")
        new_df['Deadline Date'] = deadline_dates
        new_df['Deadline Type'] = deadline_types
        return new_df
    
    # Create archive copy
    archive_df = existing_archive_df.copy()
    
    # Create a dictionary of existing rows by Reference ID
    existing_dict = {}
    if 'Reference ID' in archive_df.columns:
        for idx, row in archive_df.iterrows():
            ref_id = str(row.get('Reference ID', ''))
            if pd.notna(ref_id) and ref_id != '':
                existing_dict[ref_id] = (idx, row)
    
    # Process new data
    added_count = 0
    updated_count = 0
    skipped_count = 0
    comments_merged_from_tab1 = 0
    
    for idx, row in new_df.iterrows():
        ref_id = str(row.get('Reference ID', ''))
        
        if not ref_id or ref_id == '':
            continue
        
        # Calculate deadline date and type for new row
        deadline_date, deadline_type = determine_deadline_date(row)
        row['Deadline Date'] = deadline_date if deadline_date else ""
        row['Deadline Type'] = deadline_type if deadline_type else ""
        
        if ref_id in existing_dict:
            # Reference ID exists - check if data changed
            existing_idx, existing_row = existing_dict[ref_id]
            
            # Handle Comments: use new comment if provided, otherwise preserve old comment
            new_comment = row.get('Comments', '')
            existing_comment = existing_row.get('Comments', '')
            
            # Check if new comment is valid (not empty, not NaN)
            has_new_comment = (pd.notna(new_comment) and 
                              str(new_comment).strip() != '' and 
                              'Comments' in row.index)
            
            if has_new_comment:
                # Use new comment (from input data - highest priority)
                archive_df.at[existing_idx, 'Comments'] = str(new_comment).strip()
            else:
                # No new comment - preserve existing archive comment if it exists
                if pd.notna(existing_comment) and str(existing_comment).strip() != '':
                    archive_df.at[existing_idx, 'Comments'] = existing_comment
                elif ref_id in previous_tab1_comments:
                    # No archive comment, but previous Tab 1 had a comment - use it
                    archive_df.at[existing_idx, 'Comments'] = previous_tab1_comments[ref_id]
                    comments_merged_from_tab1 += 1
                else:
                    # No comment from any source
                    archive_df.at[existing_idx, 'Comments'] = ""
            
            if data_changed(row, existing_row):
                # Data changed - update fields (Comments already handled above)
                for col in row.index:
                    if col != 'Comments':
                        archive_df.at[existing_idx, col] = row[col]
                # Update deadline date and type (recalculate based on new data)
                archive_df.at[existing_idx, 'Deadline Date'] = deadline_date if deadline_date else ""
                archive_df.at[existing_idx, 'Deadline Type'] = deadline_type if deadline_type else ""
                updated_count += 1
            else:
                # No changes - skip update (but still update deadline date in case dates changed)
                archive_df.at[existing_idx, 'Deadline Date'] = deadline_date if deadline_date else ""
                archive_df.at[existing_idx, 'Deadline Type'] = deadline_type if deadline_type else ""
                skipped_count += 1
        else:
            # New CA - add to archive
            new_row = row.copy()
            # Use comment from input data if provided, otherwise set to empty
            if 'Comments' in row.index and pd.notna(row.get('Comments', '')) and str(row.get('Comments', '')).strip() != '':
                new_row['Comments'] = str(row.get('Comments', '')).strip()
            else:
                new_row['Comments'] = ""
            archive_df = pd.concat([archive_df, pd.DataFrame([new_row])], ignore_index=True)
            added_count += 1
    
    print_progress(f"Archive: {added_count} added, {updated_count} updated, {skipped_count} unchanged")
    if comments_merged_from_tab1 > 0:
        print_progress(f"Merged {comments_merged_from_tab1} comments from previous Next 15 Days tab into archive")
    print_progress(f"Archive now contains {len(archive_df)} CAs")
    return archive_df


def format_excel(file_path):
    """Apply formatting to Excel file"""
    print_progress("Applying Excel formatting...")
    
    wb = openpyxl.load_workbook(file_path)
    
    for sheet_name in ["Next 15 Days", "Last 7 Days", "Archive"]:
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        
        # Freeze header row
        ws.freeze_panes = 'A2'
        
        # Format header row
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        # Auto-size columns
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Conditional formatting for Tab 1 (Next 15 Days)
        if sheet_name == "Next 15 Days":
            # Find Deadline Date column
            deadline_col = None
            for idx, cell in enumerate(ws[1], 1):
                if cell.value == "Deadline Date":
                    deadline_col = get_column_letter(idx)
                    break
            
            if deadline_col:
                # Red for < 3 days (urgent)
                red_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
                ws.conditional_formatting.add(
                    f"{deadline_col}2:{deadline_col}{ws.max_row}",
                    CellIsRule(operator='lessThan', formula=[f'TODAY()+3'], fill=red_fill)
                )
                
                # Yellow for 3-7 days (approaching)
                yellow_fill = PatternFill(start_color="FFE66D", end_color="FFE66D", fill_type="solid")
                ws.conditional_formatting.add(
                    f"{deadline_col}2:{deadline_col}{ws.max_row}",
                    CellIsRule(operator='between', formula=[f'TODAY()+3', f'TODAY()+7'], fill=yellow_fill)
                )
                
                # Green for > 7 days (upcoming)
                green_fill = PatternFill(start_color="95E1D3", end_color="95E1D3", fill_type="solid")
                ws.conditional_formatting.add(
                    f"{deadline_col}2:{deadline_col}{ws.max_row}",
                    CellIsRule(operator='greaterThan', formula=[f'TODAY()+7'], fill=green_fill)
                )
    
    wb.save(file_path)
    print_progress("Formatting complete")


def backup_existing_file(file_path):
    """Create backup of existing file"""
    if not os.path.exists(file_path):
        return
    
    print_progress("Creating backup...")
    
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_path = file_path.replace(".xlsx", f"_backup_{timestamp}.xlsx")
    
    import shutil
    shutil.copy2(file_path, backup_path)
    print_progress(f"Backup created: {os.path.basename(backup_path)}")


def save_to_excel(tab1_df, tab2_df, archive_df, file_path):
    """Save dataframes to Excel with multiple tabs"""
    print_progress("Saving to Excel...")
    
    # Prepare dataframes
    if not tab1_df.empty:
        tab1_df = prepare_output_columns(tab1_df)
        tab1_df = tab1_df.sort_values(by='Deadline Date')
    else:
        tab1_df = pd.DataFrame(columns=[
            "Security ID", "Security Name", "Event Type", "Response Status(ELIG)",
            "Client", "Reference ID", "Action Class", "ISIN", "Deadline Date",
            "Deadline Type", "Comments"
        ])
    
    if not tab2_df.empty:
        tab2_df = prepare_output_columns(tab2_df)
        tab2_df = tab2_df.sort_values(by='Deadline Date')
    else:
        tab2_df = pd.DataFrame(columns=[
            "Security ID", "Security Name", "Event Type", "Response Status(ELIG)",
            "Client", "Reference ID", "Action Class", "ISIN", "Deadline Date",
            "Deadline Type", "Comments"
        ])
    
    if not archive_df.empty:
        archive_df = prepare_output_columns(archive_df)
    else:
        archive_df = pd.DataFrame(columns=[
            "Security ID", "Security Name", "Event Type", "Response Status(ELIG)",
            "Client", "Reference ID", "Action Class", "ISIN", "Deadline Date",
            "Deadline Type", "Comments"
        ])
    
    # Format dates for Excel
    for df in [tab1_df, tab2_df, archive_df]:
        if 'Deadline Date' in df.columns:
            df['Deadline Date'] = pd.to_datetime(df['Deadline Date']).dt.strftime('%Y-%m-%d')
    
    # Save to Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        tab1_df.to_excel(writer, sheet_name="Next 15 Days", index=False)
        tab2_df.to_excel(writer, sheet_name="Last 7 Days", index=False)
        archive_df.to_excel(writer, sheet_name="Archive", index=False)
    
    # Apply formatting
    format_excel(file_path)
    
    print_progress(f"Excel file saved: {file_path}")


def main():
    """Main execution function"""
    try:
        print("=" * 60)
        print("Corporate Action Tracking System")
        print("=" * 60)
        print()
        
        # Step 1: Find latest Excel file
        input_file = find_latest_file(INPUT_FOLDER)
        
        # Step 2: Load data (ALL CAs, no filtering yet)
        df = load_data(input_file)
        
        # Step 3: Extract columns
        df = extract_columns(df)
        
        # Step 4: Parse dates
        df = parse_dates(df)
        
        # Step 5: Load existing Excel archive and previous tabs
        existing_archive_df, previous_tab1_df, _ = load_existing_excel(OUTPUT_FILE)
        
        # Step 6: Merge ALL CAs with archive (add new, update if changed, preserve comments)
        # Pass previous Tab 1 to merge comments from previous "Next 15 Days" tab into archive
        archive_df = merge_with_archive(df, existing_archive_df, previous_tab1_df)
        
        # Step 7: Generate Tab 1 and Tab 2 FROM archive (apply filters + date criteria)
        # Pass previous Tab 1 to preserve comments for Next 15 Days CAs
        tab1_df, tab2_df = generate_tabs_from_archive(archive_df, previous_tab1_df)
        
        # Step 10: Backup existing file
        backup_existing_file(OUTPUT_FILE)
        
        # Step 11: Save to Excel
        save_to_excel(tab1_df, tab2_df, archive_df, OUTPUT_FILE)
        
        # Success message
        print()
        print("=" * 60)
        print("SUCCESS!")
        print("=" * 60)
        print(f"Tab 1 (Next 15 Days): {len(tab1_df)} CAs")
        print(f"Tab 2 (Last 7 Days): {len(tab2_df)} CAs")
        print(f"Archive: {len(archive_df)} CAs")
        print(f"Output file: {OUTPUT_FILE}")
        print()
        
        show_message(
            "Success",
            f"CA Tracking updated successfully!\n\n"
            f"Next 15 Days: {len(tab1_df)} CAs\n"
            f"Last 7 Days: {len(tab2_df)} CAs\n"
            f"Archive: {len(archive_df)} CAs\n\n"
            f"File saved to:\n{OUTPUT_FILE}"
        )
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        print_progress(error_msg)
        print()
        import traceback
        traceback.print_exc()
        show_message("Error", error_msg, is_error=True)


if __name__ == "__main__":
    # Initialize tkinter root (hidden) for message boxes
    root = tk.Tk()
    root.withdraw()
    
    main()
    
    root.destroy()

