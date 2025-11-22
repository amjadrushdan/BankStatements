import pandas as pd
import re
import sys
import os

# Try to import camelot and pdfplumber
try:
    import camelot
    CAMELOT_AVAILABLE = True
except ImportError:
    CAMELOT_AVAILABLE = False
    print("Warning: camelot-py not available. Will try pdfplumber instead.")

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False
    print("Warning: pdfplumber not available. Install with: pip install pdfplumber")

# ---- CONFIG ----
statement_folder = "statement_folder"
output_folder = "processed_output"

def generate_output_filename(pdf_path):
    """Generate output filename based on input PDF name."""
    # Get the base filename without extension
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # Extract meaningful parts from filename
    # Pattern examples: "Deposits-statement-1000073282-202508.pdf"
    # Try to extract account number and period
    parts = base_name.split('-')
    
    # Build a clean output name
    # Remove common prefixes like "statement", "Deposits"
    clean_parts = []
    for part in parts:
        if part.lower() not in ['statement', 'deposits', 'deposit']:
            clean_parts.append(part)
    
    if clean_parts:
        output_base = '_'.join(clean_parts)
    else:
        output_base = base_name
    
    # Sanitize filename (remove invalid characters)
    output_base = re.sub(r'[<>:"/\\|?*]', '_', output_base)
    
    return output_base

def clean_amount(value):
    """Clean amount strings and convert to float."""
    if pd.isna(value) or value == "" or value == "-":
        return 0.0
    
    # Convert to string if not already
    value_str = str(value).strip()
    
    # Check for negative sign at the start
    is_negative = value_str.startswith('-')
    
    # Remove RM, commas, and dashes (but preserve negative sign)
    value_str = re.sub(r'RM\s*', '', value_str, flags=re.IGNORECASE)
    value_str = value_str.replace(',', '')
    # Remove dashes but preserve leading negative sign
    if is_negative:
        value_str = '-' + value_str[1:].replace('-', '').replace('–', '').replace('—', '')
    else:
        value_str = value_str.replace('-', '').replace('–', '').replace('—', '')
    value_str = value_str.strip()
    
    # Return 0 if empty, otherwise convert to float
    if value_str == "" or value_str == "–" or value_str == "-":
        return 0.0
    
    try:
        return float(value_str)
    except (ValueError, TypeError):
        return 0.0

def clean_description(desc):
    """Clean description text and remove transaction date lines."""
    if pd.isna(desc):
        return ""
    desc_str = str(desc).strip()
    
    # Remove "Transaction date: XX" patterns (case insensitive)
    # This handles patterns like "Transaction date: 02 Aug 25" or "Transaction date: 05 Aug 25"
    desc_str = re.sub(r'Transaction\s+date\s*:.*?$', '', desc_str, flags=re.IGNORECASE | re.MULTILINE)
    
    # Remove extra whitespace and newlines
    desc_str = re.sub(r'\s+', ' ', desc_str)
    desc_str = desc_str.strip()
    
    return desc_str

def extract_tables_with_pdfplumber(pdf_path):
    """Extract tables using pdfplumber."""
    if not PDFPLUMBER_AVAILABLE:
        return []
    
    all_tables = []
    print("Extracting tables with pdfplumber...")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) > 0:
                        # Convert to DataFrame
                        df = pd.DataFrame(table[1:], columns=table[0] if table else None)
                        all_tables.append(df)
                        print(f"  Found table on page {page_num} with {len(df)} rows")
    except Exception as e:
        print(f"pdfplumber extraction failed: {e}")
        return []
    
    return all_tables

def extract_tables_from_pdf(pdf_path):
    """Extract tables from PDF using multiple methods."""
    tables = []
    dfs = []
    
    # Try camelot methods first if available
    if CAMELOT_AVAILABLE:
        # Try lattice method first (better for tables with clear borders)
        print("Trying camelot lattice method...")
        try:
            tables_lattice = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
            print(f"Lattice method found {len(tables_lattice)} table(s)")
            if len(tables_lattice) > 0:
                tables = tables_lattice
        except Exception as e:
            print(f"Lattice method failed: {e}")
        
        # If lattice didn't work or found no tables, try stream method
        if len(tables) == 0:
            print("Trying camelot stream method...")
            try:
                tables_stream = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
                print(f"Stream method found {len(tables_stream)} table(s)")
                if len(tables_stream) > 0:
                    tables = tables_stream
            except Exception as e:
                print(f"Stream method failed: {e}")
    
    # If camelot didn't work, try pdfplumber
    if len(tables) == 0 and PDFPLUMBER_AVAILABLE:
        print("Falling back to pdfplumber...")
        dfs = extract_tables_with_pdfplumber(pdf_path)
        if len(dfs) > 0:
            print(f"pdfplumber found {len(dfs)} table(s)")
    
    # Return camelot tables or pdfplumber dataframes
    if len(tables) > 0:
        return ('camelot', tables)
    elif len(dfs) > 0:
        return ('pdfplumber', dfs)
    else:
        return (None, [])

def process_pdf(pdf_path):
    """Process a single PDF file and return the cleaned DataFrame."""
    print(f"\n{'='*60}")
    print(f"Processing PDF: {os.path.basename(pdf_path)}")
    print(f"{'='*60}")
    
    # Extract tables from PDF
    method, tables_data = extract_tables_from_pdf(pdf_path)
    
    if method is None or len(tables_data) == 0:
        error_msg = f"No tables found in PDF: {os.path.basename(pdf_path)}. "
        if not CAMELOT_AVAILABLE and not PDFPLUMBER_AVAILABLE:
            error_msg += "Please install either camelot-py (requires Ghostscript) or pdfplumber."
        elif not CAMELOT_AVAILABLE:
            error_msg += "camelot-py is not available. Please install it or ensure pdfplumber works."
        else:
            error_msg += "Please check if the PDF contains tables or if Ghostscript is installed (for camelot)."
        print(f"ERROR: {error_msg}")
        return None
    
    print(f"Found {len(tables_data)} table(s) using {method} method")
    
    df_list = []
    
    for index, table_data in enumerate(tables_data):
        # Handle camelot vs pdfplumber differently
        if method == 'camelot':
            df = table_data.df.copy()
        else:  # pdfplumber
            df = table_data.copy()
        
        # Find the header row (look for "Date" column)
        header_row_idx = None
        
        # Check if first row looks like a header (for pdfplumber, headers are usually first)
        if method == 'pdfplumber' and len(df) > 0:
            first_row_str = ' '.join([str(cell).strip() for cell in df.iloc[0].values])
            if 'Date' in first_row_str and ('Withdrawal' in first_row_str or 'Deposit' in first_row_str or 'Balance' in first_row_str):
                header_row_idx = 0
        
        # If not found in first row, search all rows
        if header_row_idx is None:
            for i, row in df.iterrows():
                row_str = ' '.join([str(cell).strip() for cell in row.values])
                if 'Date' in row_str and ('Withdrawal' in row_str or 'Deposit' in row_str or 'Balance' in row_str):
                    header_row_idx = i
                    break
        
        if header_row_idx is None:
            print(f"Skipping table {index + 1}: No transaction header found")
            continue
        
        # Use the found header row
        if header_row_idx == 0 and method == 'pdfplumber':
            # pdfplumber already has header in first row, just use it
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
        else:
            # For camelot or other cases, set columns and remove header row
            df.columns = df.iloc[header_row_idx]
            df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
        
        # Standardize column names (handle variations)
        column_mapping = {}
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'date' in col_lower:
                column_mapping[col] = 'Date'
            elif 'description' in col_lower or 'desc' in col_lower:
                column_mapping[col] = 'Description'
            elif 'withdrawal' in col_lower or 'withdraw' in col_lower:
                column_mapping[col] = 'Withdrawal'
            elif 'deposit' in col_lower:
                column_mapping[col] = 'Deposit'
            elif 'balance' in col_lower:
                column_mapping[col] = 'Balance'
        
        df = df.rename(columns=column_mapping)
        
        # Check if we have the required columns
        required_cols = ['Date', 'Description', 'Withdrawal', 'Deposit', 'Balance']
        if not all(col in df.columns for col in required_cols):
            print(f"Skipping table {index + 1}: Missing required columns. Found: {list(df.columns)}")
            continue
        
        # Remove header row duplicates
        df = df[df["Date"].astype(str).str.strip() != "Date"]
        
        # Remove rows that are clearly not transactions (like "Important notice" or closing balance rows)
        df = df[~df["Date"].astype(str).str.contains("Important notice|Closing balance", case=False, na=False)]
        
        # Clean whitespaces
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
        
        # Handle multi-line descriptions (rows where Date is empty but Description has content)
        # Convert to list for easier look-ahead processing
        rows_list = df.to_dict('records')
        cleaned_rows = []
        
        i = 0
        while i < len(rows_list):
            row = rows_list[i]
            date_val = str(row.get("Date", "")).strip()
            desc_val = str(row.get("Description", "")).strip()
            withdrawal_val = str(row.get("Withdrawal", "")).strip()
            deposit_val = str(row.get("Deposit", "")).strip()
            balance_val = str(row.get("Balance", "")).strip()
            
            # Skip completely empty rows
            if not date_val and not desc_val and not withdrawal_val and not deposit_val and not balance_val:
                i += 1
                continue
            
            # Handle rows that are "Transaction date:" lines
            if not date_val and desc_val:
                # Check if this is a "Transaction date:" line
                if re.match(r'^\s*Transaction\s+date\s*:', desc_val, re.IGNORECASE):
                    # Extract amounts from this row before skipping it
                    withdrawal_amt = clean_amount(withdrawal_val)
                    deposit_amt = clean_amount(deposit_val)
                    balance_amt = clean_amount(balance_val)
                    
                    # If we have a previous row, merge the amounts into it
                    # But only if the previous row doesn't already have amounts (from look-ahead)
                    if cleaned_rows:
                        prev_row = cleaned_rows[-1]
                        prev_has_amounts = (prev_row.get("Withdrawal", 0.0) != 0.0 or 
                                           prev_row.get("Deposit", 0.0) != 0.0 or 
                                           prev_row.get("Balance", 0.0) != 0.0)
                        
                        # Only update if previous row has no amounts, or if transaction date row has non-zero amounts
                        if not prev_has_amounts or (withdrawal_amt != 0.0 or deposit_amt != 0.0 or balance_amt != 0.0):
                            cleaned_rows[-1]["Withdrawal"] = withdrawal_amt
                            cleaned_rows[-1]["Deposit"] = deposit_amt
                            cleaned_rows[-1]["Balance"] = balance_amt
                    
                    i += 1
                    continue
                
                # If Date is empty but Description has content (and it's not a transaction date line),
                # append to previous row
                if cleaned_rows:
                    cleaned_rows[-1]["Description"] += " " + desc_val
                    # Also check if this continuation line has amounts (use them if present)
                    withdrawal_amt = clean_amount(withdrawal_val)
                    deposit_amt = clean_amount(deposit_val)
                    balance_amt = clean_amount(balance_val)
                    # Update amounts if they exist (continuation lines might have the amounts)
                    if withdrawal_amt != 0.0 or (withdrawal_val and withdrawal_val.strip()):
                        cleaned_rows[-1]["Withdrawal"] = withdrawal_amt
                    if deposit_amt != 0.0 or (deposit_val and deposit_val.strip()):
                        cleaned_rows[-1]["Deposit"] = deposit_amt
                    if balance_amt != 0.0 or (balance_val and balance_val.strip()):
                        cleaned_rows[-1]["Balance"] = balance_amt
                i += 1
                continue
            
            # If this looks like a valid transaction row
            if date_val or desc_val:
                withdrawal_amt = clean_amount(withdrawal_val)
                deposit_amt = clean_amount(deposit_val)
                balance_amt = clean_amount(balance_val)
                
                # Check if amounts are missing (all zeros or empty) - look ahead to next row(s) for amounts
                # Also check if the raw values are empty strings (not just zeros)
                amounts_missing = (withdrawal_amt == 0.0 and deposit_amt == 0.0 and balance_amt == 0.0 and
                                  (not withdrawal_val or not withdrawal_val.strip()) and
                                  (not deposit_val or not deposit_val.strip()) and
                                  (not balance_val or not balance_val.strip()))
                
                if amounts_missing:
                    # Look ahead up to 3 rows to find amounts (increased from 2)
                    found_amounts = False
                    for look_ahead in range(1, 4):
                        if i + look_ahead >= len(rows_list):
                            break
                        
                        next_row = rows_list[i + look_ahead]
                        next_date = str(next_row.get("Date", "")).strip()
                        next_desc = str(next_row.get("Description", "")).strip()
                        next_withdrawal = str(next_row.get("Withdrawal", "")).strip()
                        next_deposit = str(next_row.get("Deposit", "")).strip()
                        next_balance = str(next_row.get("Balance", "")).strip()
                        
                        # Check if this row has amounts
                        next_withdrawal_amt = clean_amount(next_withdrawal)
                        next_deposit_amt = clean_amount(next_deposit)
                        next_balance_amt = clean_amount(next_balance)
                        
                        # If next row is a transaction date line, get amounts from it
                        # Even if amounts are zero, if the row has non-empty values, use them
                        if not next_date and next_desc and re.match(r'^\s*Transaction\s+date\s*:', next_desc, re.IGNORECASE):
                            # Use amounts from transaction date row (even if they're zero, they're the actual values)
                            withdrawal_amt = next_withdrawal_amt
                            deposit_amt = next_deposit_amt
                            balance_amt = next_balance_amt
                            found_amounts = True
                            # Skip the transaction date row since we've used its amounts
                            i += look_ahead
                            break
                        # Or if next row has amounts and no date (continuation line with amounts)
                        # This handles rows where Date and Description are empty but amounts are present
                        elif not next_date and (not next_desc or next_desc == ""):
                            # Check if this row has any non-empty amount values
                            has_amounts = (next_withdrawal_amt != 0.0 or next_deposit_amt != 0.0 or next_balance_amt != 0.0 or
                                         (next_withdrawal and next_withdrawal.strip() and next_withdrawal.strip() != '-') or
                                         (next_deposit and next_deposit.strip() and next_deposit.strip() != '-') or
                                         (next_balance and next_balance.strip() and next_balance.strip() != '-'))
                            if has_amounts:
                                withdrawal_amt = next_withdrawal_amt
                                deposit_amt = next_deposit_amt
                                balance_amt = next_balance_amt
                                found_amounts = True
                                # Skip the row(s) since we've used their amounts
                                i += look_ahead
                                break
                
                cleaned_rows.append({
                    "Date": date_val,
                    "Description": clean_description(desc_val),
                    "Withdrawal": withdrawal_amt,
                    "Deposit": deposit_amt,
                    "Balance": balance_amt
                })
            
            i += 1
        
        if cleaned_rows:
            df_clean = pd.DataFrame(cleaned_rows)
            # Remove rows where both Date and Description are empty
            df_clean = df_clean[(df_clean["Date"] != "") | (df_clean["Description"] != "")]
            # Add a table index and row index to preserve original order for sorting
            df_clean['_table_index'] = index
            df_clean['_row_index'] = range(len(df_clean))
            df_list.append(df_clean)
            print(f"Processed table {index + 1}: {len(df_clean)} transaction rows")
        else:
            print(f"Table {index + 1}: No valid transactions found after cleaning")
    
    # Check if we have any data to concatenate
    if len(df_list) == 0:
        print(f"ERROR: No valid transaction tables found in {os.path.basename(pdf_path)}")
        return None
    
    # Combine all pages
    final_df = pd.concat(df_list, ignore_index=True)
    
    # Final cleanup: Remove any rows that are completely empty or invalid
    final_df = final_df[
        (final_df["Date"].astype(str).str.strip() != "") | 
        (final_df["Description"].astype(str).str.strip() != "")
    ]
    
    # Remove duplicate rows based on all transaction fields
    # This fixes the issue where Page 2 transactions appear twice
    print(f"\nBefore deduplication: {len(final_df)} rows")
    # Create a hash of all transaction fields for deduplication
    final_df['_dedup_key'] = (
        final_df['Date'].astype(str) + '|' +
        final_df['Description'].astype(str) + '|' +
        final_df['Withdrawal'].astype(str) + '|' +
        final_df['Deposit'].astype(str) + '|' +
        final_df['Balance'].astype(str)
    )
    # Keep only the first occurrence of each duplicate
    final_df = final_df.drop_duplicates(subset=['_dedup_key'], keep='first')
    final_df = final_df.drop('_dedup_key', axis=1)
    print(f"After deduplication: {len(final_df)} rows")
    
    # Sort by date and then by original order to maintain balance continuity
    try:
        # Try to convert dates and sort
        # Format: "01 Aug 25" -> "%d %b %y"
        final_df['Date_Parsed'] = pd.to_datetime(final_df['Date'], format='%d %b %y', errors='coerce', dayfirst=True)
        # Sort by date first, then by table index, then by row index within table
        # This preserves the original PDF order within each date
        final_df = final_df.sort_values(['Date_Parsed', '_table_index', '_row_index'], na_position='last')
        final_df = final_df.drop('Date_Parsed', axis=1)
    except Exception as e:
        print(f"Warning: Date parsing failed, using fallback sorting: {e}")
        # Fallback: sort by table index and row index to preserve order
        final_df = final_df.sort_values(['_table_index', '_row_index'])
    
    # Remove the temporary sorting columns
    final_df = final_df.drop(['_table_index', '_row_index'], axis=1)
    
    # Reset index
    final_df = final_df.reset_index(drop=True)

    # ---- SAVE OUTPUT ----
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Generate output filenames based on PDF name
    output_base = generate_output_filename(pdf_path)
    output_excel = os.path.join(output_folder, f"{output_base}.xlsx")
    output_csv = os.path.join(output_folder, f"{output_base}.csv")
    
    print(f"\n{'='*60}")
    print(f"Extraction complete for: {os.path.basename(pdf_path)}")
    print(f"Total transactions: {len(final_df)}")
    print(f"\nFirst few rows:")
    print(final_df.head(10).to_string())
    print(f"\n{'='*60}")
    
    final_df.to_excel(output_excel, index=False)
    final_df.to_csv(output_csv, index=False)
    
    print(f"\nSaved to:")
    print(f"  - {output_excel}")
    print(f"  - {output_csv}")
    
    return final_df

# ---- MAIN PROCESSING LOOP ----
# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Find all PDF files in the statement folder
pdf_files = []
if os.path.exists(statement_folder):
    for file in os.listdir(statement_folder):
        if file.lower().endswith('.pdf'):
            pdf_files.append(os.path.join(statement_folder, file))
else:
    print(f"ERROR: Statement folder '{statement_folder}' does not exist!")
    sys.exit(1)

if len(pdf_files) == 0:
    print(f"ERROR: No PDF files found in '{statement_folder}' folder!")
    sys.exit(1)

print(f"\n{'='*60}")
print(f"Found {len(pdf_files)} PDF file(s) to process")
print(f"{'='*60}")

# Process each PDF file
processed_count = 0
failed_count = 0

for pdf_path in pdf_files:
    try:
        result = process_pdf(pdf_path)
        if result is not None:
            processed_count += 1
        else:
            failed_count += 1
    except Exception as e:
        print(f"\nERROR processing {os.path.basename(pdf_path)}: {e}")
        failed_count += 1
        import traceback
        traceback.print_exc()

# Summary
print(f"\n{'='*60}")
print(f"PROCESSING SUMMARY")
print(f"{'='*60}")
print(f"Total PDFs found: {len(pdf_files)}")
print(f"Successfully processed: {processed_count}")
print(f"Failed: {failed_count}")
print(f"{'='*60}")
