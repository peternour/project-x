import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# ========================================
# USER CONFIGURATION - MODIFY THESE VALUES
# ========================================

# 1. Folder containing your Excel files
SOURCE_FOLDER = r"C:\Users\Peter\Desktop\New folder"  # Use raw string or double backslashes

# 2. Output file path and name
OUTPUT_FILE = r"D:\Committee\combined_data.xlsx"

# 3. Columns to extract from source files (must exist in your Excel files)
SOURCE_COLUMNS = ["ID NUMBER", "CLIENT NAME"]  # Case-sensitive

# 4. New column names for output file (same order as SOURCE_COLUMNS)
TARGET_COLUMNS = ["ID NUMBER", "CLIENT NAME"]  # Can be same or different

# 5. Ignore blank rows (True/False) - skips rows where all extracted columns are blank
IGNORE_BLANK_ROWS = True

# 6. Ignore blank cells (True/False) - replaces blank cells with a placeholder
IGNORE_BLANK_CELLS = True
BLANK_PLACEHOLDER = ""  # What to put in blank cells (e.g., "", "N/A", or None)

# ========================================
# MAIN PROCESSING SCRIPT (DO NOT MODIFY)
# ========================================

def is_blank(value):
    """Check if a value is considered blank"""
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False

def main():
    print("Starting Excel data consolidation...")
    print(f"Source folder: {SOURCE_FOLDER}")
    print(f"Output file: {OUTPUT_FILE}")
    print(f"Source columns: {SOURCE_COLUMNS}")
    print(f"Target columns: {TARGET_COLUMNS}")
    print(f"Ignore blank rows: {IGNORE_BLANK_ROWS}")
    print(f"Ignore blank cells: {IGNORE_BLANK_CELLS}")
    print("-" * 60)
    
    # Validate configuration
    if not os.path.exists(SOURCE_FOLDER):
        messagebox.showerror("Error", f"Source folder not found:\n{SOURCE_FOLDER}")
        return
        
    if len(SOURCE_COLUMNS) != len(TARGET_COLUMNS):
        messagebox.showerror("Error", "Number of source columns must match target columns")
        return
        
    # Create empty combined DataFrame
    combined_data = pd.DataFrame(columns=TARGET_COLUMNS + ["Source File"])
    processed_files = 0
    error_files = 0
    total_rows = 0
    
    # Process each Excel file
    for filename in os.listdir(SOURCE_FOLDER):
        if filename.lower().endswith(('.xlsx', '.xls')):
            file_path = os.path.join(SOURCE_FOLDER, filename)
            try:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Create temporary DataFrame for current file
                temp_df = pd.DataFrame()
                
                # Map columns
                for src_col, tgt_col in zip(SOURCE_COLUMNS, TARGET_COLUMNS):
                    if src_col in df.columns:
                        temp_df[tgt_col] = df[src_col]
                    else:
                        # Fill missing columns with blank placeholder
                        temp_df[tgt_col] = BLANK_PLACEHOLDER
                
                # Handle blank cells
                if IGNORE_BLANK_CELLS:
                    temp_df = temp_df.map(lambda x: BLANK_PLACEHOLDER if is_blank(x) else x)
                
                # Add source file information
                temp_df['Source File'] = filename
                
                # Remove blank rows if requested
                if IGNORE_BLANK_ROWS:
                    # Create a filter for rows with at least one non-blank value in target columns
                    non_blank_filter = temp_df[TARGET_COLUMNS].apply(
                        lambda row: any(not is_blank(x) for x in row), axis=1
                    )
                    temp_df = temp_df[non_blank_filter]
                
                # Append to combined data
                combined_data = pd.concat([combined_data, temp_df], ignore_index=True)
                processed_files += 1
                total_rows += len(temp_df)
                print(f"Processed: {filename} ({len(temp_df)} rows)")
                
            except Exception as e:
                error_files += 1
                print(f"Error processing {filename}: {str(e)}")
    
    # Check if any files were processed
    if processed_files == 0:
        messagebox.showerror("Error", "No Excel files processed. Check:\n"
                              "1. Source folder path\n"
                              "2. File extensions (.xlsx/.xls)\n"
                              "3. Column names match your Excel files")
        return
    
    # Create output directory if needed
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    
    # Save combined data
    try:
        combined_data.to_excel(OUTPUT_FILE, index=False)
        result_message = (
            f"Peter its all Done Successfully: {OUTPUT_FILE}\n\n"
            f"Files processed: {processed_files}\n"
            f"Files with errors: {error_files}\n"
            f"Total records: {total_rows}"
        )
        print("\n" + result_message.replace("\n", " - "))
        messagebox.showinfo("Success", result_message)
        
    except Exception as e:
        messagebox.showerror("Output Error", f"Failed to create output file:\n{str(e)}")

if __name__ == "__main__":
    # Hide the main Tkinter window
    root = tk.Tk()
    root.withdraw()
    
    # Run the main process
    main()
    
    # Keep console open after completion
    print("\nExecution complete. Press Enter to exit...")
    input()