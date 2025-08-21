# file_processor.py
import pandas as pd
import re
from io import BytesIO
import datetime
from colorcode import apply_color_to_excel
pn_dict = {}
#temprow=r"F:\CAMERON\temp\lufty 07.03.2025\lufty 10.03.2025\List Coder.csv"
#templufty=r"F:\CAMERON\temp\LIST TEMPS AFTER 17.10.2024\10.03.2025\2\Book1.xlsx"


def process_files(row_stream, lufthansa_stream,coder_value):
    """
    This function replicates the overall flow of your 'route()' VBA sub:
      1. Read the row_file (CSV) and lufthansa_file (Excel) into DataFrames.
      2. Perform rename_heading_insert_columns, etc. on row_file DF.
      3. Perform modify_for_list_coder_remove_non_alpha_numeric, etc. on lufthansa DF.
      4. Compare & update (heart_matching_files.CompareAndUpdateColumnscheck).
      5. Separate matched & unmatched, mark rows based on date, etc.
      6. Return the final Excel as an in-memory BytesIO.
    """

    # -----------------------
    # 1) Read the input files
    # -----------------------
    # row_file is a CSV
    print(coder_value)
    df_row = pd.read_csv(row_stream)

    # lufthansa_file is an Excel
    df_lufthansa = pd.read_excel(lufthansa_stream)

    # -----------------------------------------------------
    # 2) Transform the Row File (mimicking your VBA macros)
    # -----------------------------------------------------

    # rename_heading_insert_columns
    df_row = rename_heading_insert_columns(df_row)

    # for_list_coder_remove_non_aplha_and_put_key_in_column_z
    df_row, row_pn_dict = put_keys_in_column_z_and_save_dict(df_row)

    # for_list_coder_remove_non_aplh_from_row
    df_row = remove_non_alphanumeric_column_f(df_row)

    # --------------------------------------------------------
    # 3) Transform the Lufthansa File (mimicking your VBA subs)
    # --------------------------------------------------------

    # modify_for_list_coder_remove_non_alpha_numeric
    df_lufthansa = remove_non_alphanumeric_column_a(df_lufthansa)

    # changes_for_list_coder (cut col A -> insert at col G, fill col I, etc.)
    df_lufthansa = changes_for_list_coder(df_lufthansa, coder_value) 
    # If you actually have coder.Worksheets(1).Range("d8").Value, pass it in properly.

    # SeparateDataByComma
    df_lufthansa = separate_data_by_comma(df_lufthansa)
    



    # -----------------------------------------
    # 4) Compare & update columns (heart step).
    # -----------------------------------------
    df_row = compare_and_update_columns_check(df_row, df_lufthansa)

    # -------------------------------------------------------------------
    # 5) Separate matched/unmatched, mark date-based coloring, etc. 
    # -------------------------------------------------------------------
    matched_df = separate_match_unmatch(df_row)
    matched_df = mark_rows_based_on_date(matched_df)


    # Restore original PN values in column F (index 5) using the dictionary
    matched_df.iloc[:, 5] = matched_df['Z'].map(pn_dict)  # Map row numbers (Z) back to PN
    matched_df.drop(columns=["Z"], inplace=True)  # Remove the column
    #################
    ###qty >0
    matched_df=matched_df[matched_df['QTY']>0]



    ###making of unmatch sheets
    
    original_list_df=pd.read_excel(lufthansa_stream)
    # Filter rows where the first column of original_list_df is NOT in matched_df's column 6
    unmatched_df = original_list_df[~original_list_df.iloc[:, 0].isin(matched_df.iloc[:, 5])].copy()




    # (Your code also does compare_and_paste_check, copying unmatched from Lufthansa side, etc.)
    # If you need that, implement it here:
    # unmatched_lufthansa = compare_and_paste_check(df_lufthansa, matched_df)
    # ... or something similar ...

    # --------------------------------------------------------
    # 6) Write final DFs to an Excel in memory & return it
    # --------------------------------------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write matched vs unmatched to separate sheets
        matched_df.to_excel(writer, sheet_name='Match value', index=False)
        unmatched_df.to_excel(writer, sheet_name='Not Matching Values', index=False)
        # If needed: unmatched_lufthansa.to_excel(writer, sheet_name='Unmatched_LH', index=False)
    
    output.seek(0)
    output=apply_color_to_excel(output, color_column_name="DateColor")
    return output

# ------------------------------------------------------------------
# Below are the helper functions that replicate each macro's logic.
# You must customize indexes/column references as needed.
# ------------------------------------------------------------------

def rename_heading_insert_columns(df):
    """
    VBA rename_heading_insert_coloumns => rename columns A..N.
    """
    new_headers = [
        "INTERNAL USE", "inv#", "QTY", "Allocated", "On Repair",
        "PN", "DESC", "SN", "COND", "TAG BY", "TAG DATE",
        "TRACE", "SSP", "Ext SSP","OWNER CODE","PHYSICAL LOCATION"
    ]
    df.columns = new_headers[:len(df.columns)]
    return df

def put_keys_in_column_z_and_save_dict(df):
    """
    VBA for_list_coder_remove_non_aplha_and_put_key_in_column_z
       1) place row numbers in col Z
       2) build a dict keyed by row# => column F value
    """
    row_count = len(df)
    # Column Z doesn't exist in pandas by letter, so we'll just create a new col 'Z'.
    df['Z'] = range(1, row_count + 1)  # Excel-like row numbering (2..N+1)

    # Build a dict: Key=i, Value = df.iloc[i-1, 5] (PN).
   # pn_dict = {}
    for i in range(1, row_count + 1):
        pn_dict[i] = df.iloc[i - 1, 5]  # col F => index=5
    return df, pn_dict

def remove_non_alphanumeric_column_f(df):
    """
    VBA for_list_coder_remove_non_aplh_from_row => remove non-alphanumeric in column F.
    """
    df.iloc[:, 5] = (
        df.iloc[:, 5]
        .astype(str)
        .str.replace(r'[^a-zA-Z0-9]', '', regex=True)
    )
    return df

def remove_non_alphanumeric_column_a(df):
    """
    VBA modify_for_list_coder_remove_non_alpha_numeric => remove non-alphanumeric in column A.
    """
    df.iloc[:, 0] = (
        df.iloc[:, 0]
        .astype(str)
        .str.replace(r'[^a-zA-Z0-9]', '', regex=True)
    )
    return df

def changes_for_list_coder(df, coder_value):
    """
    Corresponds to:
      - Move Column A -> Column G
      - Clear old Column A
      - Fill Column I with some value from coder (d8)
    Make sure your DF has enough columns for this reindexing to make sense.

    """
    for i in range(7):
     blank_name = f"Blank_{i+1}"
     df[blank_name] = ""  # create a new column filled with empty strings
    cols = list(df.columns)
    col_a = cols[0]
    
    
    # Remove old A from the list
    del cols[0]
    # Insert col_a at position 6 (which is G in 1-based)
   
    cols.insert(6, col_a)

    df = df[cols]

    # Clear the new col A (index=0)
    df.iloc[:, 0] = ""

    # Fill column I => index=8
   # if len(df.columns) > 8:
    df.iloc[:, 7] = coder_value
    

    return df

def separate_data_by_comma(df):
    """
    VBA SeparateDataByComma => text-to-columns on column I, 
    then create columns condition2..condition7, etc.
    """
    # Column I => index=8
    #if len(df.columns) <= 8:
        #return df  # no column I to split

    col_i = df.iloc[:, 7].astype(str)
    split_cols = col_i.str.split(',', expand=True)  # new DataFrame of splitted pieces
    
    # rename them "condition 2", "condition 3", etc.
    new_col_names = [f"condition {i+2}" for i in range(split_cols.shape[1])]
    split_cols.columns = new_col_names
    
    # Trim
    split_cols = split_cols.apply(lambda col: col.str.strip())

    # Rebuild df: columns up to I, then split_cols, then the rest
    left_part = df.iloc[:, :9]
    right_part = df.iloc[:, 9:]
    df = pd.concat([left_part, split_cols, right_part], axis=1)
    
    return df

def compare_and_update_columns_check(df_row, df_lufthansa):
    """
    VBA heart_matching_files.CompareAndUpdateColumnscheck
    We replicate your logic: if PN & COND in certain columns => mark matched.
    In your code, if found => set column O (index=14 in 1-based, 13 in 0-based) = 1.
    But your rename says column N => index=13 in 0-based. 
    We'll add a new column 'Matched' = 1 if found, else 0.
    Adjust to your actual layout.
    """
    # We'll create a column 'Matched' (like your column 15). Start all as 0.
    df_row['Matched'] = 0

    # Suppose df_row's PN is col 5, COND is col 8 or 9. We must match the code in your macros.
    # For simplicity, let's assume PN is df_row.iloc[:, 5], COND is df_row.iloc[:, 8].
    # In your macros, you do 7 checks across columns in lufthansa_data. We'll just do a set approach.

    # Gather set of (PN, possible conditions) from Lufthansa
    # e.g. PN = col 5 (F), conditions = columns 8..14 (I..O).
    if df_lufthansa.shape[1] < 15:
        # ensure it has at least 15 columns for your logic
        return df_row

    # get Lufthansa PN
    lh_pn_series = df_lufthansa.iloc[:, 6].astype(str)
    
    # get columns I..O => index=8..14
    lh_cond_matrix = df_lufthansa.iloc[:, 9:15].astype(str)

    existing_pairs = set()
    for idx in range(len(df_lufthansa)):
        the_pn = lh_pn_series.iat[idx]
        # columns I..O
        for c in range(8, 15):
            cond_val = df_lufthansa.iat[idx, c]
            pair = (the_pn, str(cond_val))
            
            existing_pairs.add(pair)

    # Now check each row in df_row
    for i in range(len(df_row)):
        row_pn = str(df_row.iat[i, 5])  # PN
        row_cond = str(df_row.iat[i, 8])  # COND, adjust if needed
        
        if (row_pn, row_cond) in existing_pairs:
            df_row.at[i, 'Matched'] = 1

    return df_row

def separate_match_unmatch(df_row):
    """
    VBA SeparateMatchAndUnmatchcheck:
    Filter matched vs unmatched. 
    In your macros you physically delete rows; in pandas we typically split into 2 DFs.
    """
    matched_df = df_row[df_row['Matched'] == 1].copy()
    #unmatched_df = df_row[df_row['Matched'] != 1].copy()
    return matched_df#, unmatched_df

def mark_rows_based_on_date(df):
    """
    VBA MarkRowsBasedOnDatecheck:
    - if difference in months > 24 => red
    - 8 <= diff <= 23 => green
    - 0 <= diff < 8 => yellow
    Because pure pandas doesn't color cells, we store a 'ColorFlag' or 'DateColor'.
    """
    today = datetime.date.today()

    def color_code(dt):
        if pd.isna(dt):
            return None
        d = pd.to_datetime(dt, errors='coerce')
        if pd.isna(d):
            return None
        # compute month diff
        d = d.date()
        month_diff = (today.year - d.year) * 12 + (today.month - d.month)
        if month_diff > 24:
            return "RED"
        elif 8 <= month_diff <= 23:
            return "GREEN"
        elif 0 <= month_diff < 8:
            return "YELLOW"
        else:
            return None

    # Suppose 'TAG DATE' is col 10. In the rename, it's the 11th column => index=10
    if len(df.columns) > 10:
        df['DateColor'] = df.iloc[:, 10].apply(color_code)
    return df
#process_files(temprow,templufty)
# If you also need "compare_and_paste_check" or others, define similarly...
