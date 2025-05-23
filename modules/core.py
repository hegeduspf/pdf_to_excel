# extract_pdf_to_excel/modules/core.py
import argparse, re
import pandas as pd
from datetime import datetime
from pathlib import Path

def parse_args():
    # Setup and read in command line args.
    parser = argparse.ArgumentParser()
    parser.add_argument("config", help="Filepath of the config file")
    parser.add_argument("-v", "--verbose",
                        help="output helpful messages while the program runs",
                        action="store_true"
    )
    # Return the stored arguments.
    return parser.parse_args()

# Get the current date and time.
def get_curr_dt():
    return datetime.now().astimezone()

# Only print if in verbose mode.
def vprint(v, msg):
    if v:
        print(msg)

# Check if a file exists and is the correct filetype.
def check_file(file_path, file_ext):
    f = Path(file_path)
    return f.is_file() and f.suffix.lower().endswith(file_ext)
        
# Check if a directory exists.
def check_dir(dir_path):
    d = Path(dir_path)
    return d.is_dir()

# Delete a file given its absolute file path.
def delete_file(file_path):
    f = Path(file_path)
    try:
        f.unlink()
        result = f"{file_path} has been deleted."
    except FileNotFoundError:
        result = f"Could not delete...{file_path} does not exist."
    except PermissionError:
        result = f"Could not delete {file_path}...permission denied."
    return result

# Check if a list contains specific text (ignoring case/special 
# characters).
def list_contains_text(input_list, text_to_check):
    # Remove special characters from string text_to_check.
    cleaned_text = re.sub(r'[^a-zA-Z0-9]', '', text_to_check)
    
    # Check if the cleaned text exists in the cleaned input.
    return any(cleaned_text.lower() in s.lower() for s in input_list)


def slice_list_to_df(text_list, cols):
    """Given a list of text and a list of column names for a dataframe,
    return a fully populated dataframe where each value from the list
    of text has been sorted into the correct column based on its
    sequential order (i.e., if there are 4 columns, the 1st value goes
    into column 1, 2nd value into column 2, 3rd value into column 3, 
    4th value into column 4, and the cycle resets: 5th value into
    column 1, 6th value into column 2, 7th value into column 3, 8th
    value into column 4, etc.). This is done via simple string slicing
    as opposed to iterating through every value in the provided text
    list.
    """
    df = pd.DataFrame(columns = cols)
    n = len(cols)
    for i in range(n):
        df[cols[i]] = text_list[i::n]
    
    return df