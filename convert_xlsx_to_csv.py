#!/usr/bin/env python3
import pandas as pd
import glob
import os

# set input folder
input_folder = "./"

# find all .xlsx files
xlsx_files = [
    f
    for f in glob.glob(os.path.join(input_folder, "*.xlsx"))
    if not os.path.basename(f).startswith("~$")
]

# Loop over each Excel file
for xlsx_file in xlsx_files:
    print(f"Reading: {xlsx_file}")
    # Read first sheet only
    df = pd.read_excel(xlsx_file)

    # Build the output .csv file
    csv_file = os.path.splitext(xlsx_file)[0] + ".csv"

    # Write to CSV
    df.to_csv(csv_file, index=False)
    print(f"Converted: {csv_file}")

print("All done.")
