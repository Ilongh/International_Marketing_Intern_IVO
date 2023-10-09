
import pandas as pd
from collections import defaultdict

def split_by_delimiters(s, delimiters):
    """Recursively split a string by multiple delimiters."""
    if not delimiters:
        return [s]
    
    delimiter = delimiters[0]
    parts = s.split(delimiter)
    
    # If there are more delimiters to process, split each part further
    if len(delimiters) > 1:
        parts = [subpart for part in parts for subpart in split_by_delimiters(part, delimiters[1:])]
    
    return parts

def process_excel_file(input_path, column_name, output_path):
    df = pd.read_excel(input_path)
    counts = defaultdict(int)
    delimiters = [';', ',', '\n']

    # Iterate through each row in the dataframe
    for index, row in df.iterrows():
        items = [item.strip() for item in split_by_delimiters(row[column_name], delimiters)]
        for item in items:
            counts[item] += row['计数']

    # Convert the dictionary to a dataframe and save to Excel
    df_processed = pd.DataFrame(counts.items(), columns=['项目', '出现次数']).sort_values(by='出现次数', ascending=False)
    df_processed.to_excel(output_path, index=False)

if __name__ == "__main__":
    process_excel_file("NEGATIVE.xlsx", "N", "processed_NEGATIVE.xlsx")
    process_excel_file("POSITIVE.xlsx", "P", "processed_POSITIVE.xlsx")
    process_excel_file("导出计数.xlsx", "R/S", "processed_导出计数.xlsx")
