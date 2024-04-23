import pandas as pd
import json
import numpy as np


excel_file_path = '/Users/georgianastan/Desktop/words-count/docs.xlsx'
excel_data = pd.read_excel(excel_file_path)

#sheet_one = list(excel_data.keys())[0]
#print(dict)
#print(dfs[sheet_one])

# Detect encoding
# with open('/Users/georgianastan/Desktop/words-count/output.json', 'rb') as f:
#     raw_data = f.read(32)  # Read the first 32 bytes for encoding detection
#     encoding = chardet.detect(raw_data)['encoding']

# Open the file with the detected encoding
# with open('/Users/georgianastan/Desktop/words-count/output.json', encoding=encoding) as f:
#     json_data = json.load(f)
#     print(json_data)

with open('/Users/georgianastan/Desktop/words-count/output.json') as f:
    json_data = json.load(f)
    #print(json_data)

# Iterate over each row in the Excel file:
for index, row in excel_data.iterrows():
    # Extract the filename and corresponding word count from the JSON data:
    for item in json_data:
        if row['Titlu document'] in item['fileName']:
            word_count = item['wordCount']
            break
    else:
        word_count = None  # None if filename not found in JSON data

    # Assign the word count to the 'Nr. cuvinte json' column
    excel_data.at[index, 'Nr. cuvinte json'] = word_count

# Rename the columns extracted from the excel file
excel_data.rename(columns={'Titlu document': 'fileName',
                            'Nr. cuvinte': 'wordCount',
                            'Nr. cuvinte json': 'wordCountOutput'}, inplace=True)

# The desired columns
excel_data = excel_data[['fileName', 'wordCount', 'wordCountOutput']]

# Difference between the words count in the json and excel file, function applied row-wise
excel_data['difference'] = excel_data.apply(lambda row: abs(row['wordCount'] - row['wordCountOutput'])
                                            if pd.notnull(row['wordCount']) and pd.notnull(row['wordCountOutput'])
                                            else None, axis = 1)

# Keep only the desired columns in a new excel file
excel_data.to_excel('/Users/georgianastan/Desktop/words-count/wordcount.xlsx', index=False)

# Read the new modified excel file, with 2 columns containing words count
new_doc = pd.read_excel('/Users/georgianastan/Desktop/words-count/wordcount.xlsx')

# Add a new column with the difference, set the value to nan if it doesn't exist
# Ensure that the difference column contains numeric values
new_doc['difference'] = new_doc['difference'].apply(lambda x: int(x) if pd.notnull(x) else np.nan)

# Difference between words count in percentage, rounded to 1 decimal only 
new_doc['difference_percentage'] = abs((new_doc['difference'] / new_doc['wordCount'])) * 100

new_doc['difference_percentage'] = new_doc['difference_percentage'].round(1)

# Keep only the rows that are not null, given the number of words in the excel file
new_doc = new_doc[new_doc['wordCountOutput'].notnull()]

# Select non-null values in the "difference percentage" column
non_null_values = new_doc['difference_percentage'].notnull()

# Sum up the non-null values
sum_non_null_values = new_doc.loc[non_null_values, 'difference_percentage'].sum()

# Count the number of non-null values
count_non_null_values = non_null_values.sum()

# Calculate the average
average_percentage_deviation = sum_non_null_values / count_non_null_values

print("Average percentage deviation:", average_percentage_deviation)

new_doc['avg_percentage_deviation'] = np.nan

new_doc = new_doc.sort_values(by='difference_percentage', ascending=False)

# Keep only the first value in the last column
new_doc['avg_percentage_deviation'].iloc[0] = round(average_percentage_deviation, 1)

print(new_doc.head())

new_doc.to_excel('/Users/georgianastan/Desktop/words-count/new_doc.xlsx', index = False)

