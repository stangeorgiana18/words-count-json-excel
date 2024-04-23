# Words Count

Check words count for the same documents, given 2 different files with the same information:

- read data from an Excel file named 'docs.xlsx' located at the specified path
- load JSON data from a file named 'output.json' 
- iterate over each row in the Excel file and match the 'Titlu document' column value with the 'fileName' field in the JSON data. If match found, extract the corresponding word count.
- rename the columns extracted from the Excel file
- calculate the difference between the word counts extracted from the Excel file and the JSON data and adds a new column 'difference' to the DataFrame
- keep only the desired columns in a new Excel file named 'wordcount.xlsx'
- read the new modified Excel file into a DataFrame
- add a new column 'difference' to the DataFrame, setting the value to NaN if it doesn't exist
- calculate the difference of word counts in percentage, rounding to one decimal place, and adds a new column 'difference_percentage' to the DataFrame
- keep only the rows that are not null in the 'wordCountOutput' column
- adds a new column 'avg_percentage_deviation' to the DataFrame, setting the value to NaN
- sort the DataFrame by 'difference_percentage' column in descending order
- assign the average percentage deviation to the first value in the 'avg_percentage_deviation' column
- saves the DataFrame to a new Excel file named 'new_doc.xlsx'
