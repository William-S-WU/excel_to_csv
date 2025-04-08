#Excel to CSV converter

import os
import pandas as pd 


directory = os.getcwd()  # Replace if working directory is not the desired directory
print("Working Directory:      ", directory)

# function creates csvs from each excel sheet
def excel_to_sheet(filename):
    excelfile = pd.ExcelFile(filename)
    print("Converting File:      ", filename)
    for sheet_name in excelfile.sheet_names:

        df = pd.read_excel(filename, sheet_name=sheet_name)
        df['Data Type'] = sheet_name #creates a new column called data type which has the sheet name in it this is useful for my project but can be removed
        df.to_csv(f"{sheet_name}.csv", index=False)




#circles through current directory serching for excel files
for filename in os.listdir(directory):
    #print("doing something")
    if filename.lower().endswith(".xlsx"):
        #print("doing something")
        filenamej = directory + "\\" + filename
        excel_to_sheet(filenamej)
        
        
print("Conversion Complete.")
