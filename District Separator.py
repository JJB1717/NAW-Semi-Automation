import pandas as pd

# Read the preprocessed Excel file
preprocessed_data = pd.read_excel('D:/CCW Project/NCRP/JANUARY 2024 Processed.xlsx')  

# Group the data by the DISTRICT column
grouped_data = preprocessed_data.groupby('DISTRICT')

# Specify the directory where you want to save the separate Excel files
output_directory = 'D:/CCW Project/NCRP/GENERATED SHEETS' 

# Loop through each group and save it to a separate Excel file
for district, group in grouped_data:
    # Define the filename for the Excel file
    if district != 0:
        filename = f'{output_directory}/{district}.xlsx'
        
        # Save the group to the Excel file
        group.to_excel(filename, index=False)

print("\n Districts separated successfully !")