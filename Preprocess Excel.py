import pandas as pd

# Read the Excel file
data = pd.read_excel('D:/CCW Project/NCRP/JANUARY 2024.xlsx')  # Replace 'your_dataset.xlsx' with the actual filename

# Replace empty or blank values in 'Final Amount' column with '0' for rows with other entries
data['Final Amount '] = data.apply(lambda row: 0 if pd.isnull(row['Final Amount ']) and any(row.iloc[:-1].notnull()) else row['Final Amount '], axis=1)

# Apply custom format '00000000000000' to cells in 'ACKNOWLEDGEMENT NO' column with entries in other columns
data.loc[data.iloc[:, :-1].notnull().any(axis=1), 'ACKNOWLEDGEMENT NO'] = data.loc[data.iloc[:, :-1].notnull().any(axis=1), 'ACKNOWLEDGEMENT NO'].apply(lambda x: str(x).zfill(14))

# Select only the specified columns
selected_columns = ['ACKNOWLEDGEMENT NO', 'DISTRICT', 'CATEGORY', 'SUB  CATEGORY', 'STATUS', 'Type of Crime', 
                    'Sub category - Crime', 'Victim LOST MONEY ?', 'REMARKS', 'Final Amount ']
preprocessed_data = data[selected_columns]

# Specify the path to the folder where you want to save the file
output_folder = 'D:/CCW Project/NCRP'  # Replace 'path/to/your/folder' with the actual folder path

# Define the full path including the filename
output_file = output_folder + '/JANUARY 2024 Processed.xlsx'

# Save the preprocessed data to the specified folder
preprocessed_data.to_excel(output_file, index=False)

print(f"Preprocessed data saved to: {output_file}")
