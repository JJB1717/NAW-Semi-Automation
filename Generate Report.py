import os
import pandas as pd

# Function to count 'STATUS' values
def count_status_values(data):
    # Define groups and corresponding keywords
    groups = {
        'FIR': ['FIR Registered'],
        'FAD': ['Closed', 'Rejected', 'Withdrawal', 'No Action'],
        'CSR': ['NC Registered'],
        'Under Process' : ['Under Process'],
        'Pending @ CCPS':['Registered']
    }
    # Initialize counts for each group
    counts = {group: 0 for group in groups}

    # Count occurrences of each group
    for group, keywords in groups.items():
        counts[group] = sum(data['STATUS'].isin(keywords))

    return counts

# Specify the directory containing district folders
district_directory = r'D:\CCW Project\NCRP\GENERATED SHEETS'     

# Initialize variables to store counts
district_counts = []

# Iterate through each district folder
for district_folder in os.listdir(district_directory):
    district_path = os.path.join(district_directory, district_folder)
    
    if os.path.isdir(district_path):
        # Initialize counts for financial and non-financial records
        financial_count = 0
        non_financial_count = 0
        final_amount_sum = 0

        # Read financial frauds Excel sheet
        financial_file_path = os.path.join(district_path, 'Financial_Frauds.xlsx')
        if os.path.exists(financial_file_path):
            financial_data = pd.read_excel(financial_file_path)
            financial_count = len(financial_data)
            final_amount_sum += financial_data['Final Amount '].sum()

        # Read non-financial frauds Excel sheet
        non_financial_file_path = os.path.join(district_path, 'Non_Financial_Frauds.xlsx')
        if os.path.exists(non_financial_file_path):
            non_financial_data = pd.read_excel(non_financial_file_path)
            non_financial_count = len(non_financial_data)
            final_amount_sum += non_financial_data['Final Amount '].sum()

        final_amount_sum = round(final_amount_sum,2)

        # Count 'STATUS' values
        status_counts = {}
        district_data = pd.concat([financial_data, non_financial_data], ignore_index=True)
        if 'STATUS' in district_data.columns:
            status_counts = count_status_values(district_data)

        # Append district name, counts, and final amount sum to the list
        district_counts.append((district_folder, financial_count, non_financial_count, status_counts, final_amount_sum))

# Generate HTML report
html_content = """
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CCW REPORT</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f2f2f2;
        }

        .container {
            max-width: 1000px;
            margin: 50px auto;
            padding: 20px;
            background-color: #fff;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }

        h1{
            text-align: center;

        }
        h4 {
            text-align: center;
            margin-top: -20px;
        }


        table {
            width: 100%;
            border-collapse: collapse;
            border : 1px solid #0e0d0e;
            margin-top: 20px;
        }

        th,
        td {
            padding: 10px;
            text-align: center;
            border-bottom: 1px solid #0e0d0e;
            border-right: 1px solid #0e0d0e;
            /* Add right border for all cells */
        }

        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }

        /* Remove right border for the last cell in each row */
        th:last-child,
        td:last-child {
            border-right: 1px solid #131212;
        }

        tr:hover {
            background-color: #f9f9f9;
        }
    </style>
</head>

<body>
    <div class="container">
        <h1>DISTRICT-WISE REPORT</h1>
        <h4>January 2024</h4>
        <table>
            <thead>
                <tr>
                    <th rowspan="2">District</th>
                    <th colspan="2">Total Portal Complaints</th>
                    <th colspan="5">Action Taken</th>
                    <th rowspan="2">Amount Lost (Rs.)</th>
                </tr>
                <tr>
                    <th>Financial</th>
                    <th>Non-Financial</th>
                    <th>FIR</th>
                    <th>FAD</th>
                    <th>CSR</th>
                    <th>Under Process</th>
                    <th>Pending @ CCPS</th>
                </tr>
            </thead>
            <tbody>
"""

for district, financial_count, non_financial_count, status_counts, final_amount_sum in district_counts:
    formatted_final_amount_sum = "{:,.2f}".format(final_amount_sum)
    html_content += f"""
            <tr>
                <td>{district}</td>
                <td>{financial_count}</td>
                <td>{non_financial_count}</td>
                <td>{status_counts['FIR']}</td>
                <td>{status_counts['FAD']}</td>
                <td>{status_counts['CSR']}</td>
                <td>{status_counts['Under Process']}</td>
                <td>{status_counts['Pending @ CCPS']}</td>
                <td>{formatted_final_amount_sum}</td>
            </tr>
"""

html_content += """
            </tbody>
        </table>
    </div>
</body>

</html>
"""

# Write HTML content to a file
with open("REPORT.html", "w") as html_file:
    html_file.write(html_content)

print("\n District-wise report generated successfully !")
