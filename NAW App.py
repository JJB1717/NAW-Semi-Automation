import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Preprocess Excel
def preprocess_excel(input_file, output_folder):
    try:
        data = pd.read_excel(input_file)

        data['Final Amount '] = data.apply(lambda row: 0 if pd.isnull(row['Final Amount ']) and any(row.iloc[:-1].notnull()) else row['Final Amount '], axis=1)
        data.loc[data.iloc[:, :-1].notnull().any(axis=1), 'ACKNOWLEDGEMENT NO'] = data.loc[data.iloc[:, :-1].notnull().any(axis=1), 'ACKNOWLEDGEMENT NO'].apply(lambda x: str(x).zfill(14))

        selected_columns = ['ACKNOWLEDGEMENT NO', 'DISTRICT', 'CATEGORY', 'SUB  CATEGORY', 'STATUS', 'Type of Crime', 
                            'Sub category - Crime', 'Victim LOST MONEY ?', 'REMARKS', 'Final Amount ']
        preprocessed_data = data[selected_columns]

        output_file = os.path.join(output_folder, 'JANUARY 2024 Processed.xlsx')
        preprocessed_data.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Preprocessed data saved to: {output_file}")
        return output_file

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while preprocessing: {str(e)}")

# District Separator
def district_separator(input_file, output_directory):
    try:
        preprocessed_data = pd.read_excel(input_file)  
        grouped_data = preprocessed_data.groupby('DISTRICT')

        for district, group in grouped_data:
            if district != 0:
                filename = f'{output_directory}/{district}.xlsx'
                group.to_excel(filename, index=False)

        messagebox.showinfo("Success", "Districts separated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while separating districts: {str(e)}")

# Category Separator
def classify_frauds(data):
    financial_frauds = data[data['Final Amount '] > 0]
    non_financial_frauds = data[data['Final Amount '] == 0]
    return financial_frauds, non_financial_frauds

def category_separator(file_paths, output_directory):
    try:
        for file_path in file_paths:
            data = pd.read_excel(file_path)
            financial_frauds, non_financial_frauds = classify_frauds(data)
            district_name = os.path.splitext(os.path.basename(file_path))[0]
            district_folder = os.path.join(output_directory, district_name)
            os.makedirs(district_folder, exist_ok=True)
            if not financial_frauds.empty:
                financial_file = os.path.join(district_folder, 'Financial_Frauds.xlsx')
                financial_frauds.to_excel(financial_file, index=False)
            if not non_financial_frauds.empty:
                non_financial_file = os.path.join(district_folder, 'Non_Financial_Frauds.xlsx')
                non_financial_frauds.to_excel(non_financial_file, index=False)
        messagebox.showinfo("Success", "Category separation completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while separating categories: {str(e)}")

# Generate Report
def count_status_values(data):
    groups = {
        'FIR': ['FIR Registered'],
        'FAD': ['Closed', 'Rejected', 'Withdrawal', 'No Action'],
        'CSR': ['NC Registered'],
        'Under Process' : ['Under Process'],
        'Pending @ CCPS':['Registered']
    }
    counts = {group: 0 for group in groups}
    for group, keywords in groups.items():
        counts[group] = sum(data['STATUS'].isin(keywords))
    return counts

def generate_report(district_directory):
    try:
        district_counts = []
        for district_folder in os.listdir(district_directory):
            district_path = os.path.join(district_directory, district_folder)
            if os.path.isdir(district_path):
                financial_count = 0
                non_financial_count = 0
                final_amount_sum = 0
                financial_file_path = os.path.join(district_path, 'Financial_Frauds.xlsx')
                if os.path.exists(financial_file_path):
                    financial_data = pd.read_excel(financial_file_path)
                    financial_count = len(financial_data)
                    final_amount_sum += financial_data['Final Amount '].sum() if 'Final Amount ' in financial_data.columns else 0
                else:
                    financial_data = pd.DataFrame()
                non_financial_file_path = os.path.join(district_path, 'Non_Financial_Frauds.xlsx')
                if os.path.exists(non_financial_file_path):
                    non_financial_data = pd.read_excel(non_financial_file_path)
                    non_financial_count = len(non_financial_data)
                    final_amount_sum += non_financial_data['Final Amount '].sum() if 'Final Amount ' in non_financial_data.columns else 0
                else:
                    non_financial_data = pd.DataFrame()
                final_amount_sum = round(final_amount_sum, 2)
                district_data = pd.concat([financial_data, non_financial_data], ignore_index=True)
                status_counts = count_status_values(district_data) if 'STATUS' in district_data.columns else {}
                district_counts.append((district_folder, financial_count, non_financial_count, status_counts, final_amount_sum))
        district_counts.sort(key=lambda x: x[4], reverse=True)

        html_content = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>DISTRICT-WISE REPORT</title>
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
                h1 {
                    text-align: center;
                }
                h4 {
                    text-align: center;
                    margin-top: -20px;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    border: 1px solid #0e0d0e;
                    margin-top: 20px;
                }
                th, td {
                    padding: 10px;
                    text-align: center;
                    border-bottom: 1px solid #0e0d0e;
                    border-right: 1px solid #0e0d0e;
                }
                th {
                    background-color: #f2f2f2;
                    font-weight: bold;
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
                            <th rowspan="2">Rank</th>
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

        for rank, (district, financial_count, non_financial_count, status_counts, final_amount_sum) in enumerate(district_counts, start=1):
            formatted_final_amount_sum = "{:,.2f}".format(final_amount_sum)
            html_content += f"""
                        <tr>
                            <td>{rank}</td>
                            <td>{district}</td>
                            <td>{financial_count}</td>
                            <td>{non_financial_count}</td>
                            <td>{status_counts.get('FIR', 0)}</td>
                            <td>{status_counts.get('FAD', 0)}</td>
                            <td>{status_counts.get('CSR', 0)}</td>
                            <td>{status_counts.get('Under Process', 0)}</td>
                            <td>{status_counts.get('Pending @ CCPS', 0)}</td>
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

        with open("REPORT.html", 'w', encoding='utf-8') as f:
            f.write(html_content)

        messagebox.showinfo("Success", "HTML report generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the report: {str(e)}")

# Main Application
class NAW_App:
    def __init__(self, root):
        self.root = root
        self.root.title("NAW SEMI-AUTOMATION")
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="10 10 10 10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="NAW Semi-Automation Tool", font=("Helvetica", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="1. Select Excel file for preprocessing:").grid(row=1, column=0, sticky=tk.W)
        self.preprocess_entry = ttk.Entry(frame, width=50)
        self.preprocess_entry.grid(row=1, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_preprocess_file).grid(row=1, column=2)

        ttk.Label(frame, text="2. Select folder to save preprocessed file:").grid(row=2, column=0, sticky=tk.W)
        self.output_entry = ttk.Entry(frame, width=50)
        self.output_entry.grid(row=2, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_output_folder).grid(row=2, column=2)

        ttk.Button(frame, text="Preprocess", command=self.run_preprocess).grid(row=3, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="3. Select preprocessed file for district separation:").grid(row=4, column=0, sticky=tk.W)
        self.district_entry = ttk.Entry(frame, width=50)
        self.district_entry.grid(row=4, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_district_file).grid(row=4, column=2)

        ttk.Label(frame, text="4. Select folder to save district files:").grid(row=5, column=0, sticky=tk.W)
        self.district_output_entry = ttk.Entry(frame, width=50)
        self.district_output_entry.grid(row=5, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_district_output_folder).grid(row=5, column=2)

        ttk.Button(frame, text="Separate Districts", command=self.run_district_separator).grid(row=6, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="5. Select district files for category separation:").grid(row=7, column=0, sticky=tk.W)
        self.category_entry = ttk.Entry(frame, width=50)
        self.category_entry.grid(row=7, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_category_files).grid(row=7, column=2)

        ttk.Label(frame, text="6. Select folder to save category files:").grid(row=8, column=0, sticky=tk.W)
        self.category_output_entry = ttk.Entry(frame, width=50)
        self.category_output_entry.grid(row=8, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_category_output_folder).grid(row=8, column=2)

        ttk.Button(frame, text="Separate Categories", command=self.run_category_separator).grid(row=9, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="7. Select folder with district files to generate report:").grid(row=10, column=0, sticky=tk.W)
        self.report_entry = ttk.Entry(frame, width=50)
        self.report_entry.grid(row=10, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.select_report_folder).grid(row=10, column=2)

        ttk.Button(frame, text="Generate Report", command=self.run_generate_report).grid(row=11, column=0, columnspan=3, pady=10)

    def select_preprocess_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.preprocess_entry.delete(0, tk.END)
            self.preprocess_entry.insert(0, file_path)

    def select_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder_path)

    def run_preprocess(self):
        input_file = self.preprocess_entry.get()
        output_folder = self.output_entry.get()
        if input_file and output_folder:
            preprocess_excel(input_file, output_folder)
        else:
            messagebox.showerror("Error", "Please select both input file and output folder.")

    def select_district_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.district_entry.delete(0, tk.END)
            self.district_entry.insert(0, file_path)

    def select_district_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.district_output_entry.delete(0, tk.END)
            self.district_output_entry.insert(0, folder_path)

    def run_district_separator(self):
        input_file = self.district_entry.get()
        output_folder = self.district_output_entry.get()
        if input_file and output_folder:
            district_separator(input_file, output_folder)
        else:
            messagebox.showerror("Error", "Please select both input file and output folder.")

    def select_category_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if file_paths:
            self.category_entry.delete(0, tk.END)
            self.category_entry.insert(0, ", ".join(file_paths))

    def select_category_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.category_output_entry.delete(0, tk.END)
            self.category_output_entry.insert(0, folder_path)

    def run_category_separator(self):
        file_paths = self.category_entry.get().split(", ")
        output_folder = self.category_output_entry.get()
        if file_paths and output_folder:
            category_separator(file_paths, output_folder)
        else:
            messagebox.showerror("Error", "Please select both input files and output folder.")

    def select_report_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.report_entry.delete(0, tk.END)
            self.report_entry.insert(0, folder_path)

    def run_generate_report(self):
        folder_path = self.report_entry.get()
        if folder_path:
            generate_report(folder_path)
        else:
            messagebox.showerror("Error", "Please select the folder containing district files.")

if __name__ == "__main__":
    root = tk.Tk()
    app = NAW_App(root)
    root.mainloop()