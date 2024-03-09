import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Function to handle file upload
def handle_upload():
    global file_paths
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Get file paths for the original Excel sheets
    file_paths = filedialog.askopenfilenames(title="Select Excel Files")

# Function to handle selecting the output directory
def handle_destination():
    global output_directory
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Get output directory for generated Excel sheets
    output_directory = filedialog.askdirectory(title="Select Output Directory")

# Function to classify financial and non-financial frauds
def classify_frauds(data):
    financial_frauds = data[data['Final Amount '] > 0]
    non_financial_frauds = data[data['Final Amount '] == 0]

    return financial_frauds, non_financial_frauds

# Function to generate Excel sheets based on the uploaded files
def generate_excel_sheets():
    global file_paths, output_directory, progress_bar

    if not file_paths:
        messagebox.showerror("Error", "Please select at least one file.")
        return

    if not output_directory:
        messagebox.showerror("Error", "Please select an output directory.")
        return

    # Initialize progress bar
    progress_bar['value'] = 0
    total_files = len(file_paths)
    progress_unit = 100 / total_files

    # Iterate over each uploaded file
    for idx, file_path in enumerate(file_paths, start=1):
        try:
            # Read the uploaded file into a pandas DataFrame
            data = pd.read_excel(file_path)

            # Classify financial and non-financial frauds
            financial_frauds, non_financial_frauds = classify_frauds(data)

            # Extract district name from file path
            district_name = os.path.splitext(os.path.basename(file_path))[0]

            # Create a folder for the district if it doesn't exist
            district_folder = os.path.join(output_directory, district_name)
            os.makedirs(district_folder, exist_ok=True)

            # Save financial frauds into separate Excel files
            if not financial_frauds.empty:
                financial_file = os.path.join(district_folder, 'Financial_Frauds.xlsx')
                financial_frauds.to_excel(financial_file, index=False)

            # Save non-financial frauds into separate Excel files
            if not non_financial_frauds.empty:
                non_financial_file = os.path.join(district_folder, 'Non_Financial_Frauds.xlsx')
                non_financial_frauds.to_excel(non_financial_file, index=False)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

        # Update progress bar
        progress_value = idx * progress_unit
        progress_bar['value'] = progress_value
        progress_bar.update()

    messagebox.showinfo("Excel Sheet Generation", "Excel sheets generated successfully!")

# Create the GUI
root = tk.Tk()
root.title("Fraud Classification")
root.configure(bg='#f0f0f0')  # Set background color

# Create a frame for better organization
frame = ttk.Frame(root, padding="20", style='Custom.TFrame')
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Create style for the frame
style = ttk.Style()
style.configure('Custom.TFrame', background='#f0f0f0')  # Set background color for the frame

# Create labels and buttons for uploading and selecting destination
upload_label = ttk.Label(frame, text="Upload Original Excel Sheets:", style='CustomLabel.TLabel')
upload_button = ttk.Button(frame, text="Upload", command=handle_upload, style='CustomButton.TButton')
destination_label = ttk.Label(frame, text="Select Output Directory:", style='CustomLabel.TLabel')
destination_button = ttk.Button(frame, text="Select", command=handle_destination, style='CustomButton.TButton')
generate_button = ttk.Button(frame, text="Generate", command=generate_excel_sheets, style='CustomButton.TButton')
progress_bar = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")

# Place labels and buttons in the frame
upload_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
upload_button.grid(row=0, column=1, padx=5, pady=5)
destination_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
destination_button.grid(row=1, column=1, padx=5, pady=5)
generate_button.grid(row=2, column=0, columnspan=2, padx=5, pady=10)
progress_bar.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

# Configure resize behavior
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

# Initialize global variables
file_paths = []
output_directory = ""

# Run the GUI
root.mainloop()
