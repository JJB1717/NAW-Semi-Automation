1) Keep excel in D:\NAW - SASTRA PROJECT , say 'JANUARY 2024.xlsx'
2) Open Preprocess Excel , open with IDLE(python)
	Edit 4rd line with file name , 
	data = pd.read_excel('D:/NAW - SASTRA PROJECT/JANUARY 2024.xlsx')

	Specify output folder in line 18
	output_folder = 'D:/NAW - SASTRA PROJECT'

	Edit the output processed excel name in line 21
	output_file = output_folder + '/JANUARY 2024 Processed.xlsx'

3) Run the Module or click F5

-----------------------------------------------------------------------------------------------------
4) Open District Separator program, click Édit with IDLE
	Line 4, open the path of processed data
	preprocessed_data = pd.read_excel('D:/NAW - SASTRA PROJECT/JANUARY 2024 Processed.xlsx')  

	Create Empty folder named 'GENERATED SHEETS'
	Refer line 10
	output_directory = 'D:/NAW - SASTRA PROJECT/GENERATED SHEETS'
	
5) Run the module or click F5

-----------------------------------------------------------------------------------------------------

6)Open Category Separator, edit with IDLE

7) Run the module, upload the district excel present in 'GENERATED SHEETS' folder

8) Select the output folder as 'GENERATED SHEETS'or any other folder

-----------------------------------------------------------------------------------------------------

9) Open Generate Report program, edit with IDLE

10) Line 24 select the directory for districts

11) Run the module, click F5

-----------------------------------------------------------------------------------------------------

12) Open Generate Ranked Report program, edit with IDLE

13) Line 24 select the directory for districts

14) Run the module, click F5
