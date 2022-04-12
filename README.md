# A2_2030_Final

Overview:
- This is an automated application that transfers meter data from Constellation to Energystar excel format.

Requirement:
- To use this application, you need both the excel sheet from constellation and the Energystar multimeter excel upload template.

Packages needed:
- pandas, numpy, openpyxl

Instructions:
1. Put the Constellation and Energystar Excel sheets in the same directory as main.py
2. Change the "energystar_excel_file" variable to be the name of the Energstay excel template
3. Change the "constellation_data_file" variable to be the name of the Constellation excel data file
4. Run the main.py file
5. The data from the Constellation file will be populated into the Energystar multimeter upload template in the "Output_file.xlsx" file



Validating:
- The validation.py file will output any warnings or errors found in your completed "Output_file.xlsx" that may be of concern and should be reviewed before uploading to Energy Star.

Packages needed:
- pandas

Instructions:
1. Ensure the constellation excel sheet and the Output file are in the same directory as validation.py
2. Change the "output_file" variable to be the name of the output excel file we just created
3. Change the "constellation_file" variable to be the name of the Constellation excel data file
4. Run the validation.py file
5. The warnings and errors found in the file will be outputted into the "warnings_and_errors.txt" file
6. These should be reviewed before uploading into Energy Star. 

