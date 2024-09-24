UNWRA Truck Entries - Script Methodology and Workflow
Workflow for Running the 0.master_script for Analyzing UNWRA Truck Entries
Step 1: Confirm Initial Setup
Ensure the following components are correctly set up:
	•	KCAL Reference File: Ensure the kcal_reference.xlsx file is located in the /Users/jackdevine/Desktop/FEWS NET/UNWRA Truck Script/data/ directory.
	•	Script Files: The following Python script files should be located in the /Users/jackdevine/Desktop/FEWS NET/UNWRA Truck Script/scripts/ directory:
	◦	1.download_raw.py
	◦	2.0processing.py
	◦	3.apply_kcal_values.py
	◦	4.calc_truck_kcals_mt.py
	◦	5.daily_totals.py
Step 2: Access Command Line Interface
	•	Mac: Open Terminal via Applications or the search function.
Step 3: Navigate to the Script Directory
In Terminal, navigate to the script directory:


cd /Users/jackdevine/Desktop/FEWS\ NET/UNWRA\ Truck\ Script/scripts/
Ensure the path matches the script storage directory.
Step 4: Execute the Master Script
To initiate the entire workflow, run the master script:


python3 0.master_script.py
This will sequentially execute all the necessary scripts for processing and analyzing truck entry data.
Step 5: Monitor Script Execution
Watch for any output or error messages in the terminal, addressing issues such as missing files, permission errors, or script failures.
Step 6: Confirm Output Integrity
After running the scripts, verify the correct data has been generated in /Users/jackdevine/Desktop/FEWS NET/UNWRA Truck Script/data/. Check for the creation of the unwra_trucks.xlsx file, which should contain the processed data.
Step 7: Troubleshooting
If issues arise, confirm the correct placement of all files, including path configurations. For unresolved issues, review the individual script logic and ensure all required input files are present.
Step 8: Documentation and Logging
Record any changes or outputs from running the scripts, noting any issues encountered or solutions implemented. This will help with future iterations and troubleshooting.

Script Documentation
Script 0: Master Execution Script
File Name: 0.master_script.py
This script manages the sequential execution of the processing scripts. Key points:
	1	Script Execution: It runs the scripts listed in the scripts array, starting with downloading raw data and continuing through to generating daily totals.
	2	Error Handling: It handles errors through try-except blocks, allowing uninterrupted execution of the subsequent scripts even if one script fails.
	3	Completion Message: Provides real-time output for each executed script and a final completion notification.
Script 1: Download Raw Data
File Name: 1.download_raw.py
This script downloads raw data from a Google Drive link using the following steps:
	1	File Download: Uses gdown to download the raw data file to the data directory.
	2	Format Conversion: If the file isn’t in Excel format, the script converts it to Excel and saves it to unwra_trucks_raw.xlsx.
Script 2.0: Data Processing
File Name: 2.0processing.py
This script processes the raw truck data and prepares it for further analysis:
	1	Data Cleaning: Renames columns, handles missing values, and converts certain columns to numeric formats.
	2	Weight Calculation: Calculates truck weight based on the type of cargo and unit.
	3	Data Storage: Saves the processed data to a new sheet, unwra_clean, in the unwra_trucks.xlsx file.
Script 3: Apply Caloric Values
File Name: 3.apply_kcal_values.py
This script integrates caloric values into the processed truck data:
	1	Data Loading: Loads processed truck data and caloric reference data.
	2	Caloric Value Assignment: Matches food items to caloric values from the reference file, adding columns for caloric data.
	3	Data Saving: Saves the enriched data with caloric values to the unwra_trucks_kcal sheet.
Script 4: Calculate Metric Tonnage and Calories
File Name: 4.calc_truck_kcals_mt.py
This script calculates the total caloric values and metric tonnage of food per truck:
	1	Truck Type Classification: Identifies the type of truck (food, non-food, or mixed).
	2	Calorie and Weight Calculations: Calculates the total caloric content and weight of food items for each truck.
	3	Data Output: Saves results to a new sheet, unwra_trucks_with_kcal, in the workbook.
Script 5: Daily Totals
File Name: 5.daily_totals.py
This script generates daily totals for truck entries:
	1	Data Aggregation: Aggregates the total truck count, caloric content, and food metric tonnage by day.
	2	Daily Breakdown: Computes daily truck counts by type (food, non-food, mixed) and sector (humanitarian or private).
	3	Crossing Points: If available, counts the truck entries by crossing point (e.g., Kerem Shalom, Rafah).
	4	Data Saving: Saves daily totals to a new sheet, unwra_daily_entries, in the unwra_trucks.xlsx file.
 
