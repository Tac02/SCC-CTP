## SCC CTP Tracker App - BETA VERSION, Still in development

This App is designed to take a single download of all CTP training on Westminster for a unit, and convert it into a readable format to allow Unit Training Officers (UTOs) to effectively plan and track training. 
This repository includes a packaged Windows executable for the CTP processor. The executable presents a small GUI where users can select the input cadet CSV/Excel file and choose the output file name and destination.

- Download the latest version of the ctp_processor.exe and Module_list.csv from [here](releases/tag/v0.1)

- **Run the EXE:** Double-click the produced `ctp_processor.exe` (or run from Command Prompt).
- **Select input file:** Browse to your cadet data file (CSV, .xls, .xlsx). 
- **Select output file:** Choose a destination and filename for the generated tracker (default `CTP_Tracker.xlsx`).
- **Master module list:** This can be downloaded alongside the App, make sure to place it in the same folder. 


## Downloading raw data

The App is designed to take a full unit input for the whole CTP. Smaller samples may work, but this has not been tested.

Follow these steps to download the data file.
- Go to the Personnel page on Westminster
- On the left, select the "Cadet Qualifications" arrow
- Select "Unit Report"
- Select "Actions" and then "Filter"
- In the "Column" Tab, under "Column" select "Syllabus" (should be the default), under "Operator" select "contains" and under Expression enter "CTP"
- This will select all CTP modules completed by all cadets in the unit
- Select "Actions" and then "Download"
- Click "Download"

- It is recommended that the CTP report you have just generated is saved for future reference. to do this select "Actions", "Save Report" and give the report a name and description. It will then appear in the dropdown under "Primary Report" for future use. 
