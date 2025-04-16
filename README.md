# PSMS_data_extractor
PSMS Data Extractor automates data collection from the BITS PSMS website using Selenium. It logs in, navigates station details, and extracts key info like branches, stipend, timings, and skills. Data is saved in an Excel file for easy access. With retry logic and randomized delays, it ensures reliable and efficient scraping.

## Features  
- Automated login and navigation  
- Extraction of station details (branches, stipend, timings, skills, etc.)  
- Retry logic for handling failures  
- Randomized delays for improved reliability  
- Data saved in an Excel file  

## Prerequisites  
- Python 3.x  
- Google Chrome  
- ChromeDriver (matching your Chrome version)  
- Required Python packages:  
  ```sh
  pip install selenium pandas openpyxl

## How to use
1. Clone this repository:  
   ```sh
   git clone https://github.com/Vivek5170/psms-data-extractor.git  
   cd psms-data-extractor
2. Place your stationss.xlsx file (containing Station IDs and names) in the project folder.
3. Fill your login credentials.
4. Run the script:
   ```sh
   python method2.py

## Note
- In drop downs in view page of a company the program auto selects the first option for Select Problem Bank and Select project.

## Disclaimer
- This script is for personal and academic use only. Ensure compliance with BITS policies before using it.
