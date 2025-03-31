from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
import random

def random_sleep(min_time=2, max_time=6):
    time.sleep(random.uniform(min_time, max_time))

input_file = "stationss.xlsx"
df = pd.read_excel(input_file)
station_ids = df.iloc[:, 0].dropna().tolist() 
station_names = df.iloc[:, 1].tolist()


driver = webdriver.Chrome()
driver.get("https://psms.bits-pilani.ac.in/stationpreference/student")

time.sleep(2)
username = driver.find_element(By.ID, "userId")
password = driver.find_element(By.ID, "password")  
login_button = driver.find_element(By.CSS_SELECTOR, ".btn")  
username.send_keys("fxxxxxxxx@hyderabad.bits-pilani.ac.in") 
password.send_keys("xxxxxxx")  # Enter your password
login_button.click()
time.sleep(5)

extracted_data = []

for station_id in station_ids[:5]:
 retry=0
 print("opening url for:")
 print(station_id)
 print("\n")
 url = f"https://psms.bits-pilani.ac.in/stationpreference/student/stationdetails/{station_id}/70078"
 while(retry<4):
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 15)
        select_elements = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "select")))
        if len(select_elements) >= 2:
            problem_bank_select = Select(driver.find_element(By.CSS_SELECTOR, "select"))
            project_select = Select(driver.find_elements(By.CSS_SELECTOR, "select")[1])
            time.sleep(5)
            problem_bank_select.select_by_index(len(problem_bank_select.options) - 1)
            time.sleep(2)
            project_select.select_by_index(len(project_select.options) - 1)
            time.sleep(2)
            wait = WebDriverWait(driver, 15)
            random_sleep()
        else:
            print(f"Dropdowns not found for Station ID {station_id}")
            retry=retry+1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error"])
            continue

        tables = driver.find_elements(By.TAG_NAME, "table")
        if len(tables) >= 2:
            matches = re.findall(r"([A-Z0-9]+) - ", tables[1].text)
            matches = sorted(matches)
            output = ", ".join(matches)
            branches=output

            tech_match = re.search(r"Technical Skills\n(.*?)(?:\nNon Technical Skills|$)", tables[5].text, re.DOTALL)
            non_tech_match = re.search(r"Non Technical Skills\n(.*)", tables[5].text, re.DOTALL)

            # Extract and clean skills
            tech = tech_match.group(1).strip() if tech_match else ""
            non = non_tech_match.group(1).strip() if non_tech_match else ""

            text = None

            for i in range(6, len(tables)):
                current_text = tables[i].text
                if current_text.startswith("Facility"):
                    text = current_text
                    break

            if text is None:
                print("Correct Table not available")

            # Extract stipend for first degree (fallback to "N/A" if not found)
            stipend_match = re.search(r"Stipend For First Degree\s*\n\s*(\d+)", text, re.IGNORECASE)
            stipend = stipend_match.group(1) if stipend_match else "N/A"

            # Extract office start and end timings
            start_time_match = re.search(r"Office Start Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", text)
            end_time_match = re.search(r"Office End Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", text)

            if start_time_match and end_time_match:
                start_hour, start_min = map(int, start_time_match.groups())
                end_hour, end_min = map(int, end_time_match.groups())

                start_period = "AM" if start_hour < 12 else "PM"
                end_period = "AM" if end_hour < 12 else "PM"

                # Convert 24-hour format to 12-hour
                start_hour = start_hour if start_hour <= 12 else start_hour - 12
                end_hour = end_hour if end_hour <= 12 else end_hour - 12

                timings = f"{start_hour}:{start_min:02d} {start_period} to {end_hour}:{end_min:02d} {end_period}"
            else:
                timings = "N/A"

            accommodation_match = re.search(r"Accommodation\s*:\s*(YES|NO)", text, re.IGNORECASE)
            accommodation = accommodation_match.group(1) if accommodation_match else "YES"

            # Extract weekly holidays
            holidays_match = re.search(r"Weekly Holidays\s*\n((?:[A-Za-z]+(?:,? ?))+)", text)

            if holidays_match:
                holidays_text = holidays_match.group(1).strip()
                holidays = re.split(r"[\n,]\s*", holidays_text)  # Split by newline or comma
                weekly_holidays = ", ".join(holidays)
            else:
                weekly_holidays = "N/A"
        else:
            retry += 1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error"])
            continue
        # Store extracted data
        extracted_data.append([station_id,station_names[station_ids.index(station_id)], branches, accommodation, timings, weekly_holidays, stipend,tech,non])
        text= None
        break

    except Exception as e:
            print(f"Retry {retry + 1} for station {station_id}: {e}")
            retry += 1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error"])


output_df = pd.DataFrame(extracted_data, columns=["Station ID","Station Name" , "Branches","Accomodation","Timings","Weekly Holidays","Stipend for Single Degree","Tech skills","Non Tech skills"])
output_file = "extracted_data.xlsx"
output_df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")
driver.quit()