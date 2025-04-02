from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
import random

def find_days(text):
    days = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
    pattern = r"\b(" + "|".join(days) + r")\b"
    matches = re.findall(pattern, text, re.IGNORECASE)
    return " ".join(map(str.lower, matches))

def random_sleep(min_time=2, max_time=6):
    time.sleep(random.uniform(min_time, max_time))

input_file = "stationss.xlsx"
df = pd.read_excel(input_file)
station_ids = df.iloc[:, 0].dropna().tolist() 
station_names = df.iloc[:, 1].tolist()
city = df.iloc[:, 3].tolist()
domain = df.iloc[:, 2].tolist()
country = df.iloc[:, 4].tolist()


driver = webdriver.Chrome()
driver.get("https://psms.bits-pilani.ac.in/stationpreference/student")

time.sleep(2)
username = driver.find_element(By.ID, "userId")
password = driver.find_element(By.ID, "password")  
login_button = driver.find_element(By.CSS_SELECTOR, ".btn")  
username.send_keys("f2023XXXX@hyderabad.bits-pilani.ac.in") 
password.send_keys("XXXXXXX")  # Enter your password
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
            problem_bank_select.select_by_index(1)
            time.sleep(2)
            project_select.select_by_index(1)
            time.sleep(2)
            wait = WebDriverWait(driver, 15)
            random_sleep()
        else:
            print(f"Dropdowns not found for Station ID {station_id}")
            retry=retry+1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)],  city[station_ids.index(station_id)], country[station_ids.index(station_id)], domain[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
            continue

        tables = driver.find_elements(By.TAG_NAME, "table")
        if len(tables) >= 2:
            ##
            page_text = driver.find_element("tag name", "body").text
            temp = re.search(r"Course Details\nFirst Degree\n(.*?)(?:\nDual|$)", page_text, re.DOTALL)
            branches = temp.group(1).strip() if temp else ""
            matches = re.findall(r"([A-Z0-9]+) - ", branches)
            matches = sorted(matches)
            output = ", ".join(matches)
            branches=output
            if(branches==''):
                branches='N/A'

            temp = re.search(r"Technical Skills\n(.*?)(?:Non Technical Skills|$)", page_text,re.DOTALL)
            tech_skills = temp.group(1).strip() if temp else "N/A"

            temp = re.search(r"Non Technical Skills\n(.*?)(?:Facility|$)", page_text,re.DOTALL)
            non_tech_skills = temp.group(1).strip() if temp else "N/A"

            temp = re.search(r"Stipend For First Degree\n(.*?)(?:Stipend For Higher Degree|$)", page_text,re.DOTALL)
            stipend = temp.group(1).strip() if temp.group(1) else "0"
            if stipend == '':
                stipend="N/A"

            temp = re.search(r"Accommodation\s*:\s*(YES|NO)", page_text, re.IGNORECASE)
            accomodation = temp.group(1).strip() if temp else "YES"  #Have to change to YES if some address is provided

            start_time_match = re.search(r"Office Start Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", page_text)
            end_time_match = re.search(r"Office End Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", page_text)

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

            temp = re.search(r"Weekly Holidays\n(.*)", page_text,re.DOTALL)
            holidays = find_days(temp.group(1).strip()).strip()
            if(holidays==""):
                holidays='N/A'
            non_tech_match = re.search(r"Non Technical Skills\n(.*)", tables[5].text, re.DOTALL)
            non_tech_skills = non_tech_match.group(1).strip() if non_tech_match else "N/A"
            if non_tech_skills == '':
                non_tech_skills = "N/A"


            print("Name: ",station_names[station_ids.index(station_id)])
            print("Branches: ",branches)
            print("Tech skills: ",tech_skills)
            print("Non-tech skills: ",non_tech_skills)
            print("Stipend for first Degree: ",stipend)
            print("Accomodation: ",accomodation)
            print("Timings: ",timings)
            print("Weekly Holidays: ",holidays)

        else:
            retry += 1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)],  city[station_ids.index(station_id)], country[station_ids.index(station_id)], domain[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
            continue
        # Store extracted data
        extracted_data.append([station_id,station_names[station_ids.index(station_id)],  city[station_ids.index(station_id)], country[station_ids.index(station_id)], domain[station_ids.index(station_id)], branches, accomodation, timings, holidays, stipend,tech_skills,non_tech_skills])
        text= None
        break

    except Exception as e:
            print(f"Retry {retry + 1} for station {station_id}: {e}")
            retry += 1
            if retry == 3:
                extracted_data.append([station_id, station_names[station_ids.index(station_id)], city[station_ids.index(station_id)], country[station_ids.index(station_id)], domain[station_ids.index(station_id)], "Error", "Error", "Error", "Error", "Error", "Error", "Error"])


output_df = pd.DataFrame(extracted_data, columns=["Station ID","Station Name", "City", "Country", "Domain", "Branches","Accomodation","Timings","Weekly Holidays","Stipend for Single Degree","Tech skills","Non Tech skills"])
output_file = "extracted_data.xlsx"
output_df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")
driver.quit()