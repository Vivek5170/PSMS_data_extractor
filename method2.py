from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
import random
import os
from multiprocessing import Process

def find_days(text):
    days = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
    pattern = r"\b(" + "|".join(days) + r")\b"
    matches = re.findall(pattern, text, re.IGNORECASE)
    return " ".join(map(str.lower, matches))

def random_sleep(min_time=2, max_time=6):
    time.sleep(random.uniform(min_time, max_time))

def run_scraper(station_ids, station_names, city, domain, country, id):
    extracted_data = []
    driver = webdriver.Chrome()
    driver.get("https://psms.bits-pilani.ac.in/stationpreference/student")

    time.sleep(2)
    username = driver.find_element(By.ID, "userId")
    password = driver.find_element(By.ID, "password")
    login_button = driver.find_element(By.CSS_SELECTOR, ".btn")
    username.send_keys("f2023xxxx@hyderabad.bits-pilani.ac.in")
    password.send_keys("xxxxxxx")  # Replace with your actual password
    login_button.click()
    time.sleep(5)
    try:
     for idx, station_id in enumerate(station_ids):
        retry = 0
        print("Opening URL for:", station_id, "\n")
        url = f"https://psms.bits-pilani.ac.in/stationpreference/student/stationdetails/{station_id}/70078"
        while retry < 4:
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
                    retry += 1
                    if retry == 3:
                        extracted_data.append([station_id, station_names[idx], city[idx], country[idx], domain[idx], "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
                    continue

                tables = driver.find_elements(By.TAG_NAME, "table")
                if len(tables) >= 2:
                    page_text = driver.find_element("tag name", "body").text

                    temp = re.search(r"Course Details\nFirst Degree\n(.*?)(?:\nDual|$)", page_text, re.DOTALL)
                    branches = temp.group(1).strip() if temp else ""
                    matches = re.findall(r"([A-Z0-9]+) - ", branches)
                    matches = sorted(matches)
                    branches = ", ".join(matches) if matches else 'N/A'

                    temp = re.search(r"Technical Skills\n(.*?)(?:Non Technical Skills|$)", page_text, re.DOTALL)
                    tech_skills = temp.group(1).strip() if temp else "N/A"
                    tech_skills = tech_skills if tech_skills!='' else "N/A"

                    temp = re.search(r"Non Technical Skills\n(.*?)(?:Facility|$)", page_text, re.DOTALL)
                    non_tech_skills = temp.group(1).strip() if temp!='' else "N/A"

                    temp = re.search(r"Project Title\n(.*?)(?:Project Description|$)", page_text, re.DOTALL)
                    project_title = temp.group(1).strip() if temp else "N/A"
                    project_title = project_title if project_title!='' else "N/A"

                    temp = re.search(r"Project Title\n(.*)", tables[0].text, re.DOTALL)
                    project_description = temp.group(1).strip() if temp else "N/A"
                    project_description = project_description if project_description!='' else "N/A"

                    temp = re.search(r"Stipend For First Degree\n(.*?)(?:Stipend For Higher Degree|$)", page_text, re.DOTALL)
                    stipend = temp.group(1).strip() if temp.group(1).strip() else 0
                    stipend = stipend if stipend!='' else 0

                    temp = re.search(r"Accommodation\s*:\s*(YES|NO)", page_text, re.IGNORECASE)
                    accomodation = temp.group(1).strip() if temp else "YES"

                    start_time_match = re.search(r"Office Start Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", page_text)
                    end_time_match = re.search(r"Office End Time\s*\n\s*(\d{2}):(\d{2}):\d{2}", page_text)
                    if start_time_match and end_time_match:
                        start_hour, start_min = map(int, start_time_match.groups())
                        end_hour, end_min = map(int, end_time_match.groups())
                        start_period = "AM" if start_hour < 12 else "PM"
                        end_period = "AM" if end_hour < 12 else "PM"
                        start_hour = start_hour if start_hour <= 12 else start_hour - 12
                        end_hour = end_hour if end_hour <= 12 else end_hour - 12
                        timings = f"{start_hour}:{start_min:02d} {start_period} to {end_hour}:{end_min:02d} {end_period}"
                    else:
                        timings = "N/A"
                    timings=timings if timings !="" else "N/A"


                    temp = re.search(r"Weekly Holidays\n(.*)", page_text, re.DOTALL)
                    holidays = find_days(temp.group(1).strip()) if temp else 'N/A'
                    holidays = holidays if holidays!='' else 'N/A'

                    non_tech_match = re.search(r"Non Technical Skills\n(.*)", tables[5].text, re.DOTALL)
                    non_tech_skills = non_tech_match.group(1).strip() if non_tech_match else "N/A"
                    non_tech_skills = non_tech_skills if non_tech_skills!='' else "N/A"

                    print("Name:", station_names[idx])
                    print("Project Title:", project_title)
                    print("Branches:", branches)
                    print("Tech skills:", tech_skills)
                    print("Non-tech skills:", non_tech_skills)
                    print("Stipend for First Degree:", stipend)
                    print("Accommodation:", accomodation)
                    print("Timings:", timings)
                    print("Weekly Holidays:", holidays)

                    extracted_data.append([station_id, station_names[idx], city[idx], country[idx], domain[idx], branches, accomodation, timings, holidays, stipend, tech_skills, non_tech_skills, project_title, project_description])
                    break
                else:
                    retry += 1
                    if retry == 3:
                        extracted_data.append([station_id, station_names[idx], city[idx], country[idx], domain[idx], "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
            except Exception as e:
                print(f"Retry {retry + 1} for station {station_id}: {e}")
                retry += 1
                if retry == 3:
                    extracted_data.append([station_id, station_names[idx], city[idx], country[idx], domain[idx], "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
    except Exception as e:
        print(e)
    finally:
        output_df = pd.DataFrame(extracted_data, columns=["Station ID","Station Name", "City", "Country", "Domain", "Branches","Accomodation","Timings","Weekly Holidays","Stipend for Single Degree","Tech skills","Non Tech skills","Project Title","Project Description"])
        output_file = f"extracted_data{id}.xlsx"
        output_df.to_excel(output_file, index=False)
        time.sleep(5)
        driver.quit()

if __name__ == "__main__":
    input_file = "stationss.xlsx"
    df = pd.read_excel(input_file)
    station_ids = df.iloc[:, 0].dropna().tolist() 
    station_names = df.iloc[:, 1].tolist()
    city = df.iloc[:, 3].tolist()
    domain = df.iloc[:, 2].tolist()
    country = df.iloc[:, 5].tolist()

    mid_index = len(station_ids) // 2

    ids1, ids2 = station_ids[:mid_index], station_ids[mid_index:]
    names1, names2 = station_names[:mid_index], station_names[mid_index:]
    cities1, cities2 = city[:mid_index], city[mid_index:]
    domains1, domains2 = domain[:mid_index], domain[mid_index:]
    countries1, countries2 = country[:mid_index], country[mid_index:]

    p1 = Process(target=run_scraper, args=(ids1, names1, cities1, domains1, countries1,1))
    p2 = Process(target=run_scraper, args=(ids2, names2, cities2, domains2, countries2,2))

    p1.start()
    p2.start()
    p1.join()
    p2.join()
    
    try:
        results1 = pd.read_excel( "extracted_data1.xlsx")
        results2 = pd.read_excel( "extracted_data2.xlsx")
        all_data = pd.concat([results1, results2], ignore_index=True)
        output_df = pd.DataFrame(all_data, columns=["Station ID","Station Name", "City", "Country", "Domain", "Branches","Accomodation","Timings","Weekly Holidays","Stipend for Single Degree","Tech skills","Non Tech skills","Project Title","Project Description"])
        output_file = "extracted_data.xlsx"
        output_df.to_excel(output_file, index=False)
        print(f"Data saved to {output_file}")
        files_to_delete = ["extracted_data1.xlsx", "extracted_data2.xlsx"]

        for file in files_to_delete:
            if os.path.exists(file):
                os.remove(file)
                print(f"{file} has been deleted.")
            else:
                print(f"{file} does not exist.")
    except Exception as e:
        print(e)