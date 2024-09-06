import enum
from token import STAR
from openpyxl import load_workbook, workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


import os

import openpyxl

import time
import datetime

datetime_now = datetime.datetime.now()


flashscore_main_site_url = 'https://www.flashscore.pl/'

excel_template = "/template.xlsx"
output_excel_name = ""
wb = openpyxl.Workbook()
ws = wb.active

home_name = []
away_name = []
link_to_details = []

last_maches_home = []
last_maches_away = []
face_to_face = []

wb = load_workbook(filename="template.xlsx")
ws = wb.active

universal_class = "_simpleText_lsjrv_4"

chromedriver_path = os.path.join(os.getcwd(), 'chromedriver/chromedriver.exe')

service = Service(chromedriver_path)

# Set Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--disable-search-engine-choice-screen")
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')


driver = webdriver.Chrome(service=service, options=chrome_options)

# driver.maximize_window()

driver.get(flashscore_main_site_url)

accept_cookies = driver.find_element(By.ID, "onetrust-accept-btn-handler")
accept_cookies.click()

saving_data = []

def reveal_all_events():
    time.sleep(5)

    buttons_children = driver.find_elements(By.CLASS_NAME, universal_class)
    
    for button_child in buttons_children:
        parent_element = button_child.find_element(By.XPATH, "..")
        driver.execute_script("arguments[0].click()", parent_element)


def scrape_all_events():
    event_divs = driver.find_elements(By.CLASS_NAME, "event__match")
    
    for event_div in event_divs:
        home_participant = event_div.find_element(By.CLASS_NAME, "event__homeParticipant")
        away_participant = event_div.find_element(By.CLASS_NAME, "event__awayParticipant")
        
        home_participant_text = home_participant.find_element(By.CLASS_NAME, universal_class).get_attribute('innerHTML')
        home_name.append(home_participant_text)

        away_participant_text = away_participant.find_element(By.CLASS_NAME, universal_class).get_attribute('innerHTML')
        away_name.append(away_participant_text)
        
        event_link = event_div.find_element(By.CLASS_NAME, "eventRowLink").get_attribute("href")
        event_link = event_link.replace("#/szczegoly-meczu", "#/h2h/overall")
        link_to_details.append(event_link)

        saving_data.append(driver.find_element(By.ID, "calendarMenu").get_attribute("textContent").strip())

def view_previous_day():
    previous_day_button = driver.find_element(By.CLASS_NAME, "calendar__navigation--yesterday")
    previous_day_button.click()
    

def write_first_part_excel():
    for i, home_name_var in enumerate(home_name, start=2):
        ws.cell(row=i, column=3, value=home_name_var)
        
    for i, away_name_var in enumerate(away_name, start=2):
        ws.cell(row=i, column=4, value=away_name_var)
        
    for i, link_to_details_var in enumerate(link_to_details, start=2):
        ws.cell(row=i, column=2, value=link_to_details_var)

    for i, saving_data_var in enumerate(saving_data, start=2):
        ws.cell(row=i, column=1, value=saving_data_var)

def scrape_from_link():
    for k, i in enumerate(link_to_details, start=2):
        tie_home = 0
        tie_away = 0
        tie_face = 0


        driver.get(i)
        
        elem = WebDriverWait(driver, 4, poll_frequency=0.1).until(EC.presence_of_element_located((By.CLASS_NAME, "h2h__section")))

        event_section_one = driver.find_elements(By.CLASS_NAME, "h2h__section")[0]
        event_section_one_events_icon = event_section_one.find_elements(By.CLASS_NAME, "h2h__icon")
        
        event_section_one_events_outcome = []

        for x in range(len(event_section_one_events_icon)):
            event_section_one_events_icon = event_section_one.find_elements(By.CLASS_NAME, "h2h__icon")[x]
            event_section_one_events_icon_title = event_section_one_events_icon.find_element(By.TAG_NAME, "div").get_attribute('title')
            event_section_one_events_outcome.append(event_section_one_events_icon_title)
         
        
        for j, one_data in enumerate(event_section_one_events_outcome, start=5):
            ws.cell(row=k, column=j, value=one_data)
            
            if one_data == "Remis":
                tie_home += 1
        
        ws.cell(row=k, column=10, value=tie_home)


        event_section_two = driver.find_elements(By.CLASS_NAME, "h2h__section")[1]
        event_section_two_events_icon = event_section_two.find_elements(By.CLASS_NAME, "h2h__icon")
        
        event_section_two_events_outcome = []

        for x in range(len(event_section_two_events_icon)):
            event_section_two_events_icon = event_section_two.find_elements(By.CLASS_NAME, "h2h__icon")[x]
            event_section_two_events_icon_title = event_section_two_events_icon.find_element(By.TAG_NAME, "div").get_attribute('title')
            event_section_two_events_outcome.append(event_section_two_events_icon_title)
            

        for j, two_data in enumerate(event_section_two_events_outcome, start=11):
            ws.cell(row=k, column=j, value=two_data)
            
            if two_data == "Remis":
                tie_away += 1
        
        ws.cell(row=k, column=16, value=tie_away)
            

       
        event_section_three = driver.find_elements(By.CLASS_NAME, "h2h__section")[2]
        event_section_three_rows = event_section_three.find_elements(By.CLASS_NAME, "h2h__row")
        
        event_section_three_events_outcome = []
        
        for x in range(len(event_section_three_rows)):
            event_section_three_result = event_section_three_rows[x].find_element(By.CLASS_NAME, "h2h__result")
            event_section_three_result_score_win = event_section_three_result.find_elements(By.TAG_NAME, "span")[0].get_attribute('innerHTML')
            event_section_three_result_score_lose = event_section_three_result.find_elements(By.TAG_NAME, "span")[1].get_attribute('innerHTML')
            event_section_three_events_outcome.append(f"{event_section_three_result_score_win} : {event_section_three_result_score_lose}")

        
        for j, three_data in enumerate(event_section_three_events_outcome, start=17):
            ws.cell(row=k, column=j, value=three_data)
            
            first_score, second_score = three_data.split(": ")

            if int(first_score) == int(second_score):
                tie_face += 1
                
        ws.cell(row=k, column=22, value=tie_face)
        
        if tie_home >= 3 and tie_away >= 3 and tie_face >= 3:
            ws.cell(row=k, column= 23, value=f"{tie_home}/{tie_away}/{tie_face}")


def main():
    for x in range(5):
        print(f"Days started {x}")
        reveal_all_events()
        scrape_all_events()
        view_previous_day()
        print(f"Scraped day {x}")
    
    print("First part to excel")
    write_first_part_excel()
    
    print("Scraping from links")
    scrape_from_link()
    
    print("Saving excel")
    wb.save(f"data {datetime_now.strftime('%d')}-{datetime_now.strftime('%m')}-{datetime_now.strftime('%Y')} {datetime_now.strftime('%H')}-{datetime_now.strftime('%M')}.xlsx")
    
    wb.close()
    driver.quit()
    
main()