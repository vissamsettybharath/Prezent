from selenium.webdriver.common.by import By
import os
from selenium import webdriver
import time
import pandas as pd
from selenium.webdriver.chrome.options import Options
import pyautogui as pa
import pymsgbox as pm
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import logging



ROOT = os.path.join(os.environ['USERPROFILE'], r"Desktop\Prezent")
Credentials_path = os.path.join(ROOT, r"LoginDetails\Credentials.xlsx")
custom_log_filename = "Prezent_Process-1"
Logs = os.path.join(ROOT, r"Logs")
log_filename = f"{custom_log_filename}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
log_file = os.path.join(Logs, log_filename)
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')



def waiting_time(driver, time_to_wait):
    global By, time, WebDriverWait, EC

    max_wait_time = time_to_wait
    try:
        WebDriverWait(driver, max_wait_time, poll_frequency=2).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        print("Page is fully loaded.")
    except Exception as e:
        print("Page load timed out:", str(e))
        driver.refresh()
    time.sleep(3)



def open_Prezent_Website(Credentials_path):
    global pa, By, Service, time, Options, webdriver,  WebDriverWait, EC, pm, pd, logging, waiting_time

    try:
        pa.hotkey("winleft","r")
        time.sleep(3)
        shortcut_chrome_browser = "C:\\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
        pa.write(shortcut_chrome_browser)
        time.sleep(2)
        pa.press("enter")
        time.sleep(5)
        service = Service()
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.switch_to.window(driver.window_handles[0])
        driver.maximize_window()
        time.sleep(5)
        pa.click(x=1400, y=800)
        time.sleep(3)
        
        
    except:
        pm.alert(f"Not able to access the chrome", "Error", timeout=3000)
        logging.info("Not able to access the chrome")
        

    try:
        df = pd.read_excel(Credentials_path, sheet_name=0)
        website_link = df.iloc[0, 3]
        print(website_link)
        
        time.sleep(2)
        driver.get(website_link)
        time.sleep(2)
        excel_UN = df.iloc[0, 1]
        excel_password = df.iloc[0, 2]

    except:
        pm.alert(f"Not able to access the file or could not be present in the desired folder", "Error", timeout=3000)
        logging.info("Not able to access the file or could not be present in the desired folder")
        
    
#     opening the website
    
    while True:
        waiting_time(driver, 30)
               
    #     Login details 
        
        try:
            user_name = WebDriverWait(driver, 30, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, "username")))
            user_name.send_keys(excel_UN)
            time.sleep(3)
        except:
            pm.alert(f"Not able to click the User Name field", "Error", timeout=3000)
            logging.info("Not able to click the User Name field")



        try:
            continue_button = WebDriverWait(driver, 30, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, "continue")))
            continue_button.click()
            time.sleep(2)
        except:
            pm.alert(f"Not able to click the Continue Button", "Error", timeout=3000)
            logging.info("Not able to click the Continue Button")

    
        waiting_time(driver, 30)
            
        try:
            password = WebDriverWait(driver, 30, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, "password")))
            password.send_keys(excel_password)
            time.sleep(2)
        except:
            pm.alert(f"Not able to click the Password field", "Error", timeout=3000)
            logging.info("Not able to click the Password field")
            
            
        try:
            login_button = WebDriverWait(driver, 30, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, "submit")))
            login_button.click()
            time.sleep(2)
            
        except:
            pm.alert(f"Not able to click the Login Button", "Error", timeout=3000)
            driver.refresh()
            time.sleep(3)
            logging.info("Not able to click the Login Button")
    

        waiting_time(driver, 60)
        
        time.sleep(5)

        try:
            Account_button = WebDriverWait(driver, 30, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@name="profile-icon"]')))
            pm.alert(f"Login is Done", "Alert", timeout=2000)
            break
        except:      
            driver.refresh()
            time.sleep(10)
        
    try:
        time.sleep(2)
        Account_button = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@name="profile-icon"]')))
        Account_button.click()
    except:
        pm.alert(f"Not able to find the Profile button on screen", "Error", timeout=3000)
        logging.info("Not able to find the Profile button on screen")

    time.sleep(5) 
    logging.info("Login to the website is done")
    return (driver)



def process_template(driver):
    global  By, time,  WebDriverWait, EC, pm, logging, waiting_time

    try:
        time.sleep(2)
        Template_tab = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, 'templates-tab')))
        Template_tab.click()
    except:
        pm.alert(f"Not able to find the Template Tab on screen", "Error", timeout=3000)
        logging.info("Not able to find the Template Tab on screen")

    waiting_time(driver, 30)


    try:
        time.sleep(5)
        selection = WebDriverWait(driver, 15, poll_frequency=2).until(EC.presence_of_all_elements_located((By.XPATH, '//span[@class="v-btn__content"]')))
        total_length = len(selection)
        for g in range(0, total_length):
            if selection[g].text.strip().lower() == "current selection":
                selection[g].click()
                time.sleep(3)
                template_name = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="v-card__title cardTitleForViewer"]')))
                print(f"Name of the Current template is {template_name.text}")
                pm.alert(f"Name of the Current template is {template_name.text} ", "Alert", timeout=5000)
                time.sleep(2)
                break
                
    except:
        pm.alert(f"Not able to find the current selection on screen", "Error", timeout=3000)
        logging.info("Not able to find the Template Tab on screen")

    time.sleep(5) 
    logging.info("Process to find template Name is done")



def Logout_website(driver):
    global pa, By, time, WebDriverWait, EC, pm, logging, waiting_time


    try:
        time.sleep(2)
        Basics_tab = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, 'basics-tab')))
        Basics_tab.click()
    except:
        pm.alert(f"Not able to find the Basics Tab on screen", "Error", timeout=3000)
        logging.info("Not able to find the Basics Tab on screen")

    waiting_time(driver, 30)


    try:
        time.sleep(5)
        logout = WebDriverWait(driver, 15, poll_frequency=2).until(EC.presence_of_all_elements_located((By.XPATH, '//span[@class="v-btn__content"]')))
        total_len = len(logout)
        for i in range(0, total_len):
            if logout[i].text.strip().lower() == "sign out":
                logout[i].click()
                time.sleep(2)
                break
    except:
        pm.alert(f"Not able to find the Sign Out button on screen", "Error", timeout=3000)
        logging.info("Not able to find the Sign Out button on screen")   

    time.sleep(5) 
    logging.info("Logout is done")
    time.sleep(2) 
    pa.hotkey("alt", "F4")
    time.sleep(2) 



drive = open_Prezent_Website(Credentials_path)
process_template(drive)
Logout_website(drive)
pm.alert(f"Task 1 is Done ", "Alert", timeout=3000)