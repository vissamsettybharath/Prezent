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
images_path = os.path.join(ROOT, r"Images")
downloads_path = os.path.join(os.environ['USERPROFILE'], r"Downloads")
custom_log_filename = "Prezent_Process-3"
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
        
    time.sleep(5) 
    logging.info("Login to the website is done")
    return (driver)



def get_latest_file(folder):
    global os

    files = [os.path.join(folder, f) for f in os.listdir(folder)]
    files = [f for f in files if os.path.isfile(f)] 
    if not files:
        return None
    latest_file = max(files, key=os.path.getctime)
    return latest_file



def process_autogenerator(driver, images_path, downloads_path):
    global By, time, WebDriverWait, EC, pm, logging, waiting_time, get_latest_file

    try:
        time.sleep(2)
        Auto_button = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, 'v-step-3')))
        Auto_button.click()
    except:
        pm.alert(f"Not able to find the autogenerator button on screen", "Error", timeout=3000)
        logging.info("Not able to find the autogenerator button on screen")


    waiting_time(driver, 30)  
    time.sleep(5)


    try:
        prompt = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@class = "v-text-field__slot"]')))
        prompt.click()
    except:
        pm.alert(f"Not able to find the prompt on screen", "Error", timeout=3000)
        logging.info("Not able to find the prompt button on screen")    

    time.sleep(5)


    try:
        suggestions = pa.locateCenterOnScreen(os.path.join(images_path,r"Prompt.png"))
        max_wait_time = 5
        start_time = time.time()
        while time.time() - start_time < max_wait_time:
            if suggestions is not None:
                pa.moveTo(suggestions)
                time.sleep(1)
                pa.click()
                time.sleep(1)
                break
    except:
        pm.alert(f"Not able to find the suggestion on screen", "Error", timeout=3000)
        logging.info("Not able to find the suggestion on screen") 


    try:
        time.sleep(5)
        generate = WebDriverWait(driver, 15, poll_frequency=2).until(EC.presence_of_all_elements_located((By.XPATH, '//span[@class="v-btn__content"]')))
        total_length = len(generate)
        for a in range(0, total_length):
            if generate[a].text.strip().lower() == "generate":
                generate[a].click()
                time.sleep(7)     
                break

    except:
        pm.alert(f"Not able to find the Generate button on screen", "Error", timeout=3000)
        logging.info("Not able to find the Generate button on screen")   


    while True:
        try:
            download_button = WebDriverWait(driver, 60, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//span[@name="download-icon"]')))
            download_button.click()
            time.sleep(7) 
            download = WebDriverWait(driver, 60, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="btnText"]')))
            download.click()
            time.sleep(3)
            pptx = WebDriverWait(driver, 60, poll_frequency=2).until(EC.element_to_be_clickable((By.ID, 'download-btn-from-list')))
            pptx.click()
            before_files = set(os.listdir(downloads_path))
            time.sleep(10)    
            timeout = 20  # seconds
            end_time = time.time() + timeout
            new_file = None

            while time.time() < end_time:
                after_files = set(os.listdir(downloads_path))
                new_files = after_files - before_files
                if new_files:
                    new_file = new_files.pop()
                    break
                time.sleep(1)

            if new_file:
                print(f"New file downloaded  {new_file}")
            else:
                print("No new file was downloaded within the timeout.")
            break

        except:
            logging.info("Not able to click the Download button on screen because waiting for the file to be downloaded")  
            time.sleep(10)


    try:
            time.sleep(2)
            Account_button = WebDriverWait(driver, 15, poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH, '//div[@name="profile-icon"]')))
            Account_button.click()
    except:
        pm.alert(f"Not able to find the Profile button on screen", "Error", timeout=3000)
        logging.info("Not able to find the Profile button on screen")

    time.sleep(5) 
    logging.info("Process to download the document is done")



def Logout_website(driver):
    global pa, By, time,  WebDriverWait, EC, pm, logging, waiting_time


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
process_autogenerator(drive, images_path, downloads_path)
Logout_website(drive)
pm.alert(f"Task 3 is Done ", "Alert", timeout=3000)
