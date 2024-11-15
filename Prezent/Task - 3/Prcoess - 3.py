from seleniumbase import BaseCase
import pandas as pd
import pyautogui as pa
import pymsgbox as pm
import os

ROOT = os.path.join(os.environ['USERPROFILE'], r"Desktop\Prezent")
Credentials_path = os.path.join(ROOT, r"LoginDetails\Credentials.xlsx")

BaseCase.main(__name__, __file__)

class TestSimpleLogin(BaseCase):
    
    def waiting_time(self, time_to_wait):
        try:
            self.wait_for_element("body", timeout=time_to_wait)
            print("Page is fully loaded.")
        except Exception as e:
            print("Page load timed out:", str(e))
            self.refresh()
        self.sleep(3) 



    def test_simple_login(self):


        try:
            df = pd.read_excel(Credentials_path, sheet_name=0)
            website_link = df.iloc[0, 3]
            print(website_link)
            excel_UN = df.iloc[0, 1]
            excel_password = df.iloc[0, 2]

        except:
            pm.alert("Not able to access the file or could not be present in the desired folder", "Error", timeout=3000)

        try:
            self.open(website_link)
            self.waiting_time(30)
            self.maximize_window()
            self.type("username", excel_UN, by="id")
            self.sleep(1)
            self.click("continue", by="id")
            self.sleep(1)
            self.type("password", excel_password, by="id")
            self.sleep(1)
            self.click("submit", by="id")
            self.waiting_time(60)
            self.wait_for_element("//div[@name='profile-icon']", timeout=60)
            pm.alert("Login is Done", "Alert", timeout=2000)
            # Downloading the document
            Auto_button = self.wait_for_element('//*[@id="v-step-3"]', timeout=60)
            Auto_button.click()
            self.waiting_time(30)
            self.sleep(2)
            prompt =  self.wait_for_element('//div[@class = "v-text-field__slot"]', timeout=60)
            prompt.click()
            self.sleep(5)
            pa.click(700, 600)
            self.sleep(5)
            generate = self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            total_length = len(generate)
            for a in range(0, total_length):
                if generate[a].text.strip().lower() == "generate":
                    generate[a].click()
                    self.sleep(7)
                    break

            while True:
                try:
                    download_button = self.wait_for_element('//span[@name="download-icon"]', timeout=60)
                    download_button.click()
                    self.sleep(7) 
                    download = self.wait_for_element('//div[@class="btnText"]', timeout=60)
                    download.click()
                    self.sleep(3)
                    pptx = self.wait_for_element('//*[@id="download-btn-from-list"]', timeout=60)
                    pptx.click()
                    self.sleep(15)    
                    break
                except:
                    pm.alert("Not able to click the Download button on screen because waiting for the file to be downloaded", "Alert", timeout=3000) 
                    self.sleep(10) 

            self.wait_for_element("//div[@name='profile-icon']", timeout=60)
            self.sleep(1)
            self.click('//div[@name="profile-icon"]', by="xpath")
            self.sleep(2)

            # Logout from website
            self.waiting_time(30)
            logout = self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            total_len = len(logout)
            for i in range(0, total_len):
                if logout[i].text.strip().lower() == "sign out":
                    logout[i].click()
                    self.sleep(1)
                    break           
        except:
            pm.alert("Error while downloading the document", "Error", timeout=3000)

        self.sleep(4)
        pm.alert("Task 3 is Done", "Alert", timeout=5000)