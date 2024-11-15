from seleniumbase import BaseCase
import pandas as pd
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
            self.click('//div[@name="profile-icon"]', by="xpath")
            # Fingerprint Configuration
            self.wait_for_element('//a[@id="fingerprint-tab"]', timeout=30)
            self.sleep(1)
            self.click('//a[@id="fingerprint-tab"]', by="id",timeout=15)
            self.waiting_time(30)
            self.click('//div[@class="btn-retake"]', by="xpath", scroll=True)
            self.waiting_time(30)
            self.wait_for_element('//span[@class="v-btn__content"]')
            self.sleep(1)
            self.click('//span[@class="v-btn__content"]', by="xpath",timeout=15)            
            self.sleep(10)
            self.waiting_time(30)
            for b in range(0, 6):
                if b % 2 == 0:
                    i = 1
                else:
                    i = 0
                selections =  self.find_elements('//div[@class="v-responsive__content"]', by="xpath")[i]                     
                selections.click()
                self.sleep(5)
            self.waiting_time(30)
            preferences = self.find_elements('//div[@class="v-responsive__content"]', by="xpath")
            total_pref = len(preferences)
            for c in range(0, total_pref - 1):
                if c % 2 == 0:
                    pref_selection = self.find_elements('//div[@class="v-responsive__content"]', by="xpath")[c]
                    pref_selection.click()
                    self.sleep(2)

            next_button = self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            next_button_length = len(next_button)
            for b in range(0, next_button_length):
                if next_button[b].text.strip().lower() == "next":
                    next_button[b].click()
                    self.sleep(2)

            self.sleep(5)
            for s in range(0, 6):
                self.wait_for_element('//div[@class="skip-button"]', timeout=30)
                self.sleep(1)
                self.click('//div[@class="skip-button"]', by="xpath",timeout=15,scroll=True)
                self.sleep(5)

            self.sleep(2)
            final_steps = self.wait_for_element('//span[@class="v-btn__content"]', timeout=30)
            final_steps.click()
            self.sleep(5)
            self.wait_for_element('//span[@class="v-btn__content"]', timeout=30)
            back_button =  self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            final_total = len(back_button)
            for t in range(0, final_total):
                if back_button[t].text.strip().lower() == "back to prezent":
                    back_button[t].click()
                    self.sleep(5)
                    break

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
            pm.alert("Error while doing the Fingerprint Configuration", "Error", timeout=3000)

        self.sleep(4)
        pm.alert("Task 2 is Done", "Alert", timeout=5000)