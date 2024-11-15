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
            # process Template
            self.wait_for_element('//a[@id="templates-tab"]', timeout=30)
            self.sleep(1)
            self.click('//a[@id="templates-tab"]', by="id",timeout=5)
            self.waiting_time(30)
            selection_elements = self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            total_length = len(selection_elements)
            for g in range(0, total_length):
                if selection_elements[g].text.strip().lower() == "current selection":
                    selection_elements[g].click()
                    self.sleep(3)
                    template_name = self.wait_for_element('//div[@class="v-card__title cardTitleForViewer"]', timeout=30)
                    template_name = template_name.text.strip()
                    print(f"Name of the Current template is {template_name}")
                    pm.alert(f"Name of the Current template is {template_name}", "Alert", timeout=5000)
                    break

            self.sleep(2)
            # Logout from website
            self.wait_for_element('//a[@id="basics-tab"]', timeout=30)
            self.sleep(1)
            self.click('//a[@id="basics-tab"]', by="xpath")
            self.waiting_time(30)
            logout = self.find_elements('//span[@class="v-btn__content"]', by="xpath")
            total_len = len(logout)
            for i in range(0, total_len):
                if logout[i].text.strip().lower() == "sign out":
                    logout[i].click()
                    self.sleep(1)
                    break
        except:
            pm.alert("Error while finding the current selection on screen", "Error", timeout=3000)

        self.sleep(4)
        pm.alert("Task 1 is Done", "Alert", timeout=5000)
   