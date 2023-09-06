# Selenium inports
import os
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from workspace.bots import secret

# --------------------------SETUP-------------------------- #
dir_path = os.path.dirname(os.path.realpath(__file__))
downloads = (
    dir_path + r"\downloads"
)  # Create downloads path for files to be downloaded using selenium
driverlocation = r"workspace\assets\drivers\chromedriver.exe"

view_button_array = []


class Browser:  # Selenium Browser Configuration
    browser, service = None, None

    # Initialise the webdriver with the path to chromedriver.exe
    def __init__(self, driver: str):
        self.service = Service(driver)

        options = Options()
        options.add_argument("start-maximized")
        options.add_experimental_option(
            "prefs", {"download.default_directory": downloads}
        )

        self.browser = webdriver.Chrome(service=self.service, chrome_options=options)

    def open_page(self, url: str):
        self.browser.get(url)

    def close_browser(self):
        self.browser.close()

    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(1)

    def clear_input(self, by: By, value: str):
        field = self.browser.find_element(by=by, value=value)
        field.clear()
        time.sleep(1)

    def click_button(self, by: By, value: str):
        button = self.browser.find_element(by=by, value=value)
        button.click()
        time.sleep(1)

    def login_papercut(self, username: str, password: str):
        self.add_input(by=By.NAME, value="inputUsername", text=username)
        self.add_input(by=By.NAME, value="inputPassword", text=password)
        self.click_button(by=By.CLASS_NAME, value="loginSubmit")

    def papercut_deposit(
        self, name: str, amount: str, paymentmethod: str, comment: str
    ):
        self.add_input(by=By.NAME, value="username", text=name)
        self.clear_input(by=By.NAME, value="creditAmount")
        self.add_input(by=By.NAME, value="creditAmount", text=amount)
        self.add_input(
            by=By.NAME, value="paymentMethod", text=paymentmethod
        )  # EFT, Cash or Other
        self.add_input(by=By.NAME, value="$TextField", text=comment)
        self.click_button(by=By.NAME, value="$Submit$0")


def deposit(ref, amount, method, comment):
    # browser = Browser_Headless(driverlocation)

    browser = Browser(driverlocation)
    browser.open_page("https://papercut.horsham-college.vic.edu.au/webcashier")
    time.sleep(2)
    try:
        browser.login_papercut(secret.papercut_username, secret.papercut_password)
    except:
        browser.open_page("https://papercut.horsham-college.vic.edu.au/webcashier")
        browser.login_papercut(secret.papercut_username, secret.papercut_password)
    browser.open_page(
        "https://papercut.horsham-college.vic.edu.au/app?service=page/WebCashierDeposit"
    )
    browser.papercut_deposit(ref, amount, method, comment)
