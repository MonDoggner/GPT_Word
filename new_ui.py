from typing import Dict

from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

class Parser:
    def __init__(self, user_data: dict, url: str, prompt:str):
        #self.url = url
        self.url = 'https://chat.deepseek.com'
        self.user_data = user_data
        self.prompt = prompt

        self.chrome_options = Options()
        self.driver = webdriver.Chrome(options=self.chrome_options)

        self.elements: Dict[str, str] = {
            'username_field': '/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[2]/div[1]/div/input',
            'password_field': '/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[3]/div[1]/div/input',
            'enter_button': '/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[5]',
            'prompt_field': '//*[@id="chat-input"]',
        }

    def wait_for_element(self, by, value, timeout=10000) -> WebElement | None:
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            return None

    def registration(self):
        try:
            self.driver.get(self.url)

            username_field = self.wait_for_element(
                By.XPATH, 
                self.elements['username_field'],
                timeout=10
            )

            password_field = self.wait_for_element(
                By.XPATH, 
                self.elements['password_field'],
                timeout=10
            )

            enter_button = self.wait_for_element(
                By.XPATH, 
                self.elements['enter_button'],
                timeout=10
            )

            if username_field and password_field:
                username_field.send_keys(self.mail)
                password_field.send_keys(self.password)
                enter_button.click()
            return None
        except Exception as e:
            print(f'Ошибка: {e}')
    
    def parse_response(self):
        ...