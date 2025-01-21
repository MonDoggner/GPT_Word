#░██████╗░██████╗░████████╗░██╗░░░░░░░██╗░█████╗░██████╗░██████╗░
#██╔════╝░██╔══██╗╚══██╔══╝░██║░░██╗░░██║██╔══██╗██╔══██╗██╔══██╗
#██║░░██╗░██████╔╝░░░██║░░░░╚██╗████╗██╔╝██║░░██║██████╔╝██║░░██║
#██║░░╚██╗██╔═══╝░░░░██║░░░░░████╔═████║░██║░░██║██╔══██╗██║░░██║
#╚██████╔╝██║░░░░░░░░██║░░░░░╚██╔╝░╚██╔╝░╚█████╔╝██║░░██║██████╔╝
#░╚═════╝░╚═╝░░░░░░░░╚═╝░░░░░░╚═╝░░░╚═╝░░░╚════╝░╚═╝░░╚═╝╚═════╝░

from selenium.webdriver.remote.webelement import WebElement

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

import time
from time import sleep

import tkinter as tk
from tkinter import messagebox

import logging

import pyperclip
import win32com.client


def setup_logging():
    """Функция для создания логов"""

    logger = logging.getLogger('GPT_Word')
    logger.setLevel(logging.DEBUG)

    # Создание файла логов 

    file_handler = logging.FileHandler('gpt_word.log', encoding='utf-8')

    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    formatter = logging.Formatter(log_format)
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)

    return logger


class Backend:
    def __init__(self):
        chrome_options = Options()

        # Устанавливаем настройки для увеличения производительности        

        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-software-rasterizer')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--allow-running-insecure-content')
        chrome_options.add_argument('--disable-gl-drawing-for-tests')
        chrome_options.add_argument('--disable-accelerated-2d-canvas')
        chrome_options.add_argument('--disable-webrtc')
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--log-level=3')

        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        self.url = 'https://ask.chadgpt.ru/'
        self.driver = webdriver.Chrome(options=chrome_options)
        self.mail = None
        self.password = None

        self.logger = setup_logging()

    def reg_gpt(self) -> None:

        # Регистрация аккаунта в gpt через конфиг 
        # Данные хранятся в файле config.txt  

        with open('config.txt', 'r', encoding='utf-8') as config_file:
            data = config_file.read().split('\n')

        self.mail = data[0][12:]
        self.password = data[1][12:]

    def wait_for_element(self, by, value, timeout=10000) -> WebElement | None:

        # Ожидание появления элемента на странице 

        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            self.logger.error(f'Элемент не найден: {value}')
            return None

    def parse_response(self) -> str | None:

        # Парсинг ответа 

        try:
            response_xpath = '/html/body/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div/div[2]'
            continue_xpath = '/html/body/div/div[2]/div/div[2]/div[2]/div[1]/div/div/div[3]/div[1]/div/button[1]'
            
            continue_button = self.wait_for_element(By.XPATH, continue_xpath)                            
            response_element = self.wait_for_element(By.XPATH, response_xpath)

            if response_element and continue_button:
                sleep(1)
                response_text = response_element.text.strip()
                return response_text
            return None
        except TimeoutException:
            self.logger.error('Ошибка ожидания ответа')
            return None
        except Exception as e:
            self.logger.error(f'Ошибка при парсинге: {str(e)}')
            return None

    def gpt_post(self, prompt: str) -> str | None:

        # Отправка запроса в gpt 

        try:
            self.driver.get(self.url)
            self.logger.info('Загружаем страницу')

            try:
                
                # Симулируем вход в аккаунт chad ai 

                self.logger.info('Входим в аккаунт')
                
                continue_button = self.wait_for_element(
                    By.XPATH, 
                    '''//div[contains(@class, 'flex') and contains(text(), 'Продолжить через')]''',
                    timeout=10
                )

                if continue_button:
                    continue_button.click()
                    self.logger.info('Выбираем вход через почту')

                username_field = self.wait_for_element(
                    By.XPATH, 
                    '''//input[@placeholder='Введите email']''',
                    timeout=10
                )

                password_field = self.wait_for_element(
                    By.XPATH, 
                    '''//input[@placeholder='Введите пароль']''',
                    timeout=10
                )

                if username_field and password_field:
                    username_field.send_keys(self.mail)
                    password_field.send_keys(self.password)

                    self.logger.info('Логин и пароль введены')

                    login_button = self.wait_for_element(By.CLASS_NAME, 'login_btn_enter', timeout=10)

                    if login_button:
                        login_button.click()
                        self.logger.info('Вход выполнен')
                    else:
                        self.logger.error('Не удалось найти кнопку подтверждения')
                        return None
                    
                else:
                    self.logger.error('Поля регистрации не найдены')
                    return None
            except Exception as e:
                self.logger.error(f'Не удалось залогиниться: {str(e)}')
                return None

            # Теперь ищем поле ввода и отправляет запрос 

            try:
                text_area = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'inp_area'))
                )
                
                if text_area:
                    time.sleep(2)
                    
                    text_area = self.driver.find_element(By.ID, 'inp_area')
                    text_area.clear()
                    text_area.send_keys(prompt)
                    text_area.send_keys(Keys.RETURN)

                    self.logger.info('Запрос успешно отправлен')
                    
                    # Вот теперь наконец-то можно парсить 

                    response_text = self.parse_response()

                    if response_text:
                        self.logger.info('Ответ получен')
                        return response_text
                    else:
                        self.logger.error('Не удалось получить ответ')
                        return None     
                    
                else:
                    self.logger.error('Поле ввода не найдено')
                    return None 
            except Exception as e:
                self.logger.error(f'Ошибка при отправке запроса или получении ответа: {str(e)}')
                return None
        except Exception as e:
            self.logger.error(f'Ошибка gpt_post: {str(e)}')
            return None
        finally:
            try:
                self.driver.quit()
            except:
                pass

    def make_request(self, prompt) -> str:
            return self.gpt_post(prompt)


class App:
    def __init__(self):
        self.backend = Backend()

        root = tk.Tk()
        root.title('Редактор текста для Word')
        root.geometry('300x100')
        root.attributes('-topmost', True) 

        button = tk.Button(root, text='Получить выделенный текст', command=self.copy_selected_text_from_word)
        button.pack(pady=20)

        root.mainloop()

    def copy_selected_text_from_word(self) -> None:
        try:

            # Получаем доступ к Word 
            
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = True
            selection = word.Selection
            
            # Проверяем, есть ли выделенный текст      
            # Если тип выделения не пустой             
            # Копируем выделенный текст в буфер обмена 
            # Извлекаем текст из буфера обмена         

            if selection.Type != 0:  
                selection.Copy()  
                time.sleep(0.1) 
                selected_text = pyperclip.paste()

                self.backend.reg_gpt()
                response = self.backend.make_request(selected_text)
                
                # Печатаем тект в вордовский файл 

                if response:
                    # Save the response to a backup file first
                    timestamp = time.strftime('%Y%m%d-%H%M%S')
                    backup_filename = f'response_backup_{timestamp}.txt'
                    
                    try:
                        with open(backup_filename, 'w', encoding='utf-8') as backup_file:
                            backup_file.write(response)
                            
                        # Try to paste into Word
                        end_point = selection.End
                        selection.Collapse(Direction=0)
                        selection.TypeText('\n' + response)
                        selection.Font.Name = word.ActiveDocument.Styles.Normal.Font.Name
                        selection.Font.Size = word.ActiveDocument.Styles.Normal.Font.Size

                    except Exception as e:
                        messagebox.showwarning(
                            "Warning",
                            f"Failed to paste into Word. The response has been saved to {backup_filename}\nError: {str(e)}"
                        )

            else:
                messagebox.showinfo('Information', 'No text selected.')
                
        except Exception as e:
            print(f'Ошибка {e}')

def main():
    App()

if __name__ == "__main__":
    main()
    