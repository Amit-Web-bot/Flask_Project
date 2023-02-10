from selenium import webdriver    
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
import pandas as pd
import pyperclip
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option("detach", True)
# Group Name to be fetched here 
#with open('group.txt', 'r', encoding='utf-8-sig') as f:
#   group_name = [group.strip() for group in f.readlines()]
#print(group_name)

data = pd.read_excel('Sample_data.xlsx')
#print(data)

all_names= data['Name']



# Message is to be fetched here 
#with open('msg.txt', 'r', encoding='utf-8-sig') as f:
#    msgs =  f.read()   
    
#print(msgs)
data = pd.read_excel('Sample_data.xlsx')
#print(data)

all_msg= data['Message']

# Opening Chrome and Maximizing
browser = webdriver.Chrome('C:\chromedriver_win32\chromedriver.exe', options=options)

browser.maximize_window()

browser.get('https://web.whatsapp.com/')


for id in range(len(all_names)):
    name = all_names[id]
    print(name)
  
    search_xpath = '//div[@data-testid="chat-list-search"][@contenteditable="true"]'

    search_box = WebDriverWait(browser, 1000).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )
    search_box.clear()

    pyperclip.copy(name)

    search_box.send_keys(Keys.CONTROL + "v")

    time.sleep(2)

    group_xpath = f'//span[@title="{name}"]'
    group_title = browser.find_element(By.XPATH, group_xpath)
    group_title.click()

    time.sleep(30)

    input_xpath = '//div[@data-testid="conversation-compose-box-input"][@role="textbox"]'
    input_box = browser.find_element(By.XPATH, input_xpath)

   
    for id in range(len(all_msg)):
        message1 = all_msg[id]
        pyperclip.copy(str(message1))
        print(name)
        input_box.send_keys(Keys.CONTROL + "v")
    
        input_box.send_keys(Keys.ENTER)

    path = r'C:\Users\Acer\Desktop\Flask Project\Flask_Proj\env\src\new.png'

    try:
         if sys.argv[1]:
            attach_xpath = '//span[@data-testid="clip"]'
            attach_box = browser.find_element(By.XPATH, attach_xpath)
            attach_box.click()
            time.sleep(10)

            #image_xpath = '//span[@button aria-lable ="Photos & Videos"][@data-animate-menu-icons-item="true"]'
            image_xpath = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
            image_box = browser.find_element(By.XPATH, image_xpath)
            image_box.send_keys(path)
            time.sleep(10)

            send_xpath = '//span[@data-icon="send"]'
            send_btn = browser.find_element(By.XPATH, send_xpath)
            send_btn.click()
            time.sleep(2)
    except IndexError:
        pass 