from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import names
from xlwt import Workbook

spreadsheet = Workbook()
sheet1 = spreadsheet.add_sheet('Sheet 1')

options = Options()
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
driver = webdriver.Chrome(chrome_options=options,
                          executable_path=r'C:\Users\Iram\Downloads\chromedriver_win32 (1)\chromedriver.exe')

#Please update driver path to your own path

def login():
    try:
        driver.get(
            "https://login.yahoo.com/?.lang=en-US&src=homepage&.done=https%3A%2F%2Fwww.yahoo.com%2F&pspid=2023538075&activity=ybar-signin")
        driver.implicitly_wait(15)

        search_bar_username = driver.find_element_by_name("username")
        search_bar_username.send_keys(username_input)

        nextButton = driver.find_elements_by_xpath('//*[@id="login-signin"]')
        nextButton[0].click()

        search_bar_password = driver.find_element_by_name("password")
        search_bar_password.send_keys(password_input)

        nextButton = driver.find_elements_by_xpath('//*[@id="login-signin"]')
        nextButton[0].click()
        print('Login Successful...!!')
    except:
        print('Login Failed')


def navigate_to_settings():
    try:
        mailbox = driver.find_elements_by_xpath('//*[@id="ybarMailLink"]')
        mailbox[0].click()
        driver.get("https://mail.yahoo.com/d/settings/1")
        driver.implicitly_wait(15)
        print('Found settings')
    except:
        print('Cannot Find Settings')


def create_catchall():
    try:
        for i in range(number_of_catchalls):
            add_button = driver.find_elements_by_xpath(
                '//*[@id="mail-app-component"]/section/div/article/div/div[1]/div/div[3]/div[3]/button')
            add_button[0].click()
            first_name = names.get_first_name()
            last_name = names.get_last_name()
            catchall = driver.find_element_by_name('KEYWORD')
            catchall.send_keys(first_name + last_name)
            driver.implicitly_wait(15)
            save_button = driver.find_elements_by_xpath(
                '//*[@id="mail-app-component"]/section/div/article/div/div[2]/div/div/div/div[4]/button[1]')
            save_button[0].click()
            driver.implicitly_wait(15)
            sheet1.write(i, 0, user_baseyahoo + '-' + first_name + last_name + '@yahoo.com')
            sheet1.write(i, 1, first_name + ' ' + last_name)
    except:
        print('There was an error')


username_input = input('Please Enter Your Yahoo Email')
password_input = input('Please Enter Your Yahoo Password')

login()
navigate_to_settings()

user_baseyahoo = input('What is Your Base Yahoo? Do not include the dash or this program won\'t work.')
number_of_catchalls = int(input('How many catchalls do you want to create? (Max is 500)'))

create_catchall()
filename = input('What do you want your file to be named?')
spreadsheetfile = filename+'.xls'
spreadsheet.save(spreadsheetfile)
print('Completed')
