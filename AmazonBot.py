#!/usr/bin/env python3
"""
    *******************************************************************************************
    AmazonBot.
    Author: Ali Toori, Python Developer [Bot Builder]
    Website: https://boteaz.com
    YouTube: https://youtube.com/@AliToori
    *******************************************************************************************
"""
import os
import time
import random
import ntplib
import datetime
import pyfiglet
import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class Amazon:

    def __init__(self):
        self.first_time = True
        self.PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))

    # Get random user-agent
    def get_random_user_agent(self):
        user_agents_list = []
        with open('user_agents.txt') as f:
            content = f.readlines()
        user_agents_list = [x.strip() for x in content]
        return random.choice(user_agents_list)

    # Login to the website for smooth processing
    def login(self, url, email, password):
        # For MacOS chromedriver path
        # PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
        # DRIVER_BIN = os.path.join(PROJECT_ROOT, "chromedriver")
        options = webdriver.ChromeOptions()
        options.add_argument("--incognito")
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument(F'--user-agent={self.get_random_user_agent()}')
        # options.add_argument('--headless')
        # driver = webdriver.Chrome(executable_path=DRIVER_BIN, options=options)
        driver = webdriver.Chrome(options=options)
        print('Signing-in to the website... ')
        # # Login to the website
        # if os.path.isfile(r'AmazonRes\\Cookies.pkl'):
        #     driver.get(url)
        #     print('Loading cookies...')
        #     try:
        #         cookies = pickle.load(open(r'AmazonRes\\Cookies.pkl', 'rb'))
        #         for cookie in cookies:
        #             driver.add_cookie(cookie)
        #     except UnableToSetCookieException as exc:
        #         driver.refresh()
        #         driver.get(url)
        # else:
        driver.get(url)
        # sleep(60)
        WebDriverWait(driver, 1000).until(EC.visibility_of_element_located((By.NAME, 'email')))
        sleep(1)
        driver.find_element_by_name('email').send_keys(email)
        sleep(1)
        driver.find_element_by_name('password').send_keys(password)
        sleep(1)
        driver.find_element_by_id('signInSubmit').click()
        print('Please fill the required fields and continue !')
        WebDriverWait(driver, 1000).until(EC.visibility_of_element_located((By.ID, 'sc-logo-asset')))
        # Store user cookies for later use
        # pickle.dump(driver.get_cookies(), open(r'AmazonRes\\Cookies.pkl', 'wb'))
        # print('Cookies has been saved...')
        return driver

    def get_data(self, driver, url, file_path, brand):
        file_path_sales = os.path.join(self.PROJECT_ROOT, 'AmazonRes/')
        today = datetime.datetime.today().strftime('a%m-a%d-%y').replace('a0', 'a').replace('a', '')
        yesterday = (datetime.datetime.now() - datetime.timedelta(1)).strftime('a%m-a%d-%y').replace('a0', 'a').replace('a', '')
        business_report_ext = '\BusinessReport-'
        file_path_today = file_path + business_report_ext + today + '.csv'
        file_path_yesterday = file_path + business_report_ext + yesterday + '.csv'
        if brand == 'Ecotero':
            driver.get(
                'https://sellercentral.amazon.com/home?cor=mmp_NA&mons_sel_dir_mcid=amzn1.merchant.d.ACOXVO2RGA3E3J3DACR3YGD3GXWA&mons_sel_mkid=ATVPDKIKX0DER')
        elif brand == 'Inspiratek':
            driver.get(
                'https://sellercentral.amazon.com/home?cor=mmp_NA&mons_sel_dir_mcid=amzn1.merchant.d.AAI6NGADS5EZQRXZADF5AB4D7DJQ&mons_sel_mkid=ATVPDKIKX0DER')
        WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.ID, 'sc-mkt-picker-switcher-select')))
        driver.get(url=url)
        try:
            WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.ID, 'export')))
        except:
            WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.ID, 'export')))
        try:
            WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CLASS_NAME, 'actionsArrow')))
            export = driver.find_element_by_class_name('actionsArrow')
            export.click()
        except:
            WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CLASS_NAME, 'actionsArrow')))
            export = driver.find_element_by_class_name('actionsArrow')
            export.click()
        WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.ID, 'downloadCSV')))
        driver.find_element_by_id('downloadCSV').click()
        print('Business Report file is being downloaded:', datetime.datetime.today())
        sleep(5)
        if os.path.isfile(file_path_today):
            file_path_final = file_path_today
        elif os.path.isfile(file_path_yesterday):
            file_path_final = file_path_yesterday
        else:
            return 0
        print('File has been downloaded to:', file_path_final)
        df = pd.read_csv(file_path_final, index_col=None)
        # Extract each row to its respective Sales File
        day_before_yesterday = (datetime.datetime.now() - datetime.timedelta(2)).strftime('%m/%d/%Y')
        for item in df.iloc:
            product_sku = str(item['SKU'])
            # Building new DataFrame from each row of downloaded/input file
            if brand == 'Inspiratek':
                df_new = {'Date': [day_before_yesterday],
                          '(Parent) ASIN': [item['(Parent) ASIN']],
                          '(Child) ASIN': [item['(Child) ASIN']], 'Title': [item['Title']], 'SKU': [item['SKU']],
                          'Sessions': [item['Sessions']], 'Session Percentage': [item['Session Percentage']],
                          'Page Views': [item['Page Views']],
                          'Page Views Percentage': [item['Page Views Percentage']],
                          'Buy Box Percentage': [item['Buy Box Percentage']],
                          'Units Ordered': [item['Units Ordered']],
                          'Units Ordered - B2B': [''],
                          'Unit Session Percentage': [item['Unit Session Percentage']],
                          'Unit Session Percentage - B2B': [''],
                          'Ordered Product Sales': [item['Ordered Product Sales']],
                          'Ordered Product Sales - B2B': [''],
                          'Total Order Items': [item['Total Order Items']]}
            else:
                df_new = {'Date': [day_before_yesterday],
                          '(Parent) ASIN': [item['(Parent) ASIN']],
                          '(Child) ASIN': [item['(Child) ASIN']], 'Title': [item['Title']], 'SKU': [item['SKU']],
                          'Sessions': [item['Sessions']], 'Session Percentage': [item['Session Percentage']],
                          'Page Views': [item['Page Views']],
                          'Page Views Percentage': [item['Page Views Percentage']],
                          'Buy Box Percentage': [item['Buy Box Percentage']],
                          'Units Ordered': [item['Units Ordered']],
                          'Units Ordered - B2B': [item['Units Ordered - B2B']],
                          'Unit Session Percentage': [item['Unit Session Percentage']],
                          'Unit Session Percentage - B2B': [item['Unit Session Percentage - B2B']],
                          'Ordered Product Sales': [item['Ordered Product Sales']],
                          'Ordered Product Sales - B2B': [item['Ordered Product Sales - B2B']],
                          'Total Order Items': ['']}
            df = pd.DataFrame(df_new)
            file_path_sales = r'AmazonRes\\' + brand + '-' + product_sku + '.csv'
            # if file does not exist write header
            if not os.path.isfile(file_path_sales) or os.path.getsize(file_path_sales) == 0:
                df.to_csv(file_path_sales, index=False)
            else:  # else if exists so append without writing the header
                df.to_csv(file_path_sales, mode='a', header=False, index=False)
        os.remove(file_path_final)
        # Put zeros for missing records
        file_path_sales = os.path.join(self.PROJECT_ROOT, 'AmazonRes/')
        for file in os.listdir(file_path_sales):
            if (file.startswith('Ecotero') or file.startswith('Inspiratek')) and file.endswith('.csv'):
                print(file)
                df = pd.read_csv(file_path_sales + file, index_col=None)
                if df.iloc[-1]['Date'] != day_before_yesterday:
                    df_new = {'Date': [day_before_yesterday],
                              '(Parent) ASIN': [df.iloc[0]['(Parent) ASIN']],
                              '(Child) ASIN': [df.iloc[0]['(Child) ASIN']], 'Title': [df.iloc[0]['Title']],
                              'SKU': [df.iloc[0]['SKU']], 'Sessions': ['0'], 'Session Percentage': ['0'],
                              'Page Views': ['0'],
                              'Page Views Percentage': ['0'], 'Buy Box Percentage': ['0'], 'Units Ordered': ['0'],
                              'Units Ordered - B2B': ['0'], 'Unit Session Percentage': ['0'],
                              'Unit Session Percentage - B2B': ['0'],
                              'Ordered Product Sales': ['0'], 'Ordered Product Sales - B2B': ['0'],
                              'Total Order Items': ['0']}
                    df = pd.DataFrame(df_new)
                    df.to_csv(file_path_sales + file, mode='a', header=False, index=False)
        print('Reports have been extracted from file:', file_path, datetime.datetime.today())

    def get_product(self, url_ecotero, url_inpiratek, file_date):
        with open(r'AmazonRes\User_Account.txt') as f:
            content = f.readlines()
        user = [x.strip() for x in content][0].split(':', maxsplit=2)
        file_path = user[2]
        # Login to the website
        driver = self.login(url_ecotero, email=user[0], password=user[1])
        ecotero_done = False
        inspiratek_done = False
        while True:
            if ecotero_done and inspiratek_done:
                print("Both Ecotero and Inspiratek reports are already extracted")
                break
            print('************************************************************************\nPlease choose an option\n')
            user_input = input('Please Enter 1 for Ecotero and 2 for Inspiratek\n')
            print('************************************************************************\n')

            if not ecotero_done:
                if str(user_input) == '1':
                    self.get_data(driver=driver, url=url_ecotero, file_path=file_path, brand='Ecotero')
                    ecotero_done = True
            elif ecotero_done and str(user_input) == '1':
                print(" Ecotero reports already extracted")
            if not inspiratek_done:
                if str(user_input) == '2':
                    self.get_data(driver=driver, url=url_inpiratek, file_path=file_path, brand='Inspiratek')
                    inspiratek_done = True
            elif inspiratek_done and str(user_input) == '2':
                print("Inspiratek reports already extracted")
            else:
                print("Wrong choice, please try again")
        print("Browser will be auto-closed after 5 minutes")
        sleep(300)
        self.finish(driver)

    def finish(self, driver):
        try:
            driver.close()
            driver.quit()
        except WebDriverException as exc:
            print('Problem occurred while closing the WebDriver instance ...', exc.args)


def main():
    amazon = Amazon()
    # Getting Day before Yesterday
    day_before_yesterday = (datetime.datetime.now() - datetime.timedelta(2)).strftime('%m/%d/%Y')
    today = datetime.datetime.today().strftime('%m/%d/%Y')
    print('Date:', today)
    print('Processing file downloading for:', day_before_yesterday)
    filter_parent = 'ByParentItem'
    filter_sku = 'BySKU'
    # The main homepage URL (AKA base URL)ere
    url_ecotero = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10/c11/c12/c13/c14&sortColumn=15&filterFromDate=' \
                  + day_before_yesterday + '&filterToDate=' + day_before_yesterday + \
                  '&fromDate=' + day_before_yesterday + '&toDate=' + day_before_yesterday + \
                  '&reportID=102:DetailSalesTraffic' + filter_sku + '&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate='
    ur_inspiratek = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10/c11/c12&sortColumn=13&filterFromDate=' \
                    + day_before_yesterday + '&filterToDate=' + day_before_yesterday + \
                    '&fromDate=' + day_before_yesterday + '&toDate=' + day_before_yesterday + \
                    '&reportID=102:DetailSalesTraffic' + filter_sku + '&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate='
    # try:
    print(url_ecotero)
    amazon.get_product(url_ecotero=url_ecotero, url_inpiratek=ur_inspiratek, file_date=day_before_yesterday)
    # except WebDriverException as exc:
    #     print('Exception in WebDriver:', exc.msg)


# Trial version logic
def trial(trial_date):
    ntp_client = ntplib.NTPClient()
    response = ntp_client.request('pool.ntp.org')
    local_time = time.localtime(response.ref_time)
    current_date = time.strftime('%Y-%m-%d %H:%M:%S', local_time)
    current_date = datetime.datetime.strptime(current_date, '%Y-%m-%d %H:%M:%S')
    return trial_date > current_date


if __name__ == '__main__':
    trial_date = datetime.datetime.strptime('2020-10-07 23:59:59', '%Y-%m-%d %H:%M:%S')
    # Print ASCII Art
    print('************************************************************************\n')
    pyfiglet.print_figlet('____________               AmazonBot ____________\n', colors='RED')
    print('Author: Ali Toori, Python Developer [Web-Automation Bot Developer]\n'
          'Author: Ali Toori, Python Developer [Bot Builder]\n',
          'Website: https://boteaz.com\n',
          'YouTube: https://youtube.com/@AliToori',
          '************************************************************************')
    # Trial version logic
    if trial(trial_date):
        # print("Your trial will end on: ",
        #       str(trial_date) + ". To get full version, please contact fiverr.com/AliToori !")
        main()
    else:
        pass
        # print("Your trial has been expired, To get full version, please contact fiverr.com/AliToori !")

