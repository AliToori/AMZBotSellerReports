#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    *******************************************************************************************
    AMZBot: Amazon reports automation bot
    Author: Ali Toori, Python Developer [Bot Builder]
    Website: https://boteaz.com
    YouTube: https://youtube.com/@AliToori
    *******************************************************************************************

"""
import csv
import logging.config
import os
import pickle
import random
import time
from datetime import datetime, timedelta
from multiprocessing import freeze_support
from pathlib import Path
from time import sleep
import gspread
import ntplib
import pandas as pd
import pyfiglet
from shutil import copyfile
import undetected_chromedriver as uc
from amazoncaptcha import AmazonCaptcha
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logging.config.dictConfig({
    "version": 1,
    "disable_existing_loggers": False,
    'formatters': {
        'colored': {
            '()': 'colorlog.ColoredFormatter',  # colored output
            # --> %(log_color)s is very important, that's what colors the line
            'format': '[%(asctime)s,%(lineno)s] %(log_color)s[%(message)s]',
            'log_colors': {
                'DEBUG': 'green',
                'INFO': 'cyan',
                'WARNING': 'yellow',
                'ERROR': 'red',
                'CRITICAL': 'bold_red',
            },
        },
        'simple': {
                'format': '[%(asctime)s,%(lineno)s] [%(message)s]',
            },
    },
    "handlers": {
        "console": {
            "class": "colorlog.StreamHandler",
            "level": "INFO",
            "formatter": "colored",
            "stream": "ext://sys.stdout"
        },
        "file": {
            "class": "logging.handlers.RotatingFileHandler",
            "level": "INFO",
            "formatter": "simple",
            "filename": "AMZBot_logs.log"
        },
    },
    "root": {"level": "INFO",
             "handlers": ["console", "file"]
             }
})

LOGGER = logging.getLogger()


class AMZBot:
    def __init__(self):
        # self.first_time = True
        self.logged_in_email = None
        self.logged_in = False
        self.driver = None
        self.stopped = False
        self.account_switched = False
        # self.update_type = 'Append'
        self.PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
        self.PROJECT_ROOT = Path(self.PROJECT_ROOT)
        # start_date = str((datetime.now() - timedelta(7)).strftime('%m/%d/%Y'))
        # end_date = str((datetime.now() - timedelta(0)).strftime('%m/%d/%Y'))
        self.AMZ_HOME_URL = 'https://www.amazon.com/'
        self.AMZ_SIGNEDIN_URL = 'https://www.amazon.com/?ref_=nav_signin&'
        self.AMZ_SELLER_SIGNIN_URL = 'https://www.amazon.com/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F%3Fref_%3Dnav_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=usflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&'
        self.AMZ_SELLER_REPORT_URL = 'https://sellercentral.amazon.com/gp/site-metrics/report.html'
        self.API_Key = 'AIzaSyBg0NfgqOTpmrNXh__zLuRjUCRMUDqBud8'
        # self.file_account = str(self.PROJECT_ROOT / 'AMZRes/Accounts.csv')
        self.file_account = str(self.PROJECT_ROOT / 'AMZRes/Accounts.xlsx')
        self.directory_downloads = str(self.PROJECT_ROOT / 'AMZRes/Downloads/')
        self.file_local_reports = str(self.PROJECT_ROOT / 'AMZRes/Reports/BusinessReports/Reports.csv')
        # self.file_ad_reports = str(self.PROJECT_ROOT / 'AMZRes/Reports/AdvertisingReports/Reports.csv')
        # self.file_fulfillment_reports = str(self.PROJECT_ROOT / 'AMZRes/Reports/FulfillmentReports/Reports.csv')
        # self.file_vendor_reports = str(self.PROJECT_ROOT / 'AMZRes/Reports/VendorReports/Reports.csv')
        # self.file_vendor_promo_reports = str(self.PROJECT_ROOT / 'AMZRes/Reports/VendorPromotionalReports/Reports.csv')
        self.file_client_secret = str(self.PROJECT_ROOT / 'AMZRes/client_secret_victor.json')


    @staticmethod
    def enable_cmd_colors():
        # Enables Windows New ANSI Support for Colored Printing on CMD
        from sys import platform
        if platform == "win32":
            import ctypes
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

    # Trial version logic
    @staticmethod
    def trial(trial_date):
        ntp_client = ntplib.NTPClient()
        try:
            response = ntp_client.request('pool.ntp.org')
            local_time = time.localtime(response.ref_time)
            current_date = time.strftime('%Y-%m-%d %H:%M:%S', local_time)
            current_date = datetime.strptime(current_date, '%Y-%m-%d %H:%M:%S')
            return trial_date > current_date
        except:
            pass

    # Get random user agent
    def get_user_agent(self):
        file_uagents = str(self.PROJECT_ROOT / 'AMZRes/user_agents.txt')
        with open(file_uagents) as f:
            content = f.readlines()
        u_agents_list = [x.strip() for x in content]
        return random.choice(u_agents_list)

    # Get web driver
    def get_driver(self, headless=False):
        LOGGER.info(f'Launching chrome browser')
        # For absolute chromedriver path
        DRIVER_BIN = str(self.PROJECT_ROOT / 'AMZRes/bin/chromedriver.exe')
        user_dir = str(self.PROJECT_ROOT / 'AMZRes/UserData')
        options = uc.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-extensions')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument(f"--user-data-dir={user_dir}")
        prefs = {f'profile.default_content_settings.popups': 0,
                 f'download.default_directory': f'{self.directory_downloads}',  # IMPORTANT - ENDING SLASH IS VERY IMPORTANT
                 "directory_upgrade": True,
                 "credentials_enable_service": False,
                 "profile.password_manager_enabled": False}
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument(f'--user-agent={self.get_user_agent()}')
        if headless:
            options.add_argument('--headless')
        driver = uc.Chrome(executable_path=DRIVER_BIN, options=options)
        return driver

    # Finish and quit browser
    def finish(self):
        LOGGER.info(f'Closing the browser instance')
        try:
            if not self.stopped:
                self.driver.close()
                self.driver.quit()
                self.stopped = True
        except:
            LOGGER.info(f'Problem occurred while closing the browser instance')

    @staticmethod
    def wait_until_clickable(driver, xpath=None, element_id=None, name=None, class_name=None, tag_name=None, css_selector=None, duration=10000, frequency=0.01):
        if xpath:
            WebDriverWait(driver, duration, frequency).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        elif element_id:
            WebDriverWait(driver, duration, frequency).until(EC.element_to_be_clickable((By.ID, element_id)))
        elif name:
            WebDriverWait(driver, duration, frequency).until(EC.element_to_be_clickable((By.NAME, name)))
        elif class_name:
            WebDriverWait(driver, duration, frequency).until(EC.element_to_be_clickable((By.CLASS_NAME, class_name)))
        elif tag_name:
            WebDriverWait(driver, duration, frequency).until(EC.visibility_of_element_located((By.TAG_NAME, tag_name)))
        elif css_selector:
            WebDriverWait(driver, duration, frequency).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector)))

    @staticmethod
    def wait_until_visible(driver, xpath=None, element_id=None, name=None, class_name=None, tag_name=None, css_selector=None, duration=10000, frequency=0.01):
        if xpath:
            WebDriverWait(driver, duration, frequency).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        elif element_id:
            WebDriverWait(driver, duration, frequency).until(EC.visibility_of_element_located((By.ID, element_id)))
        elif name:
            WebDriverWait(driver, duration, frequency).until(EC.visibility_of_element_located((By.NAME, name)))
        elif class_name:
            WebDriverWait(driver, duration, frequency).until(
                EC.visibility_of_element_located((By.CLASS_NAME, class_name)))
        elif tag_name:
            WebDriverWait(driver, duration, frequency).until(EC.visibility_of_element_located((By.TAG_NAME, tag_name)))
        elif css_selector:
            WebDriverWait(driver, duration, frequency).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))

    # Returns a range of dates between two dates e.g. 01/01/2018 to 05/21/2021
    @staticmethod
    def get_date_range(account):
            date_list = []
            date_range = str(account["DateRange"])
            LOGGER.info(f"Selecting date range as: {date_range}")
            if date_range == "Custom Date Range":
                start_date = account["StartDate"]
                end_date = account["EndDate"]
                dates = pd.date_range(start_date, end_date)
                for date in dates:
                    date_list.append(datetime.strftime(date, '%m/%d/%Y'))
                return date_list
            # Returns date for today
            elif date_range == "Today":
                today = str((datetime.now() - timedelta(0)).strftime('%m/%d/%Y'))
                date_list.append(today)
                return date_list
            # Returns dates for a week
            elif date_range == "Last 7 Days":
                start_date = str((datetime.now() - timedelta(6)).strftime('%m/%d/%Y'))
                end_date = str((datetime.now() - timedelta(0)).strftime('%m/%d/%Y'))
                dates = pd.date_range(start_date, end_date)
                for date in dates:
                    date_list.append(datetime.strftime(date, '%m/%d/%Y'))
                return date_list
            # Returns dates for a 14 days
            elif date_range == "Last 14 Days":
                start_date = str((datetime.now() - timedelta(13)).strftime('%m/%d/%Y'))
                end_date = str((datetime.now() - timedelta(0)).strftime('%m/%d/%Y'))
                dates = pd.date_range(start_date, end_date)
                for date in dates:
                    date_list.append(datetime.strftime(date, '%m/%d/%Y'))
                return date_list
            # Returns dates for a month
            elif date_range == "Last 30 Days":
                start_date = str((datetime.now() - timedelta(29)).strftime('%m/%d/%Y'))
                end_date = str((datetime.now() - timedelta(0)).strftime('%m/%d/%Y'))
                dates = pd.date_range(start_date, end_date)
                for date in dates:
                    date_list.append(datetime.strftime(date, '%m/%d/%Y'))
                return date_list

    # Logs-in through captcha
    def captcha_login(self, email, password):
        LOGGER.info(f'Waiting for possible CAPTCHA')
        try:
            self.wait_until_visible(driver=self.driver, name='password', duration=5)
            LOGGER.info(f'Empty password found')
            sleep(1)
            self.driver.find_element_by_name('password').send_keys(password)
            LOGGER.info(f'Password has been entered')
            sleep(0.5)
        except:
            pass
        try:
            self.wait_until_visible(driver=self.driver, element_id='auth-captcha-image-container', duration=5)
            LOGGER.info(f'CAPTCHA found')
        except:
            try:
                self.wait_until_visible(driver=self.driver, element_id='captchacharacters', duration=5)
                LOGGER.info(f'CAPTCHA found')
            except:
                pass
        if 'the characters you see' in str(self.driver.find_elements_by_tag_name('h4')[-1].text) or 'the characters you see' in str(self.driver.find_elements_by_tag_name('h4')[-2].text):
            LOGGER.info(f'Solving CAPTCHA')
            captcha = AmazonCaptcha.fromdriver(self.driver)
            solution = captcha.solve()
            sleep(1)
            # Check if captcha was not yet solved
            if 'Not solved' in solution:
                LOGGER.info(f'CAPTCHA not yet solved')
                LOGGER.info(f'Please enter the CAPTCHA manually')
                while 'the characters you see' in str(self.driver.find_elements_by_tag_name('h4')[-1].text) or 'the characters you see' in str(self.driver.find_elements_by_tag_name('h4')[-2].text):
                    LOGGER.info(f'Please enter the CAPTCHA characters')
                    sleep(5)
                return
            try:
                self.driver.find_element_by_id('captchacharacters').send_keys(solution)
            except:
                self.driver.find_element_by_id('auth-captcha-guess').send_keys(solution)
            sleep(2)
            try:
                self.driver.find_element_by_class_name('a-button-text').click()
            except:
                self.driver.find_element_by_id('signInSubmit').click()
            self.logged_in = True
            self.logged_in_email = email
            LOGGER.info(f'CAPTCHA has been solved successfully')

    # Logs-in to the Amazon
    def login_amazon(self, account):
        # account_num = str(account["AccountNo"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        LOGGER.info(f'Signing-in to Amazon account: {email}')
        cookies = 'Cookies_' + email + '.pkl'
        cookies_file_path = self.PROJECT_ROOT / 'AMZRes' / cookies
        # Try signing-in via cookies
        if os.path.isfile(cookies_file_path):
            LOGGER.info(f'Requesting Amazon: {str(self.AMZ_HOME_URL)} Account: {email}')
            self.driver.get(self.AMZ_HOME_URL)
            LOGGER.info(f'Loading cookies ... Account: {email}')
            with open(cookies_file_path, 'rb') as cookies_file:
                cookies = pickle.load(cookies_file)
            for cookie in cookies:
                self.driver.add_cookie(cookie)
            # self.driver.get(self.AMZ_SELLER_SIGNIN_URL)
            try:
                # self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
                self.wait_until_visible(driver=self.driver, element_id='nav-logo-sprites', duration=5)
                LOGGER.info(f'Cookies login successful ... Account: {email}')
                self.driver.refresh()
                self.logged_in = True
                self.logged_in_email = email
                return
            except:
                LOGGER.info(f'Cookies login failed ... Account: {email}')
                self.captcha_login(email=email, password=password)
                return
        # Try sign-in normally using credentials
        LOGGER.info(f'Requesting Amazon: {str(self.AMZ_SELLER_SIGNIN_URL)} Account: {email}')
        self.driver.get(self.AMZ_SELLER_SIGNIN_URL)
        LOGGER.info('Filling credentials')
        sleep(3)
        # Try and switch amazon account if logged-in email is different than the one being logged-in
        try:
            self.wait_until_visible(driver=self.driver, css_selector='.a-row.a-size-base.a-color-tertiary.auth-text-truncate', duration=5)
            logged_in_email = str(self.driver.find_element_by_css_selector('.a-row.a-size-base.a-color-tertiary.auth-text-truncate').text).strip()
            if logged_in_email != email:
                LOGGER.info(f'Switching amazon account')
                self.driver.find_element_by_id('ap_switch_account_link').click()
                sleep(1)
                self.wait_until_visible(driver=self.driver, css_selector='.a-row.a-size-base-plus.a-grid-vertical-align.a-grid-center', duration=5)
                self.driver.find_element_by_css_selector('.a-row.a-size-base-plus.a-grid-vertical-align.a-grid-center').click()
                sleep(1)
                LOGGER.info(f'Adding new account')
                sleep(1)
        except:
            pass
        # Try filling email and password in case of first time logging-in
        try:
            self.wait_until_visible(driver=self.driver, name='email', duration=5)
            sleep(1)
            self.driver.find_element_by_name('email').send_keys(email)
            sleep(1)
            self.driver.find_element_by_id('continue').click()
            self.wait_until_visible(driver=self.driver, name='password', duration=5)
            self.driver.find_element_by_name('password').send_keys(password)
            sleep(1)
            self.driver.find_element_by_name('rememberMe').click()
            self.driver.find_element_by_id('signInSubmit').click()
        except:
            pass
        # Try send OTP if asked for
        try:
            self.wait_until_visible(driver=self.driver, element_id='auth-send-code', duration=3)
            self.driver.find_element_by_id('auth-send-code').click()
            LOGGER.info('OTP has been sent to the phone number')
            LOGGER.info('Please fill the OTP and continue !')
        except:
            pass
        # Try filling password if occurs again after OTP
        try:
            self.wait_until_visible(driver=self.driver, name='password', duration=3)
            LOGGER.info('Password filling')
            self.driver.find_element_by_name('password').send_keys(password)
            LOGGER.info('Please fill the captcha and continue')
        except:
            pass
        # self.wait_until_visible(driver=driver, element_id='auth-send-code')
        # driver.find_element_by_id('auth-send-code').click()
        try:
            self.wait_until_visible(driver=self.driver, name='rememberDevice')
            self.driver.find_element_by_name('rememberDevice').click()
            # self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset')
            self.wait_until_visible(driver=self.driver, element_id='nav-logo-sprites')
            self.logged_in = True
            self.logged_in_email = email
            LOGGER.info('Successful sign-in')
            # Store user cookies for later use
            LOGGER.info(f'Saving cookies for: Account: {email}')
            with open(cookies_file_path, 'wb') as cookies_file:
                pickle.dump(self.driver.get_cookies(), cookies_file)
            LOGGER.info(f'Cookies have been saved: Account: {email}')
        except:
            pass

    # Logs-in to Seller Central
    def login_seller_central(self, email, password):
        try:
            LOGGER.info(f'Waiting for Seller Central')
            self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
            LOGGER.info(f'Seller Central login successful')
            self.logged_in = True
            self.logged_in_email = email
        except:
            # Try if seller central needs sign-in
            try:
                LOGGER.info(f'Signing-in to Seller Central')
                self.wait_until_visible(driver=self.driver, name='password', duration=5)
                sleep(3)
                self.driver.find_element_by_name('password').send_keys(password)
                sleep(1)
                self.driver.find_element_by_name('rememberMe').click()
                sleep(3)
                self.driver.find_element_by_id('signInSubmit').click()
                LOGGER.info(f'Waiting for Seller Central')
                self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
                LOGGER.info(f'Seller Central login successful')
                self.logged_in = True
                self.logged_in_email = email
            except:
                try:
                    LOGGER.info(f'Seller Central not yet logged-in')
                    self.captcha_login(email=email, password=password)
                    self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
                    LOGGER.info(f'Seller Central login successful')
                    self.logged_in = True
                    self.logged_in_email = email
                except:
                    pass

    # returns file's path in a the downloads folder/directory
    def get_file_download(self, directory_downloads):
        file_list = [str(self.PROJECT_ROOT / f'AMZRes/Downloads/{f}') for f in
                     os.listdir(directory_downloads) if
                     os.path.isfile(str(self.PROJECT_ROOT / f'AMZRes/Downloads/{f}'))]
        if len(file_list) == 1:
            file_download = file_list[0]
        else:
            file_download = file_list[-1]
        LOGGER.info(f'File has been retrieved : {file_download}')
        if os.path.isfile(file_download):
            return file_download
        else:
            LOGGER.info(f'Failed to download the file')

    # returns file's path in a the downloads folder/directory
    def clear_downloads_directory(self, directory_downloads):
        LOGGER.info(f'Checking Downloads Directory ...')
        # Check if Downloads Directory is there
        if os.path.isdir(directory_downloads):
            file_list = [str(self.PROJECT_ROOT / f'AMZRes/Downloads/{f}') for f in os.listdir(directory_downloads) if os.path.isfile(str(self.PROJECT_ROOT / f'AMZRes/Downloads/{f}'))]
            if len(file_list) > 0:
                LOGGER.info(f'Clearing Downloads Directory ...')
                [os.remove(f) for f in file_list]
                LOGGER.info(f'Downloads Directory has been cleared')
            else:
                LOGGER.info(f'Downloads Directory is clear')
        else:
            LOGGER.info(f'Downloads Directory Not Found')

    # Update reports file with StartDate and EndDate
    def update_csv_report(self, csv_path, start_date, end_date, promo=None, sales_dashboard=None):
        LOGGER.info(f'Updating reports file with {start_date} and {end_date}: {csv_path}')
        df = ''
        if csv_path.endswith('.csv'):
            # If reports are promo, read it as sep=\t as there is an extra line aggregate without its column name
            if promo:
                df = pd.read_csv(csv_path, index_col=None, sep='\t')
            # If reports are sales-dashboard, skip to row # 40
            elif sales_dashboard:
                # df = pd.read_csv(csv_path, index_col=None, skiprows=40)
                df = self.get_sales_dashboard_df(file_download=csv_path)
            else:
                df = pd.read_csv(csv_path, index_col=None)
        elif csv_path.endswith('.xlsx'):
            df = pd.read_excel(csv_path, index_col=None)
        elif csv_path.endswith('.tsv') or csv_path.endswith('.txt'):
            df = pd.read_csv(csv_path, index_col=None, sep='\t')
        # Skip, if DataFrame was empty
        if len(df) < 1:
            LOGGER.info(f'Reports file is empty: {csv_path}')
            LOGGER.info(f'Skipping SpreadSheet update')
            return
        # Extract each row to its respective Sales File
        # Using DataFrame.insert() to add a column at start of a DataFrame/CSV
        df.insert(loc=0, column="StartDate", value=[start_date for i in range(len(df))], allow_duplicates=False)
        df.insert(loc=1, column="EndDate", value=[end_date for i in range(len(df))], allow_duplicates=False)
        # Update the reports in downloads directory
        df.to_csv(csv_path, index=False)
        LOGGER.info(f'Reports have been updated: {csv_path}')

    # Uploads or Appends reports to the SpreadSheet
    def csv_to_spreadsheet(self, csv_path, spread_sheet_url, spread_sheet_name, work_sheet_name=None, advertising=None, promo=None, sales_dashboard=None):
        LOGGER.info(f'Updating reports in the SpreadSheet: {spread_sheet_name}: WorkSheet: {work_sheet_name}: URL: {spread_sheet_url}')
        df = ''
        if csv_path.endswith('.csv'):
            # If reports are promo, read it as sep=\t as there is an extra line aggregate without its column name
            if promo:
                df = pd.read_csv(csv_path, index_col=None, sep='\t')
            # If reports are sales-dashboard, skip to row # 40
            elif sales_dashboard:
                # df = pd.read_csv(csv_path, index_col=None, skiprows=40)
                df = self.get_sales_dashboard_df(file_download=csv_path)
            else:
                df = pd.read_csv(csv_path, index_col=None)
        elif csv_path.endswith('.xlsx'):
            df = pd.read_excel(csv_path, index_col=None)
        elif csv_path.endswith('.tsv') or csv_path.endswith('.txt'):
            df = pd.read_csv(csv_path, index_col=None, sep='\t')
        # Skip, if DataFrame was empty
        if len(df) < 1:
            LOGGER.info(f'Reports file is empty: {csv_path}')
            LOGGER.info(f'Skipping SpreadSheet update')
            return
        # Converts Dates to mm/dd/yy format
        try:
            df[df.columns[0]] = pd.to_datetime(df.columns[0]).dt.strftime('%m/%d/%Y')
        except:
            pass
        # Convert DataFrame to String
        try:
            df = df.applymap(str)
        except:
            pass
        # sheet_key = '1b25493dd39d82f2712c0c33fefeed1bb0d45415'
        # client_email = 'victor@river-bruin-313405.iam.gserviceaccount.com'
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.file_client_secret, scope)
        client = gspread.authorize(credentials)
        # Use create() to create a new blank spreadsheet
        # spreadsheet = client.create('A new spreadsheet')
        # If your email is otto@example.com you can share the newly created spreadsheet with yourself
        # spreadsheet.share('otto@example.com', perm_type='user', role='writer')
        spreadsheet = client.open(spread_sheet_name)
        # worksheet = sh.add_worksheet(title="A worksheet", rows="100", cols="20")
        # Uploads to worksheet by name if provided, otherwise upload to the default (first) worksheet
        if work_sheet_name:
            worksheet = spreadsheet.worksheet(work_sheet_name)
        else:
            worksheet = spreadsheet.get_worksheet(0)
        # Uploads with headers, if WorkSheet is empty
        # len(worksheet.row_values(2))
        if len(worksheet.get_all_records()) == 0:
            worksheet.resize(1)
            worksheet.update([df.columns.values.tolist()] + df.values.tolist())
            # self.update_type = 'Append'
            # account_df = pd.read_csv(self.file_account, index_col=None)
            # account_df.loc[account_df.WorkSheetName == work_sheet_name, "UpdateType"] = "Append"
            # account_df.to_csv(self.file_account, index=False)
            LOGGER.info(f'Reports have been uploaded to the SpreadSheet: {spread_sheet_name}: WorkSheet: {work_sheet_name}: URL: {spread_sheet_url}')
            LOGGER.info(f'SpreadSheet has been updated: Name: {spread_sheet_name}, URL: {spread_sheet_url}')
        # Appends the reports to the end of the spreadsheet
        else:
            # Get the spreadsheet as a DataFrame, compare with the uploading csv and drop columns accordingly
            if advertising:
                df_sheet = pd.DataFrame(worksheet.get_all_records())
                cols_sheet = [col for col in df_sheet.columns]
                for col in df.columns:
                    if col not in cols_sheet:
                        df.drop(labels=col, axis=1, inplace=True)
            worksheet.append_rows(values=df.iloc[:].values.tolist())
            LOGGER.info(f'Reports have been appended to the SpreadSheet: Name: {spread_sheet_name},  WorkSheet: {work_sheet_name}: URL: {spread_sheet_url}')
            LOGGER.info(f'SpreadSheet has been updated: Name: {spread_sheet_name}, URL: {spread_sheet_url}')
        # Drop duplicates and sort spreadsheet in-order based on first column
        self.drop_duplicates_sort_spreadsheet(spread_sheet_name=spread_sheet_name, spread_sheet_url=spread_sheet_url, work_sheet_name=work_sheet_name)

    # Drops duplicates and sorts a SpreadSheet
    def drop_duplicates_sort_spreadsheet(self, spread_sheet_url, spread_sheet_name, work_sheet_name=None):
        LOGGER.info(f'Dropping duplicates and sorting SpreadSheet: {spread_sheet_name}: WorkSheet: {work_sheet_name}: URL: {spread_sheet_url}')
        # sheet_key = '1b25493dd39d82f2712c0c33fefeed1bb0d45415'
        # client_email = 'victor@river-bruin-313405.iam.gserviceaccount.com'
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.file_client_secret, scope)
        client = gspread.authorize(credentials)
        spreadsheet = client.open(spread_sheet_name)
        # Uploads to worksheet by name if provided, otherwise upload to the default (first) worksheet
        if work_sheet_name:
            worksheet = spreadsheet.worksheet(work_sheet_name)
        else:
            worksheet = spreadsheet.get_worksheet(0)
        data_frame = pd.DataFrame(worksheet.get_all_records())
        # Sort rows based on first column
        try:
            data_frame.sort_values(data_frame.columns[0], inplace=True)
        except:
            pass
        # Drop duplicates based on all columns
        data_frame = data_frame.drop_duplicates()
        # Convert DataFrame to String
        try:
            data_frame = data_frame.applymap(str)
        except:
            pass
        worksheet.resize(1)
        worksheet.update([data_frame.columns.values.tolist()] + data_frame.values.tolist())
        LOGGER.info(f'Duplicates have be dropped and sorted the SpreadSheet: {spread_sheet_name}: WorkSheet: {work_sheet_name}: URL: {spread_sheet_url}')
        LOGGER.info(f'SpreadSheet has been updated')

    # Saves a local copy of the report file as a csv
    def save_reports_locally(self, file_download, promo=None, sales_dashboard=True):
        LOGGER.info(f'Saving report file locally: {file_download}')
        df = None
        if file_download.endswith('.csv'):
            # If reports are promo, read it as sep=\t as there is an extra line aggregate without its column name
            if promo:
                df = pd.read_csv(file_download, index_col=None, sep='\t')
            elif sales_dashboard:
                # df = pd.read_csv(file_download, index_col=None, skiprows=40)
                copyfile(file_download, self.file_local_reports)
            else:
                df = pd.read_csv(file_download, index_col=None)
        elif file_download.endswith('.xlsx'):
            df = pd.read_excel(file_download, index_col=None)
        elif file_download.endswith('.tsv') or file_download.endswith('.txt'):
            df = pd.read_csv(file_download, index_col=None, sep='\t')
        # Skip, if DataFrame was empty
        if len(df) < 1:
            LOGGER.info(f'Reports ile is empty: {file_download}')
            LOGGER.info(f'Skipping saving local copy')
            return
        # Converts Dates to mm/dd/yy format
        try:
            df[df.columns[0]] = pd.to_datetime(df.columns[0]).dt.strftime('%m/%d/%Y')
        except:
            pass
        # if file does not exist write header
        if not os.path.isfile(self.file_local_reports) or os.path.getsize(self.file_local_reports) == 0:
            df.to_csv(self.file_local_reports, index=False)
        else:  # else if exists so append without writing the header
            df.to_csv(self.file_local_reports, mode='a', header=False, index=False)

    # Removes file from the downloads directory
    @staticmethod
    def remove_file(file_download):
        LOGGER.info(f'Removing file from the downloads directory: {file_download}')
        if os.path.isfile(file_download):
            # Remove the file after extracting data from
            LOGGER.info(f'Removing temporary reports file: {file_download}')
            os.remove(file_download)
            LOGGER.info(f'Temporary reports file has been removed')
        else:
            LOGGER.info(f'Failed to remove the file')

    # Switch account under the main seller-central
    def switch_biz_account(self, account_name, location):
        # client = str(account["Client"]).strip()
        # password = str(account["Password"]).strip()
        # account_name = str(account["Brand"]).strip()
        # location = str(account["Location"]).strip()
        # report_type = str(account["ReportType"]).strip()
        # spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        # spread_sheet_name = str(account["SpreadSheetName"]).strip()
        # work_sheet_name = str(account["WorkSheetName"]).strip()
        LOGGER.info(f'Checking current selected account ...')
        # LOGGER.info(f'Switching sheet to: SpreadSheetURL: {spread_sheet_url}, SpreadSheetName: {spread_sheet_name}')
        self.wait_until_visible(driver=self.driver, element_id='partner-switcher')
        account_switcher = self.driver.find_element_by_id('partner-switcher')
        account_switcher_text = str(account_switcher.text).replace("\n", " ").strip()
        LOGGER.info(f'Current selected account: {account_switcher_text}')
        if account_name in account_switcher_text and location in account_switcher_text:
            LOGGER.info(f'Continuing with selected account')
            return
        # Switch the account if not selected
        LOGGER.info(f'Switching account to: {account_name}, Location: {location}')
        account_switcher.click()
        sleep(3)
        try:
            LOGGER.info(f'Waiting for account-switcher ...')
            self.wait_until_visible(driver=self.driver, class_name='scrollable-content', duration=5)
        except:
            account_switcher.click()
            sleep(1)
            self.wait_until_visible(driver=self.driver, class_name='scrollable-content', duration=5)
        # merchant = self.driver.find_element_by_class_name('merchant-level')
        for partner in self.driver.find_element_by_class_name('scrollable-content').find_elements_by_class_name('partner-level'):
            partner_label = str(partner.find_elements_by_tag_name('label')[1].text).strip()
            LOGGER.info(f'Account name: {partner_label}')
            if account_name in partner_label:
                LOGGER.info(f'Account found: {partner_label}')
                partner.click()
                sleep(1)
                for merchant in partner.find_elements_by_tag_name('li'):
                    if location in str(merchant.text).strip():
                        # print(f'Location: {str(merchant.text).strip()}')
                        merchant.click()
                        LOGGER.info(f'Account is being switched ...')
                        self.wait_until_visible(driver=self.driver, class_name='dropdown-button')
                        self.wait_until_visible(driver=self.driver, element_id='spacecasino-sellercentral-homepage-kpitoolbar', duration=12)
                        LOGGER.info(f'Account has been switched')
                        self.account_switched = True
                        LOGGER.info(f'Requesting: {self.AMZ_SELLER_REPORT_URL}')
                        self.driver.get(self.AMZ_SELLER_REPORT_URL)
                        sleep(3)
                        return
            else:
                LOGGER.info(f"Account not found!")

    # Switch advertising report account
    def switch_ad_account(self, account_name, location):
        # client = account["Client"]).strip()
        # account_name = str(account["Brand"]).strip()
        # location = str(account["Location"]).strip()
        # Check if account is already selected
        ad_account_switch_url = 'https://advertising.amazon.com/accounts'
        # try:
        LOGGER.info(f'Checking current selected account ...')
        sleep(1)
        self.wait_until_visible(driver=self.driver, element_id='AACChromeHeaderAccountDropdown', duration=20)
        account_selected = self.driver.find_element_by_id('AACChromeHeaderAccountDropdown').find_element_by_tag_name('button')
        account_text = str(account_selected.text).replace("\n", " ").strip()
        LOGGER.info(f'Current selected account: {account_text}')
        # print(account_name, account_text, account_name in account_text and location in account_text)
        if account_name in account_text and location in account_text:
            # account_selected.click()
            LOGGER.info(f'Continuing with selected account')
            return None
            # try:
            #     self.wait_until_visible(driver=self.driver, css_selector='.MasterAccountLinkData-sc-1bd8881-19.jWHLsn', duration=20)
            #     account_selected_id = self.driver.find_element_by_css_selector('.MasterAccountLinkData-sc-1bd8881-19.jWHLsn').text
            #     return account_selected_id
            #
            # except:
            #     try:
            #         account_selected_id = self.driver.find_element_by_xpath('//*[@id="top-bar-utilities-acc-selector"]/ul/li[1]/a/span[2]').text
            #         return account_selected_id
            #     except:
            #         pass
        # except:
        #     pass
        # Switch account
        try:
            self.driver.get(ad_account_switch_url)
            # Waits for account list
            LOGGER.info(f'Switching account to: {account_name}, Location: {location}')
            LOGGER.info(f'Waiting for accounts')
            self.wait_until_visible(driver=self.driver, css_selector='.item-row.account', duration=5)
            sleep(1)
            LOGGER.info(f'Accounts have been found')
        except:
            pass
        LOGGER.info(f'Getting account ids')
        sleep(3)
        for account in self.driver.find_elements_by_css_selector('.item-row.account'):
            account_text = str(account.text).replace("\n", " ").strip()
            if account_name in account_text and location in account_text:
                account_url = account.find_element_by_tag_name('a').get_attribute('href')
                account_id = account_url[account_url.find('ENTITY') + 6:]
                LOGGER.info(f'Account has been selected: {account_text}')
                self.account_switched = True
                return account_id
            else:
                LOGGER.info(f'Account not found!')
                first_account = self.driver.find_element_by_css_selector('.item-row.account')
                account_url = first_account.find_element_by_tag_name('a').get_attribute('href')
                account_id = account_url[account_url.find('ENTITY') + 6:]
                LOGGER.info(f'Selecting first account!: {account_text}')
                self.account_switched = True
                return account_id

    # Get filtered url for business reports
    @staticmethod
    def get_filtered_url(url, from_date, to_date):
        from_date_formatted = datetime.strptime(from_date, '%m/%d/%Y').strftime('%Y-%m-%d')
        to_date_formatted = datetime.strptime(to_date, '%m/%d/%Y').strftime('%Y-%m-%d')
        url_filter = url + '&fromDate=' + from_date_formatted + '&toDate=' + to_date_formatted
        # from_date_loc_a = url.find('FromDate')
        # to_date_loc_a = url.find('ToDate')
        # from_date_loc_b = url.find('fromDate')
        # to_date_loc_b = url.find('toDate')
        # from_date_a = url[from_date_loc_a:from_date_loc_a + 19]
        # to_date_a = url[to_date_loc_a:to_date_loc_a + 17]
        # from_date_b = url[from_date_loc_b:from_date_loc_b + 19]
        # to_date_b = url[to_date_loc_b:to_date_loc_b + 17]
        # url_filter = url.replace(from_date_a, 'FromDate=' + from_date).replace(to_date_a, 'ToDate=' + to_date).replace(
        #     from_date_b, 'fromDate=' + from_date).replace(to_date_b, 'toDate=' + to_date)
        return url_filter

    # Retrieves business reports from a seller central account
    def get_business_reports(self, account, dates):
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        biz_report_type = str(account["BizReportType"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving business reports for: {client}')
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/BusinessReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        # seller_central_url = 'https://sellercentral.amazon.com/gp/site-metrics/report.html'
        LOGGER.info(f'Navigating to Seller Central')
        LOGGER.info(f'Requesting: {self.AMZ_SELLER_REPORT_URL}')
        self.driver.get(self.AMZ_SELLER_REPORT_URL)
        sleep(3)
        # If seller central is signed-in
        try:
            LOGGER.info(f'Waiting for Seller Central')
            self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
            LOGGER.info(f'Seller Central login successful')
            self.logged_in = True
            self.logged_in_email = email
        except:
            # Login to Seller Central
            self.login_seller_central(email=email, password=password)
        sleep(3)
        # Switches account
        if not self.account_switched:
            self.switch_biz_account(account_name=account_name, location=location)
        # The filtered URL for CSV Download
        # url_filter_by_sku = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10/c11/c12/c13/c14/c15/c16&sortColumn=17&filterFromDate=' + start_date + '&filterToDate=' + end_date + '&fromDate=' + start_date + '&toDate=' + end_date + '&reportID=102:DetailSalesTraffic' + by_sku + '&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate='
        # url_filter_by_parent_item = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10/c11/c12/c13/c14&sortColumn=15&filterFromDate=' + start_date + '&filterToDate=' + end_date + '&fromDate=' + start_date + '&toDate=' + end_date + '&reportID=102:DetailSalesTraffic' + by_parent_item + '&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate='
        # url_filter_by_child_item = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10//c11/c12/c13/c14/c15&sortColumn=16&filterFromDate=' + start_date + '&filterToDate=' + end_date + '&fromDate=' + start_date + '&toDate=' + end_date + '&reportID=102:DetailSalesTraffic' + by_child_item + '&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate='
        # url_filter_sku = str(self.driver.find_element_by_id('report_DetailSalesTrafficBySKU').get_attribute('href'))
        # url_filter_parent = str(self.driver.find_element_by_id('report_DetailSalesTrafficByParentItem').get_attribute('href'))
        # url_filter_child = str(self.driver.find_element_by_id('report_DetailSalesTrafficByChildItem').get_attribute('href'))
        # url_filters = {'Detail Page Sales and Traffic': 'report_DetailSalesTrafficBySKU',
        #                'Detail Page Sales and Traffic by Parent Item': 'report_DetailSalesTrafficByParentItem',
        #                'Detail Page Sales and Traffic by Child Item': 'report_DetailSalesTrafficByChildItem'}
        # id_filter = url_filters[biz_report_type]
        try:
            LOGGER.info(f'Waiting for report filters')
            self.wait_until_visible(driver=self.driver, tag_name='kat-link', duration=5)
            LOGGER.info(f'Processing report filter: {biz_report_type}')
        except:
            pass
        # Get report filter
        for filter in self.driver.find_elements_by_tag_name('kat-link'):
            filter_link = filter.find_element_by_tag_name('a')
            if str(filter_link.text).strip().lower() == biz_report_type.lower():
                url_selected = filter_link.get_attribute('href')
                self.driver.get(url_selected)
                sleep(3)
                self.wait_until_visible(driver=self.driver, tag_name='kat-link', duration=5)
                break
        url_final = self.driver.current_url
        for date in dates:
            start_date, end_date = date, date
            LOGGER.info(f'Retrieving business reports for: {start_date} - {end_date}')
            LOGGER.info(f'Filtering business reports')
            url_filter_final = self.get_filtered_url(url=url_final, from_date=date, to_date=date)
            LOGGER.info(f'Requesting business report: {url_filter_final}')
            self.driver.get(url=url_filter_final)
            try:
                self.driver.refresh()
                sleep(3)
                LOGGER.info(f'Waiting for data table')
                self.wait_until_visible(driver=self.driver, class_name='css-1m4j3ju')
                data_table = self.driver.find_element_by_class_name('css-1m4j3ju')
                if len(data_table.text) == 0:
                    LOGGER.info(f'Report Is Empty!')
                    LOGGER.info(f'Continuing from the next date ...')
                    continue
            except:
                pass
            # Try clicking download button
            try:
                LOGGER.info(f'Waiting for download button')
                self.wait_until_visible(driver=self.driver, tag_name='kat-button', duration=3)
                self.driver.find_element_by_tag_name('kat-button').find_element_by_tag_name('button').click()
            except:
                try:
                    print(f"DOWNLOAD BUTTON SECOND TRY")
                    self.wait_until_visible(driver=self.driver, css_selector='[label="Download (.csv)"]')
                    self.driver.find_element_by_css_selector('[label="Download (.csv)"]').click()
                except:
                    pass

            LOGGER.info(f'Business Report file is being retrieved: {end_date}')
            sleep(5)
            # Get the downloaded file path
            file_download = self.get_file_download(directory_downloads=directory_downloads)
            # Update the reports by adding StartDate and EndDate columns
            self.update_csv_report(csv_path=file_download, start_date=date, end_date=date)
            # Save reports to the spreadsheet
            if not pd.isnull(work_sheet_name):
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name)
            else:
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name)
            # Check and save a local copy of the reports to Reports directory as a csv
            if str(save_local_copy).lower() == 'yes':
                self.save_reports_locally(file_download=file_download)
                self.remove_file(file_download=file_download)
            # Remove the file from the Downloads directory
            elif str(save_local_copy).lower() == 'no':
                self.remove_file(file_download=file_download)

    # Retrieves advertising reports from a seller central account
    def get_advertising_reports(self, account, dates):
        email = str(account["Email"]).strip()
        # password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        ad_campaign_type = str(account["AdCampaignType"]).strip()
        ad_report_type = str(account["AdReportType"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving advertising reports for: {client}: {ad_campaign_type}: {ad_report_type}')
        # The main URL for advertising reports
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/AdvertisingReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        ads_report_url = 'https://advertising.amazon.com/sspa/tresah?ref_=AAC_gnav_reports'
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        LOGGER.info(f'Navigating to advertising reports')
        self.driver.get(url=ads_report_url)
        sleep(1)
        # Switches account
        account_id = None
        if not self.account_switched or account_id is not None:
            account_id = self.switch_ad_account(account_name=account_name, location=location)
        if account_id:
            ads_url_filter = ads_report_url + f'&entityId=ENTITY{account_id}'
            self.driver.get(ads_url_filter)
            sleep(1)
        # for date in :
        # Get first date and last date from the date_range
        start_date, end_date = dates[0], dates[-1]
        LOGGER.info(f'Retrieving advertising reports for: {start_date} - {end_date}')
        LOGGER.info(f'Creating advertising reports')
        # Creates report
        try:
            LOGGER.info(f'Waiting for creating report')
            self.wait_until_visible(driver=self.driver, xpath='//*[@id="sspa-reports:manage-reports-page:subscriptions_table:create-report-button"]')
            self.driver.find_element_by_xpath('//*[@id="sspa-reports:manage-reports-page:subscriptions_table:create-report-button"]').click()
        except:
            self.wait_until_visible(driver=self.driver, element_id='sspa-reports:manage-reports-page:subscriptions_table:create-report-button', duration=5)
            self.driver.find_element_by_id('sspa-reports:manage-reports-page:subscriptions_table:create-report-button').click()
        # Waits for radio selectors
        try:
            LOGGER.info(f'Waiting for campaigns')
            self.wait_until_visible(driver=self.driver, tag_name='fieldset', duration=5)
        except:
            pass
        sleep(1)
        # Selects campaign type
        LOGGER.info(f'Selecting campaign type: {ad_campaign_type}')
        for campaign_type in self.driver.find_element_by_tag_name('fieldset').find_elements_by_tag_name('label'):
            # campaign_type_name = str(campaign_type.text).replace('New Feature', '').replace('Beta feature', '').replace('New', '').replace('Feature', '').replace('Report', '').replace('Beta', '').replace('feature', '').strip().lower()
            campaign_type_name = str(campaign_type.find_element_by_tag_name('span').text).strip().lower()
            if str(ad_campaign_type).strip().lower() == campaign_type_name:
                campaign_type.click()
                LOGGER.info(f'Campaign type has been selected: {ad_campaign_type}')
                sleep(1)
                break
        # Selects report type
        LOGGER.info(f'Selecting report type: {ad_report_type}')
        self.driver.find_element_by_tag_name('tbody').find_elements_by_tag_name('td')[1].find_element_by_tag_name('button').click()
        sleep(1)
        self.wait_until_visible(driver=self.driver, element_id='portal')
        # for report_type in self.driver.find_element_by_id('portal').find_elements_by_tag_name('button'):
        for report_type in self.driver.find_element_by_id('portal').find_elements_by_tag_name('button'):
            report_type_name = str(report_type.text).replace('New Report', '').replace('Beta feature', '').replace('New', '').replace('Report', '').replace('Beta', '').replace('feature', '').strip().lower()
            # print(str(ad_report_type).strip().lower() == report_type_trimmed, str(ad_report_type).strip().lower(), report_type_trimmed)
            # report_type_name = str(report_type.find_element_by_tag_name('span').text).strip().lower()
            if str(ad_report_type).strip().lower() == report_type_name:
                LOGGER.info(f'Report type has been selected: {ad_report_type}')
                report_type.click()
                sleep(1)
                break
        # Selects time unit as daily
        LOGGER.info(f'Waiting for time unit')
        self.wait_until_visible(driver=self.driver, element_id='undefined-day')
        LOGGER.info(f'Selecting time unit')
        self.driver.find_elements_by_id('undefined-day')[-1].click()
        LOGGER.info(f'Time unit has been selected as: Daily')
        # Select date range in the calender
        is_date_selected = self.select_ad_date_range(start_date=start_date, end_date=end_date)
        if is_date_selected:
            LOGGER.info(f'Running report')
            self.wait_until_visible(driver=self.driver, element_id='J_Button_NORMAL_ENABLED', duration=5)
            # Clicks on Run report
            self.driver.find_element_by_id('J_Button_NORMAL_ENABLED').click()
            sleep(1)
            # Clicks on Run report on next page
            self.wait_until_visible(driver=self.driver, element_id='J_Button_NORMAL_ENABLED', duration=5)
            self.driver.find_element_by_id('J_Button_NORMAL_ENABLED').click()
            while True:
                self.driver.refresh()
                sleep(5)
                try:
                    # Click download button
                    self.wait_until_visible(driver=self.driver, element_id='J_Button_NORMAL_ENABLED', duration=3)
                    self.driver.find_element_by_id('sspa-reports:report-settings-page:-download-button').click()
                    break
                except:
                    pass
            sleep(1)
            # Waits and click on download button
            LOGGER.info(f'Retrieving advertising report')
            while True:
                self.driver.refresh()
                sleep(5)
                try:
                    # Click download button
                    self.wait_until_visible(driver=self.driver, element_id='J_Button_NORMAL_ENABLED', duration=3)
                    self.driver.find_element_by_id('sspa-reports:report-settings-page:-download-button').click()
                    break
                except:
                    try:
                        self.wait_until_visible(driver=self.driver, css_selector='.svg-inline--fa.fa-download.fa-w-16.fa-xs', duration=5)
                        self.driver.find_element_by_css_selector('.svg-inline--fa.fa-download.fa-w-16.fa-xs').click()
                        break
                    except:
                        pass
            LOGGER.info(f'Advertising Report file is being retrieved')
            # Wait for file to be downloaded
            sleep(5)
            # Get the downloaded file path
            file_download = self.get_file_download(directory_downloads=directory_downloads)
            # Save reports to the spreadsheet
            if not pd.isnull(work_sheet_name):
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name)
            else:
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name)
            # Check and save a local copy of the reports to Reports directory as a csv
            if str(save_local_copy).lower() == 'yes':
                self.save_reports_locally(file_download=file_download)
                self.remove_file(file_download=file_download)
            # Remove the file from the Downloads directory
            elif str(save_local_copy).lower() == 'no':
                self.remove_file(file_download=file_download)
        else:
            LOGGER.info(f'Moving to next date')

    # Retrieves fulfillment reports from a seller central account
    def get_fulfillment_reports(self, account):
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving fulfillment reports for: {client}')
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/FulfillmentReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        fulfillment_url = 'https://sellercentral.amazon.com/reportcentral/FlatFileAllOrdersReport/1'
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        LOGGER.info(f'Navigating to fulfillment reports')
        sleep(3)
        self.driver.get(url=fulfillment_url)
        sleep(3)
        # Login to Seller Central
        self.login_seller_central(email=email, password=password)
        sleep(3)
        # Switches account
        if not self.account_switched:
            self.switch_biz_account(account_name=account_name, location=location)
        # Select report range as daily
        try:
            LOGGER.info(f'Waiting for fulfillment reports')
            self.wait_until_visible(driver=self.driver, class_name='button', duration=10)
        except:
            pass
        LOGGER.info(f'Retrieving fulfillment reports')
        self.driver.find_element_by_tag_name('html').send_keys(Keys.END)
        for row in self.driver.find_elements_by_tag_name('kat-table-row')[1:]:
            date_range = row.find_elements_by_tag_name('kat-table-cell')[1].text
            date_range = date_range.replace(',', '').split('-')
            start_date = date_range[0]
            end_date = date_range[1]
            row.find_element_by_tag_name('kat-button').click()
            LOGGER.info(f'fulfillment Report file is being retrieved')
            sleep(5)
            # Get the downloaded file path
            file_download = self.get_file_download(directory_downloads=directory_downloads)
            # Update the reports by adding StartDate and EndDate columns
            self.update_csv_report(csv_path=file_download, start_date=start_date, end_date=end_date)
            # Save reports to the spreadsheet
            if not pd.isnull(work_sheet_name):
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name)
            else:
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name)
            # Check and save a local copy of the reports to Reports directory as a csv
            if str(save_local_copy).lower() == 'yes':
                self.save_reports_locally(file_download=file_download)
                self.remove_file(file_download=file_download)
            # Remove the file from the Downloads directory
            elif str(save_local_copy).lower() == 'no':
                self.remove_file(file_download=file_download)

    # Selects date range from amazon date-picker calendar for advertising reports
    def select_ad_date_range(self, start_date, end_date):
        start_day_digit = start_date[3:5]
        end_day_digit = end_date[3:5]
        start_month_digit = start_date[:2]
        end_month_digit = end_date[:2]
        start_year_digit = start_date[-4:]
        start_year_int = int(start_year_digit)
        end_year_digit = end_date[-4:]
        months_converter = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 'June': '06',
                            'July': '07', 'August': '08', 'September ': '09', 'October': '10', 'November': '11',
                            'December': '12',
                            '01': 'January', '02': 'February', '03': 'March', '04': 'April', '05': 'May', '06': 'June',
                            '07': 'July', '08': 'August', '09': 'September', '1': 'October', '11': 'November',
                            '12': 'December'}
        start_month_name = months_converter[start_month_digit]
        end_month_name = months_converter[end_month_digit]
        # Removes leading zero from day_digit
        if start_day_digit.startswith('0'):
            start_day_digit = start_day_digit[-1]
        if end_day_digit.startswith('0'):
            end_day_digit = end_day_digit[-1]
        LOGGER.info(f'Date range to be selected: {start_date} - {end_date}')
        LOGGER.info(f'Selecting report period')
        # Selects period date
        try:
            self.driver.find_element_by_tag_name('tbody').find_elements_by_tag_name('td')[-1].find_element_by_tag_name('button').click()
            LOGGER.info(f'Waiting for calendar')
            # self.wait_until_visible(driver=self.driver, class_name='.CalendarMonth_caption.CalendarMonth_caption_1')
            self.wait_until_visible(driver=self.driver, element_id='portal')
        except:
            pass
        button_back = self.driver.find_element_by_css_selector('[aria-label="Move backward to switch to the previous month."]')
        button_next = self.driver.find_element_by_css_selector('[aria-label="Move forward to switch to the next month."]')
        # Selects month for start_date
        LOGGER.info(f'Selecting starting month and year: {start_month_name}, {start_year_digit}')
        while True:
            sleep(1)
            month_year_text_left = self.driver.find_elements_by_css_selector('.CalendarMonth_caption.CalendarMonth_caption_1')[1].text
            month_year_text_right = self.driver.find_elements_by_css_selector('.CalendarMonth_caption.CalendarMonth_caption_1')[2].text
            if start_month_name in month_year_text_left and start_year_digit in month_year_text_left:
                LOGGER.info(f'Start Month and year found in left table')
                # Decides which table the month was found, left == 1 or right == 2
                day_table_index = 1
                break
            elif start_month_name in month_year_text_right and start_year_digit in month_year_text_right:
                LOGGER.info(f'Start Month and year found in right table')
                day_table_index = 2
                break
            else:
                # Move backward to previous month if calender year or month is greater than the start_year or start_month
                LOGGER.info(f'Looking for month and year')
                # Converts Calender month and year to integers
                calender_month_int = int(months_converter[str(month_year_text_left[:-4]).strip()])
                calender_year_int = int(month_year_text_left[-4:])
                # print(calender_month_int, int(start_month_digit), calender_month_int > int(start_month_digit))
                # print(calender_year_int, start_year_int, calender_year_int > start_year_int)
                if calender_year_int > start_year_int or calender_month_int > int(start_month_digit):
                    LOGGER.info(f'Moving to the previous month')
                    button_back.click()
                # Move forward to next month if calender year or month is smaller than the start_year or start_month
                elif calender_year_int < start_year_int or calender_month_int < int(start_month_digit):
                    LOGGER.info(f'Moving forward to the next month')
                    button_next.click()
        # Selects day for start_date
        LOGGER.info(f'Selecting start day: {start_day_digit}')
        for day in self.driver.find_elements_by_css_selector('.CalendarMonth_table.CalendarMonth_table_1')[day_table_index].find_elements_by_tag_name('td'):
            if start_day_digit in day.text:
                # Checks and return with False, if the day is disabled
                # print(str(day.get_attribute('aria-label')))
                if 'Not available' in str(day.get_attribute('aria-label')):
                    LOGGER.info(f'Date not available')
                    LOGGER.info(f'Please select a valid date range')
                    return False
                # Selects the start day and return True
                else:
                    day.click()
                    LOGGER.info(f'Start day has been selected')
                    # Save the date button
                    sleep(1)
                    break
                    # LOGGER.info(f'Saving Date')
                    # self.wait_until_visible(driver=self.driver, css_selector='.sc-AxiKw.eZwBPK')
                    # self.driver.find_element_by_css_selector('.sc-AxiKw.eZwBPK').click()
                    # LOGGER.info(f'Date has been saved')
                    # return True
        # Selects month for end_date
        LOGGER.info(f'Selecting ending month and year: {end_month_name}, {end_year_digit}')
        while True:
            sleep(1)
            month_year_text_left = self.driver.find_elements_by_css_selector('.CalendarMonth_caption.CalendarMonth_caption_1')[1].text
            month_year_text_right = self.driver.find_elements_by_css_selector('.CalendarMonth_caption.CalendarMonth_caption_1')[2].text
            if end_month_name in month_year_text_left and end_year_digit in month_year_text_left:
                LOGGER.info(f'End Month and year found in left table')
                # Decides which table the month was found, left == 1 or right == 2
                day_table_index = 1
                break
            elif end_month_name in month_year_text_right and end_year_digit in month_year_text_right:
                LOGGER.info(f'End Month and year found in right table')
                day_table_index = 2
                break
            else:
                # Move forward to next month
                LOGGER.info(f'Moving to the previous month')
                button_next.click()
        # Selects day for end_date
        LOGGER.info(f'Selecting end day: {end_day_digit}')
        for day in self.driver.find_elements_by_css_selector('.CalendarMonth_table.CalendarMonth_table_1')[day_table_index].find_elements_by_tag_name('td'):
            # print(end_day_text, day.text, end_day_text in day.text)
            if end_day_digit in day.text:
                # Checks and return with False, if the day is disabled
                if 'Not available' in str(day.get_attribute('aria-label')):
                    LOGGER.info(f'Date not available')
                    LOGGER.info(f'Please select a valid date range')
                    return False
                # Selects the start day and return True
                else:
                    day.click()
                    LOGGER.info(f'End day has been selected')
                    # Save the date button
                    sleep(1)
                    LOGGER.info(f'Saving date range')
                    self.driver.find_element_by_id('portal').find_elements_by_tag_name('button')[-1].click()
                    LOGGER.info(f'Date range has been saved')
                    return True

    # Selects date range from amazon date-picker calendar for vendor reports
    def select_vendor_date_range(self, date):
        day_text = date[3:5]
        month = date[:2]
        year = date[-4:]
        months = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 'June': '06',
                  'July': '07', 'August': '08', 'September ': '09', 'October': '10', 'November': '11', 'December': '12'}
        for k, v in months.items():
            if month in v:
                month = k
        LOGGER.info(f'Selecting date: {date}')
        self.driver.find_element_by_xpath('//*[@id="dashboard-filter-periodPicker"]/div/div/div[1]/input').click()
        sleep(1)
        # Selects month for start_date
        while True:
            date_text = self.driver.find_elements_by_class_name('react-datepicker__current-month')[1].text
            if month in date_text and year in date_text:
                break
            else:
                try:
                    self.driver.find_element_by_css_selector('.react-datepicker__navigation.react-datepicker__navigation--previous').click()
                except:
                    LOGGER.info(f'Date out of range')
                    LOGGER.info(f'Please select a valid date range')

        # Selects day for start_date
        for day in self.driver.find_elements_by_class_name('react-datepicker__month')[1].find_elements_by_tag_name('div'):
            if day_text in day.text:
                day.click()
                break
        sleep(1)
        date_input_to = self.driver.find_element_by_xpath('//*[@id="dashboard-filter-periodPicker"]/div/div/div[3]/input')
        date_input_to.click()
        sleep(1)
        # Selects month for end_date
        while True:
            date_text = self.driver.find_elements_by_class_name('react-datepicker__current-month')[1].text
            if month in date_text and year in date_text:
                break
            else:
                try:
                    self.driver.find_element_by_css_selector('.react-datepicker__navigation.react-datepicker__navigation--previous').click()
                except:
                    LOGGER.info(f'Date out of range')
                    LOGGER.info(f'Please select a valid date range')
        # Select day for end_date
        for day in self.driver.find_elements_by_class_name('react-datepicker__month')[1].find_elements_by_tag_name('div'):
            if day_text in day.text:
                day.click()
                break
        sleep(1)
        LOGGER.info(f'Date has been selected')

    # Retrieves vendor reports from a vendor central account
    def get_vendor_reports(self, account, dates):
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving vendor reports for: {client}')
        # The main URL for vendor reports
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/VendorReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        vendor_report_url = 'https://vendorcentral.amazon.com/analytics/dashboard/salesDiagnostic'
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        # # Switches account
        # if not self.account_switched:
        #     self.switch_account(account=account)
        LOGGER.info(f'Navigating to vendor reports')
        sleep(1)
        self.driver.get(url=vendor_report_url)
        sleep(3)
        try:
            LOGGER.info(f'Waiting for vendor central]')
            self.wait_until_visible(driver=self.driver, element_id='logoCell', duration=5)
            LOGGER.info(f'Vendor central login successful]')
            self.logged_in = True
            self.logged_in_email = email
        except:
            # Try if vendor central needs sign-in
            try:
                LOGGER.info(f'Signing-in to vendor central')
                self.wait_until_visible(driver=self.driver, name='password', duration=5)
                sleep(3)
                self.driver.find_element_by_name('password').send_keys(password)
                sleep(1)
                self.driver.find_element_by_name('rememberMe').click()
                sleep(3)
                self.driver.find_element_by_id('signInSubmit').click()
                LOGGER.info(f'Waiting for vendor central]')
                self.wait_until_visible(driver=self.driver, element_id='logoCell', duration=5)
                LOGGER.info(f'Vendor central login successful]')
                self.logged_in = True
                self.logged_in_email = email
            except:
                try:
                    LOGGER.info(f'Vendor Central not yet logged]')
                    self.captcha_login(email=email, password=password)
                    self.wait_until_visible(driver=self.driver, element_id='sc-logo-top', duration=5)
                    LOGGER.info(f'Vendor Central login successful]')
                    self.logged_in = True
                    self.logged_in_email = email
                except:
                    pass
        # Select report range as daily
        try:
            LOGGER.info(f'Setting report range as daily')
            self.wait_until_visible(driver=self.driver, element_id='dashboard-filter-reportingRange', duration=10)
            self.driver.find_element_by_id('dashboard-filter-reportingRange').click()
            self.driver.find_element_by_class_name('awsui-button-dropdown-items').find_element_by_tag_name('li').click()
        except:
            pass
        try:
            LOGGER.info(f'Waiting for date picker')
            self.wait_until_visible(driver=self.driver, element_id='date-picker-container', duration=10)
            self.driver.find_element_by_id('date-picker-container').click()
        except:
            pass
        for date in dates:
            LOGGER.info(f'Setting date filter as: {date}')
            # Select date range
            self.select_vendor_date_range(date)
            try:
                # Click apply button
                self.driver.find_element_by_xpath('//*[@id="dashboard-filter-applyCancel"]/div/awsui-button[2]/button/span').click()
                sleep(3)
            except:
                pass
            self.driver.find_element_by_id('downloadButton').find_element_by_tag_name('button').click()
            sleep(1)
            LOGGER.info(f'Retrieving vendor report')
            self.wait_until_visible(driver=self.driver, xpath='//*[@id="downloadButton"]/awsui-button-dropdown/div/div/ul/li[3]/ul/li[2]')
            self.driver.find_element_by_xpath('//*[@id="downloadButton"]/awsui-button-dropdown/div/div/ul/li[3]/ul/li[2]').click()
            LOGGER.info(f'Advertising Report file is being retrieved')
            LOGGER.info(f'Download in progress ...')
            sleep(3)
            try:
                LOGGER.info(f'Waiting for alert')
                WebDriverWait(self.driver, 300).until(EC.alert_is_present(), 'Changes you made may not be saved.')
                alert = self.driver.switch_to.alert
                LOGGER.info(f'Alert: {alert.text}')
                alert.accept()
                LOGGER.info('Alert accepted')
            except:
                LOGGER.info('No alert')
            sleep(5)
            # Get the downloaded file path
            file_download = self.get_file_download(directory_downloads=directory_downloads)
            # Save reports to the spreadsheet
            if not pd.isnull(work_sheet_name):
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name)
            else:
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name)
            # Check and save a local copy of the reports to Reports directory as a csv
            if str(save_local_copy).lower() == 'yes':
                self.save_reports_locally(file_download=file_download)
                self.remove_file(file_download=file_download)
            # Remove the file from the Downloads directory
            elif str(save_local_copy).lower() == 'no':
                self.remove_file(file_download=file_download)

    # Retrieves promotional reports from a vendor central account
    def get_vendor_promo_reports(self, account, dates):
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving vendor promotional reports for: {client}')
        # The main URL for advertising reports
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/VendorPromotionalReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        promo_report_url = 'https://vendorcentral.amazon.com/hz/vendor/members/promotions/list/home?ref_=vc_xx_subNav'
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        # # Switches account
        # if not self.account_switched:
        #     self.switch_account(account=account)
        LOGGER.info(f'Navigating to vendor promotional reports')
        sleep(1)
        self.driver.get(url=promo_report_url)
        sleep(3)
        try:
            LOGGER.info(f'Waiting for vendor central]')
            self.wait_until_visible(driver=self.driver, element_id='logoCell', duration=5)
            LOGGER.info(f'Vendor central login successful]')
            self.logged_in = True
            self.logged_in_email = email
        except:
            # Try if vendor central needs sign-in
            try:
                LOGGER.info(f'Signing-in to vendor central')
                self.wait_until_visible(driver=self.driver, name='password', duration=5)
                sleep(3)
                self.driver.find_element_by_name('password').send_keys(password)
                sleep(1)
                self.driver.find_element_by_name('rememberMe').click()
                sleep(3)
                self.driver.find_element_by_id('signInSubmit').click()
                LOGGER.info(f'Waiting for vendor central]')
                self.wait_until_visible(driver=self.driver, element_id='logoCell', duration=5)
                LOGGER.info(f'Vendor central login successful]')
                self.logged_in = True
                self.logged_in_email = email
            except:
                try:
                    LOGGER.info(f'Vendor Central not yet logged]')
                    self.captcha_login(email=email, password=password)
                    self.wait_until_visible(driver=self.driver, element_id='sc-logo-top', duration=5)
                    LOGGER.info(f'Vendor Central login successful]')
                    self.logged_in = True
                    self.logged_in_email = email
                except:
                    pass
        sleep(3)
        try:
            LOGGER.info(f'Waiting for promotional reports list')
            self.wait_until_visible(driver=self.driver, element_id='promotion-list-column', duration=5)
        except:
            try:
                self.captcha_login(email=email, password=password)
            except:
                try:
                    LOGGER.info(f'Waiting for promotional reports list')
                    self.wait_until_visible(driver=self.driver, element_id='promotion-list-column', duration=5)
                except:
                    pass
        # Get all report URLs from all the pages
        report_links = []
        self.driver.find_element_by_tag_name('html').send_keys(Keys.END)
        sleep(0.5)
        # et the default pages to 1
        pages = 1
        try:
            LOGGER.info(f'Waiting for promotional reports list')
            self.wait_until_visible(driver=self.driver, element_id='promotion-list-pagination', duration=5)
            pages = int(self.driver.find_element_by_id('promotion-list-pagination').find_elements_by_tag_name('li')[-2].text)
        except:
            pass
        for i in range(pages):
            self.driver.find_element_by_tag_name('html').send_keys(Keys.END)
            self.wait_until_visible(driver=self.driver, element_id='promotion-list', duration=5)
            # Append all the links to the list
            [report_links.append(link.get_attribute('href')) for link in self.driver.find_element_by_id('promotion-list').find_elements_by_css_selector('.a-size-base.a-link-normal')]
            LOGGER.info(f'Reports found so far: {len(report_links)}')
            if pages > 1:
                try:
                    LOGGER.info(f'Navigating to next page')
                    # Wait until overlay is loading
                    WebDriverWait(self.driver, 30, 0.01).until(EC.invisibility_of_element_located((By.CLASS_NAME, 'mt-loading-overlay')))
                    sleep(0.5)
                    self.driver.find_element_by_tag_name('html').send_keys(Keys.END)
                    self.driver.find_element_by_id('promotion-list-pagination').find_elements_by_tag_name('li')[-1].click()
                except:
                    self.driver.find_element_by_id('promotion-list-pagination').find_elements_by_tag_name('a')[-1].click()
        for link in report_links:
            LOGGER.info(f'Requesting promotional report: {link}')
            self.driver.get(link)
            try:
                LOGGER.info(f'Waiting for promotional report')
                self.wait_until_visible(driver=self.driver, element_id='promotion-report-dowoload-button', duration=60)
                sleep(0.5)
                LOGGER.info(f'Retrieving vendor promotional report')
                self.driver.find_element_by_id('promotion-report-dowoload-button').click()
                sleep(1)
            except:
                self.driver.find_element_by_id('promotion-report-download-button').click()
                sleep(1)
            LOGGER.info(f'Vendor promotional Report file is being retrieved ...')
            sleep(5)
            # Get the downloaded file path
            file_download = self.get_file_download(directory_downloads=directory_downloads)
            # column_names = ["PAWS ID", "Promotion Name", "Promotion Type", "Promotion Start Date", "Promotion End Date", "Vendor Code", "Product Group", "ASIN", "Product Name", "Net Unit Demand*", "Net Sales (PCOGS)**", "Aggregate"]
            # Promotional reports has an extra column value "Aggregate" which has to be read as sep=\t
            # Save reports to the spreadsheet
            if not pd.isnull(work_sheet_name):
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name)
            else:
                self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name)
            # Check and save a local copy of the reports to the Reports directory as a csv
            if str(save_local_copy).lower() == 'yes':
                self.save_reports_locally(file_download=file_download)
                self.remove_file(file_download=file_download)
            # Remove the file from the Downloads directory
            elif str(save_local_copy).lower() == 'no':
                self.remove_file(file_download=file_download)

    @staticmethod
    def get_sales_dashboard_df(file_download):
        # Create DataFrame from Sales Dashboard
        with open(file_download) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            content_list = [row for row in csv_reader]
            for i in content_list:
                if "Compare sales - Table view" in i:
                    skip_row = content_list.index(i) + 1
                    break
            df = pd.read_csv(file_download, index_col=None, skiprows=skip_row)
            return df

    # Retrieves promotional reports from a vendor central account
    def get_sales_dashboard_reports(self, account, dates):
        account_name = str(account["Brand"]).strip()
        location = str(account["Location"]).strip()
        email = str(account["Email"]).strip()
        password = str(account["Password"]).strip()
        client = str(account["Client"]).strip()
        spread_sheet_url = str(account["SpreadSheetURL"]).strip()
        spread_sheet_name = str(account["SpreadSheetName"]).strip()
        work_sheet_name = str(account["WorkSheetName"]).strip()
        save_local_copy = str(account["SaveLocalCopy"]).strip()
        LOGGER.info(f'Retrieving sales dashboard reports for: {client}')
        # The main URL for advertising reports
        self.file_local_reports = str(self.PROJECT_ROOT / f'AMZRes/Reports/SalesDashboardReports/Reports_{account_name}_{location}.csv')
        directory_downloads = self.directory_downloads
        dashboard_report_url = 'https://sellercentral.amazon.com/gp/site-metrics/report.html#&reportID=eD0RCS'
        # Check and get browser
        if self.driver is None:
            self.driver = self.get_driver()
        # Check and login to the website if not logged_in or email hs changed
        if not self.logged_in or self.logged_in_email != email:
            self.login_amazon(account=account)
        LOGGER.info(f'Navigating to sales dashboard reports')
        self.driver.get(url=dashboard_report_url)
        sleep(3)
        # Switches account
        if not self.account_switched:
            self.switch_biz_account(account_name=account_name, location=location)
        try:
            LOGGER.info(f'Waiting for Seller Central')
            self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
            LOGGER.info(f'Seller Central login successful')
            self.logged_in = True
            self.logged_in_email = email
        except:
            # Try if seller central needs sign-in
            try:
                LOGGER.info(f'Signing-in to Seller Central')
                self.wait_until_visible(driver=self.driver, name='password', duration=5)
                sleep(3)
                self.driver.find_element_by_name('password').send_keys(password)
                sleep(1)
                self.driver.find_element_by_name('rememberMe').click()
                sleep(3)
                self.driver.find_element_by_id('signInSubmit').click()
                LOGGER.info(f'Waiting for Seller Central')
                self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
                LOGGER.info(f'Seller Central login successful')
                self.logged_in = True
                self.logged_in_email = email
            except:
                try:
                    LOGGER.info(f'Seller Central not yet logged')
                    self.captcha_login(email=email, password=password)
                    self.wait_until_visible(driver=self.driver, element_id='sc-logo-asset', duration=5)
                    LOGGER.info(f'Seller Central login successful')
                    self.logged_in = True
                    self.logged_in_email = email
                except:
                    pass
        try:
            LOGGER.info(f'Waiting for sales dashboard reports')
            self.wait_until_visible(driver=self.driver, xpath='//*[@id="timeframePseudoCell"]/span/button', duration=5)
            LOGGER.info(f'Selecting sales dashboard dropdown')
            self.driver.find_element_by_xpath('//*[@id="timeframePseudoCell"]/span/button').click()
            self.wait_until_visible(driver=self.driver, class_name='ui-menu-item', duration=5)
            LOGGER.info(f'Dropdown selected')
        except:
            pass
        # Get reports for all dates in the Date dropdown
        for i in range(len(self.driver.find_elements_by_class_name('ui-menu-item'))):
            LOGGER.info(f'Selecting sales dashboard dropdown')
            self.wait_until_visible(driver=self.driver, class_name='ui-menu-item', duration=5)
            item = self.driver.find_elements_by_class_name('ui-menu-item')[i]
            LOGGER.info(f'Selecting Date option as: {item.text}')
            if str(item.text).strip() == "Custom":
                item.click()
                for date in dates:
                    LOGGER.info(f'Retrieving sales dashboard reports for: {date}')
                    # Selects FromDate
                    sleep(3)
                    from_date = self.driver.find_element_by_id('ssrFromDate')
                    from_date.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE, date)
                    from_date.send_keys(date)
                    sleep(1)
                    # select = Select(self.driver.find_element_by_class_name('ui-datepicker-month'))
                    # select.select_by_visible_text('Banana')
                    # Selects ToDate
                    to_date = self.driver.find_element_by_id('ssrToDate')
                    to_date.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE, date)
                    to_date.send_keys(date)
                    sleep(1)
                    # Clicks Apply button
                    self.wait_until_visible(driver=self.driver, element_id='applyButton', duration=20)
                    self.driver.find_element_by_id('applyButton').click()
                    sleep(3)
                    # Clicks Download button
                    self.wait_until_visible(driver=self.driver, element_id='download', duration=20)
                    self.driver.find_element_by_id('download').click()
                    LOGGER.info(f'Sales dashboard reports file is being retrieved ...')
                    sleep(5)
                    # Get the downloaded file path
                    file_download = self.get_file_download(directory_downloads=directory_downloads)
                    # Save reports to the spreadsheet
                    if not pd.isnull(work_sheet_name):
                        self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name, sales_dashboard=True)
                    else:
                        self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, sales_dashboard=True)
                    # Check and save a local copy of the reports to the Reports directory as a csv
                    if str(save_local_copy).lower() == 'yes':
                        self.save_reports_locally(file_download=file_download, sales_dashboard=True)
                        self.remove_file(file_download=file_download)
                    # Remove the file from the Downloads directory
                    elif str(save_local_copy).lower() == 'no':
                        self.remove_file(file_download=file_download)
            else:
                item.click()
                # Clicks Apply button
                self.wait_until_visible(driver=self.driver, element_id='applyButton', duration=20)
                self.driver.find_element_by_id('applyButton').click()
                sleep(3)
                # Clicks Download button
                self.wait_until_visible(driver=self.driver, element_id='download', duration=20)
                self.driver.find_element_by_id('download').click()
                LOGGER.info(f'Sales dashboard reports file is being retrieved ...')
                sleep(5)
                # Get the downloaded file path
                file_download = self.get_file_download(directory_downloads=directory_downloads)
                # Save reports to the spreadsheet
                if not pd.isnull(work_sheet_name):
                    self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, work_sheet_name=work_sheet_name, sales_dashboard=True)
                else:
                    self.csv_to_spreadsheet(csv_path=file_download, spread_sheet_url=spread_sheet_url, spread_sheet_name=spread_sheet_name, sales_dashboard=True)
                # Check and save a local copy of the reports to the Reports directory as a csv
                if str(save_local_copy).lower() == 'yes':
                    self.save_reports_locally(file_download=file_download, sales_dashboard=True)
                    self.remove_file(file_download=file_download)
                # Remove the file from the Downloads directory
                elif str(save_local_copy).lower() == 'no':
                    self.remove_file(file_download=file_download)
            # Clicks dropdown
            self.wait_until_visible(driver=self.driver, xpath='//*[@id="timeframePseudoCell"]/span/button', duration=5)
            self.driver.find_element_by_xpath('//*[@id="timeframePseudoCell"]/span/button').click()

    def main(self):
        freeze_support()
        self.enable_cmd_colors()
        trial_date = datetime.strptime('2021-08-05 23:59:59', '%Y-%m-%d %H:%M:%S')
        # Print ASCII Art
        print('************************************************************************\n')
        pyfiglet.print_figlet('____________                   AMZBot ____________\n', colors='RED')
        print('Author: Ali Toori, Bot Developer\n'
              'Website: https://boteaz.com/\nYouTube: https://youtube.com/@AliToori\n************************************************************************')
        # Trial version logic
        if self.trial(trial_date):
            LOGGER.info(f'AMZBot launched')
            # LOGGER.warning("[Your trial will end on: ",
            #       str(trial_date) + ". To get full version, please contact fiverr.com/botflocks !')
            if os.path.isfile(self.file_account):
                # account_df = pd.read_csv(self.file_account, index_col=None)
                account_df = pd.read_excel(self.file_account, index_col=None)
                while True:
                    # Get accounts from Accounts.csv
                    for account in account_df.iloc:
                        # self.first_time = True
                        self.account_switched = False
                        report_type = str(account["ReportType"]).strip()
                        # Clear the Downloads Directory
                        self.clear_downloads_directory(directory_downloads=self.directory_downloads)
                        dates = self.get_date_range(account=account)
                        if report_type == 'Business Reports':
                            self.get_business_reports(account=account, dates=dates)
                            LOGGER.info(f'Successfully retrieved business reports')
                        elif report_type == 'Advertising Reports':
                            self.get_advertising_reports(account=account, dates=dates)
                            LOGGER.info(f'Successfully retrieved advertising reports')
                        elif report_type == 'Fulfillment Reports':
                            self.get_fulfillment_reports(account=account)
                            LOGGER.info(f'Successfully retrieved fulfillment reports')
                        elif report_type == 'Vendor Reports':
                            self.get_vendor_reports(account=account, dates=dates)
                            LOGGER.info(f'Successfully retrieved vendor reports')
                        elif report_type == 'Vendor Promotional Reports':
                            self.get_vendor_promo_reports(account=account, dates=dates)
                            LOGGER.info(f'Successfully retrieved vendor promotional reports')
                        elif report_type == 'Sales Dashboard Reports':
                            self.get_sales_dashboard_reports(account=account, dates=dates)
                            LOGGER.info(f'Successfully retrieved sales dashboard reports')
                    self.finish()
                    LOGGER.info(f'Successfully retrieved all the reports')
                    LOGGER.info(f'The bot will auto-restart after 24 hours')
                    sleep(86400)
        else:
            pass


if __name__ == '__main__':
    amz_bot = AMZBot()
    amz_bot.main()
