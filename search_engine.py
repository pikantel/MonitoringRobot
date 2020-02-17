from selenium import webdriver
from json.decoder import JSONDecodeError
import time
import json
import requests
import os
import shutil
import datetime


# this object is setting the driver and connects to the browser, opening a site
class ConnectToChrome:

    def __init__(self, chrome_path, driver_path, site, destination):
        self.chrome_path = chrome_path
        self.driver_path = driver_path
        self.site = site
        self.destination = destination

    def connection(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": self.destination}) #by destination parameter you can set destination of your download
        options.binary_location = self.chrome_path
        driver = webdriver.Chrome(executable_path=self.driver_path, chrome_options=options)
        driver.get(self.site)
        driver.maximize_window()
        return driver


# this object is to login to omprompt with given logins and passwords
class Login:
    def __init__(self, chrome_path, driver_path, site, destination):
        self.chrome_path = chrome_path
        self.driver_path = driver_path
        self.site = site
        self.destination = destination

    def login_to_site(self, log, pas, new_name):
        connect = ConnectToChrome(self.chrome_path, self.driver_path, self.site, self.destination)
        browser = connect.connection()
        browser.find_element_by_id('RememberUser').click()
        browser.find_element_by_id('username').send_keys(log)
        browser.find_element_by_id('password').send_keys(pas)
        browser.find_element_by_xpath("//*[@type='submit']").click()
        DownloadData().download(browser, self.destination, new_name)

    def close_chrome(self):
        connect = ConnectToChrome(self.chrome_path, self.driver_path, self.site, self.destination)
        browser = connect.connection()
        browser.close()


# this object downloads the data from ompt site
class DownloadData:
    def __init__(self):
        pass

    def download(self, browser, file_path, new_name):
        now = datetime.datetime.now()
        browser.find_element_by_xpath('//*[@id="MenuMonitorAuditTrail"]').click()
        if now.minute < 15:
            minutes = now.minute + 45
            hours = now.hour - 2
        else:
            minutes = now.minute - 15
            hours = now.hour - 1
        browser.find_element_by_id('selectedFromHour').send_keys('{:02d}'.format(hours))
        browser.find_element_by_id('selectedFromMin').send_keys('{:02d}'.format(minutes))
        browser.find_element_by_id('selectedToHour').send_keys('{:02d}'.format(now.hour))
        browser.find_element_by_id('selectedToMin').send_keys('{:02d}'.format(now.minute))
        browser.find_element_by_id('auditStatus').send_keys('Completed')
        browser.find_element_by_id('searchButton').click()
        time.sleep(1)
        if "No Results" in browser.find_element_by_xpath('//*[@id="dataGrid"]').text:
            browser.close()
            pass
        else:
            browser.find_element_by_id('downloadButton').click()
            time.sleep(1)
            browser.close()
            filename = max([file_path + '\\' + f for f in os.listdir(file_path)], key=os.path.getctime)
            shutil.move(os.path.join(file_path, filename), os.path.join(file_path, new_name))


# this object is taking log and pass from json file to provide as an input to another object

class PickLogin:
    def __init__(self, login_path):
        self.login_path = login_path

    def download_logins(self):
        with open(self.login_path, 'r') as loaded_json:
            data = json.load(loaded_json)
            for f in data['logins']:
                yield ((f['login'], f['password']))


def check_if_json_exists():  # if not, then it's creating a new one
    with open('error_log.json') as error_log:
        try:
            json.loads(error_log.read())
        except JSONDecodeError:
            new_file = {'failure': []}
            with open('error_log.json', 'w') as error_log2:
                json.dump(new_file, error_log2)


class LogError:
    def __init__(self, log, pas):
        self.log = log
        self.pas = pas

    def append_to_json(self):
        with open('error_log.json') as error_log:
            error_log = json.loads(error_log.read())
            try:
                new_dict = {
                    'login': self.log,
                    'password': self.pas
                }
                error_log['failure'].append(new_dict)
                with open('error_log.json', 'w') as file:
                    json.dump(error_log, file)
            except self.log in error_log:
                pass

    def check_json(self):
        check_if_json_exists()
        with open('error_log.json') as error_log:
            error_log = json.loads(error_log.read())
        for item in error_log['failure']:
            if self.log in item.values():
                return False
        return True


