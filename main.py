import datetime

from prepare_data import clear_path
from search_engine import PickLogin, LogError, Login
from selenium.common.exceptions import NoSuchElementException

from config import *
from email_handler import EmailCreation
from sap_download import SapDownload


if __name__ == '__main__':
    chrome_path = robot_credentials["chrome_path"]
    driver_path = robot_credentials["driver_path"]
    site = robot_credentials["site"]
    login_path = robot_credentials["login_path"]
    destination = robot_credentials["ompt_output"]
    clear_path(destination)
    login = Login(chrome_path, driver_path, site, destination)
    logs = PickLogin(login_path).download_logins()
    for x in logs:
        log = x[0]
        pas = x[1]

        '''Validating logins to receive unified file names'''

        if log == "PROCESS CONTROL":
            new_name = "BE_PDC.csv"
        elif log == "PROCESS CONTROL NL":
            new_name = "NL_PDC.csv"
        else:
            new_name = f"{log}.csv"
        try:
            if LogError(log, pas).check_json():
                login.login_to_site(log, pas, new_name)
        except NoSuchElementException:
            print(f"{log} or {pas} are incorrect, hitting the next one")
            LogError(log, pas).append_to_json()
            EmailCreation(log, pas, robot_credentials["email"]).send_email()
            login.close_chrome()
            continue

    now = datetime.datetime.now()
    date_from = '{:02d}.{:02d}.{:4d}'.format(now.day, now.month, now.year)
    date_to = '{:02d}.{:02d}.{:4d}'.format(now.day, now.month, now.year)
    if now.minute < 15:
        minutes = now.minute + 45
        hours = now.hour - 2
    else:
        minutes = now.minute - 15
        hours = now.hour - 1

    start_hour = '{:02d}:{:02d}:{:02d}'.format(hours, minutes, now.second)
    finish_hour = '{:02d}:{:02d}:{:02d}'.format(now.hour, now.minute, now.second)
    path = robot_credentials["sap_output"]
    sap_connection = SapDownload(date_from, date_to, start_hour, finish_hour, path)
    sap_connection.download_data()
    sap_connection.download_from_edid4()
    print('Finished')


