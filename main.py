# sortowanie importów https://www.python.org/dev/peps/pep-0008/#imports
from prepare_data import clear_path
from search_engine import PickLogin, LogError, Login
from selenium.common.exceptions import NoSuchElementException
from email_handler import EmailCreation
import datetime
from sap_download import SapDownload


# Oddzielałbym enterami ten kod tym samym tworząc sekcję.
# Taki kod dużo łatwiej się czyta.
if __name__ == '__main__':
    # Nie wiem do końca w jakim środowisku ten runner będzie działał produkcyjnie.
    # Natomiast to co mogę polecić to żeby przenieść te ścieżki do oddzielnego configa.
    # Komfort jest taki że nie ma porozrzucanych jakis stałych po skryptach i całą konfigurację mamy w jednym miejscu.
    # Do testów czy debugowania to jednak jak najbardziej jest ok :)
    chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    driver_path = r'C:\Windows\chromedriver.exe'
    site = 'https://portal1.live.omprompt.com/ompcrm/view/login.csp'
    login_path = r"C:\python\web_scrap\ompt_partners.json"
    destination = r"C:\Temp\output\OmptData"
    clear_path(destination)

    login = Login(chrome_path, driver_path, site, destination)
    logs = PickLogin(login_path).download_logins()
    for x in logs:
        log = x[0]
        pas = x[1]
        # Ty pewnie wiesz o co chodzi w tych dwóch pierwszych ifach, ale ja nie mam najmniejszego pojecia :D
        # Dodaj tu proszę komentarz co się dzieję bo porównywanie loginu do jakis kontrolek stringowych to dla mnie magia.
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
            # zamiast printowania używaj logowania - https://docs.python.org/3/library/logging.html
            # maly tutek: http://nathanielobrown.com/blog/posts/quick_and_dirty_python_logging_lesson.html
            print(f"{log} or {pas} are incorrect, hitting the next one")
            LogError(log, pas).append_to_json()
            EmailCreation(log, pas, "mikolaj.chrzan@unilever.com").send_email()
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
    path = 'C:\\Temp\\output\\SAP\\'
    sap_connection = SapDownload(date_from, date_to, start_hour, finish_hour, path)
    sap_connection.download_data()
    sap_connection.download_from_edid4()
    print('Finished')


