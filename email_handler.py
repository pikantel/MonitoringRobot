# sortowanie importów https://www.python.org/dev/peps/pep-0008/#imports
import win32com.client
import json


# zastanowiłbym się czy nie użyć tutaj dataclasses (i we wszystkich innyc obiektach też)
# pamiętaj jednak że są one dostępne dopiero od pythona 3.7 a nie wiem w jakim piszesz
# https://docs.python.org/3/library/dataclasses.html
class EmailCreation:
    # Brak anotacji typów wejsciowych: https://dev.to/dstarner/using-pythons-type-annotations-4cfe
    def __init__(self, log, pas, recipient):
        self.log = log
        self.pas = pas
        self.recipient = recipient

    # Brak anotacji typów wyjściowych: https://dev.to/dstarner/using-pythons-type-annotations-4cfe
    def send_email(self):
        o = win32com.client.Dispatch("Outlook.Application")

        Msg = o.CreateItem(0)
        Msg.To = self.recipient
        # Msg.CC = "patryk.rozenek2@unilever.com"

        Msg.Subject = "Credentials need to be checked"
        Msg.Body = f"Login {self.log} or password {self.pas} are incorrect. \n" \
                   f"Check and take action if necessary."
        Msg.Send()


# Tego typu runnerów produkcyjnie nie umieszcza się raczej w skrypcie klasy pythonowej,
# ale rozumiem ze potrzebne Ci to było do debugowania, tak tylko napominam :).
if __name__ == '__main__':
    with open('ompt_partners.json') as error_log:
        dictdump = json.loads(error_log.read())

    for p in dictdump['logins']:
        print(p['login'], p['password'])
