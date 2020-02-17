import win32com.client
import json


class EmailCreation:
    def __init__(self, log, pas, recipient):
        self.log = log
        self.pas = pas
        self.recipient = recipient

    def send_email(self):
        o = win32com.client.Dispatch("Outlook.Application")

        Msg = o.CreateItem(0)
        Msg.To = self.recipient

        Msg.Subject = "Credentials need to be checked"
        Msg.Body = f"Login {self.log} or password {self.pas} are incorrect. \n" \
                   f"Check and take action if necessary."
        Msg.Send()


if __name__ == '__main__':
    with open('ompt_partners.json') as error_log:
        dictdump = json.loads(error_log.read())

    for p in dictdump['logins']:
        print(p['login'], p['password'])
