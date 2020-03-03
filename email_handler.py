
import win32com.client


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

