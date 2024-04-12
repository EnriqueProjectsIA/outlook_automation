import win32com.client
from win32com.client import DispatchWithEvents
from typing import Dict, List
import pythoncom
import datetime as dt
from datetime import datetime, timedelta
import pytz
from pathlib import Path

class InboxEvents:
    def __init__(self, attachments_path:str|None = None, no_attachments_path:str|None = None):
        self.attachments_path = attachments_path
        self.no_attachments_path = no_attachments_path
    
    def OnItemAdd(self, item):
        try:
            if item.Attachments.Count > 0:
                self.save_attachments(item)
            else:
                self.save_body(item)
        except Exception as e:
            print("Error processing mail item:", e)

    def save_attachments(self, item):
        for attachment in item.Attachments:
            attachment.SaveAsFile(str(self.attachments_path / attachment.FileName))
            print(f"Attachment saved: {attachment.FileName}")

    def save_body(self, item):
        body_file_path = self.no_attachments_path / f"{item.Subject.replace(':', '').replace('/', '')}.txt"
        with open(body_file_path, 'w', encoding='utf-8') as file:
            file.write(item.Body)
            print(f"Body saved in file: {body_file_path}")

class Outlook:
    def __init__(self, account: str|None = None, attachments_path: str|None = None, no_attachments_path: str|None = None):
        self.outlook_app = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook_app.GetNamespace("MAPI")
        self.account = account
        self.inbox = self.namespace.GetDefaultFolder(6)  # Predeterminado a la bandeja de entrada
        self.messages = self.inbox.Items

        # Initialize paths
        self.attachments_path = Path(attachments_path) if attachments_path else Path.cwd() / "attachments"
        self.no_attachments_path = Path(no_attachments_path) if no_attachments_path else Path.cwd() / "no_attachments"
        self.attachments_path.mkdir(exist_ok=True)
        self.no_attachments_path.mkdir(exist_ok=True)

        # Ensure paths are set before setting up the event handler
        self.inbox_events = DispatchWithEvents(self.messages, InboxEvents)
        self.inbox_events.attachments_path = self.attachments_path
        self.inbox_events.no_attachments_path = self.no_attachments_path
        
        if self.account is not None:
            found_account = next((acc for acc in self.namespace.Accounts if acc.DisplayName == self.account), None)
            if found_account:
                self.inbox = found_account.DeliveryStore.GetDefaultFolder(6)
                self.messages = self.inbox.Items
            else:
                print(f"Account not found: {self.account}")

    def get_last_email(self) -> Dict[str, str]:
        try:
            last_message = self.messages.GetLast()
            return {"subject": last_message.Subject,
                    "body": last_message.Body,
                    "received_time": last_message.ReceivedTime}
        except Exception as e:
            print(f"Error retrieving the last email: {e}")
            return {"subject": "", "body": ""}
        
    
    def get_emails_from_date(self, date_str: str) -> List[Dict[str, str]]:
        """Retrieve all emails from a specific date, considering UTC."""
        # Este método tiene un bug: devuelve listas vacías para correos de las última semana
        # parece haber un problema con el filtro o el formato de las fechas que cambia en la última semana
        utc_zone = pytz.utc
        local_zone = pytz.timezone('Europe/Madrid')  # Usando la zona horaria de Madrid

        # Parse the date
        date = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=local_zone)

        # Convert to UTC
        start_date = utc_zone.normalize(date.astimezone(utc_zone)).strftime('%m/%d/%Y 12:00 AM')
        end_date = utc_zone.normalize((date + timedelta(days=1)).astimezone(utc_zone)).strftime('%m/%d/%Y 11:59 PM')

        filter = f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"

        try:
            filtered_messages = self.messages.Restrict(filter)
            emails = []

            for message in filtered_messages:
                emails.append({"subject": message.Subject,
                               "body": message.Body,
                            "sender": message.SenderEmailAddress,
                            "received_time": message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')})
            return emails
        except Exception as e:
            print(f"Error filtering emails by date {date_str}: {e}")
            return []
    def get_emails_with_attachments(self, extension: str) -> List[Dict[str, str]]:
        """Retrieve all emails that have attachments with a specific extension."""
        emails = []
        if not extension.startswith('.'):
            extension = '.' + extension  
        try:
            for message in self.messages:
                has_matching_attachment = False
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        if attachment.FileName.lower().endswith(extension.lower()):
                            has_matching_attachment = True
                            break
                if has_matching_attachment:
                    emails.append({
                        "subject": message.Subject,
                        "body": message.Body,
                        "sender": message.SenderEmailAddress,
                        "received_time": message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
                    })
        except Exception as e:
            print(f"Error retrieving emails with {extension} attachments: {e}")
        return emails

    def on_new_mail_received(self, mail_item):
        """This method is called when a new mail item is added to the Inbox."""
        print("New mail received!")
        print("Subject:", mail_item.Subject)
        print("Body:", mail_item.Body)

if __name__ == "__main__":

    outlook = Outlook()
    print("Monitoring new emails. Press Ctrl+C to exit.")
    pythoncom.PumpMessages()
