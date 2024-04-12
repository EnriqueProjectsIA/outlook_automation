import win32com.client
from win32com.client import DispatchWithEvents
from typing import Dict, List
import pythoncom
import datetime as dt
from datetime import datetime, timedelta
import pytz
from pathlib import Path
import signal
import sys

class InboxEvents:
    """
    Class to handle events for new emails in the Outlook inbox.

    Attributes:
        attachments_path (Path): Directory path where attachments will be saved.
        no_attachments_path (Path): Directory path where the bodies of emails without attachments will be saved as text files.

    Methods:
        OnItemAdd(self, item): Triggered automatically when a new mail item is added to the inbox. It checks if the item has attachments. If yes, it saves the attachments; if no, it saves the email body.
        save_attachments(self, item): Saves all attachments from the mail item to the specified attachments directory.
        save_body(self, item): Saves the body of the mail item to the specified directory as a text file if there are no attachments.
    """
    def __init__(self, attachments_path:str|None = None, no_attachments_path:str|None = None):
        self.attachments_path = attachments_path
        self.no_attachments_path = no_attachments_path
    
    def OnItemAdd(self, item):
        """
        Called when a new item is added to the inbox. It checks if the item has attachments and handles them accordingly.
        If the item has attachments, it saves them to the specified path. If not, it saves the body of the email to another path.

        Parameters:
            item (MailItem): The email item that has been added to the inbox.

        Raises:
            Exception: Logs an error if there's an issue processing the mail item.
        """
        try:
            if item.Attachments.Count > 0:
                self.save_attachments(item)
            else:
                self.save_body(item)
        except Exception as e:
            print("Error processing mail item:", e)

    def save_attachments(self, item):
        """
        Saves all attachments from the specified mail item to the defined attachments path.

        Parameters:
            item (MailItem): The email item whose attachments need to be saved.

        Notes:
            Outputs a confirmation for each attachment saved.
        """
        for attachment in item.Attachments:
            attachment.SaveAsFile(str(self.attachments_path / attachment.FileName))
            print(f"Attachment saved: {attachment.FileName}")

    def save_body(self, item):
        """
        Saves the body of the specified mail item as a text file in the defined no-attachments path.

        Parameters:
            item (MailItem): The email item whose body needs to be saved.

        Notes:
            The filename is derived from the email's subject, sanitizing characters that may cause errors in file naming.
        """
        body_file_path = self.no_attachments_path / f"{item.Subject.replace(':', '').replace('/', '')}.txt"
        with open(body_file_path, 'w', encoding='utf-8') as file:
            file.write(item.Body)
            print(f"Body saved in file: {body_file_path}")

class Outlook:
    """
    Class to interface with Microsoft Outlook for retrieving and handling emails based on various criteria.

    Attributes:
        outlook_app (COM Object): The Outlook application object.
        namespace (COM Object): The MAPI namespace used for folder and email access.
        account (str | None): Specific account name to access in Outlook. Defaults to None for the default account.
        inbox (COM Object): The default inbox folder for the specified or default account.
        messages (COM Object): Collection of messages in the inbox.
        attachments_path (Path): Path where email attachments are saved.
        no_attachments_path (Path): Path where email bodies are saved when there are no attachments.
        inbox_events (COM Event Handler): Handler for inbox events using the DispatchWithEvents method.

    Methods:
        __init__(self, account: str | None = None, attachments_path: str | None = None, no_attachments_path: str | None = None): Initializes the Outlook application, sets paths for attachments and non-attachments, and sets up event handling.
        get_last_email(self) -> Dict[str, str]: Retrieves the most recent email's details such as subject, body, and received time.
        get_emails_from_date(self, date_str: str) -> List[Dict[str, str]]: Retrieves all emails from a specified date, considering timezone adjustments.
        get_emails_with_attachments(self, extension: str) -> List[Dict[str, str]]: Retrieves all emails that contain attachments with a specific file extension.
        on_new_mail_received(self, mail_item): Method called when a new mail is received; outputs the mail's subject and body.
    """
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
        """
        Retrieves the most recent email from the inbox.

        Returns:
            Dict[str, str]: A dictionary containing the subject, body, and received time of the last email, or an empty dictionary if an error occurs.

        Exceptions:
            Exception: Captures and prints any exceptions that occur during the retrieval process.
        """
        try:
            last_message = self.messages.GetLast()
            return {"subject": last_message.Subject,
                    "body": last_message.Body,
                    "received_time": last_message.ReceivedTime}
        except Exception as e:
            print(f"Error retrieving the last email: {e}")
            return {"subject": "", "body": ""}
        
    
    def get_emails_from_date(self, date_str: str) -> List[Dict[str, str]]:
        """
        Retrieves all emails from a specific date, adjusting for UTC.

        Parameters:
            date_str (str): The date from which emails are to be retrieved, formatted as "YYYY-MM-DD".

        Returns:
            List[Dict[str, str]]: A list of dictionaries, each containing details of an email (subject, body, sender, and received time).

        Notes:
            There is a known issue where emails from the last week might not be retrieved correctly due to date formatting or filter issues.

        Exceptions:
            Exception: Captures and prints any exceptions that occur during the filtering process.
        """
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
        """
        Retrieves all emails that contain attachments with a specified file extension.

        Parameters:
            extension (str): The file extension to filter by, which should not start with a dot (e.g., 'pdf' for PDF files).

        Returns:
            List[Dict[str, str]]: A list of dictionaries, each containing details of an email (subject, body, sender, and received time) that has at least one attachment with the specified extension.

        Exceptions:
            Exception: Captures and prints any exceptions that occur during the retrieval process.
        """

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

def signal_handler(sig, frame):
    print('You pressed Ctrl+C! Stopping...')
    pythoncom.PumpWaitingMessages()
    sys.exit(0)


if __name__ == "__main__":
    signal.signal(signal.SIGINT, signal_handler)
    outlook = Outlook()
    print("Monitoring new emails. Press Ctrl+C to exit.")
    pythoncom.PumpMessages()
