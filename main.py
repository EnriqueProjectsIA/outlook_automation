import win32com.client


class Outlook:
    def __init__(self):
        self.email = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.email.GetDefaultFolder(6)
        self.messages = self.inbox.Items

    def get_last_message_subject(self):
        return self.messages.GetLast()

if __name__ == "__main__":
    outlook = Outlook()
    message = outlook.get_last_message_subject()
    print(message)