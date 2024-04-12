
# Outlook Automation System

This Python project automates the interaction with Microsoft Outlook to manage and process emails. It allows for downloading attachments, saving email bodies, retrieving emails based on date, and handling new email notifications.

## Features

- **Real time - Download Attachments**: Automatically save attachments from incoming emails to a specified directory.
- **Real time - Save Email Bodies**: Save the content of emails without attachments to a designated folder.
- **Retrieve Emails by Date**: Fetch all emails from a specified date, adjusting for time zone discrepancies.


## Installation

To set up this project, follow these steps:

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/EnriqueProjectsIA/outlook_automation.git
   cd outlook_automation
   ```

2. **Install Requirements**:
   Ensure you have Python installed, and then install the required Python packages using:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure Settings**:
   Modify the `Outlook` class initialization parameters in the script to match your Outlook configuration and desired paths for saving attachments and email bodies.

## Usage

Run the script using Python. Ensure that Microsoft Outlook is running on your system and that you have permissions configured correctly:

```bash
python main.py
```

Press `Ctrl+C` to exit the monitoring of new emails.

## Limitations

- **Outlook Dependency**: Microsoft Outlook must be installed and running on your machine.
- **Permissions and Security**: Depending on your firewall and security settings in Outlook, you may need to adjust permissions to allow the script to interact with Outlook.
- **Compatibility**: This script is intended for Windows due to its dependency on `win32com.client`.
- **Firewall Settings**: Ensure that your firewall or antivirus software allows the script to execute and interact with Outlook.

## Troubleshooting

If you encounter issues related to email fetching or attachment downloads, verify your Outlook settings and ensure that there are no network restrictions blocking the script's operations. For detailed error logs, check the console outputs when running the script.

---

**Note**: At the current date and code state. This README assumes that you have a basic understanding of operating within a Windows environment and are familiar with Python programming. Adjust the paths and settings in the script as necessary to fit your specific needs.
