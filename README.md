
# Write-GoogleSheetDataAutopilot.ps1

## Overview

The `Write-GoogleSheetDataAutopilot.ps1` script is designed to run on client PCs to collect system information required for pre-enrollment into Microsoft Windows Autopilot/Intune. The script gathers key details such as the BIOS serial number, manufacturer, and the Windows Autopilot hardware hash, and uploads this information as a new row in a specified Google Sheet.

## Prerequisites

Before running the script, ensure you have met the following prerequisites:

1. **Google Cloud Project Setup**
    - Create a Google Cloud project.
    - Enable the Google Sheets API.
    - Create a service account and download the service account key JSON file.

2. **Script and File Setup**
    - Rename the downloaded service account key JSON file to `service-account-key.json`.
    - Place the `service-account-key.json` file into a `userdata` directory located in the same directory as the script.

3. **Google Sheets**
    - Ensure you have a Google Sheet set up to receive the data. The script can use a hardcoded `spreadsheetId` or one provided as a parameter.
    - Share the Google Sheet with the service account email found in the `service-account-key.json` file.

## Instructions

### 1. Create a Google Cloud Project and Enable the Google Sheets API

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project by clicking on the project dropdown at the top and selecting "New Project."
3. Name your project and click "Create."
4. With the project selected, navigate to the **API & Services** dashboard.
5. Click "Enable APIs and Services."
6. Search for "Google Sheets API" and click "Enable."

### 2. Create a Service Account and Download the Key

1. In the Google Cloud Console, navigate to **IAM & Admin** > **Service Accounts**.
2. Click "Create Service Account."
3. Provide a name for the service account and click "Create."
4. Assign the role "Editor" to the service account and click "Continue."
5. Under "Create Key," select "JSON" and click "Create."
6. Download the JSON key file.

### 3. Prepare the Script Environment

1. Rename the downloaded JSON key file to `service-account-key.json`.
2. Create a `userdata` directory in the same location as the script.
3. Place the `service-account-key.json` file into the `userdata` directory.

### 4. Share the Google Sheet with the Service Account

1. Open the `service-account-key.json` file in a text editor.
2. Locate the `client_email` field in the JSON file. This is the email address of the service account.
3. Open your Google Sheet in a web browser.
4. Click the "Share" button in the top-right corner of the Google Sheet.
5. Enter the service account email from the JSON file in the "Add people and groups" field and set the permission to "Editor."
6. Click "Send" to share the Google Sheet with the service account.

### 5. Obtain the Google Sheets ID

1. Open your Google Sheet in a web browser.
2. Look at the URL in your browser's address bar. The Google Sheets ID is the long string of characters between `/d/` and `/edit`.
   - Example URL: `https://docs.google.com/spreadsheets/d/1E5rUayzVrFi_rjUoW0PKKEDsljlDFvvYQ-r9ztolOyA/edit`
   - In this example, the Google Sheets ID is `1E5rUayzVrFi_rjUoW0PKKEDsljlDFvvYQ-r9ztolOyA`.

### 6. Running the Script

To run the script, open an elevated PowerShell prompt (Run as Administrator) and navigate to the directory containing the script. You can execute the script with or without parameters:

```powershell
.\Write-GoogleSheetDataAutopilot.ps1
```

### 7. Refreshing the Access Token

If you encounter an error indicating that the access token has expired, you can refresh the token by running the script with the `-RefreshToken` parameter:

```powershell
.\Write-GoogleSheetDataAutopilot.ps1 -RefreshToken
```

### 8. Using Custom Spreadsheet IDs

If you want to upload the data to a different Google Sheet, you can provide a custom `spreadsheetId` when running the script:

```powershell
.\Write-GoogleSheetDataAutopilot.ps1 -SpreadsheetId "your_custom_spreadsheet_id"
```

### 9. Debugging

For additional debugging information, use the `-Debug` flag:

```powershell
.\Write-GoogleSheetDataAutopilot.ps1 -Debug
```

## Summary

This script is a powerful tool for IT administrators looking to automate the collection of system information for Windows Autopilot/Intune pre-enrollment. With robust error handling and flexible parameters, it can be easily integrated into various deployment workflows. 

NOTE: This README file as well as the scripts in this repository were created with Generative AI assistance.
