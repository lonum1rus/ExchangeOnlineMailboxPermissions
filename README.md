# ExchangeOnlineSharedMailboxPermissions

A PowerShell script to retrieve shared mailbox permissions in Exchange Online, including handling of ambiguous recipient errors and extraction of creation dates from mailbox identities. The script exports the results to a CSV file for easy analysis and reporting.

## Features

- Connects to Exchange Online using the V3 module.
- Retrieves shared mailbox permissions.
- Handles ambiguous recipient errors gracefully.
- Extracts creation dates from mailbox identities (if present).
- Exports results to a CSV file.

## Requirements

- Windows PowerShell
- Exchange Online Management Module
- Proper permissions to access shared mailbox data in Exchange Online

## Installation

1. **Clone the repository:**
    ```bash
    git clone https://github.com/lonum1rus/ExchangeOnlineSharedMailboxPermissions.git
    cd ExchangeOnlineSharedMailboxPermissions
    ```

2. **Install the Exchange Online Management module:**
    ```powershell
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    ```

## Usage

1. **Open PowerShell ISE as Administrator:**
    - Right-click on the PowerShell ISE icon and select "Run as Administrator".

2. **Run the Script:**
    - Load the script into PowerShell ISE:
      ```powershell
      .\Get-SharedMailboxPermissions.ps1
      ```

3. **Follow the Prompts:**
    - Enter your Office 365 admin credentials when prompted.

4. **Script Output:**
    - The script will export the permissions to a CSV file located at `C:\Temp\SharedMailboxes.csv`. Ensure this directory exists or modify the path in the script as needed.