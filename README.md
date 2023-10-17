# PST Attachment Extractor

This script automates the process of extracting email attachments from PST files. It uses the `win32com` module to interact with Microsoft Outlook, enabling the extraction of attachments based on specific criteria from the email subjects.

## Requirements

- `win32com` module

- Install the required dependencies using the following command:

    ```
    pip install pywin32
    ```

## Installation

1. Clone the repository or download the script.


2. Run the script by executing the following command:

    ```
    py main.py
    ```

3. Follow the prompts to enter the folder path containing PST files.

## Functionality

The script performs the following actions:

- Opens PST files and adds them to the Outlook application.
- Processes the folders within the PST files to extract email attachments.
- Saves the attachments to a specific directory based on the criteria provided.

## Usage

1. Run the script and provide the folder path containing the PST files.
2. Check the specified destination folder for the extracted attachments.

## Note

Ensure there is necessary permissions to access the PST files and MUST have Microsoft Outlook installed on system.
