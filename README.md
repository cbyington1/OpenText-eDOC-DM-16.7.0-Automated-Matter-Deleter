# OpenText eDOC DM 16.7.0 Automated Matter Deleter

## Overview
This Python script automates the process of deleting files within the OpenText eDOC DM 16.7.0 software environment (Also called Hummingbird).

## Features
- **Automated Deletion**: The script interacts with the OpenText eDOC DM software to delete files based on provided matter numbers.
- **Error Handling**: It includes error handling mechanisms to address internal errors and popup windows encountered during the deletion process.
- **Restart Mechanism**: In case of issues, the script restarts the application with appropriate time delays.
- **File Processing**: It reads matter numbers from a .txt file, processes them sequentially, and deletes corresponding files within the application.
- **Image Recognition**: The script uses .png files to identify specific elements within the application. These images are reference points for the script to perform actions like clicking buttons or entering text.

## Usage
- Ensure Python and the required libraries (pyautogui, win32gui, pygetwindow, psutil, PyPDF2) are installed.
- Place the .png files in the same directory as the script.
- Run the `Auto.py` script to initiate the automated matter deletion process.
- The script will prompt for matter numbers stored in a text file. Ensure the file is formatted correctly with one matter number per line.
- The script will interact with the OpenText eDOC DM software environment, deleting files associated with the provided matter numbers.
- Monitor the script for any errors or issues. If encountered, the script will attempt to restart the application automatically.

## Additional Notes
- The `matterGrabber.py` script is included to extract matter numbers from PDF files. It is part of the workflow to prepare matter numbers for deletion.
- Customize the script as needed to fit specific requirements or configurations of the OpenText eDOC DM software environment.

## Disclaimer
- This script is provided as-is without any warranty. Use it at your own risk.
- The `matterGrabber.py` script is designed to extract matter numbers from a PDF file that adheres to a specific format or layout. It may require adjustments or modifications to work with PDFs of different structures.

## Why I Made This
As an intern, one of my task was repeatedly deleting files within the OpenText eDOC DM 16.7.0 software environment. The manual process was time-consuming and prone to errors as the software was written 2004, leading to frustration and inefficiency. To streamline this task I created this automated matter deletion tool. By automating the process, I was able to significantly reduce the time and effort required for file deletion, allowing me to focus on more meaningful work tasks

