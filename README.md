# OpenText eDOC DM 16.7.0 Automated Matter Deleter

## Overview
This Python script automates the process of deleting files within the OpenText eDOC DM 16.7.0 software environment. It streamlines the deletion of files associated with specific matter numbers or identifiers, enhancing efficiency and reducing manual effort.

## Features
- **Automated Deletion**: The script interacts with the OpenText eDOC DM software to delete files based on provided matter numbers.
- **Error Handling**: It includes error handling mechanisms to address internal errors and popup windows encountered during the deletion process.
- **Restart Mechanism**: In case of issues, the script restarts the application with appropriate time delays, ensuring smooth operation.
- **File Processing**: It reads matter numbers from a text file, processes them sequentially, and deletes corresponding files within the application environment.
- **Image Recognition**: The script uses .png files to identify specific elements within the application interface. These images serve as reference points for the script to perform actions such as clicking buttons or entering text.

## Usage
1. Ensure Python and the required libraries (pyautogui, win32gui, pygetwindow, psutil, PyPDF2) are installed.
2. Place the .png files in the same directory as the script. These images serve as reference points for interaction with the OpenText eDOC DM software interface.
3. Run the `Auto.py` script to initiate the automated matter deletion process.
4. The script will prompt for matter numbers stored in a text file. Ensure the file is formatted correctly with one matter number per line.
5. The script will interact with the OpenText eDOC DM software environment, deleting files associated with the provided matter numbers.
6. Monitor the script for any errors or issues. If encountered, the script will attempt to restart the application automatically.

## Additional Notes
- The `matterGrabber.py` script is included to extract matter numbers from PDF files. It is part of the workflow to prepare matter numbers for deletion.
- Customize the script as needed to fit specific requirements or configurations of the OpenText eDOC DM software environment.
- Ensure proper permissions and access rights are granted for the script to interact with the software and perform file deletion operations.

## Disclaimer
- This script is provided as-is without any warranty. Use it at your own risk.
- Ensure compliance with organizational policies and legal regulations when using the script to delete files within the OpenText eDOC DM software environment.
