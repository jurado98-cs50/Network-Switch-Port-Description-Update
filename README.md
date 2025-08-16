# Network-Switch-Port-Description-Update
Cisco Interface Description Updater

Author: Daniel Jurado
License: MIT License (2025)

Overview

The Cisco Interface Description Updater is a Windows desktop tool built with Python and Tkinter. It allows network engineers to bulk update interface descriptions on Cisco switches using data from an Excel file.

The tool supports SSH and Telnet connections and provides real-time progress updates.

Features

Bulk update interface descriptions from an Excel spreadsheet.

Automatic SSH connection, with Telnet fallback.

Real-time status output and progress tracking.

Saves success/failure results back to the Excel file.

Progress bar and logging for clarity.

Fully compiled as a standalone .exe using PyInstaller (no Python installation required).

Requirements

Windows 10 or newer.

No Python required if using the .exe version.

If running from source:

Python 3.12+

netmiko

openpyxl

tkinter (usually included with Python)

Installation
Using the Executable

Download the Cisco_Interface_Description_Updater.exe.

Place it in a convenient folder.

Double-click to run.

Using Python Source (optional)

Clone the repository or download the .py files.

Install dependencies:

pip install netmiko openpyxl


Run the script:

python updater.py

Usage

Launch the application (EXE or Python script).

Accept the MIT license at startup.

Enter the switch IP address, username, and password.

Click Browse to select your Excel file containing interface and description data.

Click Run Update:

The program will attempt SSH first, then Telnet if SSH fails.

The progress bar will show completed interfaces.

Results will be saved back to the Excel file. The output box will show success or failure for each interface.

Click Exit to close the application.

Excel File Format
Column	Description
A	Interface Description (new)
B	Interface Name
C	Status (auto-updated)

The first row should contain headers. Data starts at row 2.

Notes

The tool requires network connectivity to the target Cisco switch.

Long-running updates are handled in a separate thread to prevent the GUI from freezing.

The .exe file was created using PyInstaller. No additional Python setup is needed for end users.

License

MIT License © 2025 Daniel Jurado

Full license included in the application.
