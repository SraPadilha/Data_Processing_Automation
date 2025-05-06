# Data_Processing_Automation
Automates data processing and organization.
A lightweight and flexible automation tool for processing, organizing, and transforming structured data files into standardized outputs.

## Overview

This project was designed to streamline the processing of data files, automating repetitive tasks such as:
- Selecting and preparing the latest available files.
- Cleaning and formatting structured data.
- Generating organized outputs ready for further use or analysis.

Its modular structure allows easy integration into various workflows where data preparation is required.

## Features

- Automated file selection and processing.
- Data cleaning and normalization routines.
- Organized output generation in defined folder structures.
- Simple to deploy and extend for custom needs.

## Technologies Used

Language:
- Python 3.x

External Libraries:
- pandas
- openpyxl

Standard Library Modules:
- pathlib
- subprocess
- shutil

> *(Libraries can be installed via pip install pandas openpyxl.)*

## Getting Started

1. **Prepare Input Files**:  
   Place your data files (e.g., .xlsx) into the designated input folder.

2. **Run the Automation Script**:  
   Execute the controller script to process the files:
   
bash
   python process_controller.py

## Packaging (Optional)

This project can be converted into a standalone executable using tools like [PyInstaller](https://pyinstaller.org/en/stable/), allowing it to run on systems without requiring a Python installation.

However, since this script was designed for internal or scripted use and does not require broad distribution, no executable has been generated for now.

   
