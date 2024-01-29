# SurveyCTO Extracts filter
## Author: Ayuba Ahmed Bayugo
## Overview
This Python script is designed to filter and extract records associated with a specified client name from multiple sheets within an Excel file. The filtered records are then saved into a single workbook, maintaining the original sheet names. The script prompts the user for input such as the Excel file name, client name, output path, and file name for saving the processed Excel files.

## Requirements
Python (3.x recommended)
pandas library (pip install pandas)
openpyxl library (pip install openpyxl)
Usage
Clone the Repository:

## Clone this repository to your local machine:

git clone https://github.com/your-username/your-repository.git
cd your-repository


## Install Required Libraries:

Ensure you have Python installed on your machine.

## Install the required Python libraries using:

pip install pandas openpyxl
Run the Script:

## Execute the script by running the following command:

python CTO_extracts_fiterer.py
Follow the prompts to enter the required information (Excel file name, client name, output path, and file name).
Review Output:

The script will inform you about the progress, including sheets where no records were found for the specified client.
Check Output Workbook:

The script will save the filtered records into a single workbook with the specified file name and path. Check the specified output path for the generated workbook.
Notes:
Ensure the Excel file exists in the specified path.
The script uses the 'pandas' library for data manipulation and the 'openpyxl' library for Excel reading/writing.
