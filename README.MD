# Data Cleaning Script
## Introduction
This script is designed to clean a specific excel sheet named itau_convertido.xlsx and save it as a new cleaned sheet named cleaned_itau_convertido.xlsx in the processado folder.

## Usage
- Clone the repository
- Run pip install pandas and pip install numpy to install the necessary dependencies
- Run the script by executing python <script_name>.py in the terminal or command prompt.

# Libraries Used
- pandas
- numpy

## Key steps in the Cleaning Process
1) Read the itau_convertido.xlsx excel sheet using pandas
2) Drop the first 8 rows and last row of the sheet
3) Drop columns that have all NaN values
4) Rename columns with descriptive names
5) Combine the values of the Parenteses and Atalho columns into a single column Atalho
6) Remove the Parenteses column
7) Reset the index of the data frame
8) Clean the data in the Data and Histórico columns
9) If a row in the Data column is NaN, concatenate the value of the current row's Histórico with the Histórico of the previous non-NaN row
10) Drop all NaN rows
11) Reset the index of the data frame
12) Extract the check number from the Histórico column and store it in a new column Cheques
13) Save the cleaned data frame as an excel sheet in the processado folder.

## Prerequisites
- Python 3.x
- pandas
- numpy
