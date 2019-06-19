# WSUS-Scripts

Various Scripts to automate or aid the WSUS project.

Contains the following scripts:

## WSUS-Format
PowerShell Script that gathers necessary information and formats it to the WSUS Spreadsheet specification. Includes information from GLPI, Altiris, Active Directory, and Windows Update Services (WSUS).

## Create_Secure_Login
PowerShell Script that creates a 256-AES (32 Byte) keys that are encrypted into .key and .txt files based on inputted username and password. Used in WSUS-Format to run the script without inputting credentials in a secure manner.

## SortMonth
JavaScript code that updates the current number of tracked resolved issues and unresolved issues through a pre-defined formatted cell area. Updates all graph reports based on data inputted.

Example of Spreadsheet:
https://docs.google.com/spreadsheets/d/1N128ni7SrhEFjTvoPAAWvzvE7n352bH8xHMSWlZ3IvY/
