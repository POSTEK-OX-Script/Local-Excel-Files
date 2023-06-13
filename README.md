# Local-Excel-Files

This code is designed to print labels on a printer based on data retrieved from a or many excel files stored inside the printer

## Demo video

A demonstration of what this script can do can be found here

https://youtu.be/8b-7YVsjUfY

## Installation

To run this code, you will need to have Python 3.8+ installed on your printer(which comes standard on all POSTEK Printers that support OX Script after May 2023). You will be able to execute this code directly on your computer but the behavior is probably going to be different from behaviors on the printer as it is intended to be executed on the printer
  
You will also need the free POSTEK Companion app which can be downloaded here:

https://www.postekus.com/service-support/download/software/
    
 

## Usage

Currently the code requires two files to execute, an Excel file and a command file that holds all the information in regards to the actual label design. 

Both files are provided in this repository. Please note the command file is provided for 300 dpi machines, you can easily generate your own design using a label editing software like bartender and the Excel data can be easily replaced with your own Excel file. Be sure to change the part of the code that iterates through the excel file if you are using a different data structure for your excel file.

To execute this demo, simply load the .py file, Fixed Assets Data.xlsx(Source for the data) and command4-en.txt(file for the label design) into your printer through the POSTEK App. Then to run the program you can initiate it from the printer touch screen or the POSTEK App. 

- Initiating the program from printer touch screen
    - On your printer's touch screen, go to settings>Ox Script>[your file name].py. Press it and select run from the bottom right of the pop-up window
 
- Initiating the program from the POSTEK App
    - Inside the App, select Ox Script from the left hand side. Connect to the printer that you just moved the files to and select the file from the left hand side drop down, click run on the top right hand corner

## License

This code is released under the MIT License. Please see the LICENSE file for more information.

## Disclaimer

This software is provided "as is" and the author of the software is not liable for any damages or losses that may arise from the use of the software. Use at your own risk.
