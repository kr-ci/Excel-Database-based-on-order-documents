# Excel-Database-based-on-order-documents

Purpose
The program is designed to create a database of contractors based on issued orders. It then allows checking the distance of the assigned carrier route using the Google Maps API.

THE REPOSITORY CONTAINS SAMPLE INPUT AND OUTPUT DATA

1. Description of pdfToExcelDataBase_v1.02.py
The script iterates through all files in the working directory, extracts desired information from PDF files using regular expressions, and places them in an XLSX file according to a predefined schema.

2. Description of distanceMaps.py
You need to insert an active Google Maps API key in the code. The program creates a distance matrix for destinations listed in the resulting database and returns the distance in kilometers, which is then updated in the database.

3. Built with:

  -Python 3.12.3

  -openpyxl 3.1.4

  -PyPDF2 3.0.1

  -googlemaps 4.10.0
