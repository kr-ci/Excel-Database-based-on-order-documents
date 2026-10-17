# Excel Database from Freight Orders

> My first real data-engineering project: turning a pile of freight-forwarding order PDFs into a structured database of **which carrier drove which route, and at what rate**.

## Background

In freight forwarding, the history of transport orders holds valuable business knowledge: which carriers have already driven a given route and how much they charged. That knowledge was locked inside many individual PDF orders, so answering a simple question such as *"who has already done this lane, and for how much?"* meant opening documents one by one.

This tool extracts that history automatically and builds an Excel database, enriched with real road distances from Google Maps, so that carrier rates can be compared **per kilometre** – a common basis for pricing new orders and choosing carriers.

## How it works

1. **Extract** – `pdfToExcelDataBase_v1.02.py` scans all PDF orders in the working directory and uses regular expressions to pull out the order date, the carrier's tax ID (NIP) and name, loading and unloading countries and places, the loading reference and the freight price.
2. **Load** – the data is written into `database.xlsx` (openpyxl) using a fixed schema.
3. **Enrich** – `distanceMaps.py` sends every loading/unloading pair to the Google Maps Distance Matrix API and fills in the driving distance in km (`Database_with_Distance.xlsx`).
4. **Analyse** – the sheet calculates the **price per km** for every order, making rates on similar routes directly comparable.

## Output schema

| Column | Meaning |
|---|---|
| Nr pliku | Source PDF file |
| Data zlecenia | Order date |
| NIP | Carrier tax ID |
| Nazwa zleceniobiorcy | Carrier name and address |
| Kierunek z / Kierunek do | Loading / unloading country |
| Ref załadunku | Loading reference |
| Miejscowość załadunku / rozładunku | Loading / unloading place |
| Cena USD | Freight price |
| Dystans km | Driving distance (Google Maps) |
| Cena/km | Price per kilometre |

## Sample data

`input_data_samples/` contains two **anonymised** sample orders (fictional companies, people and addresses). `output_data_samples/` shows the resulting database before and after distance enrichment.

## Usage

1. Put the PDF orders in the working directory and run `python pdfToExcelDataBase_v1.02.py` – this creates `database.xlsx`.
2. Insert your Google Maps API key in `distanceMaps.py` and run it – this creates `Database_with_Distance.xlsx`.

## Tech stack

Python 3.12 · PyPDF2 · openpyxl · googlemaps (Distance Matrix API) · regular expressions

## Limitations

The extraction patterns are tailored to one specific order template; other document layouts require adapting the regular expressions. Google Maps API usage may incur costs above the free tier.

## License

GPL-3.0
