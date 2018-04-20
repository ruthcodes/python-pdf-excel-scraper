<h1>Python PDF/Excel Scraper</h1>

A quick python3 script I wrote to automate a task at work.

- Compares incoming data (.pdf file) to data already in the system (.xlsx file). 
- Returns a text file listing the data that is already in the system, and the necessary associated information about it.

I used PDFMiner.six to convert the pdf to a string (which is then written to a txt file), and openpyxl to parse the data in the excel file.
I also used pprint when exporting the data to make it more readable.
