# Lenovo-web-scraping

Small application wrote for planning department which improve buying spare parts process.The code take part number (one by one), paste them to Lenvo website and extract all another alternative part numbers suitable for the PN. It's really useful when some main part number is unvaiable for suppliers.

* selenium is used to walking through tabs,
* bs4 extracts HTML code which is need for further text processing,
* openpyxl to load workbook,
* xlswriter to save final file;
