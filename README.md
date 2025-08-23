# FCA Daily Prices Scraper

This project automates the scraping of **daily price data** from the [FCA Info Web](https://fcainfoweb.nic.in/reports/report_menu_web.aspx) portal.  
It uses **Python + Selenium** for scraping and **Pandas + OpenPyXL** for saving results into Excel.

---

## 📂 Project Structure

- `scraper.py` → Script for bulk scraping date ranges (Feb–Apr 2020).  
- `retry_scraper.py` → Script for scraping missing dates individually.  
- `sort_and_clean.py` → Script for sorting, cleaning, and assigning column headers to the dataset.  
- `report.tex` → LaTeX file describing the workflow and methodology.  
- `daily_prices_feb_apr2020.xlsx` → Output dataset (scraped results).  

---

## ⚙️ Requirements

- Python 3.8+
- Chrome browser
- ChromeDriver (managed automatically with `webdriver_manager`)
- MiKTeX / TeX Live (if you want to build the LaTeX report)

Install dependencies:

```bash
pip install -r requirements.txt
