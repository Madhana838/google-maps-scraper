# Google Maps Event Organizer Scraper

A Python + Selenium script that scrapes business listings (name, address, phone number) 
for a given search query and location from Google Maps, and saves the results to an Excel file.

## Features
- Automated search on Google Maps using Selenium (Microsoft Edge)
- Scrolls through results and collects business links
- Extracts business name, address, and phone number for each listing
- Handles Google's consent/cookie screens automatically
- Pierces shadow DOM elements to reliably locate the search box
- Saves results to an Excel (.xlsx) file, with automatic backup if the file is locked
- Configurable: search query, location, and max number of results to process

## Example Output
Currently configured to scrape **event organizers in Chennai, Tamil Nadu**. 
Results are saved to `event_organizers_chennai.xlsx` with columns:
`Location | Business Name | Address | Phone Number | Link`

## Requirements
- Python 3.x
- Microsoft Edge browser (msedgedriver auto-managed by Selenium 4)
- Install dependencies:

```
pip install selenium openpyxl
```

## Usage
1. Edit the CONFIG section at the top of `mad.py` to set your search query, location, 
   and max results:

```python
SEARCH_LOCATION = "Chennai, Tamil Nadu, India"
SEARCH_QUERY = "event organizers"
MAX_RESULTS = 15
```

2. Run the script:

```
python mad.py
```

3. Results will be saved to the Excel file specified in `EXCEL_FILE`.

## Notes
- This project is for educational/personal use. Scraping Google Maps may be subject 
  to Google's Terms of Service, so use responsibly and avoid excessive request volume.