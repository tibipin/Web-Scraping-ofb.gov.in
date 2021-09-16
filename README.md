# OFB India WebScraper

Project developed for Sustainalytics to find indian suppliers of controversial weapons in the largest [indian vendor database](https://ofb.gov.in/vendor/general_reports/show/registered_vendors/search)

### Prerequisites

The scraper requires Microsoft Edge webdriver which can be downloaded from [here](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) \
Required python libraries can be installed using `pip install -r requirements.txt`

---

The scraper runs based on an Excel file: `kewords_ofb_india.xlsx` which has two columns: `Indicator` and `Keyword` \
The `Keyword` is used on the indian vendor database. \
The `Indicator` is used in the output file generation.


### Functionality

The scraper downloads all the results from the `Keyword` search in the folder `output`. \
The results are saved as `.txt`  files. \
The `.txt` files are then parsed and compiled in a results Excel file called `OFB_India_scrape.xlsx`


