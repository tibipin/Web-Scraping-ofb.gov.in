from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from bs4 import BeautifulSoup
import pandas
import os
import time

start_time = time.time()

# ===> Instantiate selenium webdriver and go to ofb search link
try:
    driver = webdriver.Edge('msedgedriver.exe')
except WebDriverException:
    print('Please visit README.md => msedgedriver executable is missing')

driver.get(url='http://ddpdoo.gov.in/vendor/general_reports/show/registered_vendors')

# ===> Read cwr keywords int a dataframe

keywords_file = 'keywords_ofb_india.xlsx'
keywords_df = pandas.read_excel(keywords_file)

# ===> Set output folder of page extractors

output_folder = fr'{os.getcwd()}\output'


# ===> Function to search based on keyword

def search_page(keyword):
    # Search based on keyword from excel file

    product_box = driver.find_element_by_xpath('//*[@id="wrd_frm"]/div[1]/div[4]/input')
    product_box.clear()
    product_box.send_keys(keyword)

    search_button = driver.find_element_by_xpath('//*[@id="wrd_frm"]/div[2]/div/button')
    search_button.click()

    # Extract html page of search results on 1st page

    html_page = driver.page_source
    with open(fr'{output_folder}\{keyword}_page_1.txt', 'w', encoding='utf-8') as f:
        f.write(html_page)
        f.close()

    # Parse 1st page for total number of pages and extract each html

    soup = BeautifulSoup(html_page, 'html.parser')
    if soup.findAll('ul', {'class': 'pagination'}):
        total_pages = [i.a for i in soup.findAll('ul', {'class': 'pagination'})[0].children]
        i = 2
        for page in total_pages:
            if 'data-ci-pagination-page' in page.attrs:
                if page.text.isdigit():
                    driver.get(page.attrs['href'])
                    html_page = driver.page_source
                    with open(fr'{output_folder}\{keyword}_page_{i}.txt', 'w', encoding='utf-8') as f:
                        f.write(html_page)
                        f.close()
                i += 1


# ===> Loop which extracts the html pages as txt files

for keyword in keywords_df['Keyword'].values:
    search_page(keyword=keyword)

# ===> Close the webdriver

driver.close()


# ===> Function to parse txt files

def file_parser(file):
    supa = BeautifulSoup(open(file, encoding='utf-8'), 'html.parser')

    # Extract main table with data found based on keyword

    main_table = supa.findAll('table')[0]

    # Extract all table rows in variable b

    rows = []
    for i in main_table.tbody.contents:
        try:
            if len(i.findAll('td')) > 1:
                a = [j for j in i.findAll('td')]
                rows.append(a)
        except AttributeError:
            pass

    # Clean up rows so the table has a homogenuous structure

    for i in range(len(rows)):
        if not rows[i][0].text.isdigit():
            rows[i].insert(0, rows[i - 1][0])
            rows[i].insert(1, rows[i - 1][1])
            rows[i].insert(2, rows[i - 1][2])

    for i in rows:
        products_table = pandas.DataFrame()
        for j in i:
            if j.table:
                products_table = products_table.append(pandas.read_html(str(j.table), header=0)[0])
                i.remove(j)
        i.insert(-2, products_table.to_dict(orient='index'))

    # Create lists to populate final dataframe
    numbers = []
    names = []
    products = []
    for i in rows:
        numbers.append(i[0].text)
        names.append(i[1].text.split('\n')[0])
        products.append(str(i[-3]))

    # Generate final data frame

    df = pandas.DataFrame(
        {'Index on OFB India Page': numbers,
         'Company Name': names,
         'Products': products,
         'Keyword': os.path.basename(file).split('_')[0]
         })

    return df


# ===> Loop to generate final report

final_df = pandas.DataFrame()

for filename in os.listdir(output_folder):
    filename_level_df = file_parser(file=fr'{output_folder}\{filename}')
    final_df = final_df.append(filename_level_df)

final_df = final_df.merge(keywords_df, on='Keyword')
final_df.to_excel('OFB_India_scrape.xlsx', index=False)

print(f'===> OFB India Scraper has run in: {(time.time() - start_time)} seconds')
print(f'===> Number of keywords searched: {len(keywords_df)}')
