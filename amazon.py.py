import openpyxl
from bs4 import BeautifulSoup

# Read the saved html file
with open('amaz1.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

#parse the HTML with BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

#create a new excel workbook and add a sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Find all div elements with the specified class
product_divs = soup.find_all('div', class_='puis-card-container s-card-container s-overflow-hidden aok-relative puis-expand-height puis-include-content-margin puis puis-v1qlj3vrpms3m32klujlujtgcs8 s-latency-cf-section puis-card-border')

# Iterate through each div and extract information
for div in product_divs:
    # Find the h2 element with the specified class for Product Name
    product_name_elem = div.find('h2', class_='a-size-mini a-spacing-none a-color-base s-line-clamp-4')
    product_name = product_name_elem.text.strip() if product_name_elem else ''

    # Find the span element with the specified class for Product Price
    product_price_elem = div.find('span', class_='a-price-whole')
    product_price = product_price_elem.text.strip() if product_price_elem else ''

    # Find the span element with the specified class for Reviews
    reviews_elem = div.find('span', class_='a-icon-alt')
    reviews = reviews_elem.text.strip() if reviews_elem else ''

    # Write data to the Excel sheet
    sheet.append([product_name, product_price, reviews])

# Save the Excel file
workbook.save('output1.xlsx')
