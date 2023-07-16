import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import datetime
import time

# Define file name
file_name = 'prices.xlsx'

# Check if file exists
if os.path.exists(file_name):
    wb = load_workbook(file_name)
else:
    wb = Workbook()
    sheet1 = wb.active

    headers = ['Link', 'Megjegyzes', 'Product ID', 'Location']  # Add 'Location' to headers
    sheet1.append(headers)
    wb.save(filename=file_name)

sheet1 = wb['Sheet'] if 'Sheet' in wb.sheetnames else wb.create_sheet('Sheet')
sheet2 = wb['Sheet2'] if 'Sheet2' in wb.sheetnames else wb.create_sheet('Sheet2')

options = Options()
options.add_argument("--headless")
options.add_argument("start-maximized")
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")
driver = webdriver.Firefox(options=options)

user_agent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"

previous_prices = {}
current_date = datetime.datetime.now().strftime('%Y.%m.%d')

date_column = None
product_id_column = None
location_column = None  # Add a new column variable for the location data

for column in sheet1.columns:
    if column[0].value == current_date:
        date_column = column[0].column
        print(f"Using existing column for date: {current_date}")
        break

if date_column is None:
    date_column = sheet1.max_column + 1
    sheet1.cell(row=1, column=date_column, value=current_date)
    print(f"Created new column for date: {current_date}")

headers = sheet1[1]
for index, header in enumerate(headers):
    if header.value == 'Product ID':
        product_id_column = index + 1
        print("Product ID column already exists.")
        break

if product_id_column is None:
    product_id_column = sheet1.max_column + 1
    sheet1.cell(row=1, column=product_id_column, value='Product ID')
    print("Created new Product ID column.")

# Check if the location column exists
for index, header in enumerate(headers):
    if header.value == 'Location':
        location_column = index + 1
        print("Location column already exists.")
        break

# Add the location column if it doesn't exist
if location_column is None:
    location_column = product_id_column + 1  # Add it after the 'Product ID' column
    sheet1.cell(row=1, column=location_column, value='Location')
    print("Created new Location column.")

screenshots_folder = os.path.join(os.getcwd(), 'screenshots')
if not os.path.exists(screenshots_folder):
    os.makedirs(screenshots_folder)

current_date_folder = os.path.join(screenshots_folder, current_date)
if not os.path.exists(current_date_folder):
    os.makedirs(current_date_folder)

for index, row in enumerate(sheet1.iter_rows(min_row=2, values_only=True), start=2):
    url = row[0]

    try:
        response = requests.get(url, headers={"User-Agent": user_agent})
        response.raise_for_status()
    except requests.HTTPError as http_err:
        print(f'HTTP error occurred: {http_err}')
        price = "Nincs"
        sheet1.cell(row=index, column=date_column, value=price)
        continue

    soup = BeautifulSoup(response.text, 'html.parser')

    if "A keresett oldal nem található!" in soup.text:
        price = "Nincs"
    else:
        price_div = soup.find('div', {
            'class': 'listing-property justify-content-around col-12 col-sm col-print d-flex flex-column text-center print-border-end border-sm-end border-1 border-ash fs-7 font-family-secondary'})

        if price_div:
            price = price_div.find('span', {'class': 'fw-bold fs-5 text-nowrap'}).text

            try:
                if "millió Ft" in price:
                    price = price.replace(" millió Ft", "").replace(",", ".")
                    price = float(price) * 1000000
                else:
                    price = float(price)
            except ValueError:
                print(f"Could not convert price to number for URL: {url}")
                continue
        else:
            price = "Nincs"

    previous_price = previous_prices.get(url)

    if price != previous_price or sheet1.cell(row=index, column=date_column).value is None:
        sheet1.cell(row=index, column=date_column, value=price)
        previous_prices[url] = price
        print(f"Updated price for URL: {url}")

    driver.get(url)

    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll"))
        ).click()
    except Exception as e:
        print(f"Could not click on the cookie consent button. Error: {e}")

    total_width = driver.execute_script("return document.body.offsetWidth")
    total_height = driver.execute_script("return document.body.parentNode.scrollHeight")
    driver.set_window_size(total_width, total_height)
    time.sleep(2)
    screenshot_filename = f"{url.split('/')[-1]}_{current_date}.png"
    screenshot_path = os.path.join(current_date_folder, screenshot_filename)
    driver.save_screenshot(screenshot_path)
    print(f"Screenshot saved: {screenshot_path}")

    product_id = url.split('/')[-1]
    sheet1.cell(row=index, column=product_id_column, value=product_id)
    print(f"Product ID for URL: {url} is {product_id}")

    # Extract and write the location data to Excel
    location = soup.find('span', {
        'class': 'card-title px-0 fw-bold fs-4 mb-0 font-family-secondary'}).text
    sheet1.cell(row=index, column=location_column, value=location)
    print(f"Location for URL: {url} is {location}")

wb.save(filename=file_name)
driver.quit()
