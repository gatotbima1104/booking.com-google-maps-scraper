import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from openpyxl import load_workbook
import os

def scrape_hotels():
    try:
        # Configure the ChromeOptions
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')  # To run Chrome in headless mode

        # Start the WebDriver
        driver = webdriver.Chrome(options=chrome_options)

        url = "https://www.google.com/maps"
        driver.get(url)

        # Make new object excel
        workbook = load_workbook('./hotel-list/hotel1-address.xlsx')
        worksheet = workbook.active

        column_b_values = [cell.value for cell in worksheet['B'][1:]]

        # Wait for search selector
        search_input = driver.find_element(By.XPATH, '//input[@id="searchboxinput"]')

        data = []
        hotel_count = 0
        default_delay = 1  # seconds

        for name in column_b_values:
            try:
                print(f"Processing hotel {hotel_count + 1}: {name}")
                hotel_count += 1

                # Click on the search input and clear any existing text
                search_input.click()
                search_input.clear()

                # Type the hotel name and press Enter
                search_input.send_keys(name)
                search_input.send_keys(Keys.ENTER)

                time.sleep(default_delay * 10)  # Wait for the page to load (adjusted to 10 seconds)

                # Extract data from the page
                try:
                    phone_element = driver.find_element(By.XPATH, '//div[@class="rogA2c"]')
                    phone = phone_element.text.strip()
                except NoSuchElementException:
                    phone = ''

                try:
                    website_element = driver.find_element(By.XPATH, '//a[@class="CsEnBe"]')
                    website = website_element.get_attribute('href')
                except NoSuchElementException:
                    website = ''

                result = {'name': name, 'phone': phone, 'website': website}
                data.append(result)
                print(result)
            except Exception as error:
                print(error)

        output_directory = "./"
        output_file_path = os.path.join(output_directory, 'maps-hotel1.xlsx')

        # Save the data to an Excel file
        workbook.save(output_file_path)

        driver.quit()

    except Exception as error:
        print(error)

if __name__ == "__main__":
    scrape_hotels()
