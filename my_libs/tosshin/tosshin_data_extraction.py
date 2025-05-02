import logging
import os
from typing import Any, Optional

import my_libs.utils as Utils
from my_libs.tosshin.tosshin_xlsx_writer import MyTosshinExcel, TosshinData
from selenium import webdriver
from selenium.webdriver.common.by import By

SCREENSHOTS_LIST = []


def read_keywords(
    driver: webdriver.Chrome, keywords: list[str], output_folder: str
) -> None:
    """
    Scrape and store product data for a list of keywords.

    Args:
        driver (webdriver.Chrome): The Selenium WebDriver instance used for web interaction.
        keywords (list[str]): A list of search keywords to fetch data for.
        output_folder (str): The directory where images and workbooks will be saved.

    Returns:
        None: This function does not return a value. It initializes a workbook and processes each keyword.
    """
    if not keywords:
        logging.warning("No keywords provided. Skipping data fetch.")
        return

    logging.info("Fetching and saving data for keywords: %s", ", ".join(keywords))

    # Initialize the workbook
    workbook = MyTosshinExcel(output_folder)
    screenshot_folder_path = Utils.create_subfolder(
        output_folder, "Tosshin Screenshots"
    )

    for search_keyword in keywords:
        search_keyword = search_keyword.strip()
        if search_keyword:
            logging.info("Fetching data for search keyword: %s", search_keyword)
            scrape_keyword_data(
                driver,
                search_keyword,
                workbook,
                screenshot_folder_path,
            )
        else:
            logging.warning("Empty search keyword encountered. Skipping.")

    # for row, file in enumerate(SCREENSHOTS_LIST):
    #     workbook.add_screenshot(file, row)
    workbook.save_workbook()


def scrape_keyword_data(
    driver: webdriver.Chrome,
    keyword: str,
    workbook: MyTosshinExcel,
    screenshot_folder_path: str,
) -> None:
    """
    Scrape data for a keyword and write it to the worksheet.

    Args:
        driver (webdriver.Chrome): The Selenium WebDriver instance for web interaction.
        keyword (str): The search keyword for fetching data.
        workbook (MyTosshinExcel): The workbook instance to write the data to.
        image_dir (str): The os path of the image folder that stored screenshots

    Returns:
        None: This function does not return a value. It handles data scraping and writing to the workbook.
    """
    url = Utils.build_tosshin_url(keyword)
    logging.info("Fetching data from URL: %s", url)
    driver.get(url)

    data = fetch_data(driver, keyword)

    screenshot_path = os.path.join(screenshot_folder_path, f"{keyword} screenshot.png")
    if Utils.take_screenshot(screenshot_path, driver):
        SCREENSHOTS_LIST.append(screenshot_path)

    if data:
        data[TosshinData.URL] = url
        workbook.write_data_row(data)
    else:
        logging.warning("No result found for the keyword: %s", keyword)


def fetch_data(
    driver: webdriver.Chrome, keyword: str
) -> Optional[dict[TosshinData, Any]]:
    """
    Extract data from the first row of the OEM table on the Tosshin page.

    Args:
        driver (webdriver.Chrome): The Selenium WebDriver instance used for web interaction.
        keyword (str): The search keyword for which data is being fetched.

    Returns:
        (Optional[dict[TosshinDataKey, Any]]): A dictionary containing the extracted data, or None if no results are found.
    """
    try:
        # Check if the "Nothing found!" message is present
        no_results_message = driver.find_elements(
            By.CSS_SELECTOR, "div.parts-search__result__nothing strong"
        )
        if no_results_message:
            logging.warning("No results found for the search query.")
            return None

        # Find the OEM table. Assuming the first table is OEM.
        oem_table = driver.find_element(
            By.CSS_SELECTOR, "table.parts-search__result__table"
        )
        first_row = oem_table.find_element(By.CSS_SELECTOR, "tbody > tr:first-child")

        # Extract data from the first row
        maker = first_row.find_element(By.CSS_SELECTOR, "td:nth-child(2)").text.strip()
        weight = first_row.find_element(By.CSS_SELECTOR, "td:nth-child(3)").text.strip()
        price = first_row.find_element(By.CSS_SELECTOR, "td:nth-child(4)").text.strip()

        maker = " ".join(maker.split())

        logging.info(
            f"Extracted Data - Keyword: {keyword}, Maker: {maker}, Weight: {weight}, Price: {price}"
        )

        return {
            TosshinData.KEYWORD: keyword,
            TosshinData.MAKER: maker,
            TosshinData.WEIGHT: weight,
            TosshinData.PRICE: price,
        }

    except Exception as e:
        logging.error(f"Failed to extract OEM data: {e}")
        return None
