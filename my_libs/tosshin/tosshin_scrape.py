import logging
from time import perf_counter

from my_libs.tosshin.tosshin_data_extraction import read_keywords
from my_libs.web_driver import initialize_driver


def scrape(keywords: list[str], output_folder: str) -> None:
    """
    Scrape function to execute the web scraping process.

    This function sets up logging, initializes the WebDriver, handles cookies,
    creates a new workbook, fetches and saves product data based on provided
    keywords, and saves the workbook to the specified output folder.

    Args:
        keywords (list[str]): List of search keywords to scrape data for.
        output_folder (str): Path to the folder where the output files will be saved.
    """
    # Configure logging
    start_time = perf_counter()
    logging.info("Starting the Tosshin scrape execution...")

    # Initialize the WebDriver
    driver = initialize_driver(headless=True)

    try:
        # Fetch and store product data
        logging.info("Fetching and saving product data...")
        read_keywords(driver, keywords, output_folder)

        logging.info("Workbook saved successfully to '%s'.", output_folder)

    except Exception as e:
        logging.error("An error occurred during execution: %s", e)

    finally:
        # Close the WebDriver properly
        logging.info("Closing WebDriver...")
        driver.quit()
        logging.info("WebDriver closed.")

        # Log completion message with additional context
        end_time = perf_counter()
        run_time = end_time - start_time
        logging.info("===================================================")
        logging.info("Tosshin scraping has successfully completed.")
        logging.info(f"Total runtime: {run_time:.6f} seconds")
        logging.info("Results saved to: %s", output_folder)
        logging.info("===================================================")
