import errno
import logging
import os
from enum import Enum
from typing import Any

import xlsxwriter
import xlsxwriter.format
import xlsxwriter.worksheet

import my_libs.utils as Utils
from my_libs.xlsxwriter_formats import DataAttr, FormatType, initialize_formats


class TosshinData(Enum):
    """
    Enum representing the columns in the Tosshin data Excel worksheet.

    Members:
        MAKER
        KEYWORD
        WEIGHT
        PRICE
        URL
    """

    MAKER = DataAttr(header="Maker", column=0)
    KEYWORD = DataAttr(header="Keyword", column=1)
    WEIGHT = DataAttr(header="Weight", column=2)
    PRICE = DataAttr(header="Price", column=3)
    URL = DataAttr()


class MyTosshinExcel:
    """
    Manages an Excel workbook for Tosshin data, including creating, saving, and writing data to the workbook.

    Attributes:
        workbook (xlsxwriter.Workbook): The Excel workbook instance.
        formats (dict[FormatType, xlsxwriter.format.Format]): Formatting styles for the workbook.
        worksheet (xlsxwriter.worksheet.Worksheet): The worksheet for Tosshin data.
        row_count (int): The current row count in the worksheet, starting from 1 after the header.

    Methods:
        __init__(output_dir: str) -> None:
            Initializes the workbook, formats, and worksheet. Adds headers to the worksheet.

        create_workbook(output_dir: str) -> xlsxwriter.Workbook:
            Creates and returns a new Excel workbook, logging the creation process.

        save_workbook() -> None:
            Saves and closes the workbook, handling file permission errors and logging success or failure.

        add_headers() -> None:
            Writes the header row to the worksheet and increments the row count.

        write_data_row(data: dict[TosshinData, Any]) -> None:
            Writes a row of Tosshin data to the worksheet based on the provided dictionary.
    """

    def __init__(self, output_dir: str) -> None:
        """
        Initialize a new Excel workbook for storing Tosshin data.

        Args:
            output_dir (str): The directory where the workbook will be saved.
        """
        self.workbook: xlsxwriter.Workbook = self.create_workbook(output_dir)
        self.formats: dict[FormatType, xlsxwriter.format.Format] = initialize_formats(
            self.workbook
        )
        self.worksheet: xlsxwriter.worksheet.Worksheet = self.workbook.add_worksheet(
            "Tosshin Data"
        )
        # self.screenshot_sheet: xlsxwriter.worksheet.Worksheet = (
        #     self.workbook.add_worksheet("Screenshots")
        # )
        self.row_count = 0
        self.add_headers()

    def create_workbook(self, output_dir: str) -> xlsxwriter.Workbook:
        """
        Create a new Excel workbook and add a worksheet for Tosshin data.

        Returns:
            xlsxwriter.Workbook: A new workbook instance.
        """
        output_file = os.path.join(output_dir, "Tosshin data.xlsx")
        logging.info("Creating new workbook at %s", output_file)
        workbook = xlsxwriter.Workbook(output_file)
        logging.info("Workbook created successfully")
        return workbook

    def save_workbook(self) -> None:
        """
        Save the workbook by closing it.

        Notes:
            - Autofits columns in the worksheet.
            - Saves and closes the workbook, handling any errors related to file permissions.
            - Logs a message indicating success or an error if the workbook cannot be saved.
            - Retries if the file is open elsewhere, prompting the user to close it and retry.
        """
        logging.info("Finalizing workbook by applying autofit and saving...")
        self.worksheet.autofit()
        while True:
            try:
                self.workbook.close()
                logging.info("Workbook successfully saved.")
                break  # Exit the loop if the workbook is saved successfully
            except OSError as e:
                if e.errno == errno.EACCES:  # Permission denied error
                    logging.error(
                        "PermissionError: Please close the Excel file if it is open and press Enter to retry."
                    )
                    input("Please close the Excel file and press Enter to retry...")
                else:
                    logging.error(
                        "An OSError occurred while saving the workbook: %s", e
                    )
                    input("An unexpected error occurred. Press Enter to retry...")
            except Exception as e:
                logging.error(
                    "An unexpected error occurred while saving the workbook: %s", e
                )
                input("An unexpected error occurred. Press Enter to retry...")

    def add_headers(self) -> None:
        """
        Writes the header row to the worksheet and increments the row count.
        """
        headers = Utils.get_enum_headers_row(TosshinData)
        self.worksheet.write_row(0, 0, headers, self.formats[FormatType.HEADER])
        self.row_count += 1

    def write_data_row(
        self,
        data: dict[TosshinData, Any],
    ) -> None:
        """
        Write a row of Tosshin data to the worksheet.

        Args:
            data (dict[TosshinDataKey, Any]): A dictionary containing Tosshin data to be written to the worksheet.
        """
        try:
            # Write data to the corresponding columns

            fields_to_write = [
                (TosshinData.MAKER, TosshinData.URL),
                (TosshinData.KEYWORD, None),
                (TosshinData.WEIGHT, None),
                (TosshinData.PRICE, None),
            ]

            for data_key, url_key in fields_to_write:
                is_currency = data_key == TosshinData.PRICE

                Utils.write_data(
                    self.worksheet,
                    self.formats,
                    self.row_count,
                    Utils.get_enum_col(data_key),
                    data,
                    data_key,
                    url_key=url_key,
                    is_currency=is_currency,
                )

            self.row_count += 1

        except Exception as e:
            logging.error(f"Failed to write data to row {self.row_count}: {e}")
            raise

    # def add_screenshot(self, file_path: str, row_idx: int) -> None:
    #     logging.info(f"Embedding screenshot at row {row_idx+1}: {file_path}")
    #     self.screenshot_sheet.set_column_pixels(0, 0, 500)
    #     self.screenshot_sheet.set_row_pixels(row_idx, 500)
    #     self.screenshot_sheet.embed_image(row_idx, 0, file_path)
