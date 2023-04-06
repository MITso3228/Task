# This file is used to complete 1-3 items from "Acceptance criteria"

from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from helpers.common_helper import *
from helpers.xlsx_handler import *
from config.project_config import *
from loguru import logger

# Clear folders
clear_folder("downloads")
clear_folder("temp")

# Run ChromeDriver to download fresh WP template
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": os.path.join(project_dir, "downloads")
}
options.add_experimental_option("prefs", prefs)

logger.info("Downloading file...")
browser = webdriver.Chrome(options=options)

browser.get('https://datacollection.jrc.ec.europa.eu/guidelines/wp-table-template')

browser.find_element(By.XPATH, "//p[text()='Files can be downloaded here after:']/..//a").click()

time.sleep(10)

browser.close()

# Copy WP template to temp file
copied_file = copy_xlsx_document(os.path.join(downloads_dir, default_file_name), os.path.join(temp_dir, "Test1.xlsx"))

# Copy specific row in specific table
copy_row_by_direction(target_row=4, length=10, direction="Down", sheet="Table 1.1 Data availability", source_file=copied_file)

# Fill specific table with random data
fill_table_with_random_data(start_row=4, length=10, sheet="Table 1.1 Data availability", source_file=copied_file)
