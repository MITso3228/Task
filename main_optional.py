# This file is used to show 4th item from "Acceptance criteria"
# In order to see the result:
# 1. WP template should be in "downloads" folder
# 2. Table "Table 1.1 Data availability" should be filled with different data, including date format, formulas, DDL, etc. (4h row!)

from helpers.common_helper import *
from helpers.xlsx_handler import *
from config.project_config import *

# Clear folders
clear_folder("temp")

# Copy WP template to temp file
copied_file = copy_xlsx_document(os.path.join(downloads_dir, default_file_name), os.path.join(temp_dir, "Test1.xlsx"))

# Copy specific row for specific table
copy_row_by_direction(target_row=4, length=10, direction="Down", sheet="Table 1.1 Data availability", source_file=copied_file)
