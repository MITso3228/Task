import os
import xlwings
import openpyxl
from helpers.common_helper import project_dir
from config.tables_config import *
import random
import string
from loguru import logger

downloads_dir = os.path.join(project_dir, "downloads")
temp_dir = os.path.join(project_dir, "temp")


def copy_xlsx_document(source_file, target_file, sheet=None):
    # Copy xlsx document in temp dir with all formatting and validations
    # Copying Data Validation only for "Table 1.1 Data availability"

    # Open source_file
    logger.info(f"Opening [{source_file}] ...")
    wb_source = xlwings.Book(source_file)
    wb_source_sheets = reversed(wb_source.sheet_names)

    # Create temp_file
    logger.info(f"Creating [{target_file}] ...")
    wb = openpyxl.Workbook()
    wb.save(target_file)

    # Open temp file
    wb_target = xlwings.Book(target_file)

    # Copy all data and formatting from source_file to temp_file
    for s in wb_source_sheets:
        logger.info(f"Copying [{s}] Sheet to [{target_file}] ...")
        ws = wb_source.sheets(s)
        ws.api.Copy(After=wb_target.sheets(1).api)
        wb_target.save()

    # Delete default "Sheet" sheet from temp_file
    wb_target_sheets = wb_target.sheets
    for s in wb_target_sheets:
        if s.name == "Sheet":
            s.delete()
            wb_target.save()
            logger.info(f"Deleted default empty [Sheet] from [{target_file}]")

    # Copy table Data Validation for table (in this example used only for "Table 1.1 Data availability" table)
    wb_target_sheets = wb_target.sheet_names
    for s in wb_target_sheets:
        if s == "Table 1.1 Data availability":
            logger.info(f"Copying cells validation for [{s}] table ...")
            for column in range(65, (65 + int(tables[s].get("table_width")))):
                for row in range(4, (4 + int(tables[s].get("default_table_length")))):
                    ws_source_validation = wb_source.sheets(s).range(f"{chr(column)}{row}").api.Validation
                    try:
                        type = ws_source_validation.Type
                        alert_style = ws_source_validation.AlertStyle
                        operator = ws_source_validation.Operator
                        formula1 = ws_source_validation.Formula1
                        formula2 = ws_source_validation.Formula2
                        wb_target.sheets(s).range(f"{chr(column)}{row}").api.Validation.Delete()
                        wb_target.sheets(s).range(f"{chr(column)}{row}").api.Validation.Add(type,
                                                                                            alert_style,
                                                                                            operator,
                                                                                            formula1,
                                                                                            formula2)
                    except BaseException:
                        pass
                    wb_target.save()

    return target_file


def copy_row_by_direction(length, direction, sheet, source_file, target_row):
    # Copy full specific row (e.g. A2, B2, ..., K2) in specific direction <length> times
    wb_source = xlwings.Book(source_file)
    sheet_source = wb_source.sheets[sheet]

    logger.info(f"Copying [{target_row}] row for [{sheet}] table for [{length}] rows in [{direction}] direction with theirs validation ...")
    for column in range(65, (65 + int(tables[sheet_source.name].get("table_width")))):
        target_cell_validation = sheet_source.range(f"{chr(column)}{target_row}").api.Validation

        # Here should be <direction> validation

        for row in range(target_row + 1, (target_row + length + 1)):
            sheet_source.range(f"{chr(column)}{target_row}").copy(sheet_source.range(f"{chr(column)}{row}"))
            try:
                type = target_cell_validation.Type
                alert_style = target_cell_validation.AlertStyle
                operator = target_cell_validation.Operator
                formula1 = target_cell_validation.Formula1
                formula2 = target_cell_validation.Formula2
                sheet_source.range(f"{chr(column)}{row}").api.Validation.Delete()
                sheet_source.range(f"{chr(column)}{row}").api.Validation.Add(type, alert_style, operator, formula1, formula2)
            except BaseException:
                pass
            wb_source.save()


def fill_table_with_random_data(start_row, length, sheet, source_file):
    # Fill specific table with random values based on table specific (tables_config)
    # Start with specific row
    # This method fills only one sheet (table)
    wb_source = xlwings.Book(source_file)
    sheet_source = wb_source.sheets[sheet]

    logger.info(f"Filling [{sheet}] table with random data from [{start_row}] row to [{length}] rows Down ...")
    for column in range(65, (65 + int(tables[sheet_source.name].get("table_width")))):
        column_title = get_column_title(sheet_source, chr(column))
        for row in range(start_row, (start_row + length + 1)):
            if column_title in list(available_tables_data[sheet].keys()):
                sheet_source.range(f"{chr(column)}{row}").value = random.choice(available_tables_data[sheet].get(column_title))
            else:
                letters = string.ascii_lowercase
                sheet_source.range(f"{chr(column)}{row}").value = ''.join(random.choice(letters) for _ in range(10))
            wb_source.save()

    wb_source.app.quit()


def get_column_title(sheet, column):
    # Get available data for specific column based on table's specific (tables_config)
    title = sheet.range(f"{column}{tables[sheet.name].get('table_titles_row')}").value

    return title
