import os
from pathlib import Path
from loguru import logger

project_dir = Path(__file__).parent.parent


def clear_folder(folder_name):
    # Method to clear specific directory in project_dir

    files = os.listdir(os.path.join(project_dir, folder_name))
    for f in files:
        logger.info(f"Deleting [{f}] file from [{folder_name}] folder...")
        os.remove(os.path.join(project_dir, folder_name, f))
