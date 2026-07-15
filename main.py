import os
import sys
import time
import logging
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

logging.basicConfig(
    level=logging.INFO,
    format="> %(message)s at %(asctime)s.",
    datefmt="%H:%M:%S"
)

class Main(FileSystemEventHandler):
    def on_created(self, event): 
        path = Path(event.src_path)

        # guards to ensure the right data comes through
        is_file = path.is_file()
        is_csv = path.suffix == '.csv'
        is_primary = '_2' not in path.stem

        if is_file and is_csv and is_primary and self._check(path):
            logging.info(pd.read_csv(path, skiprows=[0,2], index_col=0))

    # a method to verify that the csv data file is loaded before reading it.
    def _check(self, data_file, max_attemps = 20):
        attempts = 0
        while attempts < max_attemps:
            try:
                with open(data_file, 'r+'):
                    return True
            except (PermissionError, FileNotFoundError, IOError):
                attempts += 1
                time.sleep(0.5)
        return False

if __name__ == "__main__":    
    load_dotenv()
    DATA_DIRECTORY = os.getenv("DATA_DIRECTORY")
    REPORTS_DIRECTORY = os.getenv("REPORTS_DIRECTORY")
        
    observer = Observer()
    observer.schedule(Main(), DATA_DIRECTORY, recursive = True)
    observer.start()

    logging.info(f"observation has begun")

    try:
        while observer.is_alive():
            observer.join(1)
    finally:
        observer.stop()
        observer.join()