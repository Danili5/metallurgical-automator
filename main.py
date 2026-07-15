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
        if Path(event.src_path).is_file() and Path(event.src_path).suffix == '.csv' and '_2' not in Path(event.src_path).stem:
            data_file = Path(event.src_path)

            # a method to verify that the csv data file is loaded before reading it.
            if self._check(data_file) == 0:
                return

            data = pd.read_csv(data_file, skiprows=[0,2], index_col=0)

            logging.info(data)

    def _check(self, data_file, max_attemps = 20):
        attempts = 0

        while attempts < max_attemps:
            try:
                with open(data_file, 'r+'):
                    return 1
            except (PermissionError, FileNotFoundError, IOError):
                attempts += 1
                time.sleep(0.5)

        return 0

if __name__ == "__main__":    
    load_dotenv()

    DATA_DIRECTORY = os.getenv("DATA_DIRECTORY")
    REPORTS_DIRECTORY = os.getenv("REPORTS_DIRECTORY")
    
    logging.info(f"observation has begun")

    observer = Observer()

    observer.schedule(Main(), DATA_DIRECTORY, recursive = True)
    observer.start()

    try:
        while observer.is_alive():
            observer.join(1)
    finally:
        observer.stop()
        observer.join()