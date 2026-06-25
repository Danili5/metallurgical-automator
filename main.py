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

load_dotenv()

# requires a .env
DATA = os.getenv("DATA")
REPORTS = os.getenv("REPORTS")

class Main(FileSystemEventHandler):
    def on_created(self, event):    
        """set up"""

        # writes into the template file from the data file
        # for row in template_form.tables[1].rows:
        #         for cell in row.cells:
        #             for paragraph in cell.paragraphs:        
        #                 for key, value in data.items():
        #                     if key in paragraph.text:
        #                         try:
        #                             value = f"{int(value):,}"
        #                         except Exception:
        #                             print(Exception)

        #                         new_text = paragraph.text.replace(key, value)
        #                         paragraph.text = ""
        #                         paragraph.text = new_text

        #                         for run in paragraph.runs:
        #                             run.font.name = "Bookman Old Style"
        #                             run.font.size = Pt(10)

        #                 if "+" in paragraph.text:
        #                     paragraph.text = ""

        #     template_form.save(str(Path(REPORT_PATH) / report_title))

if __name__ == "__main__":
    observer = Observer()

    observer.schedule(Main(), DATA)
    observer.start()

    try:
        if observer.is_alive():
            print("started", flush = True)

        while observer.is_alive():
            observer.join(1)
    finally:
        observer.stop()
        observer.join()