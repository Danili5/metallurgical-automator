import os
import sys
import time
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

load_dotenv()

DATA = os.getenv("DATA")
FORMS = os.getenv("FORMS")

class Main(FileSystemEventHandler):
    def on_created(self, event):    
        data = max(Path(DATA).glob("*.csv"), key = os.path.getmtime)

        print(data, flush = True)

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