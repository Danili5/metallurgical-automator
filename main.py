import os
import time
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

load_dotenv()

TEMPLATE_PATH = os.getenv("TEMPLATE_DIRECTORY")
REPORT_PATH = os.getenv("REPORT_DIRECTORY")
DATA_PATH = os.getenv("DATA_DIRECTORY")

class Main(FileSystemEventHandler):
    def on_created(self, event):
        try:
            time.sleep(0.5)

            report_template = Document(list(Path(TEMPLATE_PATH).glob("*.docx"))[0])
            report_title = None
            test_data = {f"+{key} {sub_key}+": sub_value for key, value in pd.read_csv(max([file for file in Path(DATA_PATH).glob("*.csv") if file.is_file()], key = lambda file: file.stat().st_ctime), skiprows = [0,2], index_col = 0).astype(str).to_dict().items() for sub_key, sub_value in value.items()}


            # finds and replaces the placeholders with the corresponding data
            for row in template.tables[1].rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:        
                        for key, value in data.items():
                            if key in paragraph.text:
                                try:
                                    value = f"{int(value):,}"
                                except Exception:
                                    print(Exception)

                                new_text = paragraph.text.replace(key, value)
                                paragraph.text = ""
                                paragraph.text = new_text

                                for run in paragraph.runs:
                                    run.font.name = "Bookman Old Style"
                                    run.font.size = Pt(10)

                        if "+" in paragraph.text:
                            paragraph.text = ""

            template.save(str(Path(REPORT_PATH)) + report_title)
        except Exception as exception:
            print(exception)

observer = Observer()

observer.schedule(Main(), Path(DATA_PATH))
observer.start()

try:
    while observer.is_alive():
        observer.join(1)
finally:
    observer.stop()
    observer.join()