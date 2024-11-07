import os
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from DocConverter import DocConverter
from JsonConverter import JsonConverter
import Checker as check
from Logger import Logger

def convert_to_json(filepath, filename, root, output_dir):
    try:
        document = JsonConverter(filepath)
        json_filename = f'{filename}.json'
        document.save_data_to_json(os.path.join(output_dir, json_filename))
    except UnicodeError:
        document.save_data_to_json(os.path.join(output_dir, json_filename))

def process_file(entry, root, doc, log, output_dir, count):
    if entry.name.startswith(("A_", "a_")):
        if not entry.name.endswith((".docx", ".DOCX")):
            filepath = os.path.join(root, entry.name)
            try:
                count += 1
                root_name = root.split("\\")[6]
                filepath = doc.convert_doc_to_docx_sp(filepath, f'{count}_{root_name}_{entry.name}.docx')
            except Exception as e:
                log.write_error_messages(root, entry.name, e, "DocConverter")
                return
            if check.search_word_in_docx(filepath, "Stammdatei", entry.name, root):
                convert_to_json(filepath, f'{count}_{root_name}_{entry.name}', root_name, output_dir)
        else:
            if check.search_word_in_docx(filepath, "Stammdatei", entry.name, root):
                root_name = root.split("\\")[6]
                convert_to_json(filepath, f'{count}_{root_name}_{entry.name}', root, output_dir)
                count += 1

def scan_directory(directory, doc, log, output_dir):
    count = 0
    with ProcessPoolExecutor() as executor:
        futures = []
        for root, dirs, files in os.walk(directory):
            if "\\NHF" in root or "\\SF" in root:
                with os.scandir(root) as it:
                    for entry in it:
                        if entry.is_file():
                            futures.append(executor.submit(process_file, entry, root, doc, log, output_dir, count))
        for future in as_completed(futures):
            future.result()

def main(directory="T:\\ima\\Alle\\Temp\\SCN\\_05_Grundstücke", output_dir="JSON"):
    log = Logger()
    doc = DocConverter()
    directory = Path(directory)
    os.makedirs(output_dir, exist_ok=True)
    scan_directory(directory, doc, log, output_dir)

if __name__ == "__main__":
    main()
