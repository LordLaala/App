import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
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

def process_file(root, filename, doc, log, output_dir, count):
    if filename.startswith(("A_", "a_")):
        if not filename.endswith((".docx", ".DOCX")):
            filepath = os.path.join(root, filename)
            try:
                count += 1
                root_name = root.split("\\")[6]
                filepath = doc.convert_doc_to_docx_sp(filepath, f'{count}_{root_name}_{filename}.docx')
            except Exception as e:
                log.write_error_messages(root, filename, e, "DocConverter")
                return
            if check.search_word_in_docx(filepath, "Stammdatei", filename, root):
                convert_to_json(filepath, f'{count}_{root_name}_{filename}', root_name, output_dir)
        else:
            if check.search_word_in_docx(filepath, "Stammdatei", filename, root):
                root_name = root.split("\\")[6]
                convert_to_json(filepath, f'{count}_{root_name}_{filename}', root, output_dir)
                count += 1

def file_generator(directory):
    for root, dirs, files in os.walk(directory):
        if "\\NHF" in root or "\\SF" in root:
            for filename in files:
                yield root, filename

def main(directory="T:\\ima\\Alle\\Temp\\SCN\\_05_Grundstücke", output_dir="JSON"):
    log = Logger()
    doc = DocConverter()
    directory = Path(directory)
    os.makedirs(output_dir, exist_ok=True)
    count = 0

    with ThreadPoolExecutor() as executor:
        futures = []
        for root, filename in file_generator(directory):
            futures.append(executor.submit(process_file, root, filename, doc, log, output_dir, count))

        for future in as_completed(futures):
            future.result()

if __name__ == "__main__":
    main()
