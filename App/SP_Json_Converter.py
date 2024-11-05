import os
from pathlib import Path
from DocConverter import DocConverter
from JsonConverter import JsonConverter
import Checker as check
from Logger import Logger

def convert_to_json(filepath,filename,root,output_dir):
        try:
            document = JsonConverter(filepath)
            json_filename = f'{filename}.json'
            document.save_data_to_json(os.path.join(output_dir, json_filename))
        except UnicodeError:
            document.save_data_to_json(os.path.join(output_dir, json_filename))

def main(directory="T:\ima\Alle\Temp\SCN\_05_Grundstücke", output_dir="JSON"):
    log = Logger()
    doc = DocConverter()
    directory = Path(directory)
    os.makedirs(output_dir, exist_ok=True)
    count = 0
    for root, dirs, files in os.walk(directory):
        if "\\NHF" in root or "\\SF" in root:
            for filename in files:
                if (filename.startswith("A_") or filename.startswith("a_")):
                    if not (filename.endswith(".docx") or filename.endswith(".DOCX")):
                        filepath = os.path.join(root,filename)
                        try:
                            count += 1
                            root_name = root.split("\\")[6]
                            filepath = doc.convert_doc_to_docx_sp(filepath, f'{count}_{root_name}_{filename}.docx')
                        except Exception as e:
                            log.write_error_messages(root,filename,e,"DocConverter")
                            continue
                        if check.search_word_in_docx(filepath, "Stammdatei", filename, root):
                            convert_to_json(filepath, f'{count}_{root_name}_{filename}', root_name, output_dir)

                    else:
                        if check.search_word_in_docx(filepath, "Stammdatei", filename, root):
                            root_name = root.split("\\")[6]
                            convert_to_json(filepath, f'{count}_{root_name}_{filename}',root, output_dir)
                            count += 1

