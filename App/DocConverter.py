import doc2docx
from spire.doc import *
class DocConverter:

    @staticmethod
    def turn_off_processes(word):
        word.ScreenUpdating = False
        word.DisplayAlerts = False
        word.Visible = False

    @staticmethod
    def convert_doc_to_docx(doc_path, name):
        doc2docx.convert(input_path=doc_path, output_path="Words\\"+name)
        return "Words\\"+name

    def convert_doc_to_docx_sp(self,doc_path, name):
        doc = Document()
        doc.LoadFromFile(doc_path)
        doc.SaveToFile("Words\\"+name,FileFormat.Docx2016)
        doc.Close()
        return "Words\\"+name
