import docx
def search_word_in_docx(filepath, word,name,root):
    # Open the .docx filep
    doc = docx.Document(filepath)

    # Search through the paragraphs in the document
    for para in doc.paragraphs:
        if word in para.text:
            return True  # Word found, return True
    return False  # Word not found
