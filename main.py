from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import os

###sets file title in properties to existing filename (excluding extensions like ".pdf" or ".docx")
def main():
    dir_path = os.getcwd() + "\src"
    files = []
    not_supported = []
    for filename in os.listdir(dir_path):
        file_path = os.path.join(dir_path, filename)
        if file_path.lower().endswith(".pdf"):            
            reader = PdfReader(file_path)
            writer = PdfWriter()

            writer.append_pages_from_reader(reader)
            writer.add_metadata({
            '/Title': os.path.splitext(os.path.basename(file_path))[0]
            })
            files.append(filename)
        elif file_path.lower().endswith(".docx"):
            document = Document(file_path)
            document.core_properties.title = os.path.splitext(os.path.basename(file_path))[0]
            document.save(file_path)
            files.append(filename)
        else:
            not_supported.append(f)
    print(f"{len(files)} file(s) retitled, listed below.")
    for f in files:
        print(f)
    print(f"{len(not_supported)} file(s) not retitled because filetypes not supported.")
    for f in not_supported:
        print(f)
if __name__ == "__main__":
   main()