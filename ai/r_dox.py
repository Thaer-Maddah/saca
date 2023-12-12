from docx import Document

def convert_docx_to_txt(docx_filename, txt_filename):
    try:
        # Load the Word document
        doc = Document(docx_filename)

        # Create a text file and write the content
        with open(txt_filename, 'w', encoding='utf-8') as txt_file:
            for paragraph in doc.paragraphs:
                txt_file.write(paragraph.text + '\n')

        print(f'Successfully converted {docx_filename} to {txt_filename}')
    except Exception as e:
        print(f'Error: {e}')
# Example usage
docx_file = "../test/test1.docx"
txt_file = "../test/test.txt"
convert_docx_to_txt(docx_file, txt_file)
