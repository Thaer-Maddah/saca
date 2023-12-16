# Writtin By Thaer Maddah
from docx import Document

# Check if there is a cover page
def has_cover_page(doc):
    # Iterate through the sections of the document
    for section in doc.sections:
        # Check if the section has a header
        if section.header:
            header = section.header
            # Iterate through the paragraphs in the header
            for paragraph in header.paragraphs:
                # Check if the paragraph contains text
                if paragraph.text.strip():
                    return True  # If there is text in the header, consider it as a cover page
    return False


# Search for word in text
def has_string(doc):
    for paragraph in doc.paragraphs:
        # Check if the paragraph contains a table of contents field
        if "ثائر" in paragraph.text:
            return True
    return False


# This is the true function that's detect table of contents
def has_table_of_contents(doc):
    body_element = doc._body._body
    #print(body_element.xml)
    # Search for table of contents (TOC)
    if "TOC" in body_element.xml:
        return True
    return False

def main():
    # Specify the path to your Word document
    doc_path = '../test/test.docx'

    # Load the Word document
    doc = Document(doc_path)

    # Check if the document has a cover page
    if has_cover_page(doc):
        print("The document has a cover page.")
    else:
        print("No cover page found in the document.")

    # Check if the document has a table of contents
    if has_table_of_contents(doc):
        print("The document has a table of contents.")
    else:
        print("No table of contents found in the document.")

    #has_table_of_contents3(doc)

if __name__ == "__main__":
    main()
