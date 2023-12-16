# Writtin By Thaer Maddah

from docx import Document
from docx.shared import RGBColor

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

# Font section
def check_font(doc):
    body_element = doc._body._body
    #print(body_element.xml)
    return "w:val=\"40\"" in body_element.xml 

def is_font_justified(paragraph):
    # Check if the paragraph alignment is justified
    return paragraph.alignment == 3  # 3 corresponds to justified alignment

def check_justified_font(doc):
    for paragraph in doc.paragraphs:
        if is_font_justified(paragraph):
            print(f"Paragraph '{paragraph.text}' \nhas justified font.")

# Check font color
def is_red_color(color):
    # return red color
    return color == RGBColor(255, 0, 0)

def has_gradient_font_color(color):
    # Check if the color has variations in RGB components
    return len(set(color) - {0, 255}) > 1

def is_red_variant(color):
    # Define a range of RGB values corresponding to red variants
    red_range = range(180, 256)     # Adjust the range as needed
    green_range = range(0, 75)      # Adjust the range as needed
    blue_range = range(0, 75)       # Adjust the range as needed

    # Check if the color is a red variant
    return color[0] in red_range and color[1] in green_range and color[2] in blue_range

def check_font_color(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color.rgb if run.font.color else None

            if font_color is not None and is_red_variant(font_color):
                print(f"Red variant font color found in paragraph '{paragraph.text}'")


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

    if check_font(doc):
        print("Font size: 20")
    else:
        print("Font size it's not correct")

    check_font_color(doc)
    check_justified_font(doc)
if __name__ == "__main__":
    main()
