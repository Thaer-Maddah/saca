#!/usr/bin/python3
# Writtin By Thaer Maddah

from docx import Document
from docx.shared import RGBColor
#import zipfile
#import xml.etree.ElementTree as ET
import re
#from docx.shared import Inches
from PIL import Image
from io import BytesIO


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

# find table of contents using regular expression revised by claude 2.0
def has_table_of_contents1(doc):
    """Check if a document contains a table of contents"""
    body_text = doc._body._body.text    # we can replace _body.text with _body.xml
    
    # Use a regular expression to search for 'Table of Contents' or 'Contents'
    toc_regex = re.compile(r'(Table of Contents|Contents)', re.IGNORECASE)
    if toc_regex.search(body_text):
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
            return True
    return False


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


def has_images(doc):
    for rel in doc.part.rels:
        if 'image' in doc.part.rels[rel].target_ref:
            return True
    return False


def get_image_dimensions(image_bytes):
    with Image.open(BytesIO(image_bytes)) as img:
        return img.size


def find_images_with_dimensions(doc, target_dimensions_cm):
    matching_images = []

    # Convert target dimensions from cm to pixels
    target_dimensions_px = (int(target_dimensions_cm[0] / 2.54 * 96), int(target_dimensions_cm[1] / 2.54 * 96))
    #print(target_dimensions_px)

    for rel_id in doc.part.rels:
        rel = doc.part.rels[rel_id]
        if 'image' in rel.target_ref:
            image_part = rel.target_part
            image_dimensions = get_image_dimensions(image_part.blob)
            print(image_dimensions, rel_id)
            
            if image_dimensions == target_dimensions_px:
                matching_images.append(rel_id)

    return matching_images

# References

def has_hyperlinks(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.element.findall('.//w:hyperlink', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}) is not None:
                return True  # Document has at least one hyperlink

    return False  # No hyperlinks found in the document

def print_hyperlinks(doc):
    hyperlink_count = 0

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            hyperlink = run.element.find('.//w:hyperlink', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if hyperlink is not None:
                hyperlink_count += 1
                #url = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                #print(f"Hyperlink {hyperlink_count}: {url}")

    print(f"Total Hyperlinks Found: {hyperlink_count}")



from docx.opc.constants import RELATIONSHIP_TYPE as RT
def links(rels):
    for rel in rels.items():
        if rels[rel].reltype == RT.HYPERLINK:
            yield rels[rel]._target


def hyperlinks(doc):
    links = []
    for paragraph in doc.paragraphs:
        if not paragraph.hyperlinks:
            continue
        for hyperlink in paragraph.hyperlinks:
            links.append(hyperlink.address)
    return links


def is_a4_word_document(doc):
    section = doc.sections[0]  # Assuming the first section defines the paper size
    page_width, page_height = section.page_width, section.page_height
    print(page_width, page_height)
    return page_width == 7560310  and page_height == 10692130    # A4 dimensions in twips

def check_margins(doc):
    section = doc.sections[0]  # Assuming the first section defines the margins
    left_margin = section.left_margin
    right_margin = section.right_margin
    top_margin = section.top_margin
    bottom_margin = section.bottom_margin
    print(top_margin, bottom_margin)
    return (
        left_margin == right_margin == 720090  and  # 1440 twips per inch
        top_margin == bottom_margin == 720090 
    )




def has_watermark(doc):

    # Iterate through sections
    for shape in doc.sections[0].footer.paragraphs[0].runs:
        if shape.text.lower().startswith('watermark'):
            return True
    return False

def main():

    data = []
    # Specify the path to your Word document
    doc_path = '../test/test.docx'

    # Load the Word document
    doc = Document(doc_path)

    # Check if the document has a cover page
    if has_cover_page(doc):
        print("The document has a cover page.")
        data.append(5)
    else:
        print("No cover page found in the document.")
        data.append(0)

    # Check if the document has a table of contents
    if has_table_of_contents(doc):
        print("The document has a table of contents.")
        data.append(5)
    else:
        print("No table of contents found in the document.")
        data.append(0)

    if check_font(doc):
        print("Font size: 20")
    else:
        print("Font size it's not correct")

    if has_string(doc):
        print(f"The document has string {has_string(doc)}")
    else:
        print(f"The document has string {has_string(doc)}")

    check_font_color(doc)
    if check_justified_font(doc):
        print('The paragraph is justified')
        data.append(4)
    else:
        print('The paragraph is not justified')
        data.append(0)

    if has_images(doc):
        print("The document contains images.")
    else:
        print("No images found in the document")

    target_dimensions_cm = (6, 6)  # Assuming default DPI is 96, adjust if needed
    try:
        matching_images = find_images_with_dimensions(doc, target_dimensions_cm)
        print(matching_images)

        if matching_images:
            print(f"Images with dimensions {target_dimensions_cm} cm found:")
            for i, rel_id in enumerate(matching_images, start=1):
                print(f"Image {i}: Relationship ID = {rel_id}")
        else:
            print(f"No images with dimensions {target_dimensions_cm} cm found in the document.")
    except Exception as e:
        print(f"An error occurred: {e}")


    all_hyperlinks = hyperlinks(doc)
    if all_hyperlinks: 
        print(f"The document has links\n {all_hyperlinks}")
        data.append(2)
    else:
        print('No hyperlinks found in the document')
        data.append(0)

    if is_a4_word_document(doc):
        print("The Word document has A4 paper size.")
        data.append(2)
    else:
        print("The Word document does not have A4 paper size.")
        data.append(0)

    if check_margins(doc):
        print("The Word document has 2 cm margins on all sides.")
        data.append(2)
    else:
        print("The Word document does not have 2 cm margins on all sides.")
        data.append(0)
  
    if has_watermark(doc):
        print("The Word document contains a watermark.")
    else:
        print("The Word document does not contain a watermark.") 
    

    print(data)
if __name__ == "__main__":
    main()
