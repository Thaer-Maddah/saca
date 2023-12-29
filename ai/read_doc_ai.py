import os
import win32com.client

def read_word_file(word_path, text_path):
    # Open Microsoft Word application
    word = win32com.client.Dispatch("Word.Application")

    # Load the specified Word file
    doc = word.Documents.Open(word_path)

    # Write the contents of the Word document to the text file
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(doc.Content.Text)

    # Close the Word document
    doc.Close()


if __name__ == "__main__":
    # Specify the path of the input Microsoft Word file
    word_path = "C:\\code/test/test1.docx"

    # Specify the path where you want to save the text file
    text_path = "C:\\code/test/test.txt"

    # Call the read_word_file function
    read_word_file(word_path, text_path)
