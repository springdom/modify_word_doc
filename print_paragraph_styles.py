from docx import Document

def print_document_info(document_path):
    # Create a Document object from the specified document path
    document = Document(document_path)

    # Print general document information
    print(f"Number of paragraphs: {len(document.paragraphs)}")
    print(f"Number of tables: {len(document.tables)}")
    print(f"Number of sections: {len(document.sections)}")

    # Print paragraph details
    print("Paragraphs:")
    for i, paragraph in enumerate(document.paragraphs, start=1):
        print(f"Paragraph {i}:")
        print(f"Text: {paragraph.text}")  # Print the text content of the paragraph
        print(f"Style: {paragraph.style.name}")  # Print the name of the paragraph's style
        print()

    # Print table details
    print("Tables:")
    for i, table in enumerate(document.tables, start=1):
        print(f"Table {i}:")
        print(f"Number of rows: {len(table.rows)}")  # Print the number of rows in the table
        print(f"Number of columns: {len(table.columns)}")  # Print the number of columns in the table
        print()

    # Print section details
    print("Sections:")
    for i, section in enumerate(document.sections, start=1):
        print(f"Section {i}:")
        print(f"Header: {section.header}")  # Print the content of the section's header
        print(f"Footer: {section.footer}")  # Print the content of the section's footer
        print()

def main():
    # Path to the document file
    document_path = r'word_doc\Course Design Guide AU Module 1.docx'
    # Call the function to print the document information
    print_document_info(document_path)

if __name__ == "__main__":
    main()
