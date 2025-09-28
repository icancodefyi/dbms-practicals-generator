from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import time

practical_dml_data = {
    'title': 'Perform Data Manipulation Operations LO3',
    'aim': 'To demonstrate various Data Manipulation Language (DML) commands in SQL such as INSERT, UPDATE, DELETE, and SELECT for manipulating and retrieving data from database tables.',
    'theory': (
        "Data Manipulation Language (DML) is a critical component of SQL, enabling users to interact with, modify, and retrieve the data stored in relational databases. "
        "It serves as the primary means by which data is inserted, updated, deleted, and queried. Understanding DML commands is fundamental for anyone working with databases, as these commands allow the management of the actual content within database tables.\n\n"

        "The INSERT command is used to add new records to a table. By specifying the target table and the data values, users can populate their tables with information crucial for their applications. "
        "It supports inserting single rows, as well as batch inserts for multiple records in one command, improving efficiency.\n\n"

        "The UPDATE command allows modification of existing data. This operation is essential for maintaining accuracy and relevance in the database by enabling changes to specific fields based on conditions. "
        "Without a WHERE clause, an UPDATE statement affects all records, so careful use of conditions is required to avoid unintended data alteration.\n\n"

        "The DELETE command removes data from tables. It can delete specific rows based on criteria or all rows if used without a WHERE clause. "
        "Deletion is permanent and often requires proper permissions and safeguards, as it affects database integrity.\n\n"

        "SELECT is the query language component of DML, used to retrieve data. It is one of the most powerful commands, supporting filters (WHERE), sorting (ORDER BY), and joins to combine data across multiple tables. "
        "SELECT statements can be simple or complex, forming the basis of database reporting and analysis.\n\n"

        "DML commands interact closely with transactions to maintain consistency and integrity of data. Changes made by INSERT, UPDATE, or DELETE operations can be committed permanently or rolled back to maintain stable database states. "
        "This transactional control allows for safe concurrent access and prevents data corruption.\n\n"

        "Mastering DML commands empowers users to build dynamic, responsive applications and provides the ability to manage and analyze data effectively. Through practical use of INSERT, UPDATE, DELETE, and SELECT, one gains a comprehensive understanding of how databases operate at the data level.\n\n"

        "Additionally, proper indexing, constraints, and optimization techniques can enhance the performance of DML operations, making them faster and more efficient. "
        "DML combined with best practices in database design ensures reliable, scalable, and secure data management essential for modern applications. "
    ),
    'execution': [
        {
            'step': '1.      Create Database and Table',
            'code': (
                "CREATE DATABASE LibraryDB;\n"
                "USE LibraryDB;\n"
                "CREATE TABLE Books (\n"
                "  BookID INT PRIMARY KEY,\n"
                "  Title VARCHAR(100),\n"
                "  Author VARCHAR(50),\n"
                "  PublishedYear INT\n"
                ");"
            )
        },
        {
            'step': '2.      Alter Table to Add Column',
            'code': "ALTER TABLE Books ADD COLUMN Genre VARCHAR(30);"
        },
        {
            'step': '3.      Insert Data Into Table',
            'code': (
                "INSERT INTO Books (BookID, Title, Author, PublishedYear, Genre) VALUES\n"
                "(1, 'DBMS Fundamentals', 'A. Kumar', 2022, 'Education'),\n"
                "(2, 'Learn SQL', 'S. Sharma', 2020, 'Reference');"
            )
        },
        {
            'step': '4.      Update Table Data',
            'code': "UPDATE Books SET Genre = 'Academic' WHERE BookID = 1;"
        },
        {
            'step': '5.      Select Table Data',
            'code': "SELECT * FROM Books WHERE Genre = 'Education';"
        },
        {
            'step': '6.      Delete Table Row',
            'code': "DELETE FROM Books WHERE BookID = 2;"
        },
        {
            'step': '7.      Truncate Table Rows (keep structure)',
            'code': "DELETE FROM Books;"
        },
        {
            'step': '8.      Drop Table',
            'code': "DROP TABLE Books;"
        }
    ],
    'output_description': (
        "Each step’s output demonstrates the proper execution of SQL DML commands. Screenshots show successful table creation, column alteration, record insertion, data updates, filtered selections, row deletions, and complete table removal for verified understanding."
    ),
    'images': [
        ('images/image.png', 'Insert Query Output'),
        ('images/image.png', 'Update Query Output'),
        ('images/image.png', 'Delete Query Output'),
        ('images/image.png', 'Select Query Output'),
    ],
    'outcomes': [
        'Practiced table creation and structure modification.',
        'Inserted new records, updated values, and deleted unwanted rows.',
        'Retrieved specific information using SELECT queries based on filter conditions.',
        'Gained thorough understanding of core DML operations in SQL.'
    ],
    'conclusion': (
        'This practical reinforced the fundamental role of DML in SQL, highlighting efficient manipulation and retrieval of relational data. Mastery of these commands supports reliable and explainable database operations.'
    )
}

class PracticalDocGenerator:
    def __init__(self, practical_number):
        self.doc = Document()
        self.practical_number = practical_number
        self.figure_count = 0
        self.set_page_margins()
        self.set_default_style()
        self.add_page_number_footer()

    def set_page_margins(self):
        section = self.doc.sections[-1]
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    def set_default_style(self):
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = Pt(12)

    def add_page_number_footer(self):
        section = self.doc.sections[-1]
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = 'PAGE'
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar)
        run._r.append(instrText)
        run._r.append(fldChar2)

    def add_main_heading(self, text):
        para = self.doc.add_paragraph()
        run = para.add_run(text)
        run.font.size = Pt(19.5)
        run.font.name = 'Times New Roman'
        run.bold = True
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        run.underline = True
        para.space_after = Pt(12)

    def add_sub_heading(self, text):
        para = self.doc.add_paragraph()
        run = para.add_run(text)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'
        run.bold = True
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        para.space_after = Pt(8)

    def add_paragraph(self, text):
        para = self.doc.add_paragraph(text)
        para.style = 'Normal'
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        para.space_after = Pt(10)
        para.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def add_bullet_list(self, items):
        for item in items:
            para = self.doc.add_paragraph(item, style='List Bullet')
            para.space_after = Pt(6)

    def add_code_block(self, code_text):
        para = self.doc.add_paragraph()
        run = para.add_run(code_text)
        run.font.name = 'Courier New'
        run.font.size = Pt(12)

        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:val'), 'clear')
        shading_elm.set(qn('w:color'), 'auto')
        shading_elm.set(qn('w:fill'), 'F2F2F2')  # light gray background
        para._element.get_or_add_pPr().append(shading_elm)

        pBorders = OxmlElement('w:pBdr')
        leftBorder = OxmlElement('w:left')
        leftBorder.set(qn('w:val'), 'single')
        leftBorder.set(qn('w:sz'), '6')
        leftBorder.set(qn('w:space'), '4')
        leftBorder.set(qn('w:color'), 'D9D9D9')
        pBorders.append(leftBorder)
        para._element.get_or_add_pPr().append(pBorders)

        para.paragraph_format.space_before = Pt(8)
        para.paragraph_format.space_after = Pt(8)
        para.paragraph_format.left_indent = Cm(0.75)
        para.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def add_image_with_caption(self, image_path, caption):
        self.figure_count += 1
        self.doc.add_picture(image_path, width=Inches(5))
        last_paragraph = self.doc.paragraphs[-1]
        last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        caption_text = f'Fig {self.practical_number}.{self.figure_count} {caption}'
        cap_para = self.doc.add_paragraph(caption_text)
        cap_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cap_run = cap_para.runs[0]
        cap_run.italic = True
        cap_run.font.name = 'Times New Roman'
        cap_run.font.size = Pt(12)
        cap_para.space_after = Pt(19.5)

    def save(self, filename):
        # Add a small delay to ensure any file handles are released
        time.sleep(0.2)
        
        # Try to save with error handling
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                self.doc.save(filename)
                print(f"Successfully saved: {filename}")
                break
            except PermissionError as e:
                if attempt < max_attempts - 1:
                    print(f"File is in use, waiting... (attempt {attempt + 1}/{max_attempts})")
                    print("Please close the file if it's open in Word or another application.")
                    time.sleep(2)
                else:
                    # If still can't save, try with a different name
                    timestamp = int(time.time())
                    backup_name = f"{filename.split('.')[0]}_{timestamp}.docx"
                    try:
                        self.doc.save(backup_name)
                        print(f"Saved as {backup_name} instead (original file was in use)")
                    except Exception as backup_error:
                        print(f"Error saving file: {backup_error}")
                        raise
            except Exception as e:
                print(f"Unexpected error: {e}")
                raise

    def close_open_file(self, filename):
        """Check if file exists and warn user"""
        if os.path.exists(filename):
            print(f"Warning: {filename} already exists. Please close it if open in Word or other applications.")
            print("Attempting to save in 3 seconds... Press Ctrl+C to cancel if you need to close the file.")
            time.sleep(3)

def generate_practical_doc(practical_num, data):
    doc_gen = PracticalDocGenerator(practical_num)
    doc_gen.add_main_heading(f"Practical {practical_num}: {data['title']}")

    doc_gen.add_sub_heading("Aim:")
    doc_gen.add_paragraph(data['aim'])

    doc_gen.add_sub_heading("Theory:")
    doc_gen.add_paragraph(data['theory'])

    doc_gen.add_sub_heading("Execution:")
    for step in data['execution']:
        # step is a dict with 'step' and 'code' keys
        doc_gen.add_paragraph(step['step'])  # Normal text for step description
        doc_gen.add_code_block(step['code'])  # Code block for sql commands

    doc_gen.add_sub_heading("Output:")
    doc_gen.add_paragraph(data['output_description'])

    for img_path, caption in data['images']:
        doc_gen.add_image_with_caption(img_path, caption)

    doc_gen.add_sub_heading("Lab Outcomes Achieved:")
    doc_gen.add_bullet_list(data['outcomes'])

    doc_gen.add_sub_heading("Conclusion:")
    doc_gen.add_paragraph(data['conclusion'])

    filename = f"Practical_{practical_num}.docx"
    doc_gen.save(filename)
    return filename

if __name__ == "__main__":
    filename = generate_practical_doc(3, practical_dml_data)
    print(f"Generated document: {filename}")
