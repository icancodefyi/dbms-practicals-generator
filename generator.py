from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

practical_3_data = {
    'title': 'Write a Program to Constraints in SQL Language',
    'aim': 'To study and implement various constraints in SQL such as NOT NULL, UNIQUE, PRIMARY KEY, FOREIGN KEY, and CHECK constraints to enforce data integrity and rules at the database level.',
    'theory': (
        'Constraints in SQL are rules applied to columns in tables to ensure the accuracy and reliability of the data within the database. '
        'They enforce data integrity by restricting the type of data that can be inserted or updated in a table. Constraints can be specified '
        'either during table creation or afterwards using ALTER commands.\n\n'
        '- NOT NULL: Ensures columns cannot have NULL values, making input mandatory.\n'
        '- UNIQUE: Guarantees all values in a column are distinct, avoiding duplicates.\n'
        '- PRIMARY KEY: Combines uniqueness and non-nullability for row identification. Only one primary key per table.\n'
        '- FOREIGN KEY: Creates a relation between two tables; values must match primary key in referenced table ensuring referential integrity.\n'
        '- CHECK: Limits allowed values in a column based on a condition (e.g., age >=18).\n\n'
        'Implementing constraints maintains data consistency and enforces business rules inside the database.'
    ),
    'execution': [
        "CREATE DATABASE LibraryDB;",
        "USE LibraryDB;",
        ("CREATE TABLE Students ("
         " StudentID INT NOT NULL PRIMARY KEY,"
         " LastName VARCHAR(50) NOT NULL,"
         " FirstName VARCHAR(50),"
         " Age INT CHECK (Age >= 18),"
         " Email VARCHAR(100) UNIQUE,"
         " DepartmentID INT,"
         " FOREIGN KEY (DepartmentID) REFERENCES Departments(DepartmentID)"
         ");"),
        "ALTER TABLE Students ADD CONSTRAINT chk_Age CHECK (Age <= 60);"
    ],
    'output_description': (
        'Fig 3.1 shows the creation of the Students table with various constraints applied. '
        'Fig 3.2 displays the addition of the CHECK constraint on the Age column.'
    ),
    'images': [
        ('images/image.png', 'Students Table Creation with Constraints'),
        ('images/image.png', 'Age CHECK Constraint Added on Students Table')
    ],
    'outcomes': [
        'Successfully created tables with various constraints to enforce integrity.',
        'Applied NOT NULL, UNIQUE, PRIMARY KEY, FOREIGN KEY, and CHECK constraints practically.',
        'Matured ability to structure tables that maintain consistent and valid data entries.'
    ],
    'conclusion': (
        'This practical enhanced understanding of enforcing rules at the database level using constraints, '
        'improving data quality and consistency for better application reliability.'
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
        self.doc.save(filename)

def generate_practical_doc(practical_num, data):
    doc_gen = PracticalDocGenerator(practical_num)

    doc_gen.add_main_heading(f"Practical {practical_num}: {data['title']}")

    doc_gen.add_sub_heading("Aim:")
    doc_gen.add_paragraph(data['aim'])

    doc_gen.add_sub_heading("Theory:")
    doc_gen.add_paragraph(data['theory'])

    doc_gen.add_sub_heading("Execution:")
    for step in data['execution']:
        doc_gen.add_code_block(step)

    doc_gen.add_sub_heading("Output:")

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
    filename = generate_practical_doc(3, practical_3_data)
    print(f"Generated document: {filename}")
