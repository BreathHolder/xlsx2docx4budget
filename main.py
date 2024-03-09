from openpyxl import load_workbook
from docx import Document

# Load the workbook
excel_path = './working_files/budget.xlsx'
workbook = load_workbook(filename=excel_path)
sheet = workbook.active

# Create a new Word document# Load the workbook
excel_path = './working_files/budget.xlsx'
workbook = load_workbook(filename=excel_path, data_only=True)  # Add data_only=True
sheet = workbook.active
doc = Document()

# Iterate through each row in the Excel file
for row in sheet.iter_rows(min_row=2, values_only=True):
    project_title, project_description, business_driver, business_value, business_risk, \
    budget_expense, internal_hours, external_hours, solutions_involvement, \
    solutions_hours, pmo_involvement, pmo_hours, total_expected_hours = row

    # Add data to the Word document
    doc.add_paragraph(f'Project title: {project_title}')
    doc.add_paragraph(f'Project description: {project_description}')
    doc.add_paragraph(f'Business driver: {business_driver}')
    doc.add_paragraph(f'Business value: {business_value}')
    doc.add_paragraph(f'Business risk: {business_risk}')
    doc.add_paragraph(f'Budget expense: {budget_expense}')
    doc.add_paragraph(f'Internal hours: {internal_hours}')
    doc.add_paragraph(f'External hours: {external_hours}')
    doc.add_paragraph(f'Solutions involvement: {solutions_involvement}')
    doc.add_paragraph(f'Solutions hours: {solutions_hours}')
    doc.add_paragraph(f'PMO involvement: {pmo_involvement}')
    doc.add_paragraph(f'PMO hours: {pmo_hours}')
    doc.add_paragraph(f'Total expected hours: {total_expected_hours}')

    # Add a page break after each project
    doc.add_page_break()

# Save the Word document
word_document_path = './working_files/budget.docx'
doc.save(word_document_path)