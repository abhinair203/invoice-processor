from docx import Document
import os
import openpyxl

def generate_report():
    excel_file = "invoice_report.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(['Invoice ID', 'Total Quantity', 'Subtotal', 'Tax', 'Total'])

    for file_name in os.listdir('.'):
        if file_name.endswith('.docx'):
            doc = Document(file_name)

            # Extract Invoice ID
            invoice_id = file_name.replace(".docx", '')

            # Extract product quantities
            products = doc.paragraphs[1].text.split('\n')
            quantity = sum([int(product.split(':')[1]) for product in products if ':' in product])

            # Extract subtotal, tax, and total
            lines = doc.paragraphs[2].text.split('\n')
            subtotal = float(lines[1].split(':')[1])
            tax = float(lines[2].split(':')[1])
            total = subtotal + tax

            worksheet.append([invoice_id, quantity, subtotal, tax, total])

    workbook.save(excel_file)
    print(f"Report saved as {excel_file}")

generate_report()