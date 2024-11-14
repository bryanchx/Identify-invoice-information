import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import os
import openpyxl
import re

# 设置Tesseract的路径
# pytesseract.pytesseract.tesseract_cmd = r'D:\ProgramData\Tesseract-OCR\tesseract.exe'

# 设置发票文件夹路径和输出Excel文件路径
folder_path = r'E:\workspace_py\Identify invoice information\Invoices'
output_path = r'E:\workspace_py\Identify invoice information\invoice_data.xlsx'
output_images_path = r'E:\workspace_py\Identify invoice information\ProcessedImages'  # 保存图片的路径

# 创建保存图片的文件夹（如果不存在）
os.makedirs(output_images_path, exist_ok=True)

# 创建一个工作簿
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(['发票文件名称', '文件格式', '发票号码', '发票所有信息'])

invoice_numbers = {}


def extract_invoice_number(text):
    match = re.search(r'发\s*票\s*号\s*码\s*:\s*(\S+)', text)
    return match.group(1) if match else None


# 遍历文件夹中的文件
for filename in os.listdir(folder_path):
    if filename.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png')):
        file_path = os.path.join(folder_path, filename)

        # 如果是PDF文件，使用PyMuPDF转换为图片
        if filename.lower().endswith('.pdf'):
            pdf_document = fitz.open(file_path)
            for i in range(len(pdf_document)):
                page = pdf_document[i]
                pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))  # 300 DPI
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # 保存处理后的图片
                img_filename = f"{os.path.splitext(filename)[0]}_page_{i + 1}.png"
                img.save(os.path.join(output_images_path, img_filename))

                text = pytesseract.image_to_string(img, lang='chi_sim')
                invoice_number = extract_invoice_number(text)

                # 打印提取的图片信息
                print(f'Processed PDF file: {filename}, Page: {i + 1}')
                print(f'Saved image: {img_filename}')
                print(f'Extracted text: {text}')
                print(f'Invoice number: {invoice_number}\n')

                sheet.append([filename, 'PDF', invoice_number, text])

                if invoice_number:
                    invoice_numbers[invoice_number] = invoice_numbers.get(invoice_number, 0) + 1

        else:  # 处理图片文件
            image = Image.open(file_path)
            text = pytesseract.image_to_string(image, lang='chi_sim')
            invoice_number = extract_invoice_number(text)

            # 保存处理后的图片
            img_filename = filename
            image.save(os.path.join(output_images_path, img_filename))

            # 打印提取的图片信息
            print(f'Processed Image file: {filename}')
            print(f'Saved image: {img_filename}')
            print(f'Extracted text: {text}')
            print(f'Invoice number: {invoice_number}\n')

            sheet.append([filename, 'Image', invoice_number, text])

            if invoice_number:
                invoice_numbers[invoice_number] = invoice_numbers.get(invoice_number, 0) + 1

# 保存Excel文件
workbook.save(output_path)

# 打印重复的发票号码
for number, count in invoice_numbers.items():
    if count > 1:
        print(f'发票号码: {number} 重复次数: {count}')
