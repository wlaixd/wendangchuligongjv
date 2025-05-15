import os
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path, docx_path):
	cv = Converter(pdf_path)
	cv.convert(docx_path, start=0, end=None)
	cv.close()

def batch_convert_pdf_to_docx(folder_path, output_folder):
	# 检查输出文件夹是否存在，如果不存在则创建
	if not os.path.exists(output_folder):
		os.makedirs(output_folder)

	# 遍历文件夹中的PDF文件
	for filename in os.listdir(folder_path):
		if filename.endswith('.pdf'):
			pdf_path = os.path.join(folder_path, filename)
			docx_filename = filename.replace('.pdf', '.docx')
			docx_path = os.path.join(output_folder, docx_filename)
			convert_pdf_to_docx(pdf_path, docx_path)

# 设置输入文件夹和输出文件夹的路径
input_folder = r'C:\Users\PKF\Desktop\实习\ycq\新建文件夹'
output_folder = r'C:\Users\PKF\Desktop\实习\ycq\新建文件夹'

# 执行批量转换
batch_convert_pdf_to_docx(input_folder, output_folder)