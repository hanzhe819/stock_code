"""
请注意替换 'path/to/your/excel/file.xlsx' 和 'path/to/your/word/file.docx' 为实际的文件路径。
此外，如果你的工作表名称不是默认的 Sheet1，请修改 read_excel_cell 函数中的 sheet_name 参数。
这个脚本首先读取 Excel 中指定单元格的值，然后在 Word 文档中查找并替换指定的占位符。在你的例子中，将 [1] 替换为 Excel 单元格的值。
"""
import openpyxl
from docx import Document


def read_excel_cell(excel_file, sheet_name, cell_address):
    # 打开 Excel 文件
    workbook = openpyxl.load_workbook(excel_file)

    # 选择工作表
    sheet = workbook[sheet_name]

    # 读取指定单元格的值
    cell_value = sheet[cell_address].value

    # 关闭 Excel 文件
    workbook.close()

    return cell_value


def update_word_document(word_file, placeholder, replacement):
    # 打开 Word 文档
    doc = Document(word_file)

    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        # 查找并替换占位符
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement)

    # 保存更新后的 Word 文档
    doc.save(word_file)


# 设置你的文件路径
excel_file_path = './my_excel.xlsx'
word_file_path = './my_word.docx'

# 读取 Excel 单元格的值
excel_value = read_excel_cell(excel_file_path, 'Sheet1', 'A2')

# 更新 Word 文档中的占位符
update_word_document(word_file_path, '[1]', excel_value)

print("操作完成。")
