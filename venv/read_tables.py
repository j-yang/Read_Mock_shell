from docx import Document

def get_table_text(docx_path):
    """
        获取Word文档中的表格内容
    """
    result = []
    document = Document(docx_path)  # 读入文件
    tables = document.tables  # 获取Word文件中的表格集
    for table in tables:  # 遍历每个表格
        table_content = []
        for row in table.rows:  # 从表格第一行开始循环读取表格数据
            row_content = get_cell_content(row.cells)
            table_content.append(row_content)
        result.append(table_content)

    return result


def get_cell_content(cells):
    """
        获取每一行中每一列的内容
    """
    row_content = []
    for cell in cells:  # 遍历每一行的每一个单元格
        # cell数量为表格最大列数+1，故对于较少列的行存在重复值，需去重
        if cell.text and cell.text not in row_content:
            row_content.append(cell.text)

    return row_content


docx_path = "D:\KL296-Ⅰ-01-TLF_V0.5-清洁版.docx"
result = get_table_text(docx_path)

print(result[2][1])


