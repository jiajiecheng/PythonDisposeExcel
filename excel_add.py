from xlutils.copy import copy
import pandas as pd
import os
import xlrd

data_dir = r"./data"  # 数据区
excel_model = r'model/doc1.xls' # 模板
new_file = r'./new_file/doc1.xls' # 输出文件
# 找到文件路径下的所有表格名称，返回列表
file_list = os.listdir(data_dir)
new_list = []

for file in file_list:
    file_path = os.path.join(data_dir, file)  # join 拼接完整的路径(需要合并的excel文件)
    dataframe = pd.read_excel(file_path,
                              keep_default_na=False)  # keep_default_na 对空数据不进行默认处理
    print(dataframe)
    dataframe.loc[len(dataframe.index)] = ""  # 添加空行
    new_list.append(dataframe)

df = pd.concat(new_list)  # 总数据


# 追加数据到新的Excel文件中
def write_excel_xls_append(file_path, out_file, df):
    data = xlrd.open_workbook(file_path)
    table = data.sheet_by_index(0)  # 拿到第一页
    nrows = table.nrows  # 为行数，整形
    ncolumns = table.ncols  # 为列数，整形

    book = xlrd.open_workbook(file_path, formatting_info=True)  # 保持Excel原格式
    workbook = copy(book)
    worksheet = workbook.get_sheet(0)
    for i in range(0, len(df)):  # 行
        for j in range(25):  # 列
            worksheet.write(nrows + 1 + i, j, str(df.iloc[i][j]))  # 追加数据，从3行开始写入
    workbook.save(out_file)


write_excel_xls_append(excel_model, new_file, df)
