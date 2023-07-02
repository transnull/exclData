import os
from datetime import time

from tools.tool import *
from tools.mydict import *

# 初始化处理
path = r"./"
os.chdir(path)  # 修改工作路径

workbook = openpyxl.load_workbook('test.xlsx')  # 返回一个workbook数据类型的值
# 优化表格
sheet = workbook['应用目录']

# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':

    # 等待函数执行完成
    initialize_table(filename="test.xlsx", sheet_name="应用目录")

    unin_list = get_col_data_unique(sheet, "统计标识")

    backup_rate_dict = MyDict({})
    three_fine_rate_dict = MyDict({})

    for unin in unin_list:
        print("正在统计:\t" + unin + "\t")
        backup_rate_dict.add(name=unin, content=calculate_backup_rate(sheet, unin))
        three_fine_rate_dict.add(name=unin, content=calculate_level_three_fine_rate(sheet, unin))
    print("保存数据中.........")
    for unin in unin_list:
        add_row_to_excel(stat_name=unin, data=backup_rate_dict.get(name=unin), sheet_name="备案率",
                         file_name="汇总数据.xlsx")
        add_row_to_excel(stat_name=unin, data=three_fine_rate_dict.get(name=unin), sheet_name="三级优良率",
                         file_name="汇总数据.xlsx")
