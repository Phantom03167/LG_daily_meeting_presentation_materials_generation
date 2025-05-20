import xlwings as xw
import sys

from datetime import datetime

DEFAULT_FILE_PATH = r"..\2025年一分公司立管改造日情况统计表.xlsx"     # 日统计表文件名
TODAY_DATE = datetime.today().date()    # 当前日期
MODIFY_DATE = [5.19]  # 修改日期列表，后面会转成工作表名

def determine_modify_date():
    global MODIFY_DATE
    if MODIFY_DATE:
        MODIFY_DATE = [datetime.strptime(format(d, ".2f"), r"%m.%d").strftime(r"%#m月%#d日") for d in MODIFY_DATE]
    else:
        MODIFY_DATE = [TODAY_DATE.strftime(r"%#m月%#d日")]

def modify_excel_xlwings():
    app = xw.App(visible=False)
    wb = app.books.open(DEFAULT_FILE_PATH)

    for sheet_name in MODIFY_DATE:
        print(f"{sheet_name}")
        ws = wb.sheets[sheet_name]
        
        # 读取前一天日期和指定日期
        previous_date = ws.range('AE2').value
        specified_date = ws.range('AF2').value

        # 计算最大行数，假设B列有数据到最后一行
        last_row = sys.maxsize

        for row in range(3, last_row + 1):
            cell_b = ws.range(f'B{row}')

            # 判断是否是合并单元格
            if cell_b.api.MergeCells:
                break

            val_b = cell_b.value
            val_d = ws.range(f'D{row}').value

            if val_b not in ("小计", "合计"):
                formula_p = f"=IFERROR($O{row} - _xlfn.XLOOKUP($B{row}, '{previous_date}'!$B:$B, '{previous_date}'!$O:$O), 0)"
                formula_s = f"=IFERROR($O{row} - _xlfn.XLOOKUP($B{row}, '{specified_date}'!$B:$B, '{specified_date}'!$O:$O), 0)"
            elif val_b == "小计" and val_d:
                print(val_b)
                print(val_d)
                formula_p = "=IFERROR($O{} - _xlfn.XLOOKUP($B{} & $D{}, '{}'!$B:$B & '{}'!$D:$D, '{}'!$O:$O), 0)".format(row, row, row, previous_date, previous_date, previous_date)
                formula_s = "=IFERROR($O{} - _xlfn.XLOOKUP($B{} & $D{}, '{}'!$B:$B & '{}'!$D:$D, '{}'!$O:$O), 0)".format(row, row, row, specified_date, specified_date, specified_date)
            else:
                print("*************")
                print(val_b)
                print(val_d)
                formula_p = "=IFERROR($O{} - _xlfn.XLOOKUP($B{} & $C{}, '{}'!$B:$B & '{}'!$C:$C, '{}'!$O:$O), 0)".format(row, row, row, previous_date, previous_date, previous_date)
                formula_s = "=IFERROR($O{} - _xlfn.XLOOKUP($B{} & $C{}, '{}'!$B:$B & '{}'!$C:$C, '{}'!$O:$O), 0)".format(row, row, row, specified_date, specified_date, specified_date)

            # 使用公式写入到 AB 和 AC 列
            print(formula_p)
            print(formula_s)
            ws.range(f'AB{row}').formula = formula_p
            ws.range(f'AC{row}').formula = formula_s

    wb.save()
    wb.close()
    app.quit()

if __name__ == "__main__":
    determine_modify_date()
    modify_excel_xlwings()
