from datetime import datetime

import pandas as pd
import os, sys


RES_TEXT_FILE = os.path.join(os.environ['USERPROFILE'], 'Downloads', '当日汇报材料文本.txt')

def load_current_day_data():
    # 获取当前日期
    current_date = datetime.now().date().strftime("%#m月%#d日")
    print(current_date)
    
    df = pd.read_excel(r"..\2024年一分公司立管改造日情况统计表.xlsx", sheet_name=current_date, header=[0, 1], skiprows=1)
    # df = pd.read_excel(r"..\2024年一分公司立管改造日情况统计表.xlsx", sheet_name='9月29日', header=[0, 1], skiprows=1)
    
    # header = ['序号', '开片小区', '施工队伍', '施工人数', '当日打眼数量', '累计打眼数量', '当日立管串数', '累计立管串数', '当日置换串数', '累计置换串数', '当日实际完成量', '累计实际完成量', '当日PMS系统录入量', '累计PMS系统录入量']
    df.columns = [col[0].replace('\n', '') if 'Unnamed' in col[1] else col[1]+col[0] for col in df.columns.values]
    df = df.dropna(subset=['施工人数'])
    df = df.astype({'施工人数': 'int', '当日打眼数量': 'int', '累计打眼数量': 'int', '当日立管串数': 'int', '累计立管串数': 'int'})
    df = df.set_index('开片小区')

    # print(df)
    return df
    

def get_format_text(df:pd.DataFrame):
    previous_situation_text = ""
    current_situation_text = ""
    number_of_workers = []
    number_of_holes = []

    # 进度概况
    # 第一段
    previous_situation_text += "昨日计划施工4公里，实际完成立管{:d}串，完成{}公里，PMS系统内录入工程量{}公里。累计完成立管{:d}串，累计完成{}公里，PMS系统内累计录入工程量{}公里。\n".format(
        df.at['合计', '当日立管串数'],
        round(df.at['合计', '当日实际完成量'] / 1000, 3),
        round(df.at['合计', '当日PMS系统录入量'] / 1000, 3),
        df.at['合计', '累计立管串数'],
        round(df.at['合计', '累计实际完成量'] / 1000, 3),
        round(df.at['合计', '累计PMS系统录入量'] / 1000, 3),
    )
    # 第二段
    previous_situation_text += "其中，"
    for row in df[(df['管理营业所'].isna())].loc['小计'].itertuples():
        previous_situation_text += "{}昨日立管{:d}串，工程量{}公里，累计立管{:d}串，累计工程量{}公里；".format(
             row.施工队伍 if len(row.施工队伍) >= 4 else row.施工队伍 + '公司',
             row.当日立管串数,
             round(row.当日实际完成量 / 1000, 3),
             row.累计立管串数,
             round(row.累计实际完成量 / 1000, 3),
         )
    previous_situation_text = previous_situation_text[:-1].replace("0.0公里", "0公里") + "。\n"
    # 第三段
    current_situation_text += "今日计划进场施工人数{:d}人，实际{:d}人，其中罡世公司{:d}人，中石化建{:d}人。".format(
        (df['序号'].count() - 1) * 12,
        df.at['合计', '施工人数'],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '施工人数'].values[0],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '施工人数'].values[0],
    )

    # 输出文本
    # print(previous_situation_text)
    # print(current_situation_text)
    
    # 写入文件
    with open(RES_TEXT_FILE, 'w', encoding='utf-8') as f:
        f.write(previous_situation_text)
        f.write(current_situation_text)


if __name__ == "__main__":
    os.chdir(sys.path[0])
    current_day_data = load_current_day_data()
    get_format_text(current_day_data)