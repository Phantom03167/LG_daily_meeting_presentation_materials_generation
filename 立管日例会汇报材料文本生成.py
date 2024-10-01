from datetime import datetime

import pandas as pd
import os, sys


def load_current_day_data():
    # 获取当前日期
    current_date = datetime.now().date().strftime("%#m月%#d日")
    
    df = pd.read_excel(r"..\2024年一分公司立管改造日情况统计表.xlsx", sheet_name=current_date, header=[0, 1], skiprows=1)
    # df = pd.read_excel(r"..\2024年一分公司立管改造日情况统计表.xlsx", sheet_name='9月29日', header=[0, 1], skiprows=1)
    
    # header = ['序号', '开片小区', '施工队伍', '施工人数', '当日打眼数量', '累计打眼数量', '当日立管串数', '累计立管串数', '当日置换串数', '累计置换串数', '当日实际完成量', '累计实际完成量', '当日PMS系统录入量', '累计PMS系统录入量']
    df.columns = [col[0] if 'Unnamed' in col[1] else col[1]+col[0] for col in df.columns.values]
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

    # 前一日完成情况
    # 第一段
    previous_situation_text += "昨日计划施工6.45公里，昨日实际完成{:.3g}公里，累计完成{:.3g}公里。\n".format(
        round(df.at['合计', '当日实际完成量'] / 1000, 3),
        round(df.at['合计', '累计实际完成量'] / 1000, 3),
    )
    # 第二段
    previous_situation_text += "昨日"
    for row in df.itertuples():
        if not pd.isna(row.序号):
            previous_situation_text += "{}立管{:d}串、累计完成{:d}串、累计完成{:.3g}公里；".format(
                row.Index,
                row.当日立管串数,
                row.累计立管串数,
                round(row.累计实际完成量 / 1000, 3),
            )
    previous_situation_text = previous_situation_text[:-1] + "。"
    previous_situation_text += "开片小区累计完成立管{:d}串，共计{:.3g}公里。PMS系统内录入工程量{:.3g}公里。".format(
        df.at['合计', '累计立管串数'],
        round(df.at['合计', '累计实际完成量'] / 1000, 3),
        round(df.at['合计', '累计PMS系统录入量'] / 1000, 3),
    )

    # 当日完成情况
    # 表格数据
    number_of_workers = df['施工人数'][(df.index != '小计') & (df.index != '合计') & (df.index != '平昌楼')].to_list()
    number_of_holes = df['累计打眼数量'][(df.index != '小计') & (df.index != '合计') & (df.index != '平昌楼')].to_list()
    
    # 第一段
    current_situation_text += "今日计划进场施工人数{:d}人，实际{:d}人，其中罡世公司{:d}人，中石化建{:d}人，累计打眼数{:d}个。\n".format(
        df['序号'].count() * 12,
        df.at['合计', '施工人数'],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '施工人数'].values[0],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '施工人数'].values[0],
        df.at['合计', '累计打眼数量'],
    )
    # 第二段
    current_situation_text += "罡世公司昨日立管{:d}串，累计立管{:d}串，昨日工程量{:.3g}公里，累计工程量{:.3g}公里。\n".format(
        df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '当日立管串数'].values[0],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '累计立管串数'].values[0],
        round(df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '当日实际完成量'].values[0] / 1000, 3),
        round(df.loc[(df.index == '小计') & (df['施工队伍'] == '罡世'), '累计实际完成量'].values[0] / 1000, 3),
    )
    current_situation_text += "中石化建昨日立管{:d}串，累计立管{:d}串，昨日工程量{:.3g}公里，累计工程量{:.3g}公里。".format(
        df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '当日立管串数'].values[0],
        df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '累计立管串数'].values[0],
        round(df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '当日实际完成量'].values[0] / 1000, 3),
        round(df.loc[(df.index == '小计') & (df['施工队伍'] == '中石化建'), '累计实际完成量'].values[0] / 1000, 3),
    )

    # 输出文本
    print(previous_situation_text)
    print("=" * 150)
    print(' '.join(map(str, number_of_workers)))
    print(' '.join(map(str, number_of_workers)))
    print("-" * 100)
    print(current_situation_text)


if __name__ == "__main__":
    os.chdir(sys.path[0])
    current_day_data = load_current_day_data()
    get_format_text(current_day_data)