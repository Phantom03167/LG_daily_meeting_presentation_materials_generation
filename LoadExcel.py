from datetime import date
from datetime import datetime
from datetime import timedelta

import pandas as pd

DEFAULT_FILE_PATH = r"..\2025年一分公司立管改造日情况统计表.xlsx"     # 日统计表文件名
Y24_FILE_PATH = r"..\2024年一分公司立管改造日情况统计表.xlsx"
TODAY_DATE = date.today()    # 当前日期
INTERVAL_DAYS = 1 if TODAY_DATE.weekday() != 0 else 1     # 2025年立管改造任务剩余天数

def load_specific_day_data(date:datetime = TODAY_DATE, y24date:bool = True) -> tuple[pd.DataFrame, pd.DataFrame]:
    global TODAY_DATE
    global INTERVAL_DAYS
    # 获取当前日期
    # TODAY_DATE = date.today().replace(day=17)     # 自定义读取日期
    current_date = date.strftime(r"%#m月%#d日")
    current_date_df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name=current_date, header=[0,], skiprows=1)
    current_date_df = dateframe_preprocessing(current_date_df)

    while True:
        try:
            previous_date = (date - timedelta(days=INTERVAL_DAYS)).strftime(r"%#m月%#d日")    
            previous_date_df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name=previous_date, header=[0,], skiprows=1)
            break
        except:
            INTERVAL_DAYS += 1
            continue
    previous_date_df = dateframe_preprocessing(previous_date_df)

    if y24date:
        y24_date_df = pd.read_excel(Y24_FILE_PATH, sheet_name=current_date, header=[0, 1], skiprows=1)
        # df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name='9月29日', header=[0, 1], skiprows=1)
        y24_date_df = dateframe_preprocessing(y24_date_df, True)
    
        return current_date_df, previous_date_df, y24_date_df
    else:
        return current_date_df, previous_date_df, None

def dateframe_preprocessing(df:pd.DataFrame, multi_header: bool = False) -> pd.DataFrame:
    # header = ['序号', '开片小区', '施工队伍', '施工人数', '当日打眼数量', '累计打眼数量', '当日立管串数', '累计立管串数', '当日置换串数', '累计置换串数', '当日实际完成量', '累计实际完成量', '当日PMS系统录入量', '累计PMS系统录入量']
    if multi_header:
        df.columns = [col[0].replace('\n', '') if 'Unnamed' in col[1] else col[1]+col[0] for col in df.columns.values]
    else:
        df.columns = [col.replace('\n', '') for col in df.columns.values]
    df = df.dropna(subset=['开片小区', '施工人数', '当日打眼数量', '当日立管串数', '当日实际完成量'])
    df = df.astype({'施工人数': 'int', '当日打眼数量': 'int', '累计打眼数量': 'int', '当日立管串数': 'int', '累计立管串数': 'int', '当日置换串数': 'int', '累计置换串数': 'int'})
    df = df.set_index('开片小区')
    # print(df.index)
    return df
