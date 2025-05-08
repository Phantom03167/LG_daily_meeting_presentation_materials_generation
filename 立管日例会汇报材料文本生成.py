from datetime import datetime
from datetime import date
from datetime import timedelta

import math
import pandas as pd
import os, sys

DEFAULT_FILE_PATH = r"..\2025年一分公司立管改造日情况统计表.xlsx"     # 日统计表文件名
RES_TEXT_FILE = os.path.join(os.environ['USERPROFILE'], 'Downloads', '当日汇报材料文本.txt')
COMPLETED_AREA_COUNT = 38     # 完工小区数量
PAUSE_AREA_COUNT = 0     # 停工小区数量
Global_Digital_Precision = 2     # 全局数字精度
Global_Percentage_Precision = Global_Digital_Precision + 2     # 全局百分比精度

FHY_END_DATE = date(2025, 6, 30)     # 2025年立管改造任务上半年结束日期
TODAY_DATE = datetime.today().date()    # 当前日期
REMINDER_DAYS = (FHY_END_DATE - TODAY_DATE).days + 1     # 2025年立管改造任务剩余天数
REDUCTION_FACTOR = 0.85     # 2025年立管改造任务时间缩减系数
INTERVAL_DAYS = 1 if TODAY_DATE.weekday() != 0 else 3     # 2025年立管改造任务剩余天数

def load_current_day_data():
    # 获取当前日期
    # TODAY_DATE = datetime.now().date().replace(day=23)
    current_date = TODAY_DATE.strftime("%#m月%#d日")
    previous_date = (TODAY_DATE - timedelta(days=INTERVAL_DAYS)).strftime("%#m月%#d日")
    print(current_date)
    print(previous_date)
    
    current_date_df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name=current_date, header=[0,], skiprows=1)
    previous_date_df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name=previous_date, header=[0,], skiprows=1)
    # df = pd.read_excel(DEFAULT_FILE_PATH, sheet_name='9月29日', header=[0, 1], skiprows=1)

    current_date_df = dateframe_preprocessing(current_date_df)
    previous_date_df = dateframe_preprocessing(previous_date_df)
    return current_date_df, previous_date_df

def dateframe_preprocessing(df:pd.DataFrame) -> pd.DataFrame:
    # header = ['序号', '开片小区', '施工队伍', '施工人数', '当日打眼数量', '累计打眼数量', '当日立管串数', '累计立管串数', '当日置换串数', '累计置换串数', '当日实际完成量', '累计实际完成量', '当日PMS系统录入量', '累计PMS系统录入量']
    # df.columns = [col[0].replace('\n', '') if 'Unnamed' in col[1] else col[1]+col[0] for col in df.columns.values]
    df.columns = [col.replace('\n', '') for col in df.columns.values]
    df = df.dropna(subset=['施工人数', '当日打眼数量', '当日立管串数', '当日实际完成量'])
    df = df.astype({'施工人数': 'int', '当日打眼数量': 'int', '累计打眼数量': 'int', '当日立管串数': 'int', '累计立管串数': 'int'})
    df = df.set_index('开片小区')
    # print(df.index)
    return df

def get_format_text(cdate_df:pd.DataFrame, pdate_df:pd.DataFrame):
    previous_situation_text = str()
    current_situation_text = str()
    comtent = []
    number_of_workers = []
    number_of_holes = []

    # 进度概况
    # 昨日施工总体情况
    crow = cdate_df.loc['总计']
    prow = pdate_df.loc['总计']
    previous_situation_text += (
        "一分公司" +
        "，{}".format("前{}日".format(INTERVAL_DAYS) if INTERVAL_DAYS > 1 else "昨日") + 
        "计划完成{}公里".format(round((crow.上半年计划工程量 - prow.累计实际完成量) / math.ceil((REMINDER_DAYS + INTERVAL_DAYS) * REDUCTION_FACTOR) / 1000 * INTERVAL_DAYS, Global_Digital_Precision)) +
        "，实际" +
        # "完成立管{:d}串，".format(crow.累计立管串数 - prow.累计立管串数) +
        "完成{}公里".format(round((crow.累计实际完成量 - prow.累计实际完成量) / 1000, Global_Digital_Precision)) +
        # "，PMS系统内录入工程量{}公里".format(round((crow.累计PMS系统录入量 - prow.累计PMS系统录入量) / 1000, Global_Digital_Precision)) +
        # "，累计完成立管{:d}串".format(crow.累计立管串数) +
        "，累计完成{}公里".format(round(crow.累计实际完成量 / 1000, Global_Digital_Precision)) +
        # "，PMS系统内累计录入工程量{}公里".format(round(crow.累计PMS系统录入量 / 1000, Global_Digital_Precision)) +
        "。" +
        "其中，" +
        "民心工程昨日完成{}公里，".format(round((cdate_df.loc["总计"].民心工程累计完成量 - pdate_df.loc["总计"].民心工程累计完成量) / 1000, Global_Digital_Precision)) +
        "累计完成{}公里".format(round(cdate_df.loc["总计"].民心工程累计完成量 / 1000, Global_Digital_Precision)) +
        "。" +
        "按2025年立管改造上半年任务量{}公里计算".format(round(crow.上半年计划工程量 / 1000, Global_Digital_Precision)) +
        "，当前完成率为{:.2%}".format(round(crow.累计实际完成量 / crow.上半年计划工程量, Global_Percentage_Precision)) +
        "；按全年任务量{}公里计算".format(round(crow.全年计划工程量 / 1000, Global_Digital_Precision)) +
        "，当前完成率为{:.2%}".format(round(crow.实际完成率, Global_Percentage_Precision)) +
        "；按2025年民心工程任务量{}公里计算，".format(round(cdate_df.loc["总计"].民心工程计划工程量 / 1000, Global_Digital_Precision)) +
        "当前完成率为{:.2%}".format(cdate_df.loc["总计"].民心工程完成率) +
        "。" +
        # "当前PMS系统录入率为{:.2%}".format(round(crow.PMS录入率, Global_Percentage_Precision)) +
        # "，立管置换率为{:.2%}".format(round(crow.立管置换率, Global_Percentage_Precision)) +
        # "。" +
        "距上半年立管改造任务截止时间（{}）还剩{:d}天".format(FHY_END_DATE.strftime("%Y年%#m月%#d日"), REMINDER_DAYS) +
        "，按{:d}天计算倒排工期".format(math.ceil(REMINDER_DAYS * REDUCTION_FACTOR)) +
        "，每天需完成{}公里".format(round((crow.上半年计划工程量 - crow.累计实际完成量) / math.ceil(REMINDER_DAYS * REDUCTION_FACTOR) / 1000, Global_Digital_Precision)) +
        "。\n"
    )
    for crow in cdate_df.loc['合计'].itertuples():
        prow = pdate_df.loc['合计'].query('管理单位 == @crow.管理单位').iloc[0]
        previous_situation_text += (
            "{}区域".format(crow.管理单位) + 
            "，{}".format("前{}日".format(INTERVAL_DAYS) if INTERVAL_DAYS > 1 else "昨日") + 
            "计划完成{}公里".format(round((crow.上半年计划工程量 - prow.累计实际完成量) / math.ceil((REMINDER_DAYS + INTERVAL_DAYS) * REDUCTION_FACTOR) / 1000 * INTERVAL_DAYS, Global_Digital_Precision)) +
            # "，实际完成立管{:d}串".format(crow.累计立管串数 - prow.累计立管串数) +
            "，实际完成{}公里".format(round((crow.累计实际完成量 - prow.累计实际完成量) / 1000, Global_Digital_Precision)) +
            # "，PMS系统内录入工程量{}公里".format(round((crow.累计PMS系统录入量 - prow.累计PMS系统录入量) / 1000, Global_Digital_Precision)) +
            # "，累计完成立管{:d}串".format(crow.累计立管串数) +
            "，累计完成{}公里".format(round(crow.累计实际完成量 / 1000, Global_Digital_Precision)) +
            # "，PMS系统内累计录入工程量{}公里".format(round(crow.累计PMS系统录入量 / 1000, Global_Digital_Precision)) +
            "。" +
            "其中，" +
            "民心工程昨日完成{}公里，".format(round((crow.民心工程累计完成量 - prow.民心工程累计完成量) / 1000, Global_Digital_Precision)) +
            "累计完成{}公里".format(round(crow.民心工程累计完成量 / 1000, Global_Digital_Precision)) +
            "。"
        )
        
        previous_situation_text += (
            "按2025年立管改造上半年任务量{}公里计算".format(round(crow.上半年计划工程量 / 1000, Global_Digital_Precision)) +
            "，当前完成率为{:.2%}".format(round(crow.累计实际完成量 / crow.上半年计划工程量, Global_Percentage_Precision)) +
            "；按全年任务量{}公里计算".format(round(crow.全年计划工程量 / 1000, Global_Digital_Precision)) +
            "，当前完成率为{:.2%}".format(round(crow.实际完成率, Global_Percentage_Precision)) +
            "；按2025年民心工程任务量{}公里计算，".format(round(crow.民心工程计划工程量 / 1000, Global_Digital_Precision)) +
            "当前完成率为{:.2%}".format(crow.民心工程完成率) +
            "。" +
            # "当前PMS系统录入率为{:.2%}".format(round(crow.PMS录入率, Global_Percentage_Precision)) +
            # "，立管置换率为{:.2%}".format(round(crow.立管置换率, Global_Percentage_Precision)) +
            # "。" +
            "上半年立管改造任务倒排工期每天需完成{}公里".format(round((crow.上半年计划工程量 - crow.累计实际完成量) / math.ceil(REMINDER_DAYS * REDUCTION_FACTOR) / 1000, Global_Digital_Precision)) +
            "。\n"
        )
        
    previous_situation_text += "\n"
    
    # 昨日施工队工程量
    # previous_situation_text += "其中，"
    for crow in cdate_df[(cdate_df['施工队伍'].notnull())].loc['小计'].itertuples():
        prow = pdate_df[(pdate_df['施工队伍'].notnull())].loc['小计'].query('施工队伍 == @crow.施工队伍').iloc[0]
        previous_situation_text += (
            "{}".format(crow.施工队伍 if len(crow.施工队伍) >= 4 else crow.施工队伍 + '公司') +
            "昨日上岗{}人".format(crow.施工人数)
        )
        if (crow.累计立管串数 - prow.累计立管串数) or (crow.累计实际完成量 - prow.累计实际完成量):
            previous_situation_text += (
                "，{}".format("前{}日".format(INTERVAL_DAYS) if INTERVAL_DAYS > 1 else "") + 
                "计划完成{}公里".format(round((crow.上半年计划工程量 - prow.累计实际完成量) / math.ceil((REMINDER_DAYS + INTERVAL_DAYS) * REDUCTION_FACTOR) / 1000 * INTERVAL_DAYS, Global_Digital_Precision)) +
                # "，实际完成立管{:d}串".format(crow.累计立管串数 - prow.累计立管串数) +
                "，实际完成{}公里".format(round((crow.累计实际完成量 - prow.累计实际完成量) / 1000, Global_Digital_Precision))
            )
        else:
            previous_situation_text += (
                "，{}".format("前{}日".format(INTERVAL_DAYS) if INTERVAL_DAYS > 1 else "") + 
                "计划完成{}公里".format(round((crow.上半年计划工程量 - prow.累计实际完成量) / math.ceil((REMINDER_DAYS + INTERVAL_DAYS) * REDUCTION_FACTOR) / 1000 * INTERVAL_DAYS, Global_Digital_Precision)) +
                "，实际无工程量"
            )
            
        previous_situation_text += (
            # "，累计立管{:d}串".format(crow.累计立管串数) +
            "，累计完成{}公里".format(round(crow.累计实际完成量 / 1000, Global_Digital_Precision)) +
            "。"
            "按2025年立管改造上半年任务量{}公里计算".format(round(crow.上半年计划工程量 / 1000, Global_Digital_Precision)) +
            "，当前完成率为{:.2%}".format(round(crow.累计实际完成量 / crow.上半年计划工程量, Global_Percentage_Precision)) +
            # "；按全年任务量{}公里计算".format(round(crow.全年计划工程量 / 1000, Global_Digital_Precision)) +
            # "，当前完成率为{:.2%}".format(round(crow.实际完成率, Global_Percentage_Precision)) +
            "。"
        )
        
        previous_situation_text += "其中，"
        for scrow in cdate_df[(cdate_df['施工队伍'] == crow.施工队伍) & (cdate_df['监理单位'].notnull())].itertuples():
            try:
                sprow = pdate_df.loc[scrow.Index]
            except KeyError:
                sprow = pd.Series(0, index=pdate_df.columns)
                sprow.name = scrow.Index
            if scrow.施工状态 == "完工":
                continue
            previous_situation_text += "{}".format(scrow.Index)     # 开片小区名称
            if scrow.施工状态 == "待置换":
                previous_situation_text += "立管已完成，待置换；"
                continue
            if scrow.施工状态 == "停工" or not scrow.施工人数:
                previous_situation_text += "未施工；"
                continue
            previous_situation_text += "上岗{}人".format(scrow.施工人数)
            if not scrow.累计打眼数量 - sprow.累计打眼数量 and not scrow.累计实际完成量 - sprow.累计实际完成量:
                previous_situation_text += "，无工程量；"
                continue
            if not scrow.累计实际完成量 - sprow.累计实际完成量:
                previous_situation_text += (
                    "，打眼{:d}个".format(scrow.累计打眼数量 - sprow.累计打眼数量) +
                    "，累计打眼{:d}个".format(scrow.累计打眼数量) +
                    "；"
                )
            else:
                previous_situation_text += (
                    # "，立管{:d}串".format(scrow.累计立管串数 - sprow.累计立管串数) +
                    "，完成{}公里".format(round((scrow.累计实际完成量 - sprow.累计实际完成量) / 1000, Global_Digital_Precision)) +
                    # "，累计立管{:d}串".format(scrow.累计立管串数) +
                    "，累计完成{}公里".format(round(scrow.累计实际完成量 / 1000, Global_Digital_Precision)) +
                    "；"
                )
        previous_situation_text = previous_situation_text[:-1] + (
            "。" +
            "上半年立管改造任务倒排工期每天需完成{}公里".format(round((crow.上半年计划工程量 - crow.累计实际完成量) / math.ceil(REMINDER_DAYS * REDUCTION_FACTOR) / 1000, Global_Digital_Precision)) +
            "。\n"
        )
    
    # 后期处理
    previous_situation_text = previous_situation_text[:-1]
    previous_situation_text = previous_situation_text.replace("0.0公里", "0公里")
    previous_situation_text = previous_situation_text.replace("实际完成立管0串，完成0公里", "实际无工程量")
    previous_situation_text = previous_situation_text.replace("完成0公里", "实际无工程量")
    # if INTERVAL_DAYS > 1:
    #     previous_situation_text = previous_situation_text.replace("昨日", "前{}日".format(INTERVAL_DAYS))
    
    # 今日施工人数
    # current_situation_text += \
    # "今日" + \
    # "计划进场施工人数{:d}人".format((df['序号'].count() - COMPLETED_AREA_COUNT - PAUSE_AREA_COUNT) * 12) + \
    # "，实际{:d}人".format(df.at['合计', '施工人数'].sum()) + \
    # "，其中"
    # for row in df[(df['施工队伍'].notnull())].loc['小计'].itertuples():
    #     current_situation_text += "{}{:d}人，".format(
    #         row.施工队伍 if len(crow.施工队伍) >= 4 else row.施工队伍 + '公司',
    #         row.施工人数,
    #     )
    # current_situation_text = current_situation_text[:-1] + "。"

    # 输出文本
    # print(previous_situation_text)
    # print(current_situation_text)
    
    # 写入文件
    with open(RES_TEXT_FILE, 'w', encoding='utf-8') as f:
        f.write(previous_situation_text)
        f.write(current_situation_text)


if __name__ == "__main__":
    # os.chdir(sys.path[0])
    if len(sys.argv) > 1:
        DEFAULT_FILE_PATH = sys.argv[1]
    interval_days = input("请输入间隔天数（默认{}天）：".format(INTERVAL_DAYS))
    if interval_days:
        INTERVAL_DAYS = int(interval_days)
    current_day_data, previous_day_data = load_current_day_data()
    get_format_text(current_day_data, previous_day_data)