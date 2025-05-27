from LoadExcel import *

import math
import os, sys, traceback

RES_TEXT_PATH = os.path.join(os.environ['USERPROFILE'], 'Downloads')
COMPLETED_AREA_COUNT = 38     # 完工小区数量
PAUSE_AREA_COUNT = 0     # 停工小区数量
Global_Digital_Precision = 2     # 全局数字精度
Global_Percentage_Precision = Global_Digital_Precision + 2     # 全局百分比精度

FHY_END_DATE = date(2025, 6, 30)     # 2025年立管改造任务上半年结束日期
REMINDER_DAYS = (FHY_END_DATE - TODAY_DATE).days + 1     # 2025年立管改造任务剩余天数
REDUCTION_FACTOR = 0.85     # 2025年立管改造任务时间缩减系数


def get_format_text(cdate_df:pd.DataFrame, pdate_df:pd.DataFrame, y24data_df:pd.DataFrame, date:str):
    global TODAY_DATE
    global INTERVAL_DAYS
    previous_situation_text = str()
    current_situation_text = str()

    # 进度概况
    # 昨日施工总体情况
    crow = cdate_df.loc['总计']
    prow = pdate_df.loc['总计']
    previous_situation_text += (
        "一分公司" +
        "{}".format(date) + 
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
        "民心工程" +
        "{}".format(date) +
        "完成{}公里，".format(round((cdate_df.loc["总计"].民心工程累计完成量 - pdate_df.loc["总计"].民心工程累计完成量) / 1000, Global_Digital_Precision)) +
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
        "距上半年立管改造任务截止时间（{}）还剩{:d}天".format(FHY_END_DATE.strftime(r"%Y年%#m月%#d日"), REMINDER_DAYS) +
        "，按{:d}天计算倒排工期".format(math.ceil(REMINDER_DAYS * REDUCTION_FACTOR)) +
        "，每天需完成{}公里".format(round((crow.上半年计划工程量 - crow.累计实际完成量) / math.ceil(REMINDER_DAYS * REDUCTION_FACTOR) / 1000, Global_Digital_Precision)) +
        "。\n"
    )
    for crow in cdate_df.loc['合计'].itertuples():
        prow = pdate_df.loc['合计'].query('管理单位 == @crow.管理单位').iloc[0]
        previous_situation_text += (
            "{}区域".format(crow.管理单位) + 
            "{}".format(date) +
            "计划完成{}公里".format(round((crow.上半年计划工程量 - prow.累计实际完成量) / math.ceil((REMINDER_DAYS + INTERVAL_DAYS) * REDUCTION_FACTOR) / 1000 * INTERVAL_DAYS, Global_Digital_Precision)) +
            # "，实际完成立管{:d}串".format(crow.累计立管串数 - prow.累计立管串数) +
            "，实际完成{}公里".format(round((crow.累计实际完成量 - prow.累计实际完成量) / 1000, Global_Digital_Precision)) +
            # "，PMS系统内录入工程量{}公里".format(round((crow.累计PMS系统录入量 - prow.累计PMS系统录入量) / 1000, Global_Digital_Precision)) +
            # "，累计完成立管{:d}串".format(crow.累计立管串数) +
            "，累计完成{}公里".format(round(crow.累计实际完成量 / 1000, Global_Digital_Precision)) +
            # "，PMS系统内累计录入工程量{}公里".format(round(crow.累计PMS系统录入量 / 1000, Global_Digital_Precision)) +
            "。" +
            "其中，" +
            "民心工程" +
            "{}".format(date) +
            "完成{}公里，".format(round((crow.民心工程累计完成量 - prow.民心工程累计完成量) / 1000, Global_Digital_Precision)) +
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
            "{}".format(date) +
            "上岗{}人".format(crow.施工人数)
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
    previous_situation_text += "\n"
    previous_situation_text = previous_situation_text.replace("0.0公里", "0公里")
    previous_situation_text = previous_situation_text.replace("实际完成立管0串，完成0公里", "实际无工程量")
    previous_situation_text = previous_situation_text.replace("完成0公里", "实际无工程量")
    
    # 置换情况
    y24data_df = y24data_df[(y24data_df['累计置换串数'] != 0)]
    cdate_df = cdate_df[(cdate_df['累计置换串数'] != 0)]
    replacement_situation_text = ""
    
    if not y24data_df.empty:
        replacement_situation_text += (
            "24年改造小区{}".format(date) +
            "置换{}串".format(y24data_df.loc['合计'].当日置换串数.sum()) +
            "，累计置换{}（+178={}）串。".format(y24data_df.loc['合计'].累计置换串数.sum() - 178, y24data_df.loc['合计'].累计置换串数.sum())
        )
        
        if not y24data_df.loc['合计'].当日置换串数.sum() == 0:
            replacement_situation_text += "其中，"
            for row in y24data_df[(y24data_df['监理单位'].notnull())].itertuples():
                if not row.当日置换串数:
                    continue
                replacement_situation_text += (
                    "{}".format(row.Index) +
                    "置换{}串".format(row.当日置换串数) +
                    "，"
                )
        replacement_situation_text = replacement_situation_text[:-1] + "。\n"
    
    if not cdate_df.empty:
        replacement_situation_text += (
            "25年改造小区{}".format(date) +
            "置换{}串".format(cdate_df.loc['总计'].当日置换串数.sum()) +
            "，累计置换{}串。".format(cdate_df.loc['总计'].累计置换串数.sum())
        )
        
        if not cdate_df.loc['总计'].当日置换串数.sum() == 0:
            replacement_situation_text += "其中，"
            for row in cdate_df[(cdate_df['监理单位'].notnull())].itertuples():
                if not row.当日置换串数:
                    continue
                replacement_situation_text += (
                    "{}".format(row.Index) +
                    "置换{}串".format(row.当日置换串数) +
                    "，"
                )
        replacement_situation_text = replacement_situation_text[:-1] + "。\n"
    
    # 输出文本
    # print(previous_situation_text)
    # print(current_situation_text)
    # print(replacement_situation_text)
    
    # 写入文件
    date = date.replace("月", ".").replace("日", "")
    with open(os.path.join(RES_TEXT_PATH, '{}汇报材料文本.txt'.format(date)), 'w', encoding='utf-8') as f:
        f.write(previous_situation_text)
        f.write(current_situation_text)
        f.write(replacement_situation_text)


if __name__ == "__main__":
    os.chdir(sys.path[0])
    if len(sys.argv) > 1:
        DEFAULT_FILE_PATH = sys.argv[1]
    # interval_days = input("请输入间隔天数（默认{}天）：".format(INTERVAL_DAYS))
    # if interval_days:
    #     INTERVAL_DAYS = int(interval_days)
        
    # 获取需要生成文本的日期
    try:
        days = input("请输入需要生成文本的日期（默认今天{}）：".format(TODAY_DATE.strftime(r"%#m.%#d")))
        if not days:
            days = [TODAY_DATE.strftime(r"%#m月%#d日")]
        else:
            days = days.split(' ')
        days = [datetime.strptime(d, r"%m.%d").strftime(r"%#m月%#d日") for d in days]
    except:
        days = [TODAY_DATE.strftime(r"%#m月%#d日")]
    # 获取工期缩减系数
    try:
        reduction_factor = input("请输入工期缩减系数（默认{}）：".format(REDUCTION_FACTOR))
        reduction_factor = float(reduction_factor)
        if reduction_factor > 0 and reduction_factor < 1:
            REDUCTION_FACTOR = reduction_factor
    except:
        pass
    
    for day in days:
        print(day)
        try:
            current_day_data, previous_day_data, y24_day_data = load_specific_day_data(datetime.strptime(day, r"%m月%d日"))
            get_format_text(current_day_data, previous_day_data, y24_day_data, day)
        except ValueError:
            print("没有找到{}工作表".format(day))
            continue
        except Exception:
            traceback.print_exc()
            continue