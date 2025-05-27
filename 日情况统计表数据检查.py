from LoadExcel import *

import os, sys, traceback


def check_data(current_day_data: pd.DataFrame, previous_day_data: pd.DataFrame) -> None:
    """
    检查当前日期数据的准确性
    """
    checked_items = ["累计打眼数量", "累计立管串数", "累计置换串数", "累计实际完成量", "累计PMS系统录入量"]
    checked_results = dict()
    
    for crow in current_day_data.iterrows():
        crow = crow[1]
        row_name = crow.name
        manager  = crow['管理单位']
        construction_team = crow['施工队伍']
        
        # 获取前一日对应数据行
        try:
            prow = previous_day_data.loc[row_name]
            if not isinstance(prow, pd.Series):
                if construction_team:
                    prow = prow.query("施工队伍 == @construction_team").iloc[0]
                    row_name += f"_{construction_team}"
                elif manager:
                    prow = prow.query("管理单位 == @manager").iloc[0]
                    row_name += f"_{manager}"
            # prow = prow
        except KeyError:
            prow = pd.Series(0, index=crow.index)
            prow.name = row_name
        
        checked_results[row_name] = []
        for item in checked_items:
            # 检查累计完成量
            if (prow[item] != 0 and crow[item] == 0) or (crow[item.replace('累计', '当日')] != 0 and crow[item] == 0):
                checked_results[row_name].append(f"{item}为0，累计数据错误")
                
            # 检查当日完成量
            cday_volume = (crow[item] - prow[item]).round(2)
            item = item.replace('累计', '当日')
            if cday_volume < 0:
                checked_results[row_name].append(f"{item}小于0，累计数据错误")
            elif cday_volume == 0 and crow[item] != 0:
                checked_results[row_name].append(f"{item}为0，与表中数据不符")
            elif cday_volume != crow[item]:
                checked_results[row_name].append(f"{item}与表中数据不符")
                
        if row_name[:2] not in ("小计", "合计", "总计"):
            # 检查数据逻辑错误
            if crow["当日立管串数"] != 0 and crow["当日实际完成量"] == 0:
                checked_results[row_name].append("有立管但当日完成量为0")
            if crow["累计立管串数"] != 0 and crow["累计打眼数量"] == 0:
                checked_results[row_name].append("有立管但累计打眼数为0")
            if crow["累计置换串数"] != 0 and crow["累计立管串数"] == 0:
                checked_results[row_name].append("有置换串但累计立管为0")
            if crow["当日打眼数量"] == 0 and crow["当日立管串数"] == 0 and crow["当日实际完成量"] == 0 and crow["当日置换串数"] == 0 and crow["施工人数"] != 0:
                checked_results[row_name].append("无工作量但施工人数不为0")
            if (crow["当日打眼数量"] != 0 or crow["当日立管串数"] != 0 or crow["当日实际完成量"] != 0 or crow["当日置换串数"] != 0) and crow["施工人数"] == 0:
                checked_results[row_name].append("有工作量但施工人数为0")
            
            # 检查施工状态
            if (crow["施工人数"] == 0 and crow["施工状态"] not in ("停工", "待置换", "完工")) or \
            (crow["施工人数"] != 0 and crow["施工状态"] in ("在施")):
                checked_results[row_name].append("施工状态错误")
        
    # 输出检查结果
    for name, results in checked_results.items():
        if results:
            results = f"\033[4m{'；'.join(results)}\033[0m".replace("；", "\033[0m；\033[4m")
            print("\033[31m{}\033[0m: {}".format(name, results))

if __name__ == "__main__":
    os.chdir(sys.path[0])
    if len(sys.argv) > 1:
        DEFAULT_FILE_PATH = sys.argv[1]
    # interval_days = input("请输入间隔天数（默认{}天）：".format(INTERVAL_DAYS))
    # if interval_days:
    #     INTERVAL_DAYS = int(interval_days)
        
    # 获取需要生成文本的日期
    try:
        days = input("请输入需要检查数据的日期（默认今天{}）：".format(TODAY_DATE.strftime(r"%#m.%#d")))
        if not days:
            days = [TODAY_DATE.strftime(r"%#m月%#d日")]
        else:
            days = days.split(' ')
        days = [datetime.strptime(d, r"%m.%d").strftime(r"%#m月%#d日") for d in days]
    except:
        days = [TODAY_DATE.strftime(r"%#m月%#d日")]
    
    for day in days:
        print(day)
        try:
            current_day_data, previous_day_data, _ = load_specific_day_data(datetime.strptime(day, r"%m月%d日"))
            check_data(current_day_data, previous_day_data)
        except ValueError:
            print("没有找到{}工作表".format(day))
            # traceback.print_exc()
        except Exception:
            traceback.print_exc()
            print("检查{}数据时发生错误".format(day))            
        print("=" * 100)
