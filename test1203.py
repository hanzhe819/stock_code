import os
import pandas as pd
from datetime import datetime

ret_list_1 = []
ret_list_2 = []
ret_list_3 = []
ret_list_4 = []
ret_list_5 = []
ret_list_6 = []
ret_dic_1 = {
        "日期":"",
        "金额(亿)":"",
        "流通市值(亿)":"",
        "换手%":"",
        "3日换手%":"",
        "涨幅%":"",
        "3日涨幅%":"",
        "6日涨幅%":"",
        "振幅%":"",
        "最新_股价":"",
        "内外比":"",
        "金额(亿)":"",
        "a":"",
        "a":"",
        }
ret_dic_2 = {
        "日期":"",
        "3日涨幅%":"",
        "涨幅%":"",
        "换手%":"",
        "振幅%":"",
        "最新_股价":"",
        "内外比":"",
        "金额(亿)":"",
        "3日换手%":"",
        "流通市值(亿)":"",
        "a":"",
        "a":"",
        }
ret_dic_3 = {
        "日期":"",
        "3日涨幅%":"",
        "涨幅%":"",
        "换手%":"",
        "振幅%":"",
        "最新_股价":"",
        "内外比":"",
        "金额(亿)":"",
        "3日换手%":"",
        "流通市值(亿)":"",
        "a":"",
        "a":"",
        }
ret_dic_4 = {
        "日期":"",
        "3日涨幅%":"",
        "涨幅%":"",
        "换手%":"",
        "振幅%":"",
        "最新_股价":"",
        "内外比":"",
        "金额(亿)":"",
        "3日换手%":"",
        "流通市值(亿)":"",
        "a":"",
        "a":"",
        }
ret_dic_5 = {
        "日期":"",
        "-100, -10":"",
        "-10, -7":"",
        "-7, -5":"",
        "-5, -3":"",
        "-3, 0":"",
        "0, 3":"",
        "3, 5":"",
        "5, 7":"",
        "7, 10":"",
        "10, 100":"",
        "-100, 0":"",
        "0, 100":"",
        "+/- rate":"",
        "a":"",
        }
ret_dic_6 = {
        "日期":"",
        "3 0, 10":"",
        "3 10, 20":"",
        "3 20, 30":"",
        "3 30, 99":"",
        "6 0, 10":"",
        "6 10, 20":"",
        "6 20, 30":"",
        "6 30, 99":"",
        "a":"",
        }
STERT_DATE = '20231203'
#CURRENT_DATE = datetime.now().strftime("%Y%m%d")


def f(sort_by, line_number, average_column, sort_by_filter=31, special_=False):
    # 读取 Excel 文件
    excel_file_path = './%s.xlsx' % CURRENT_DATE
    df = pd.read_excel(excel_file_path)
    
    # 定义一个函数来将亿转化为阿拉伯数字
    def convert_to_arabic(value):
        # 假设输入的格式是 x亿
        if '万亿' in value:
            return float(value.replace('万亿', '')) * 10000
        elif '亿' in value:
            return float(value.replace('亿', ''))
        elif '万' in value:
            return float(value.replace('万', '')) / 10000
        else:
            return float(value)  # 如果不包含亿，直接转换为浮点数
    # 包含亿, 万亿, 万的列进行转换
    num_list_column = ["流通市值", "金额",]
    if average_column in num_list_column:
        df[average_column] = df[average_column].apply(convert_to_arabic)

    print(df.columns)
    if special_ == False:
        # 按照  涨幅 列进行排序 降序添加参数
        df_sorted = df.sort_values(by=sort_by, ascending=False)
    else:
        # 按照  涨幅 列进行排序 默认升序, 不需要添加参数
        df_sorted = df.sort_values(by=sort_by)


    # 剔除 B 列值超过 0.3 的数据
    if sort_by == "涨幅%" and special_ == False:
        df_filtered = df_sorted[df_sorted[sort_by] <= int(sort_by_filter)]
    elif sort_by == "涨幅%" and special_ == True:
        df_filtered = df_sorted[df_sorted[sort_by] >= int(sort_by_filter) * -1]
    elif sort_by == "金额":
        # 则不进行过滤, 将排序后的变量赋值给过滤变量
        df_filtered = df_sorted
    
    # 获取排序后的前 100 行
    top_100_rows = df_filtered.head(int(line_number))

    # 计算 AH 列 (3 日涨幅))的平均值
    average_ah_column = top_100_rows[average_column].mean()

    # 打印结果
#    print("排序后的前 %s 行：" % line_number)
#    print(top_100_rows)
    
#    print("\nAH 列(3 日涨幅 %)的平均值：", average_ah_column)
    return average_ah_column

def f1():
    ret = f("金额", 100, "3日涨幅%")
    ret_dic_1["3日涨幅%"] = ret
    ret = f("金额", 100, "6日涨幅%")
    ret_dic_1["6日涨幅%"] = ret
    ret = f("金额", 100, "涨幅%")
    ret_dic_1["涨幅%"] = ret
    ret = f("金额", 100, "换手%")
    ret_dic_1["换手%"] = ret
    ret = f("金额", 100, "振幅%")
    ret_dic_1["振幅%"] = ret
    ret = f("金额", 100, "振幅%")
    ret_dic_1["振幅%"] = ret
    ret = f("金额", 100, "最新")
    ret_dic_1["最新_股价"] = ret
    ret = f("金额", 100, "内外比")
    ret_dic_1["内外比"] = ret
    ret = f("金额", 100, "金额")
    ret_dic_1["金额(亿)"] = ret
    ret = f("金额", 100, "3日换手%")
    ret_dic_1["3日换手%"] = ret
    ret = f("金额", 100, "流通市值")
    ret_dic_1["流通市值(亿)"] = ret
    
    ret_dic_1["日期"] = CURRENT_DATE
    print(ret)
    ret_list_1.append(ret_dic_1)
    print(ret_list_1)

    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output1.xlsx')):
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_1)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output1.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
    # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output1.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_1])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)


def f2():
    ret = f("涨幅%", 100, "3日涨幅%")
    ret_dic_2["3日涨幅%"] = ret
    ret = f("涨幅%", 100, "涨幅%")
    ret_dic_2["涨幅%"] = ret
    ret = f("涨幅%", 100, "换手%")
    ret_dic_2["换手%"] = ret
    ret = f("涨幅%", 100, "振幅%")
    ret_dic_2["振幅%"] = ret
    ret = f("涨幅%", 100, "振幅%")
    ret_dic_2["振幅%"] = ret
    ret = f("涨幅%", 100, "最新")
    ret_dic_2["最新_股价"] = ret
    ret = f("涨幅%", 100, "内外比")
    ret_dic_2["内外比"] = ret
    ret = f("涨幅%", 100, "金额")
    ret_dic_2["金额(亿)"] = ret
    ret = f("涨幅%", 100, "3日换手%")
    ret_dic_2["3日换手%"] = ret
    ret = f("涨幅%", 100, "流通市值")
    ret_dic_2["流通市值(亿)"] = ret
    ret_dic_2["日期"] = CURRENT_DATE
    print(ret)
    ret_list_2.append(ret_dic_2)
    print(ret_list_2)

    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output2.xlsx')):
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_2)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output2.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output2.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_2])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)


def f3():
    ret = f("涨幅%", 200, "3日涨幅%")
    ret_dic_3["3日涨幅%"] = ret
    ret = f("涨幅%", 200, "涨幅%")
    ret_dic_3["涨幅%"] = ret
    ret = f("涨幅%", 200, "换手%")
    ret_dic_3["换手%"] = ret
    ret = f("涨幅%", 200, "振幅%")
    ret_dic_3["振幅%"] = ret
    ret = f("涨幅%", 200, "振幅%")
    ret_dic_3["振幅%"] = ret
    ret = f("涨幅%", 200, "最新")
    ret_dic_3["最新_股价"] = ret
    ret = f("涨幅%", 200, "内外比")
    ret_dic_3["内外比"] = ret
    ret = f("涨幅%", 200, "金额")
    ret_dic_3["金额(亿)"] = ret
    ret = f("涨幅%", 200, "3日换手%")
    ret_dic_3["3日换手%"] = ret
    ret = f("涨幅%", 200, "流通市值")
    ret_dic_3["流通市值(亿)"] = ret
    ret_dic_3["日期"] = CURRENT_DATE
    print(ret)
    ret_list_3.append(ret_dic_3)
    print(ret_list_3)

    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output3.xlsx')):
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_3)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output3.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output3.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_3])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)

def f4():
    ret = f("涨幅%", 100, "3日涨幅%", special_=True)
    ret_dic_4["3日涨幅%"] = ret
    ret = f("涨幅%", 100, "涨幅%", special_=True)
    ret_dic_4["涨幅%"] = ret
    ret = f("涨幅%", 100, "换手%", special_=True)
    ret_dic_4["换手%"] = ret
    ret = f("涨幅%", 100, "振幅%", special_=True)
    ret_dic_4["振幅%"] = ret
    ret = f("涨幅%", 100, "振幅%", special_=True)
    ret_dic_4["振幅%"] = ret
    ret = f("涨幅%", 100, "最新", special_=True)
    ret_dic_4["最新_股价"] = ret
    ret = f("涨幅%", 100, "内外比", special_=True)
    ret_dic_4["内外比"] = ret
    ret = f("涨幅%", 100, "金额", special_=True)
    ret_dic_4["金额(亿)"] = ret
    ret = f("涨幅%", 100, "3日换手%", special_=True)
    ret_dic_4["3日换手%"] = ret
    ret = f("涨幅%", 100, "流通市值", special_=True)
    ret_dic_4["流通市值(亿)"] = ret
    ret_dic_4["日期"] = CURRENT_DATE
    print(ret)
    ret_list_4.append(ret_dic_4)
    print(ret_list_4)

    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output4.xlsx')):
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_4)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output4.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output4.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_4])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)

def f5():
    '''统计各个涨幅范围内的股票数量'''
    excel_file_path = '%s.xlsx' % CURRENT_DATE
    df = pd.read_excel(excel_file_path)
    # 需要新增删掉非法数据的方法
    pass

    count_minus_100_to_10 = df[(df['涨幅%'] >= -100) & (df['涨幅%'] < -10)].shape[0]
    count_minus_10_to_7 = df[(df['涨幅%'] >= -10) & (df['涨幅%'] < -7)].shape[0]
    count_minus_7_to_5 = df[(df['涨幅%'] >= -7) & (df['涨幅%'] < -5)].shape[0]
    count_minus_5_to_3 = df[(df['涨幅%'] >= -5) & (df['涨幅%'] < -3)].shape[0]
    count_minus_3_to_0 = df[(df['涨幅%'] >= -3) & (df['涨幅%'] < 0)].shape[0]
    count_0_to_3 = df[(df['涨幅%'] >= 0) & (df['涨幅%'] <= 3)].shape[0]
    count_3_to_5 = df[(df['涨幅%'] >= 3) & (df['涨幅%'] <= 5)].shape[0]
    count_5_to_7 = df[(df['涨幅%'] >= 5) & (df['涨幅%'] <= 7)].shape[0]
    count_7_to_10 = df[(df['涨幅%'] >= 7) & (df['涨幅%'] <= 10)].shape[0]
    count_10_to_100 = df[(df['涨幅%'] >= 10) & (df['涨幅%'] <= 100)].shape[0]

    # 统计全部涨与跌的数量并计算涨跌比
    count_minus_100_to_0 = df[(df['涨幅%'] >= -100) & (df['涨幅%'] < 0)].shape[0]
    count_0_to_100 = df[(df['涨幅%'] >= 0) & (df['涨幅%'] < 100)].shape[0]

    ret_dic_5["日期"] = CURRENT_DATE
    ret_dic_5["-100, -10"] = count_minus_100_to_10
    ret_dic_5["-10, -7"] = count_minus_10_to_7
    ret_dic_5["-7, -5"] = count_minus_7_to_5
    ret_dic_5["-5, -3"] = count_minus_5_to_3
    ret_dic_5["-3, 0"] = count_minus_3_to_0
    ret_dic_5["0, 3"] = count_0_to_3
    ret_dic_5["3, 5"] = count_3_to_5
    ret_dic_5["5, 7"] = count_5_to_7
    ret_dic_5["7, 10"] = count_7_to_10
    ret_dic_5["10, 100"] = count_10_to_100
    ret_dic_5["-100, 0"] = count_minus_100_to_0
    ret_dic_5["0, 100"] = count_0_to_100
    ret_dic_5["+/- rate"] = round(int(count_0_to_100) / int(count_minus_100_to_0), 2)

    print(ret_dic_5)
    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output5.xlsx')):
        ret_list_5.append(ret_dic_5)
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_5)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output5.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output5.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_5])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)

    
def f6():
    '''统计各个3日 6日涨幅范围内的股票数量'''
    excel_file_path = '%s.xlsx' % CURRENT_DATE
    df = pd.read_excel(excel_file_path)

    count_3_0_to_10 = df[(df['3日涨幅%'] >= 0) & (df['3日涨幅%'] <= 10)].shape[0]
    count_3_10_to_20 = df[(df['3日涨幅%'] >= 10) & (df['3日涨幅%'] <= 20)].shape[0]
    count_3_20_to_30 = df[(df['3日涨幅%'] >= 20) & (df['3日涨幅%'] <= 30)].shape[0]
    count_3_30_to_99 = df[(df['3日涨幅%'] >= 30)].shape[0]
    count_6_0_to_10 = df[(df['6日涨幅%'] >= 0) & (df['6日涨幅%'] <= 10)].shape[0]
    count_6_10_to_20 = df[(df['6日涨幅%'] >= 10) & (df['6日涨幅%'] <= 20)].shape[0]
    count_6_20_to_30 = df[(df['6日涨幅%'] >= 20) & (df['6日涨幅%'] <= 30)].shape[0]
    count_6_30_to_99 = df[(df['6日涨幅%'] >= 30)].shape[0]

    ret_dic_6["日期"] = CURRENT_DATE
    ret_dic_6["3 0, 10"] = count_3_0_to_10
    ret_dic_6["3 10, 20"] = count_3_10_to_20
    ret_dic_6["3 20, 30"] = count_3_20_to_30
    ret_dic_6["3 30, 99"] = count_3_30_to_99
    ret_dic_6["6 0, 10"] = count_6_0_to_10
    ret_dic_6["6 10, 20"] = count_6_10_to_20
    ret_dic_6["6 20, 30"] = count_6_20_to_30
    ret_dic_6["6 30, 99"] = count_6_30_to_99

    print(ret_dic_6)
    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output6.xlsx')):
        ret_list_6.append(ret_dic_6)
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_6)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output6.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output6.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_6])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)


def f7():
    '''统计各个胡神经全部的涨幅范围内的股票数量'''
    excel_file_path = '%s.xlsx' % CURRENT_DATE
    df = pd.read_excel(excel_file_path)
    # 将中文替换掉,

    count_3_0_to_10 = df[(df['3日涨幅%'] >= 0) & (df['3日涨幅%'] <= 10)].shape[0]
    count_3_10_to_20 = df[(df['3日涨幅%'] >= 10) & (df['3日涨幅%'] <= 20)].shape[0]
    count_3_20_to_30 = df[(df['3日涨幅%'] >= 20) & (df['3日涨幅%'] <= 30)].shape[0]
    count_3_30_to_99 = df[(df['3日涨幅%'] >= 30)].shape[0]
    count_6_0_to_10 = df[(df['6日涨幅%'] >= 0) & (df['6日涨幅%'] <= 10)].shape[0]
    count_6_10_to_20 = df[(df['6日涨幅%'] >= 10) & (df['6日涨幅%'] <= 20)].shape[0]
    count_6_20_to_30 = df[(df['6日涨幅%'] >= 20) & (df['6日涨幅%'] <= 30)].shape[0]
    count_6_30_to_99 = df[(df['6日涨幅%'] >= 30)].shape[0]

    ret_dic_6["日期"] = CURRENT_DATE
    ret_dic_6["3 0, 10"] = count_3_0_to_10
    ret_dic_6["3 10, 20"] = count_3_10_to_20
    ret_dic_6["3 20, 30"] = count_3_20_to_30
    ret_dic_6["3 30, 99"] = count_3_30_to_99
    ret_dic_6["6 0, 10"] = count_6_0_to_10
    ret_dic_6["6 10, 20"] = count_6_10_to_20
    ret_dic_6["6 20, 30"] = count_6_20_to_30
    ret_dic_6["6 30, 99"] = count_6_30_to_99

    print(ret_dic_6)
    # 如果是第一次跑, 当前路径下没有产出Excel, 则生成一次
    if not os.path.exists(os.path.join(os.getcwd(), 'output6.xlsx')):
        ret_list_6.append(ret_dic_6)
        # 使用 pandas 创建 DataFrame
        df = pd.DataFrame(ret_list_6)
        # 将 DataFrame 写入 Excel 文件
        excel_file_path = 'output6.xlsx'
        df.to_excel(excel_file_path, index=False)
    else:
        # 将 DataFrame 更新到 Excel 文件
        excel_file_path = 'output6.xlsx'
        df = pd.read_excel(excel_file_path)
        df = pd.concat([df, pd.DataFrame([ret_dic_6])], ignore_index=True)
        df.to_excel(excel_file_path, index=False)

    
if __name__ == '__main__':
    l = ['20231206', '20231207', '20231208']
    for i in l:
        CURRENT_DATE = str(i)
        f1()
        f2()
        f3()
        f4()
        f5()
        f6()
