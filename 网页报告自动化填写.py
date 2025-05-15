# 根据用户的要求，我们将从输入的文本中提取特定的键值对，并创建一个新的Excel文件，通过巧妙复制粘贴可以解决繁琐的一些工作
# Re研究报告自动化点击工具
# author: yuchenqiang
#使用步骤：1.点开dow系统报告生成第一界面 2.复制属于自己的那一块内容 3.系统会自动生成一份excel并自动打开 4.对于该excel 按win+←左，分屏向左，选择浏览器分屏向右边，第一界面停留8s如果有误可以自己调整参数，随后等待即可
#注意事项：1.wps的路径wps_path需要修改 2.字典、excel、第一界面、第二界面4个函数可以按需调用



import pyautogui
import time
import re


import pandas as pd
import os
import subprocess
import sys

# 输入数据
# input_data = "课题编号：PRO6604，产品名称：氯氮平片，样本量：3790，样本范围：167050-170840，协议编号：TB-II-MDS-2023-01-01-17，企业名称：宁波大红鹰药业股份有限公司，报告日期：2024年12月，研究目的：监测不良反应/探索新适应症/指导特殊人群用药/变更给药途径/研制复方制剂/支持医保用药，报告名称：氯氮平片产品上市后研究报告"


# 定义转换输入数据为字典的函数
def convert_input_to_dict(input_data):
    # 以下适用粘贴：产品名称：喜炎平注射液，样本量：200000，协议编号：OK-B-I-2024--06-28--011，企业名称：江西青峰药业有限公司 ，报告日期：2025年2月，研究目的：监测不良反应/探索新适应症/指导特殊人群用药/变更给药途径/研制复方制剂/支持医保用药,报告名称：喜炎平注射液产品上市后研究报告

    # data_pairs = input_data.split("，")
    # data_dict = {}
    # for pair in data_pairs:
    #     if "：" in pair:
    #         key, value = pair.split("：", 1)
    #         data_dict[key.strip()] = value.strip()

    # 以下适用粘贴：四川维奥制药有限公司        OK-B-Ⅱ-2024-03-01-021        醋氯芬酸肠溶片        3367        2024年11月        于xx
    # 定义固定的字典标题
    params = {'研究目的':"监测不良反应/探索新适应症/指导特殊人群用药/变更给药途径/研制复方制剂/支持医保用药"}
    headers = ["企业名称", "协议编号", "产品名称", "样本量", "报告日期", "报告出具"]
    # 使用正则表达式分割输入的数据，分割符为4个或更多的空格
    values = re.split(r'\s{4,}', input_data.strip())
    # 使用zip函数将标题和值配对，并转换为字典
    datadict = dict(zip(headers, values))
    # 打印字典
    datadict["报告名称"] = datadict["产品名称"] + "产品上市后研究报告"
    params.update(datadict)

    params['报告日期'] = params['报告日期'].replace('年', '-').replace('月', '')
    # Adding a leading zero to the month if necessary
    month = params['报告日期'].split('-')[1]
    month_with_leading_zero = f"{int(month):02d}"  # Format the month with leading zero
    params['报告日期'] = params['报告日期'].replace(month, month_with_leading_zero)

    # 指定新的键顺序
    new_order = [
        '报告名称',  # 报告名称
        '企业名称',  # 企业名称
        '产品名称',  # 产品名称
        '研究目的',  # 研究目的
        '协议编号',  # 协议编号
        '样本量',  # 样本量
        "报告日期",
        # 如果还有其他键，请继续添加
    ]

    # 创建一个新的字典，按照指定的顺序插入键值对
    ordered_dict = {key: params[key] for key in new_order}

    return ordered_dict

# 定义从字典中选出特定键值对的函数
def select_specific_data(data_dict):
    selected_keys = ["报告名称", "企业名称", "产品名称", "研究目的", "协议编号", "样本量","报告日期",]
    selected_data = {key: data_dict.get(key, '') for key in selected_keys}
    return selected_data

# 定义创建Excel文件并填入数据的函数
def create_excel(data_dict):
    excel_file_path = 'h1.xlsx'
    df = pd.DataFrame(list(data_dict.items()), columns=["Key", "Value"])
    df = df.set_index("Key")
    df.to_excel(excel_file_path)
    wps_path = r"C:\Users\yuchenqiang\AppData\Local\Kingsoft\WPS Office\ksolaunch.exe"
    # 设定当前文件夹的路径
    current_folder = os.getcwd()
    file_to_open = os.path.join(current_folder, excel_file_path)
    # 运行wps.exe打开指定的Excel文件
    subprocess.run([wps_path, file_to_open])
    time.sleep(5)

    pyautogui.hotkey('win', 'left')
    return df

def automate_click():
    time.sleep(4)

    pyautogui.click(1800,400)
    pyautogui.scroll(-2000)
    #起始样本
    pyautogui.click(1664,436)


    pyautogui.click(1400,830)

    time.sleep(8)
    pyautogui.click(1400,920)
    pyautogui.press('0')

    #样本量
    pyautogui.click(126,327)
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(1616,925)
    pyautogui.hotkey('ctrl', 'v')



    pyautogui.click(1855,1000)

# 垂直向下滚动鼠标1000个单位
def automate_click_and_copy_paste():


    time.sleep(4)


    pyautogui.click(1800,400)

    # #向上平移5000确保下一个界面
    pyautogui.scroll(2000)
    pyautogui.scroll(-108)


    # 定义一个列表，包含所有自定义的坐标
    click_positions = [
        (138, 235),  # 第1次点击的坐标
        (138, 255),
        (138, 275),
        (138, 290),
        (138, 305),
        (138, 325),
        (138, 346),
    ]

    # 定义一个列表，包含所有自定义的粘贴坐标
    paste_positions = [
        (1331, 330),  # 第1次粘贴的坐标
        (1331, 415),  # 第2次粘贴的坐标
        (1331, 500),  # 第3次粘贴的坐标
        (1331, 600),  # 第4次粘贴的坐标
        (1331, 715),  # 第5次粘贴的坐标
        (1331, 800),  # 第5次粘贴的坐标
        (1331, 888)  # 第5次粘贴的坐标

    ]

    # 确保两个列表的长度相同
    assert len(click_positions) == len(paste_positions), "点击坐标和粘贴坐标的数量必须相同"


    # 循环次数与坐标列表长度相同
    for click_pos, paste_pos in zip(click_positions, paste_positions):
        # 移动鼠标到指定位置并点击宁波大红鹰药业股份有限公司
        pyautogui.click(click_pos)

        # 使用ctrl+c复制文本
        pyautogui.hotkey('ctrl', 'c')

        # 稍作延迟，确保复制操作完成
        time.sleep(0.1)

        # 移动鼠标到粘贴位置并点击
        pyautogui.click(paste_pos)

        # 使用ctrl+v粘贴文本
        pyautogui.hotkey('ctrl', 'v')

        # 如果需要，在此处添加额外的延迟
        time.sleep(0.1)  # 例如，每次操作后暂停1秒

    #确认生成
    pyautogui.click(1831,1002)





if __name__ == "__main__":
    '''
    四川维奥制药有限公司        OK-B-Ⅱ-2024-03-01-021        米格列醇片        18000        2024年6月        于郴强
    '''
    input_data = input("输入老板任务:")
    data_dict = convert_input_to_dict(input_data)
    print(data_dict)
    selected_data = select_specific_data(data_dict)
    excel_data = create_excel(selected_data)
    print(f"h1成功生成，请打开")
    # time.sleep(5)df2
    print("10s打开工作文档(左)和dow系统(右)，屏幕比例100%")
    # time.sleep(5)
    #第一个界面点击
    automate_click()
    # 第二个界面点击fruits.csv
    automate_click_and_copy_paste()


