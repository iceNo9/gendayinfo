import pyautogui
import pyperclip
import time
from datetime import datetime, timedelta

# 设置起始日期和序号
start_date = "2025/02/01"
start_sequence = 4415

# 转换起始日期为 datetime 对象
start_date = datetime.strptime(start_date, "%Y/%m/%d")

# 获取本月的最后一天
# 计算本月最后一天，方法是下个月的第一天减去一天
next_month = start_date.replace(day=28) + timedelta(days=4)  # 先跳到下个月
last_day_of_month = next_month - timedelta(days=next_month.day)

# 打印最后一天（用于调试）
print(f"本月最后一天是: {last_day_of_month.strftime('%Y/%m/%d')}")

# 星期几的中文映射
weekday_map = {
    'Mon': '周一',
    'Tue': '周二',
    'Wed': '周三',
    'Thu': '周四',
    'Fri': '周五',
    'Sat': '周六',
    'Sun': '周日'
}

# 定义一个函数，用于切换到 OneNote 软件
def switch_to_onenote():
    # 使用 Alt + Tab 切换到 OneNote
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.5)  # 等待切换完成

# 定义一个函数，用于输入日期和序号
def input_date_and_sequence(current_date, sequence):
    # 格式化日期为 "月日 周几" 格式，例如 "01-01 周三"
    date_str = current_date.strftime("%m-%d")
    weekday_str = current_date.strftime("%a")  # 获取星期几的英文简写
    # 使用中文星期几
    formatted_date = f"{date_str} {weekday_map[weekday_str]} #{sequence}"

    # 输出当前日期和序号
    print(f"输入: {formatted_date}")

    # 将文本复制到剪贴板
    pyperclip.copy(formatted_date)

    # 按下 Ctrl + N 打开新的输入框
    pyautogui.hotkey('ctrl', 'n')
    time.sleep(0.5)  # 等待输入框激活

    # 粘贴文本（包含中文）
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)  # 等待粘贴完成

# 初次输入从起始日期开始
current_date = start_date
sequence = start_sequence

while True:
    # 执行时切换到 OneNote
    print("切换到 OneNote...")
    switch_to_onenote()

    # 循环日期直到本月最后一天
    while current_date <= last_day_of_month:
        input_date_and_sequence(current_date, sequence)
        
        # 增加日期和序号
        current_date += timedelta(days=1)
        sequence += 1

    # 等待用户按下回车键来继续
    print("\n本月日期和序号已完成。按回车键继续输入下个月的数据。")
    input()  # 等待用户按下回车键

    # 重新计算下个月的最后一天
    start_date = current_date
    next_month = start_date.replace(day=28) + timedelta(days=4)  # 先跳到下个月
    last_day_of_month = next_month - timedelta(days=next_month.day)

    # 打印下个月最后一天
    print(f"下个月最后一天是: {last_day_of_month.strftime('%Y/%m/%d')}")
