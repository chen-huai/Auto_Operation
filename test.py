import pandas as pd
from pandas.tseries.offsets import BDay
from datetime import datetime, timedelta

# 读取人员数据
file_name = "人员.xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(file_name, sheet_name=sheet_name)

# 初始化日历和假期信息
start_date = datetime(2025, 4, 1)
end_date = datetime(2025, 12, 31)
holidays = [
    datetime(2025, 1, 1),
    *pd.date_range(start='2025-01-29', end='2025-02-04'),
    *pd.date_range(start='2025-04-04', end='2025-04-06'),
    *pd.date_range(start='2025-05-01', end='2025-05-03'),
    *pd.date_range(start='2025-06-07', end='2025-06-09'),
    *pd.date_range(start='2025-10-01', end='2025-10-07')
]

# 生成工作日期
work_dates = pd.date_range(start=start_date, end=end_date, freq='B')

# 初始化任务安排表
schedule_rows = []

# 分配任务
daily_task_count = {name: 0 for name in df["人员名称"]}
after_holiday_task_count = {name: 0 for name in df["人员名称"]}
supplement_task_count = {name: 0 for name in df["人员名称"]}

def assign_members(exclude_list, count, task_count):
    candidates = [name for name in df["人员名称"] if name not in exclude_list and "怀孕" not in name and "离岗" not in name]
    candidates.sort(key=lambda x: task_count[x])
    return candidates[:count]

# 检查日期是否是节假日或周末后的工作日
def is_after_holiday_or_weekend(date, holidays):
    if date.weekday() == 0:  # Monday
        return (date - timedelta(days=2)) not in work_dates or (date - timedelta(days=1)) in holidays
    return (date - timedelta(days=1)) in holidays

for date in work_dates:
    # 分配日常组成员，确保公平性
    daily_members = assign_members([], 5, daily_task_count)
    for member in daily_members:
        daily_task_count[member] += 1

    # 检查是否为节假日或周末后的工作日，并记录
    if is_after_holiday_or_weekend(date, holidays):
        for member in daily_members:
            after_holiday_task_count[member] += 1

    # 分配增补组成员，确保公平性
    supplement_members = []
    if is_after_holiday_or_weekend(date, holidays):
        supplement_members = assign_members(daily_members, 2, supplement_task_count)
        for member in supplement_members:
            supplement_task_count[member] += 1

    # 合并所有成员
    all_members = daily_members + supplement_members

    # 确保同一天内没有同一个人被分配多个任务
    if len(set(all_members)) != len(all_members):
        continue

    # 尽量避免同一个组的成员在同一天内同时分配任务，但不强制
    assigned_groups = df[df['人员名称'].isin(all_members)]['小组'].tolist()
    # If having members from the same group, consider using a relaxed constraint
    # if len(assigned_groups) != len(set(assigned_groups)):
    #     continue

    schedule_rows.append({
        "日期": date,
        "日常组成员": "; ".join(daily_members),
        "增补组成员": "; ".join(supplement_members),
    })

# 更新人员表的次数
df["日常次数"] = df["人员名称"].map(daily_task_count)
df["增补次数"] = df["人员名称"].map(supplement_task_count)
df["节假日后一天"] = df["人员名称"].map(after_holiday_task_count)

# 导出任务安排表
schedule = pd.DataFrame(schedule_rows)
schedule.to_excel("排班表.xlsx", index=False)

# 导出更新后的人员表
df.to_excel("更新后的人员.xlsx", index=False)
