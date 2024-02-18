import pandas as pd
import datetime


class Logger:
    def __init__(self, log_file, columns):
        self.log_file = log_file
        self.columns = columns
        self.log_df = pd.DataFrame(columns=columns)

    def log(self, data):
        if len(data) != len(self.columns) - 1:
            raise ValueError("Data length does not match the number of columns.")
        timestamp = datetime.datetime.now()
        log_data = {'Update': timestamp, **data}
        self.log_df.loc[len(self.log_df)] = log_data

    def save_log_to_csv(self):
        self.log_df.to_csv(self.log_file, index=False, encoding='utf_8_sig')

# # 创建Logger对象，传递列名作为参数
# logger = Logger("log.csv", ["Timestamp", "Message", "Value"])
#
# # 记录日志
# logger.log(["This is a log message", 42])
# logger.log(["Another log message", 123])
#
# # 保存日志到CSV文件
# logger.save_log_to_csv()
