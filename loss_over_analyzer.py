import pandas as pd
from tkinter import messagebox
from base_analyzer import BaseAnalyzer


class LossOverAnalyzer(BaseAnalyzer):
    def __init__(self, original_columns):
        super().__init__(original_columns)
        self.total_valid_rows = 0  # 有效亏损金额行数

    def analyze(self, df, threshold=1000, parent=None, **kwargs):
        try:
            self.logs.clear()
            self._log(f"开始执行亏损金额> {threshold} 的数据分析...")

            # 定位亏损金额列
            loss_col = self._find_col(df, "亏损金额")
            self._log(f"匹配亏损金额列：{loss_col}")

            # 转换为数值类型
            df_copy = df.copy()
            df_copy["亏损金额_数值"] = pd.to_numeric(df_copy[loss_col], errors="coerce")
            self._log("已将亏损金额列转换为数值类型（非数值转为空值）")

            # 筛选有效数据
            valid_data = df_copy[df_copy["亏损金额_数值"].notna()]
            self.total_valid_rows = len(valid_data)
            self._log(f"有效亏损金额数据：{self.total_valid_rows} 行")

            # 核心筛选：亏损金额 > 阈值
            filtered_data = valid_data[valid_data["亏损金额_数值"] > threshold]
            self.analyzed_data = filtered_data.drop(columns=["亏损金额_数值"])[self.original_columns]

            # 计算占比
            ratio = round(len(self.analyzed_data) / self.total_valid_rows * 100, 2) if self.total_valid_rows > 0 else 0
            self._log(
                f"符合条件（亏损金额> {threshold}）的数据：{len(self.analyzed_data)} 行\n"
                f"占有效数据比例：{ratio}%"
            )


            return True

        except ValueError as ve:
            err_msg = f"分析失败：{str(ve)}"
            self._log(err_msg)
            messagebox.showerror("分析错误", err_msg, parent=parent)
            return False
        except Exception as e:
            err_msg = f"分析失败：{str(e)}"
            self._log(err_msg)
            messagebox.showerror("分析错误", err_msg, parent=parent)
            return False

    def get_valid_rows_count(self):
        return self.total_valid_rows