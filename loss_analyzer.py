from tkinter import messagebox

import pandas as pd

from base_analyzer import BaseAnalyzer


class LossDataAnalyzer(BaseAnalyzer):
    def __init__(self, original_columns):
        super().__init__(original_columns)
        self.valid_rows_count = 0  # 有效比较行数

    def analyze(self, df, parent=None, **kwargs):
        try:
            self.logs.clear()
            self._log("开始执行成本构成异常分析...")

            # 定位必要列
            loss_col = self._find_col(df, "亏损金额")
            settlement_col = self._find_col(df, "项目结算金额")
            contract_col = self._find_col(df, "合同金额")
            lwf_col = self._find_col(df, "项目主要成本情况_劳务费_结算")
            clf_col = self._find_col(df, "项目主要成本情况_材料费_结算")
            jxf_col = self._find_col(df, "项目主要成本情况_设备机械租赁费_结算")
            zxf_col = self._find_col(df, "项目主要成本情况_技术服务、咨询费_结算")
            fbf_col = self._find_col(df, "项目主要成本情况_专业分包_结算")

            # 转换为数值类型
            df_copy = df.copy()
            df_copy["亏损金额_数值"] = pd.to_numeric(df_copy[loss_col], errors="coerce")
            df_copy["合同金额_数值"] = pd.to_numeric(df_copy[contract_col], errors="coerce")
            df_copy["项目结算金额_数值"] = pd.to_numeric(df_copy[settlement_col], errors="coerce")
            df_copy["劳务费_数值"] = pd.to_numeric(df_copy[lwf_col], errors="coerce")
            df_copy["材料费_数值"] = pd.to_numeric(df_copy[clf_col], errors="coerce")
            df_copy["设备费_数值"] = pd.to_numeric(df_copy[jxf_col], errors="coerce")
            df_copy["技术费_数值"] = pd.to_numeric(df_copy[zxf_col], errors="coerce")
            df_copy["分包费_数值"] = pd.to_numeric(df_copy[fbf_col], errors="coerce")

            # 筛选有效行
            valid_rows = df_copy[
                (df_copy["亏损金额_数值"].notna() & df_copy["项目结算金额_数值"].notna()) |
                (df_copy["亏损金额_数值"].notna() & df_copy["合同金额_数值"].notna()) |
                (df_copy["劳务费_数值"].notna() & df_copy["合同金额_数值"].notna()) |
                (df_copy["材料费_数值"].notna() & df_copy["合同金额_数值"].notna()) |
                (df_copy["设备费_数值"].notna() & df_copy["合同金额_数值"].notna()) |
                (df_copy["技术费_数值"].notna() & df_copy["合同金额_数值"].notna()) |
                (df_copy["分包费_数值"].notna() & df_copy["合同金额_数值"].notna())
                ]
            invalid_rows = len(df_copy) - len(valid_rows)
            self.valid_rows_count = len(valid_rows)
            self._log(
                f"过滤无效行：{invalid_rows} 行\n"
                f"有效比较行：{self.valid_rows_count} 行"
            )

            # 核心筛选条件
            filtered_df = valid_rows[
                (valid_rows["亏损金额_数值"] >= valid_rows["项目结算金额_数值"]) |
                (valid_rows["亏损金额_数值"] >= valid_rows["合同金额_数值"]) |
                (valid_rows["劳务费_数值"] / valid_rows["合同金额_数值"] >= 0.5) |
                (valid_rows["材料费_数值"] / valid_rows["合同金额_数值"] >= 0.5) |
                (valid_rows["设备费_数值"] / valid_rows["合同金额_数值"] >= 0.5) |
                (valid_rows["技术费_数值"] / valid_rows["合同金额_数值"] >= 0.5) |
                (valid_rows["分包费_数值"] / valid_rows["合同金额_数值"] >= 0.5)
                ]

            # 整理结果
            self.analyzed_data = filtered_df.drop(
                columns=["亏损金额_数值", "合同金额_数值", "项目结算金额_数值",
                         "劳务费_数值", "材料费_数值", "设备费_数值",
                         "技术费_数值", "分包费_数值"]
            )[self.original_columns]
            self._log(f"符合成本异常条件的数据：{len(self.analyzed_data)} 行")
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
        return self.valid_rows_count
