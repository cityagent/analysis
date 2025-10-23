from tkinter import messagebox
import pandas as pd
import json
from pathlib import Path
from base_analyzer import BaseAnalyzer


class ConstructionAnalyzer(BaseAnalyzer):
    def __init__(self, original_columns):
        super().__init__(original_columns)
        # 加载外部配置
        config_path = Path(__file__).parent / "config/categories.json"
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        self.target_categories = config["construction_categories"]
        self.category_stats = {}  # 类别统计

    def analyze(self, df, parent=None, **kwargs):
        try:
            self.logs.clear()
            self._log("开始执行施工类项目亏损分析...")

            # 定位必要列
            category_col = self._find_col(df, "项目类别")
            loss_col = self._find_col(df, "亏损金额")
            contract_col = self._find_col(df, "合同金额")
            self._log(f"匹配列：项目类别={category_col}, 亏损金额={loss_col}, 合同金额={contract_col}")

            # 清理项目类别数据
            df_clean = df.copy()
            df_clean["项目类别_清理后"] = df_clean[category_col].astype(str).str.strip()

            # 筛选目标类别数据
            in_target = df_clean["项目类别_清理后"].isin(self.target_categories)
            category_data = df_clean[in_target]
            self._log(f"目标类别总数据量：{len(category_data)} 行")
            if len(category_data) == 0:
                self._log("未找到属于目标类别的数据，分析终止")
                messagebox.showinfo("提示", "未找到属于目标类别的数据", parent=parent)
                return True

            # 转换金额列为数值
            category_data = category_data.copy()
            category_data["亏损金额_数值"] = pd.to_numeric(category_data[loss_col], errors="coerce")
            category_data["合同金额_数值"] = pd.to_numeric(category_data[contract_col], errors="coerce")

            # 筛选有效金额数据（合同金额>0）
            valid_amount_data = category_data[
                category_data["亏损金额_数值"].notna() &
                category_data["合同金额_数值"].notna() &
                (category_data["合同金额_数值"] > 0)
                ]
            invalid_amount = len(category_data) - len(valid_amount_data)
            self._log(
                f"目标类别中金额有效数据：{len(valid_amount_data)} 行\n"
                f"排除无效金额数据：{invalid_amount} 行"
            )

            # 核心筛选：亏损金额/合同金额 > 30%
            filtered_data = valid_amount_data[
                (valid_amount_data["亏损金额_数值"] / valid_amount_data["合同金额_数值"]) > 0.3
                ]
            self.analyzed_data = filtered_data.drop(
                columns=["项目类别_清理后", "亏损金额_数值", "合同金额_数值"]
            )[self.original_columns]

            # 统计各目标类别的符合条件数量
            self.category_stats = filtered_data["项目类别_清理后"].value_counts().to_dict()
            self._log(f"各目标类别符合条件数量：{self.category_stats}")
            self._log(f"最终符合条件数据：{len(self.analyzed_data)} 行")
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

    def get_category_stats(self):
        return self.category_stats