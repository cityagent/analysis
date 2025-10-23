from tkinter import messagebox
from base_analyzer import BaseAnalyzer

class LeaderFrequencyAnalyzer(BaseAnalyzer):
    def __init__(self, original_columns):
        super().__init__(original_columns)
        self.leader_stats = None  # 负责人出现次数统计

    def analyze(self, df, min_count=3, parent=None,** kwargs):
        try:
            self.logs.clear()
            self._log("开始执行项目负责人频次分析...")

            # 定位项目负责人列
            leader_col = self._find_col(df, "项目负责人")
            self._log(f"匹配项目负责人列：{leader_col}")

            # 清理负责人名称
            df_clean = df.copy()
            df_clean[leader_col] = df_clean[leader_col].astype(str).str.strip()
            df_clean = df_clean[
                (df_clean[leader_col] != "") &
                (df_clean[leader_col].str.lower() != "nan")
            ]
            clean_rows = len(df_clean)
            self._log(f"清理后有效数据：{clean_rows} 行（排除空值/无效负责人）")

            # 统计负责人出现次数
            self.leader_stats = df_clean[leader_col].value_counts().to_dict()
            qualified_leaders = [
                leader for leader, count in self.leader_stats.items()
                if count >= min_count
            ]
            self._log(
                f"负责人总数：{len(self.leader_stats)} 人\n"
                f"出现≥{min_count}次的负责人：{len(qualified_leaders)} 人"
            )
            if qualified_leaders:
                self._log(
                    f"高频负责人列表（次数）：\n" +
                    "\n".join([f"- {l}: {self.leader_stats[l]}次" for l in qualified_leaders])
                )
            else:
                self._log(f"无出现≥{min_count}次的负责人")

            # 提取高频负责人的所有项目数据
            self.analyzed_data = df_clean[df_clean[leader_col].isin(qualified_leaders)][self.original_columns]
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

    def get_leader_stats(self):
        return self.leader_stats