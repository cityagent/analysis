class BaseAnalyzer:
    def __init__(self, original_columns):
        self.original_columns = original_columns  # 原始列名（保持结果顺序）
        self.analyzed_data = None  # 分析结果数据
        self.logs = []  # 分析过程日志

    def _find_col(self, df, col_name):
        """通用列查找方法（适配多级表头）"""
        target_clean = str(col_name).strip().replace(" ", "").replace("　", "")
        matches = []
        for col in df.columns:
            col_clean = str(col).strip().replace(" ", "").replace("　", "")
            if col_clean == target_clean or col_clean.endswith(f"_{target_clean}"):
                matches.append(col)
        if not matches:
            raise ValueError(f"未找到「{col_name}」列（候选列示例：{df.columns[:5]}）")
        if len(matches) > 1:
            raise ValueError(f"找到多个「{col_name}」列：{matches}，请确认唯一列")
        return matches[0]

    def _log(self, msg):
        """内部日志记录"""
        self.logs.append(msg)

    def analyze(self, df, parent=None, **kwargs):
        """子类必须实现的分析方法"""
        raise NotImplementedError("子类需实现analyze方法")

    def get_analyzed_data(self):
        return self.analyzed_data

    def get_logs(self):
        return self.logs
