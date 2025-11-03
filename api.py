import json
import os
import pandas as pd
import numpy as np
from fastapi import FastAPI, File, UploadFile, HTTPException
from io import BytesIO
from openpyxl import load_workbook, Workbook
from fastapi.responses import FileResponse, JSONResponse  # <-- 导入 JSONResponse
from fastapi.encoders import jsonable_encoder
import logging
from datetime import datetime
# 导入所需的分析器类
from leader_analyzer import LeaderFrequencyAnalyzer
from construction_analyzer import ConstructionAnalyzer
from design_analyzer import DesignAnalyzer
from loss_over_analyzer import LossOverAnalyzer
from loss_analyzer import LossDataAnalyzer
from excel_saver import ExcelResultSaver

# 设置日志配置
logging.basicConfig(level=logging.INFO)  # 设置日志级别为 INFO
logger = logging.getLogger(__name__)

app = FastAPI()
def convert_datetime_to_string(obj):
    """Recursively convert datetime objects to string in the format of 'YYYY-MM-DD HH:MM:SS'."""
    if isinstance(obj, datetime):
        return obj.strftime('%Y-%m-%d %H:%M:%S')  # Adjust the format as per your requirement
    elif isinstance(obj, dict):
        return {key: convert_datetime_to_string(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_datetime_to_string(item) for item in obj]
    return obj
class AnalysisAPI:
    def __init__(self):
        self.original_columns = [] 
        self.raw_data = None  
        self.analyzers = []

        self.analyzers_config = [
            {
                "class": LeaderFrequencyAnalyzer,
                "sheet_name": "附表1 亏损3个项目项目负责人",
                "analyze_kwargs": {"min_count": 3}
            },
            {
                "class": DesignAnalyzer,
                "sheet_name": "附表2 亏损大于合同",
                "analyze_kwargs": {}
            },
            {
                "class": ConstructionAnalyzer,
                "sheet_name": "附表3 施工项目亏损金额占合同金额30%",
                "analyze_kwargs": {}
            },
            {
                "class": LossOverAnalyzer,
                "sheet_name": "附表4 亏损大于1000万",
                "analyze_kwargs": {"threshold": 1000}
            },
            {
                "class": LossDataAnalyzer,
                "sheet_name": "附表5 成本费用异常情况",
                "analyze_kwargs": {}
            }
        ]

    def upload_excel(self, file: UploadFile):
        """解析上传的Excel文件并读取数据"""
        try:
            content = file.file.read()
            wb = load_workbook(BytesIO(content), data_only=True)
            sheet = wb.active
            max_col = sheet.max_column
            self.original_columns = []
            for col_idx in range(1, max_col + 1):
                main_header = self.get_merged_value(sheet, row=3, col=col_idx)
                sub1_header = self.get_merged_value(sheet, row=4, col=col_idx)
                sub2_header = self.get_merged_value(sheet, row=5, col=col_idx)

                parts = []
                seen = set()
                for p in [main_header, sub1_header, sub2_header]:
                    if p is not None:
                        p_str = str(p)
                        if p_str not in seen:
                            seen.add(p_str)
                            parts.append(p_str)

                col_name = "_".join(parts) if parts else f"未知列_{col_idx}"
                self.original_columns.append(col_name)

            data = []
            for row in sheet.iter_rows(min_row=6, values_only=True):
                row_data = row[:len(self.original_columns)]
                if row_data and row_data[0] is None:
                    break
                data.append(row_data)

            self.raw_data = pd.DataFrame(data, columns=self.original_columns)

            self.analyzers = [
                cfg["class"](original_columns=self.original_columns)
                for cfg in self.analyzers_config
            ]
            logger.info(f"Excel 文件上传并读取成功，共 {len(self.raw_data)} 行数据")
            return True
        except Exception as e:
            self.raw_data = None 
            logger.error(f"Excel 读取失败：{str(e)}")
            raise HTTPException(status_code=400, detail=f"Excel 读取失败：{str(e)}")

    def run_analysis(self):
        """执行所有分析器"""
        if self.raw_data is None or not self.analyzers:
            raise HTTPException(status_code=400, detail="请先上传Excel文件！")

        logger.info("开始执行数据分析...")
        results = []
        for i, cfg in enumerate(self.analyzers_config):
            analyzer = self.analyzers[i]
            logger.info(f"开始执行分析器：{analyzer.__class__.__name__}")
            success = analyzer.analyze(df=self.raw_data, **cfg["analyze_kwargs"])
            
            if success:
                analyzed_df = analyzer.get_analyzed_data()

                # 清理特殊的 float 值（NaN, Infinity 等）
                cleaned_df = analyzed_df.applymap(self.replace_invalid_floats)

                # Convert datetime objects to string
                cleaned_data = cleaned_df.to_dict(orient="records")
                cleaned_data = convert_datetime_to_string(cleaned_data)  # Apply the conversion

                results.append({
                    "analyzer": analyzer.__class__.__name__,
                    "sheet_name": cfg["sheet_name"],
                    "status": "success",
                    "data": cleaned_data  # Now, this data is JSON serializable
                })
                logger.info(f"{analyzer.__class__.__name__} 分析完成，包含 {len(cleaned_df)} 条数据")
            else:
                results.append({
                    "analyzer": analyzer.__class__.__name__,
                    "sheet_name": cfg["sheet_name"],
                    "status": "failed",
                    "message": f"{analyzer.__class__.__name__} 分析失败或无结果"
                })
                logger.error(f"{analyzer.__class__.__name__} 分析失败或无结果")

        logger.info("所有分析器执行完毕")
        return results

    def replace_invalid_floats(self, value):
        """替换无效的 float 值（NaN, Infinity）"""
        if isinstance(value, float):
            if pd.isna(value) or value == float('inf') or value == float('-inf'):
                return ""  # 或者返回一个替代字符串 "NaN"
        return value

    def save_results_to_excel(self, results: list):
        """将分析结果保存为新的Excel文件"""
        if not results:
            raise HTTPException(status_code=400, detail="没有结果需要保存！")

        logger.info("开始保存分析结果到 Excel 文件...")
        excel_saver = ExcelResultSaver()

        for result in results:
            if result['status'] == 'success' and result['data']:
                try:
                    data_df = pd.DataFrame.from_dict(result['data'])
                    sheet_name = result.get('sheet_name', result['analyzer'])
                    
                    excel_saver.save_dataframe_to_sheet(df=data_df, sheet_name=sheet_name)
                    
                except Exception as e:
                    logger.error(f"将 {result['analyzer']} 结果保存到 Excel 失败: {e}")

        file_name = "分析报告.xlsx"
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        file_path = os.path.join(output_dir, file_name)
        
        excel_saver.save(file_path)
        logger.info(f"分析结果已保存到：{file_path}")
        return file_path

    def get_merged_value(self, sheet, row, col):
        """获取指定行列的单元格值（处理合并单元格）"""
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
        return sheet.cell(row=row, column=col).value


analysis_api = AnalysisAPI()    

@app.post("/upload_and_analyze_json/", tags=["一站式API"])
async def upload_and_analyze_json(file: UploadFile = File(..., description="要分析的项目数据 Excel 文件")):
    """
    【一站式】上传 Excel 文件，立即执行所有分析，并返回 JSON 格式的结果。
    """
    try:
        logger.info("开始上传并分析 Excel 文件...")
        analysis_api.upload_excel(file)
        results = analysis_api.run_analysis()
        logger.info("文件上传并分析成功")
        logger.info(results)
        return JSONResponse(content=results)
    except HTTPException as e:
        logger.error(f"HTTP 错误：{e.detail}")
        raise e
    except Exception as e:
        logger.error(f"发生未知错误：{str(e)}")
        if "Out of range float values" in str(e):
            raise HTTPException(status_code=500, detail="分析结果包含非标准的浮动数 (NaN/Inf)，请检查数据清洗。")
        raise HTTPException(status_code=500, detail=f"处理请求时发生未知错误: {str(e)}")


@app.post("/upload_and_download_excel/", tags=["一站式API"])
async def upload_and_download_excel(file: UploadFile = File(..., description="要分析的项目数据 Excel 文件")):
    """
    【一站式】上传 Excel 文件，执行分析，并将结果保存为 Excel 文件并直接返回下载。
    """
    try:
        logger.info("开始上传并分析 Excel 文件...")
        analysis_api.upload_excel(file)
        results = analysis_api.run_analysis()
        file_path = analysis_api.save_results_to_excel(results)
        logger.info(f"分析结果已保存到：{file_path}")
        return FileResponse(
            file_path, 
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename="分析报告.xlsx"
        )
    except HTTPException as e:
        logger.error(f"HTTP 错误：{e.detail}")
        raise e
    except Exception as e:
        logger.error(f"发生未知错误：{str(e)}")
        raise HTTPException(status_code=500, detail=f"处理请求时发生未知错误: {str(e)}")


# 运行FastAPI服务器
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8004)
