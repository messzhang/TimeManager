# storage_excel.py
import os
import pandas as pd

class ExcelLogger:
    def __init__(self, out_dir=".", filename_template="工作记录_{year}.xlsx"):
        """
        out_dir: 存放 Excel 的目录
        filename_template: 文件名模板，{year} 会被替换为年份
        """
        self.out_dir = out_dir
        self.template = filename_template

    def _path_for_date(self, date_str: str) -> str:
        # date_str 格式 "YYYY-MM-DD"
        year = date_str[:4]
        fname = self.template.format(year=year)
        return os.path.join(self.out_dir, fname)

    def log(self,
            date: str,
            t_start: str,
            t_end: str,
            work: str,
            effect: str,
            remark: str):
        """
        date: "2025-05-07"
        t_start: "09:00"
        t_end:   "12:30"
        work:    工作内容
        effect:  效果
        remark:  备注
        """
        path = self._path_for_date(date)

        # 如果不存在，则先创建带表头的空表
        if not os.path.exists(path):
            df0 = pd.DataFrame(columns=[
                "日期", "时间段", "工作内容", "效果", "备注"
            ])
            df0.to_excel(path, index=False)

        # 读取当年所有记录
        df = pd.read_excel(path, parse_dates=["日期"])

        # 构造新行的 DataFrame
        new_row = pd.DataFrame([{
            "日期": pd.to_datetime(date),
            "时间段": f"{t_start}-{t_end}",
            "工作内容": work,
            "效果": effect,
            "备注": remark
        }])

        # 使用 concat 追加
        df = pd.concat([df, new_row], ignore_index=True)

        # 排序：先按日期，再按 时间段
        df.sort_values(by=["日期", "时间段"], inplace=True)

        # 写回（覆盖原文件）
        df.to_excel(path, index=False)
