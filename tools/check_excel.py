#!/usr/bin/env python3
"""
检查Excel文件内容
"""

import os
import pandas as pd
import glob

def 检查最新excel文件(输出目录: str):
    """查找并读取最新的Excel文件"""
    # 查找所有Excel文件
    excel文件列表 = glob.glob(os.path.join(输出目录, "cdn链接_*.xlsx"))

    if not excel文件列表:
        print("未找到Excel文件")
        return

    # 按修改时间排序，获取最新文件
    最新文件 = max(excel文件列表, key=os.path.getmtime)
    print(f"最新Excel文件: {最新文件}")
    print(f"文件大小: {os.path.getsize(最新文件)} 字节")
    print(f"修改时间: {os.path.getmtime(最新文件)}")

    # 读取Excel文件
    try:
        df = pd.read_excel(最新文件, engine='openpyxl')
        print(f"\n数据形状: {df.shape}")
        print(f"列名: {list(df.columns)}")

        print("\n前5行数据:")
        print(df.head())

        print("\nCDN链接示例:")
        for idx, row in df.head(3).iterrows():
            print(f"{row['文件名']} -> {row['CDN 链接']}")

    except Exception as e:
        print(f"读取Excel文件时出错: {e}")

if __name__ == "__main__":
    输出目录 = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
    检查最新excel文件(输出目录)