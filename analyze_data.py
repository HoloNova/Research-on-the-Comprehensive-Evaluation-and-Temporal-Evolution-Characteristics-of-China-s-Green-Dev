"""
查看数据文件结构
"""

import pandas as pd
import os

data_file = r'国民消费水平\国民经济总量数据.xlsx'

if os.path.exists(data_file):
    df = pd.read_excel(data_file)
    print("="*80)
    print("数据文件结构")
    print("="*80)
    print(f"\n形状: {df.shape}")
    print(f"\n列名: {df.columns.tolist()}")
    print("\n前10行数据:")
    print(df.head(10))
    print("\n数据类型:")
    print(df.dtypes)
    print("\n统计信息:")
    print(df.describe())
else:
    print(f"文件不存在: {data_file}")
