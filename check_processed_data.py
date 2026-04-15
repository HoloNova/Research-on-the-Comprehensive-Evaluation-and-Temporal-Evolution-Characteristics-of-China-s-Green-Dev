"""
查看处理后的数据文件
"""

import pandas as pd
import os

processed_file = r'国民消费水平\国民经济总量数据_已处理_20260415_081527.xlsx'

if os.path.exists(processed_file):
    df = pd.read_excel(processed_file)
    print("="*80)
    print("处理后数据文件结构")
    print("="*80)
    print(f"\n形状: {df.shape}")
    print(f"\n列名: {df.columns.tolist()}")
    print("\n前20行数据:")
    print(df.head(20))
    print("\n数据类型:")
    print(df.dtypes)
else:
    print(f"文件不存在: {processed_file}")
