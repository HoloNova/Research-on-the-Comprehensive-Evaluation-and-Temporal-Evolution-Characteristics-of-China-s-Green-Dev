import os
import pandas as pd
import numpy as np

print("调试熵权法代码")
print("="*80)

file_path = r'.\国民消费水平\居民消费水平数据_已处理_20260415_080630.xlsx'

if os.path.exists(file_path):
    print(f"读取文件: {file_path}")
    df = pd.read_excel(file_path)
    print(f"\n数据形状: {df.shape}")
    print(f"\n列名: {list(df.columns)}")
    print(f"\n前5行数据:")
    print(df.head())
    
    year_cols = [col for col in df.columns if '年' in col and '标准化' not in col]
    print(f"\n年份列: {year_cols}")
    
    if len(year_cols) >= 2:
        data_matrix = df[year_cols].values
        print(f"\n数据矩阵形状: {data_matrix.shape}")
        print(f"\n数据矩阵:")
        print(data_matrix)
        
        print(f"\n检查NaN:")
        print(np.isnan(data_matrix).sum())
        
        data_clean = data_matrix[~np.isnan(data_matrix).any(axis=1)]
        print(f"\n清理后形状: {data_clean.shape}")
        print(f"\n清理后数据:")
        print(data_clean)
        
        indicators_clean = df['指标'].values[~np.isnan(data_matrix).any(axis=1)]
        print(f"\n清理后指标: {indicators_clean}")
        
        if len(indicators_clean) >= 2 and data_clean.shape[1] >= 2:
            print(f"\n开始计算熵权...")
            
            data = np.array(data_clean, dtype=np.float64)
            m, n = data.shape
            
            print(f"m={m}, n={n}")
            
            k = 1 / np.log(m) if m > 1 else 1
            print(f"k={k}")
            
            P = data / data.sum(axis=0)
            print(f"\nP:")
            print(P)
            
            e = -k * (P * np.log(P + 1e-10)).sum(axis=0)
            print(f"\ne: {e}")
            
            d = 1 - e
            print(f"d: {d}")
            
            w = d / d.sum() if d.sum() != 0 else np.ones(n) / n
            print(f"w: {w}")
            
            print(f"\n指标权重:")
            for ind, weight in zip(indicators_clean, w):
                print(f"  {ind}: {weight:.6f}")
else:
    print(f"文件不存在: {file_path}")
