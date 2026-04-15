"""
我国绿色发展水平综合评价与时间演化特征研究
课题分析脚本
"""

import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
import seaborn as sns
sns.set_style('whitegrid')

# 读取分析结果
print("="*80)
print("步骤1：读取分析结果数据")
print("="*80)

df_result = pd.read_excel('analysis_results/分析结果.xlsx')
print(f"数据形状: {df_result.shape}")
print("\n数据列名:")
print(df_result.columns.tolist())
print("\n前5行数据:")
print(df_result.head())

# 检查是否有原始数据文件
print("\n" + "="*80)
print("检查原始数据文件")
print("="*80)

if os.path.exists('sample_data.xlsx'):
    df_raw = pd.read_excel('sample_data.xlsx')
    print(f"原始数据形状: {df_raw.shape}")
    print("\n原始数据列名:")
    print(df_raw.columns.tolist())
    print("\n原始数据前5行:")
    print(df_raw.head())
else:
    print("未找到原始数据文件")

# 生成简要统计
print("\n" + "="*80)
print("关键指标统计")
print("="*80)

key_metrics = ['综合得分', '经济子系统得分', '环境子系统得分', '耦合协调度']
for metric in key_metrics:
    if metric in df_result.columns:
        print(f"\n{metric}:")
        print(f"  平均值: {df_result[metric].mean():.4f}")
        print(f"  最大值: {df_result[metric].max():.4f}")
        print(f"  最小值: {df_result[metric].min():.4f}")
        print(f"  标准差: {df_result[metric].std():.4f}")

if '协调等级' in df_result.columns:
    print("\n协调等级分布:")
    print(df_result['协调等级'].value_counts())

print("\n" + "="*80)
print("数据检查完成")
print("="*80)
