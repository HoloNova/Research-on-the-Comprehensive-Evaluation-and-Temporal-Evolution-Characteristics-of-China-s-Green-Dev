"""
处理现有数据并运行标准化分析流程
"""

import pandas as pd
import numpy as np
import os
import sys
sys.path.insert(0, '.')

from code.standard_analysis import StandardAnalysis

def prepare_national_data():
    """
    准备国民经济总量数据 - 转换为标准化格式
    """
    print("="*80)
    print("准备国民经济总量数据")
    print("="*80)
    
    processed_file = r'国民消费水平\国民经济总量数据_已处理_20260415_081527.xlsx'
    
    if not os.path.exists(processed_file):
        print("处理后文件不存在，使用原始文件")
        return None
    
    df = pd.read_excel(processed_file)
    
    print(f"\n原始数据形状: {df.shape}")
    print(f"原始列名: {df.columns.tolist()}")
    
    df_long = df.melt(
        id_vars=['指标'],
        value_vars=[col for col in df.columns if '年' in col and '标准化' not in col],
        var_name='年份',
        value_name='数值'
    )
    
    df_long['年份'] = df_long['年份'].str.extract('(\d+)').astype(int)
    
    df_pivot = df_long.pivot(index='年份', columns='指标', values='数值').reset_index()
    
    df_pivot['省份'] = '全国'
    
    print(f"\n转换后数据形状: {df_pivot.shape}")
    print(f"转换后列名: {df_pivot.columns.tolist()}")
    print("\n转换后数据预览:")
    print(df_pivot.head())
    
    output_file = 'prepared_national_data.xlsx'
    df_pivot.to_excel(output_file, index=False)
    print(f"\n准备好的数据已保存到: {output_file}")
    
    return output_file, df_pivot

def main():
    print("="*80)
    print("现有数据处理与标准化分析")
    print("="*80)
    
    data_file, df_data = prepare_national_data()
    
    if data_file is None:
        print("\n无法找到合适的数据文件")
        return
    
    print("\n" + "="*80)
    print("运行标准化分析流程")
    print("="*80)
    
    analysis = StandardAnalysis()
    
    numeric_cols = [col for col in df_data.columns if col not in ['年份', '省份']]
    
    results = analysis.run(
        input_file=data_file,
        output_dir='analysis_results',
        positive_cols=numeric_cols,
        negative_cols=[]
    )
    
    print("\n" + "="*80)
    print("分析完成！")
    print("="*80)
    print("\n结果文件位置: analysis_results/")

if __name__ == '__main__':
    main()
