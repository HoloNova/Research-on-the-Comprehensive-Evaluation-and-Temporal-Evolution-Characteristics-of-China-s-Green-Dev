"""
生成示例数据并运行完整分析流程
"""

import pandas as pd
import numpy as np
import os

from standard_analysis import StandardAnalysis

def generate_sample_data():
    """
    生成符合标准化格式的示例数据
    """
    print("="*80)
    print("生成示例数据")
    print("="*80)
    
    provinces = [
        '北京', '天津', '河北', '山西', '内蒙古',
        '辽宁', '吉林', '黑龙江', '上海', '江苏',
        '浙江', '安徽', '福建', '江西', '山东',
        '河南', '湖北', '湖南', '广东', '广西',
        '海南', '重庆', '四川', '贵州', '云南',
        '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆'
    ]
    
    years = [2018, 2019, 2020, 2021, 2022]
    
    np.random.seed(42)
    
    data = []
    for year in years:
        year_factor = (year - 2018) * 0.05
        
        for idx, province in enumerate(provinces):
            province_factor = idx / len(provinces)
            
            data.append({
                '年份': year,
                '省份': province,
                '人均GDP': 50000 + province_factor * 100000 + year_factor * 20000 + np.random.normal(0, 15000),
                '第三产业占比': 40 + province_factor * 30 + year_factor * 5 + np.random.normal(0, 8),
                'R&D占比': 1.0 + province_factor * 3.0 + year_factor * 0.5 + np.random.normal(0, 0.5),
                '单位GDP能耗': 1.2 - province_factor * 0.6 - year_factor * 0.1 + np.random.normal(0, 0.2),
                '单位GDP水耗': 60 - province_factor * 30 - year_factor * 5 + np.random.normal(0, 10),
                '污水处理率': 75 + province_factor * 20 + year_factor * 3 + np.random.normal(0, 5),
                '垃圾处理率': 80 + province_factor * 15 + year_factor * 2 + np.random.normal(0, 4),
                '森林覆盖率': 20 + province_factor * 30 + np.random.normal(0, 10),
                '优良天数比例': 60 + province_factor * 30 + year_factor * 2 + np.random.normal(0, 8),
                '环保投资占比': 1.2 + province_factor * 1.5 + year_factor * 0.2 + np.random.normal(0, 0.4)
            })
    
    df = pd.DataFrame(data)
    
    numeric_cols = [
        '人均GDP', '第三产业占比', 'R&D占比', '单位GDP能耗', 
        '单位GDP水耗', '污水处理率', '垃圾处理率', 
        '森林覆盖率', '优良天数比例', '环保投资占比'
    ]
    
    for col in numeric_cols:
        df[col] = df[col].clip(lower=0)
    
    output_file = 'sample_data.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"\n示例数据已生成: {output_file}")
    print(f"数据形状: {df.shape}")
    print("\n数据预览:")
    print(df.head(10))
    
    return output_file, df

def main():
    print("="*80)
    print("数据处理与分析 - 完整流程")
    print("="*80)
    
    sample_file, df_sample = generate_sample_data()
    
    print("\n" + "="*80)
    print("运行标准化分析流程")
    print("="*80)
    
    analysis = StandardAnalysis()
    
    positive_cols = [
        '人均GDP', '第三产业占比', 'R&D占比',
        '污水处理率', '垃圾处理率', '森林覆盖率',
        '优良天数比例', '环保投资占比'
    ]
    
    negative_cols = [
        '单位GDP能耗', '单位GDP水耗'
    ]
    
    results = analysis.run(
        input_file=sample_file,
        output_dir='analysis_results',
        positive_cols=positive_cols,
        negative_cols=negative_cols
    )
    
    print("\n" + "="*80)
    print("分析完成！")
    print("="*80)
    print("\n结果文件位置:")
    print("- 分析结果: analysis_results/分析结果.xlsx")
    print("- 分析报告: analysis_results/分析报告.md")
    print("- 可视化图表: analysis_results/visualizations/")

if __name__ == '__main__':
    main()
