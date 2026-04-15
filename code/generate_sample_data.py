"""
示例数据生成脚本

生成符合标准化流程要求的示例数据
"""

import pandas as pd
import numpy as np
import os

def generate_sample_data(output_file: str = 'sample_data.xlsx'):
    """
    生成示例数据
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
    
    print(f"\n生成 {len(years)} 年 × {len(provinces)} 个省份的示例数据...")
    
    data = []
    np.random.seed(42)
    
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
    
    df.to_excel(output_file, index=False)
    print(f"\n示例数据已保存到: {output_file}")
    print(f"数据形状: {df.shape}")
    
    print("\n" + "="*80)
    print("数据预览")
    print("="*80)
    print(df.head(10))
    
    print("\n" + "="*80)
    print("统计信息")
    print("="*80)
    print(df.describe())
    
    return df


def test_standard_workflow(sample_file: str = 'sample_data.xlsx'):
    """
    测试标准化流程
    """
    print("\n" + "="*80)
    print("测试标准化流程")
    print("="*80)
    
    try:
        from standard_analysis import StandardAnalysisWorkflow
        
        workflow = StandardAnalysisWorkflow()
        results = workflow.run_full_workflow(
            input_file=sample_file,
            output_dir='sample_results'
        )
        
        print("\n" + "="*80)
        print("测试完成！")
        print("="*80)
        print("\n结果文件保存在: sample_results/")
        
        return results
        
    except Exception as e:
        print(f"\n测试过程中出现错误: {str(e)}")
        print("\n注意: 这可能是因为缺少依赖库或数据格式问题")
        return None


def main():
    sample_file = 'sample_data.xlsx'
    
    if not os.path.exists(sample_file):
        df = generate_sample_data(sample_file)
    else:
        print(f"示例数据文件已存在: {sample_file}")
        df = pd.read_excel(sample_file)
        print(f"数据形状: {df.shape}")
    
    run_test = input("\n是否运行标准化流程测试？(y/n): ").strip().lower()
    
    if run_test == 'y':
        test_standard_workflow(sample_file)
    else:
        print("\n跳过测试。要运行测试，请执行:")
        print("python generate_sample_data.py")
        print("或直接使用:")
        print("from standard_analysis import StandardAnalysisWorkflow")
        print("workflow = StandardAnalysisWorkflow()")
        print("results = workflow.run_full_workflow('sample_data.xlsx', 'results')")


if __name__ == '__main__':
    main()
