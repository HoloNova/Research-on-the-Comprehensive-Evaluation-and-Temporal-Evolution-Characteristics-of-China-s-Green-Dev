import os
import pandas as pd
import numpy as np
import json
from datetime import datetime

def check_file_integrity():
    print("="*80)
    print("步骤1: 文件完整性检查")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    check_results = {
        'original_files': 0,
        'processed_xlsx': 0,
        'processed_csv': 0,
        'reports': 0,
        'errors': []
    }
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            files = os.listdir(folder_path)
            original = [f for f in files if f.endswith('.xlsx') and '已处理' not in f]
            processed_xlsx = [f for f in files if '已处理' in f and f.endswith('.xlsx')]
            processed_csv = [f for f in files if '已处理' in f and f.endswith('.csv')]
            
            check_results['original_files'] += len(original)
            check_results['processed_xlsx'] += len(processed_xlsx)
            check_results['processed_csv'] += len(processed_csv)
            
            print(f"\n{folder}:")
            print(f"  原始文件: {len(original)}个")
            print(f"  处理后Excel: {len(processed_xlsx)}个")
            print(f"  处理后CSV: {len(processed_csv)}个")
            
            for f in processed_xlsx:
                try:
                    df = pd.read_excel(os.path.join(folder_path, f))
                    if len(df) == 0 or len(df.columns) == 0:
                        check_results['errors'].append(f"{folder}/{f}: 空文件")
                except Exception as e:
                    check_results['errors'].append(f"{folder}/{f}: {str(e)}")
            for f in processed_csv:
                try:
                    df = pd.read_csv(os.path.join(folder_path, f), encoding='utf-8-sig')
                    if len(df) == 0 or len(df.columns) == 0:
                        check_results['errors'].append(f"{folder}/{f}: 空文件")
                except Exception as e:
                    check_results['errors'].append(f"{folder}/{f}: {str(e)}")
    
    report_files = ['数据字典.json', '熵权法权重分析报告.json', '熵权法权重分析报告.md', '使用说明.md', '项目整理总结.md']
    for rf in report_files:
        if os.path.exists(rf):
            check_results['reports'] += 1
            try:
                if rf.endswith('.json'):
                    with open(rf, 'r', encoding='utf-8') as f:
                        json.load(f)
            except Exception as e:
                check_results['errors'].append(f"{rf}: {str(e)}")
    
    print(f"\n检查完成！")
    print(f"原始文件总数: {check_results['original_files']}")
    print(f"处理后Excel总数: {check_results['processed_xlsx']}")
    print(f"处理后CSV总数: {check_results['processed_csv']}")
    print(f"报告文件总数: {check_results['reports']}")
    
    if check_results['errors']:
        print(f"\n发现 {len(check_results['errors'])} 个错误:")
        for err in check_results['errors'][:10]:
            print(f"  - {err}")
    else:
        print("\n所有文件检查通过！")
    
    return check_results

def fix_entropy_report():
    print("\n" + "="*80)
    print("步骤2: 修复熵权法报告")
    print("="*80)
    
    report_path = '熵权法权重分析报告.json'
    if os.path.exists(report_path):
        with open(report_path, 'r', encoding='utf-8') as f:
            report = json.load(f)
        
        fixed_count = 0
        for file_name, result in report.get('results', {}).items():
            weight_dict = result.get('weight_dict', {})
            for key, value in weight_dict.items():
                if value != value or value is None:
                    weight_dict[key] = 0.0
                    fixed_count += 1
        
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        
        print(f"已修复 {fixed_count} 个NaN值")

def calculate_entropy_weight(data):
    data = np.array(data, dtype=np.float64)
    m, n = data.shape
    
    k = 1 / np.log(m) if m > 1 else 1
    
    P = data / (data.sum(axis=0) + 1e-10)
    
    e = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    
    d = 1 - e
    
    w = d / (d.sum() + 1e-10) if d.sum() != 0 else np.ones(n) / n
    
    return w

def topsis_analysis():
    print("\n" + "="*80)
    print("步骤3: TOPSIS综合评价分析")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    all_results = {}
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            processed_files = [f for f in os.listdir(folder_path) 
                             if '已处理' in f and f.endswith('.xlsx')]
            
            for pf in processed_files:
                file_path = os.path.join(folder_path, pf)
                
                try:
                    df = pd.read_excel(file_path)
                    year_cols = [col for col in df.columns 
                                 if '年' in col and '标准化' not in col]
                    
                    if len(year_cols) >= 2:
                        data_matrix = df[year_cols].values
                        
                        data_clean = data_matrix[~np.isnan(data_matrix).any(axis=1)]
                        indicators_clean = df['指标'].values[~np.isnan(data_matrix).any(axis=1)]
                        
                        if len(indicators_clean) >= 2 and data_clean.shape[1] >= 2:
                            weights = calculate_entropy_weight(data_clean)
                            
                            normalized_data = data_clean / np.sqrt((data_clean ** 2).sum(axis=0))
                            
                            weighted_normalized = normalized_data * weights
                            
                            ideal_best = np.max(weighted_normalized, axis=0)
                            ideal_worst = np.min(weighted_normalized, axis=0)
                            
                            distance_best = np.sqrt(((weighted_normalized - ideal_best) ** 2).sum(axis=1))
                            distance_worst = np.sqrt(((weighted_normalized - ideal_worst) ** 2).sum(axis=1))
                            
                            relative_closeness = distance_worst / (distance_best + distance_worst + 1e-10)
                            
                            result_df = pd.DataFrame({
                                '指标': indicators_clean,
                                '与正理想解距离': distance_best,
                                '与负理想解距离': distance_worst,
                                '相对接近度': relative_closeness,
                                '综合得分': relative_closeness * 100
                            })
                            
                            result_df = result_df.sort_values('综合得分', ascending=False)
                            result_df['排名'] = range(1, len(result_df)+1)
                            
                            base_name = os.path.splitext(pf)[0].replace('_已处理', '')
                            
                            all_results[f"{folder}_{base_name}"] = {
                                'folder': folder,
                                'file': pf,
                                'indicators': indicators_clean.tolist(),
                                'years': year_cols,
                                'weights': weights.tolist(),
                                'topsis_results': result_df.to_dict('records'),
                                'topsis_df': result_df
                            }
                            
                            print(f"\n  文件: {pf}")
                            print(f"  评价指标数: {len(indicators_clean)}")
                            print(f"  前5名指标:")
                            for i, row in result_df.head(5).iterrows():
                                print(f"    {row['排名']}. {row['指标']}: {row['综合得分']:.4f}")
                            
                            result_excel = os.path.join(folder_path, f"{base_name}_TOPSIS分析.xlsx")
                            result_df.to_excel(result_excel, index=False)
                            
                except Exception as e:
                    print(f"  处理 {pf} 时出错: {e}")
    
    return all_results

def generate_topsis_report(all_results):
    print("\n" + "="*80)
    print("步骤4: 生成TOPSIS分析报告")
    print("="*80)
    
    report = {
        'report_title': 'TOPSIS综合评价分析报告',
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'methodology': {
            'name': 'TOPSIS法',
            'description': '逼近理想解排序法，通过计算与正理想解和负理想解的距离进行综合评价',
            'steps': [
                '1. 数据标准化',
                '2. 确定权重（熵权法）',
                '3. 计算加权标准化矩阵',
                '4. 确定正理想解和负理想解',
                '5. 计算欧氏距离',
                '6. 计算相对接近度和综合得分'
            ]
        },
        'results': {}
    }
    
    for key, result in all_results.items():
        report['results'][key] = {
            'folder': result['folder'],
            'file': result['file'],
            'indicators': result['indicators'],
            'years': result['years'],
            'topsis_results': result['topsis_results']
        }
    
    report_path = 'TOPSIS综合评价报告.json'
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    print(f"JSON报告已保存: {report_path}")
    
    md_report = '# TOPSIS综合评价分析报告\n\n'
    md_report += f'**生成时间**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n'
    md_report += '## 方法说明\n\n'
    md_report += '### TOPSIS法简介\n'
    md_report += 'TOPSIS（Technique for Order Preference by Similarity to an Ideal Solution）是一种常用的多属性决策分析方法，通过计算各评价对象与正理想解和负理想解的距离来进行综合排序。\n\n'
    md_report += '### 计算步骤\n'
    md_report += '1. **数据标准化**: 采用向量归一化法\n'
    md_report += '2. **确定权重**: 采用熵权法确定各指标权重\n'
    md_report += '3. **加权标准化矩阵**: 计算加权后的标准化数据\n'
    md_report += '4. **确定理想解**: 正理想解（各指标最大值）和负理想解（各指标最小值）\n'
    md_report += '5. **计算距离**: 各评价对象与正、负理想解的欧氏距离\n'
    md_report += '6. **相对接近度**: 计算综合得分并排序\n\n'
    md_report += '## 分析结果\n\n'
    
    for key, result in all_results.items():
        md_report += f'### {result["folder"]} - {result["file"]}\n\n'
        md_report += f'- **评价指标数**: {len(result["indicators"])}\n'
        md_report += f'- **评价年份**: {", ".join(result["years"])}\n\n'
        md_report += '#### 综合排名（前10位）\n\n'
        md_report += '| 排名 | 指标 | 综合得分 |\n'
        md_report += '|------|------|----------|\n'
        
        topsis_df = result['topsis_df']
        for _, row in topsis_df.head(10).iterrows():
            md_report += f'| {row["排名"]} | {row["指标"]} | {row["综合得分"]:.4f} |\n'
        if len(topsis_df) > 10:
            md_report += '| ... | ... | ... |\n'
        md_report += '\n'
    
    with open('TOPSIS综合评价报告.md', 'w', encoding='utf-8') as f:
        f.write(md_report)
    
    print(f"Markdown报告已保存: TOPSIS综合评价报告.md")

def main():
    print("="*100)
    print("文件检查与TOPSIS综合评价工具")
    print("="*100)
    
    check_file_integrity()
    fix_entropy_report()
    all_results = topsis_analysis()
    generate_topsis_report(all_results)
    
    print("\n" + "="*100)
    print("所有工作完成！")
    print("="*100)

if __name__ == "__main__":
    main()
