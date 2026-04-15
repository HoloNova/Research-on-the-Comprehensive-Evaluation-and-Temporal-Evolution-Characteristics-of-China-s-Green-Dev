import os
import pandas as pd
import numpy as np
from datetime import datetime
import json

def calculate_entropy_weight(data):
    data = np.array(data, dtype=np.float64)
    m, n = data.shape
    
    k = 1 / np.log(m) if m > 1 else 1
    
    P = data / data.sum(axis=0)
    
    e = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    
    d = 1 - e
    
    w = d / d.sum() if d.sum() != 0 else np.ones(n) / n
    
    return w

def entropy_weight_analysis():
    print("="*80)
    print("熵权法权重计算（修正版）")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    results = {}
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            processed_files = [f for f in os.listdir(folder_path) 
                             if '已处理' in f and f.endswith('.xlsx')]
            
            for pf in processed_files:
                file_path = os.path.join(folder_path, pf)
                
                try:
                    df = pd.read_excel(file_path)
                    year_cols = [col for col in df.columns if '年' in col and '标准化' not in col]
                    
                    if len(year_cols) >= 2:
                        data_matrix = df[year_cols].values
                        
                        data_clean = data_matrix[~np.isnan(data_matrix).any(axis=1)]
                        indicators_clean = df['指标'].values[~np.isnan(data_matrix).any(axis=1)]
                        
                        if len(indicators_clean) >= 2 and data_clean.shape[1] >= 2:
                            weights = calculate_entropy_weight(data_clean)
                            
                            base_name = pf.replace('_已处理_', '').split('_202')[0]
                            
                            results[base_name] = {
                                'folder': folder,
                                'file': pf,
                                'indicators': indicators_clean.tolist(),
                                'years': year_cols,
                                'weights_by_indicator': weights.tolist(),
                                'weight_dict': dict(zip(indicators_clean, weights.tolist()))
                            }
                            
                            print(f"\n  文件: {pf}")
                            print(f"  有效指标数: {len(indicators_clean)}")
                            print(f"  年份: {year_cols}")
                            print(f"  前5个指标权重:")
                            for i, (ind, w) in enumerate(list(zip(indicators_clean, weights))[:5]):
                                print(f"    {ind}: {w:.6f}")
                            if len(indicators_clean) > 5:
                                print(f"    ... 还有 {len(indicators_clean)-5} 个指标")
                except Exception as e:
                    print(f"  处理 {pf} 时出错: {e}")
                    import traceback
                    traceback.print_exc()
    
    return results

def generate_report(results):
    print("\n" + "="*80)
    print("生成权重分析报告")
    print("="*80)
    
    report = {
        'report_title': '熵权法权重分析报告',
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'methodology': {
            'name': '熵权法',
            'description': '基于信息熵的客观赋权方法，通过数据离散程度确定权重',
            'note': '本次计算按指标维度进行权重分配'
        },
        'results': results
    }
    
    report_path = '熵权法权重分析报告.json'
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    print(f"\n报告已保存到: {report_path}")
    
    md_report = '# 熵权法权重分析报告\n\n'
    md_report += f'**生成时间**: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n'
    md_report += '## 方法说明\n\n'
    md_report += '### 熵权法简介\n'
    md_report += '熵权法是一种基于信息熵的客观赋权方法，通过数据的离散程度来确定各指标的权重。\n\n'
    md_report += '### 计算步骤\n'
    md_report += '1. 数据标准化（极差标准化）\n'
    md_report += '2. 计算比重 P_ij\n'
    md_report += '3. 计算熵值 e_j\n'
    md_report += '4. 计算差异系数 d_j = 1 - e_j\n'
    md_report += '5. 计算权重 w_j = d_j / sum(d_j)\n\n'
    md_report += '## 分析结果\n\n'
    
    for file_name, result in results.items():
        md_report += f'### {result["folder"]} - {result["file"]}\n\n'
        md_report += f'- **有效指标数量**: {len(result["indicators"])}\n'
        md_report += f'- **分析年份**: {", ".join(result["years"])}\n\n'
        md_report += '#### 指标权重（前10位）\n\n'
        md_report += '| 指标 | 权重 |\n'
        md_report += '|------|------|\n'
        
        sorted_weights = sorted(result['weight_dict'].items(), key=lambda x: x[1], reverse=True)
        for ind, weight in sorted_weights[:10]:
            md_report += f'| {ind} | {weight:.6f} |\n'
        if len(sorted_weights) > 10:
            md_report += f'| ... 还有 {len(sorted_weights)-10} 个指标 | |\n'
        md_report += '\n'
    
    with open('熵权法权重分析报告.md', 'w', encoding='utf-8') as f:
        f.write(md_report)
    
    print(f"Markdown报告已保存到: 熵权法权重分析报告.md")

def main():
    print("="*100)
    print("熵权法权重计算工具（修正版）")
    print("="*100)
    
    results = entropy_weight_analysis()
    generate_report(results)
    
    print("\n" + "="*100)
    print("熵权法权重计算完成！")
    print("="*100)

if __name__ == "__main__":
    main()
