import os
import shutil
import pandas as pd
import numpy as np
from datetime import datetime
import json

def delete_duplicate_processed_files():
    print("="*80)
    print("步骤1: 清理重复的已处理文件")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    deleted_count = 0
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            files = os.listdir(folder_path)
            processed_files = [f for f in files if '已处理' in f]
            
            file_groups = {}
            for f in processed_files:
                base_name = '_'.join(f.split('_')[:-2])
                if base_name not in file_groups:
                    file_groups[base_name] = []
                file_groups[base_name].append(f)
            
            for base_name, group in file_groups.items():
                if len(group) > 1:
                    group.sort(reverse=True)
                    for f in group[1:]:
                        os.remove(os.path.join(folder_path, f))
                        deleted_count += 1
                        print(f"  删除: {folder}/{f}")
    
    print(f"\n共删除 {deleted_count} 个重复文件")
    return deleted_count

def create_code_folder_and_move_scripts():
    print("\n" + "="*80)
    print("步骤2: 创建code文件夹并迁移脚本文件")
    print("="*80)
    
    code_folder = 'code'
    if not os.path.exists(code_folder):
        os.makedirs(code_folder)
        print(f"  创建文件夹: {code_folder}")
    
    scripts = [
        'data_processor.py',
        'data_processor_fast.py',
        'final_processor.py',
        'run_processor.py',
        'analyze_files.py'
    ]
    
    moved_count = 0
    for script in scripts:
        if os.path.exists(script):
            shutil.move(script, os.path.join(code_folder, script))
            moved_count += 1
            print(f"  移动: {script} -> {code_folder}/{script}")
    
    print(f"\n共移动 {moved_count} 个脚本文件")
    return moved_count

def get_file_mapping():
    return {
        '国民消费水平': {
            '年度数据 (1).xlsx': '居民消费水平数据.xlsx',
            '年度数据 (2).xlsx': 'GDP及产业增长指数_上年=100.xlsx',
            '年度数据 (3).xlsx': 'GDP及产业增长指数_1978=100.xlsx',
            '年度数据 (4).xlsx': '三次产业贡献率数据.xlsx',
            '年度数据 (5).xlsx': 'GDP增长拉动数据.xlsx',
            '年度数据.xlsx': '国民经济总量数据.xlsx'
        },
        '科技': {
            '年度数据 (19).xlsx': 'R&D研究与试验发展数据.xlsx',
            '年度数据 (20).xlsx': '高等学校科技数据.xlsx'
        },
        '能源': {
            '年度数据 (10).xlsx': '能源消费总量与结构数据.xlsx',
            '年度数据 (11).xlsx': '能源供需平衡数据.xlsx',
            '年度数据 (6).xlsx': '平均每天能源消费量数据.xlsx',
            '年度数据 (7).xlsx': '生活能源消费量数据.xlsx',
            '年度数据 (8).xlsx': '人均能源生产与消费数据.xlsx',
            '年度数据 (9).xlsx': '人均生活能源消费数据.xlsx'
        },
        '资源与环境': {
            '年度数据 (12).xlsx': '水资源总量数据.xlsx',
            '年度数据 (13).xlsx': '供用水总量数据.xlsx',
            '年度数据 (14).xlsx': '水环境污染物排放数据.xlsx',
            '年度数据 (15).xlsx': '大气环境污染物排放数据.xlsx',
            '年度数据 (16).xlsx': '森林资源数据.xlsx',
            '年度数据 (17).xlsx': '环境污染治理投资数据.xlsx',
            '年度数据 (18).xlsx': '工业污染治理投资数据.xlsx'
        }
    }

def rename_files():
    print("\n" + "="*80)
    print("步骤3: 重命名文件")
    print("="*80)
    
    file_mapping = get_file_mapping()
    renamed_count = 0
    
    for folder, files in file_mapping.items():
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            for old_name, new_name in files.items():
                old_path = os.path.join(folder_path, old_name)
                new_path = os.path.join(folder_path, new_name)
                
                if os.path.exists(old_path):
                    shutil.move(old_path, new_path)
                    renamed_count += 1
                    print(f"  重命名: {folder}/{old_name} -> {new_name}")
                
                base_name = os.path.splitext(old_name)[0]
                processed_files = [f for f in os.listdir(folder_path) 
                                 if base_name in f and '已处理' in f]
                
                for pf in processed_files:
                    new_base = os.path.splitext(new_name)[0]
                    timestamp = '_'.join(pf.split('_')[-2:])
                    new_pf = f"{new_base}_已处理_{timestamp}"
                    old_pf_path = os.path.join(folder_path, pf)
                    new_pf_path = os.path.join(folder_path, new_pf)
                    
                    if os.path.exists(old_pf_path):
                        shutil.move(old_pf_path, new_pf_path)
                        renamed_count += 1
                        print(f"  重命名: {folder}/{pf} -> {new_pf}")
    
    print(f"\n共重命名 {renamed_count} 个文件")
    return renamed_count

def calculate_entropy_weight(data):
    data = np.array(data, dtype=np.float64)
    m, n = data.shape
    
    k = 1 / np.log(m)
    
    P = data / data.sum(axis=0)
    
    e = -k * (P * np.log(P + 1e-10)).sum(axis=0)
    
    d = 1 - e
    
    w = d / d.sum()
    
    return w

def entropy_weight_analysis():
    print("\n" + "="*80)
    print("步骤4: 熵权法权重计算")
    print("="*80)
    
    file_mapping = get_file_mapping()
    results = {}
    
    for folder, files in file_mapping.items():
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            for old_name, new_name in files.items():
                base_name = os.path.splitext(new_name)[0]
                processed_files = [f for f in os.listdir(folder_path) 
                                 if base_name in f and '已处理' in f and f.endswith('.xlsx')]
                
                if processed_files:
                    pf = processed_files[0]
                    file_path = os.path.join(folder_path, pf)
                    
                    try:
                        df = pd.read_excel(file_path)
                        year_cols = [col for col in df.columns if '年' in col and '标准化' not in col]
                        
                        if len(year_cols) >= 2:
                            data_matrix = df[year_cols].values
                            
                            data_clean = data_matrix[~np.isnan(data_matrix).any(axis=1)]
                            
                            if len(data_clean) >= 2 and data_clean.shape[1] >= 2:
                                weights = calculate_entropy_weight(data_clean.T)
                                
                                results[new_name] = {
                                    'folder': folder,
                                    'file': new_name,
                                    'indicators': df['指标'].tolist(),
                                    'years': year_cols,
                                    'weights_by_year': weights.tolist(),
                                    'weight_dict': dict(zip(year_cols, weights.tolist()))
                                }
                                
                                print(f"\n  文件: {new_name}")
                                print(f"  指标数: {len(df)}")
                                print(f"  年份: {year_cols}")
                                print(f"  年份权重: {dict(zip(year_cols, [f'{w:.4f}' for w in weights]))}")
                    except Exception as e:
                        print(f"  处理 {new_name} 时出错: {e}")
    
    return results

def generate_report(results):
    print("\n" + "="*80)
    print("步骤5: 生成权重分析报告")
    print("="*80)
    
    report = {
        'report_title': '熵权法权重分析报告',
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'methodology': {
            'name': '熵权法',
            'description': '基于信息熵的客观赋权方法，通过数据离散程度确定权重',
            'formula': 'X\' = (X - X_min) / (X_max - X_min); P = X\' / sum(X\'); e = -k*sum(P*lnP); w = (1-e)/sum(1-e)'
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
        md_report += f'### {result["folder"]} - {file_name}\n\n'
        md_report += f'- **指标数量**: {len(result["indicators"])}\n'
        md_report += f'- **分析年份**: {", ".join(result["years"])}\n\n'
        md_report += '#### 年份权重\n\n'
        md_report += '| 年份 | 权重 |\n'
        md_report += '|------|------|\n'
        for year, weight in result['weight_dict'].items():
            md_report += f'| {year} | {weight:.6f} |\n'
        md_report += '\n'
    
    with open('熵权法权重分析报告.md', 'w', encoding='utf-8') as f:
        f.write(md_report)
    
    print(f"Markdown报告已保存到: 熵权法权重分析报告.md")

def main():
    print("="*100)
    print("项目文件系统化整理工具")
    print("="*100)
    
    delete_duplicate_processed_files()
    create_code_folder_and_move_scripts()
    rename_files()
    results = entropy_weight_analysis()
    generate_report(results)
    
    print("\n" + "="*100)
    print("所有整理工作完成！")
    print("="*100)

if __name__ == "__main__":
    main()
