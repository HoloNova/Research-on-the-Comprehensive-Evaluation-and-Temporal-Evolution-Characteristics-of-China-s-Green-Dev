import os
import shutil
import pandas as pd
import numpy as np
from datetime import datetime
import json

def delete_processed_files():
    print("="*80)
    print("步骤1: 清理所有已处理文件")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    deleted_count = 0
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            files = os.listdir(folder_path)
            for f in files:
                if '已处理' in f:
                    os.remove(os.path.join(folder_path, f))
                    deleted_count += 1
                    print(f"  删除: {folder}/{f}")
    
    print(f"\n共删除 {deleted_count} 个已处理文件")
    return deleted_count

def rename_original_files():
    print("\n" + "="*80)
    print("步骤2: 重命名原始文件")
    print("="*80)
    
    file_mapping = {
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
    
    print(f"\n共重命名 {renamed_count} 个原始文件")
    return renamed_count

def parse_and_transform_data(df):
    header_row_idx = 2
    data_start_idx = 3
    
    df.columns = df.iloc[header_row_idx]
    df = df.iloc[data_start_idx:].reset_index(drop=True)
    
    df = df[df.iloc[:, 0].notna()]
    df = df[~df.iloc[:, 0].astype(str).str.contains('数据来源', na=False)]
    
    df = df.rename(columns={df.columns[0]: '指标'})
    
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df

def reprocess_all_files():
    print("\n" + "="*80)
    print("步骤3: 重新处理所有数据文件")
    print("="*80)
    
    folders = ['国民消费水平', '科技', '能源', '资源与环境']
    processed_count = 0
    
    for folder in folders:
        folder_path = os.path.join('.', folder)
        if os.path.exists(folder_path):
            original_files = [f for f in os.listdir(folder_path) 
                             if f.endswith('.xlsx') and not f.startswith('~$')]
            
            for file_name in original_files:
                file_path = os.path.join(folder_path, file_name)
                
                try:
                    df = pd.read_excel(file_path, header=None)
                    df_transformed = parse_and_transform_data(df)
                    
                    for col in df_transformed.columns[1:]:
                        if pd.api.types.is_numeric_dtype(df_transformed[col]):
                            df_transformed[col] = df_transformed[col].fillna(df_transformed[col].median())
                    
                    df_transformed = df_transformed.drop_duplicates(subset=['指标'], keep='first')
                    
                    numeric_cols = df_transformed.select_dtypes(include=[np.number]).columns
                    for col in numeric_cols:
                        Q1 = df_transformed[col].quantile(0.25)
                        Q3 = df_transformed[col].quantile(0.75)
                        IQR = Q3 - Q1
                        lower_bound = Q1 - 1.5 * IQR
                        upper_bound = Q3 + 1.5 * IQR
                        df_transformed[col] = np.where(df_transformed[col] < lower_bound, lower_bound, df_transformed[col])
                        df_transformed[col] = np.where(df_transformed[col] > upper_bound, upper_bound, df_transformed[col])
                    
                    df_transformed['指标'] = df_transformed['指标'].astype(str).str.strip()
                    
                    for col in numeric_cols:
                        min_val = df_transformed[col].min()
                        max_val = df_transformed[col].max()
                        if max_val - min_val != 0:
                            df_transformed[f"{col}_标准化"] = (df_transformed[col] - min_val) / (max_val - min_val)
                        else:
                            df_transformed[f"{col}_标准化"] = 0
                    
                    base_name = os.path.splitext(file_name)[0]
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    
                    new_excel_name = f"{base_name}_已处理_{timestamp}.xlsx"
                    new_csv_name = f"{base_name}_已处理_{timestamp}.csv"
                    
                    excel_path = os.path.join(folder_path, new_excel_name)
                    csv_path = os.path.join(folder_path, new_csv_name)
                    
                    df_transformed.to_excel(excel_path, index=False)
                    df_transformed.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=',', quotechar='"')
                    
                    processed_count += 1
                    print(f"  处理完成: {folder}/{file_name} -> {new_excel_name}")
                    
                except Exception as e:
                    print(f"  处理错误: {folder}/{file_name} - {e}")
                    import traceback
                    traceback.print_exc()
    
    print(f"\n共处理 {processed_count} 个文件")
    return processed_count

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
    print("\n" + "="*80)
    print("步骤4: 熵权法权重计算")
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
                            
                            base_name = os.path.splitext(pf)[0].replace('_已处理', '')
                            
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
            'note': '按指标维度进行权重分配'
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
    print("项目文件系统化整理工具（完整版）")
    print("="*100)
    
    delete_processed_files()
    rename_original_files()
    reprocess_all_files()
    results = entropy_weight_analysis()
    generate_report(results)
    
    print("\n" + "="*100)
    print("所有整理工作完成！")
    print("="*100)

if __name__ == "__main__":
    main()
