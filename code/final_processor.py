import os
import pandas as pd
import numpy as np
from datetime import datetime
import json
import warnings
warnings.filterwarnings('ignore')

def delete_processed_files():
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.xlsx') and '已处理' in file:
                os.remove(os.path.join(root, file))
            if file.endswith('.csv') and '已处理' in file:
                os.remove(os.path.join(root, file))
    if os.path.exists('数据字典.json'):
        os.remove('数据字典.json')
    print("已清理所有之前处理的文件")

def find_original_excel_files():
    excel_files = []
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.xlsx') and '已处理' not in file and not file.startswith('~$'):
                excel_files.append(os.path.join(root, file))
    return excel_files

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

def process_single_file(file_path, data_dictionary):
    print(f"\n正在处理: {os.path.basename(file_path)}")
    
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
    
    directory = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    folder_name = os.path.basename(directory) if directory else '数据'
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    new_excel_name = f"{folder_name}_{base_name}_已处理_{timestamp}.xlsx"
    new_csv_name = f"{folder_name}_{base_name}_已处理_{timestamp}.csv"
    
    excel_path = os.path.join(directory, new_excel_name)
    csv_path = os.path.join(directory, new_csv_name)
    
    df_transformed.to_excel(excel_path, index=False)
    df_transformed.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=',', quotechar='"')
    
    file_dict = {
        'file_name': os.path.basename(file_path),
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'columns': {}
    }
    
    for col in df_transformed.columns:
        file_dict['columns'][col] = {
            'data_type': str(df_transformed[col].dtype),
            'non_null_count': int(df_transformed[col].count()),
            'null_count': int(df_transformed[col].isnull().sum()),
            'unique_count': int(df_transformed[col].nunique())
        }
        if pd.api.types.is_numeric_dtype(df_transformed[col]):
            file_dict['columns'][col]['min'] = float(df_transformed[col].min())
            file_dict['columns'][col]['max'] = float(df_transformed[col].max())
            file_dict['columns'][col]['mean'] = float(df_transformed[col].mean())
    
    data_dictionary[os.path.basename(file_path)] = file_dict
    
    print(f"  完成 - 列数: {len(df_transformed.columns)}, 行数: {len(df_transformed)}")
    return excel_path, csv_path

def main():
    print("="*100)
    print("Excel数据系统化处理工具")
    print("="*100)
    
    delete_processed_files()
    
    excel_files = find_original_excel_files()
    
    if not excel_files:
        print("未找到原始Excel文件！")
        return
    
    print(f"\n找到 {len(excel_files)} 个原始Excel文件")
    print(f"\n开始处理...\n")
    
    data_dictionary = {}
    processed_count = 0
    
    for file_path in excel_files:
        try:
            process_single_file(file_path, data_dictionary)
            processed_count += 1
        except Exception as e:
            print(f"  错误: {e}")
            import traceback
            traceback.print_exc()
    
    dict_path = "数据字典.json"
    with open(dict_path, 'w', encoding='utf-8') as f:
        json.dump(data_dictionary, f, ensure_ascii=False, indent=2)
    print(f"\n数据字典已保存到: {dict_path}")
    
    print("\n" + "="*100)
    print(f"所有文件处理完成！共处理 {processed_count} 个文件")
    print("="*100)

if __name__ == "__main__":
    main()
