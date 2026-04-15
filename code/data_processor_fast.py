import os
import pandas as pd
import numpy as np
from datetime import datetime
import json
import warnings
warnings.filterwarnings('ignore')

class DataProcessorFast:
    def __init__(self):
        self.data_dictionary = {}
        self.processed_files = []
        
    def find_excel_files(self):
        excel_files = []
        for root, dirs, files in os.walk('.'):
            for file in files:
                if file.endswith('.xlsx') and not file.startswith('~$'):
                    excel_files.append(os.path.join(root, file))
        return excel_files
    
    def parse_and_transform_data(self, df):
        header_row = 1
        data_start_row = 2
        
        years = df.iloc[header_row, 1:].values
        years = [str(y).replace('年', '') for y in years if pd.notna(y)]
        
        indicators = df.iloc[data_start_row:, 0].values
        indicators = [str(ind) for ind in indicators if pd.notna(ind) and '数据来源' not in str(ind)]
        
        data_rows = []
        for i, ind in enumerate(indicators):
            row_data = df.iloc[data_start_row + i, 1:len(years)+1].values
            data_dict = {'指标': ind}
            for j, year in enumerate(years):
                if j < len(row_data):
                    data_dict[year] = row_data[j]
            data_rows.append(data_dict)
        
        df_transformed = pd.DataFrame(data_rows)
        
        for col in df_transformed.columns[1:]:
            df_transformed[col] = pd.to_numeric(df_transformed[col], errors='coerce')
        
        return df_transformed
    
    def load_data(self, file_path):
        df = pd.read_excel(file_path, header=None)
        df_transformed = self.parse_and_transform_data(df)
        return df_transformed
    
    def clean_missing_values(self, df):
        for col in df.columns[1:]:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(df[col].median())
        return df
    
    def remove_duplicates(self, df):
        df = df.drop_duplicates(subset=['指标'], keep='first')
        return df
    
    def handle_outliers(self, df, method='iqr'):
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            return df
        
        for col in numeric_cols:
            if method == 'iqr':
                Q1 = df[col].quantile(0.25)
                Q3 = df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                df[col] = np.where(df[col] < lower_bound, lower_bound, df[col])
                df[col] = np.where(df[col] > upper_bound, upper_bound, df[col])
            elif method == 'zscore':
                z_scores = np.abs((df[col] - df[col].mean()) / df[col].std())
                df[col] = np.where(z_scores > 3, df[col].median(), df[col])
        return df
    
    def standardize_formats(self, df):
        df['指标'] = df['指标'].astype(str).str.strip()
        return df
    
    def min_max_normalization(self, df):
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            return df
        
        for col in numeric_cols:
            min_val = df[col].min()
            max_val = df[col].max()
            if max_val - min_val != 0:
                df[f"{col}_标准化"] = (df[col] - min_val) / (max_val - min_val)
            else:
                df[f"{col}_标准化"] = 0
        return df
    
    def create_data_dictionary(self, df, file_name):
        data_dict = {
            'file_name': file_name,
            'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'columns': {}
        }
        
        for col in df.columns:
            data_dict['columns'][col] = {
                'data_type': str(df[col].dtype),
                'non_null_count': int(df[col].count()),
                'null_count': int(df[col].isnull().sum()),
                'unique_count': int(df[col].nunique())
            }
            if pd.api.types.is_numeric_dtype(df[col]):
                data_dict['columns'][col]['min'] = float(df[col].min())
                data_dict['columns'][col]['max'] = float(df[col].max())
                data_dict['columns'][col]['mean'] = float(df[col].mean())
        
        self.data_dictionary[file_name] = data_dict
        return data_dict
    
    def save_data(self, df, original_path):
        directory = os.path.dirname(original_path)
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        folder_name = os.path.basename(directory) if directory else '数据'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        new_excel_name = f"{folder_name}_{base_name}_已处理_{timestamp}.xlsx"
        new_csv_name = f"{folder_name}_{base_name}_已处理_{timestamp}.csv"
        
        excel_path = os.path.join(directory, new_excel_name)
        csv_path = os.path.join(directory, new_csv_name)
        
        df.to_excel(excel_path, index=False)
        df.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=',', quotechar='"')
        
        self.processed_files.append({
            'original': original_path,
            'excel': excel_path,
            'csv': csv_path
        })
        
        return excel_path, csv_path
    
    def save_data_dictionary(self):
        dict_path = "数据字典.json"
        with open(dict_path, 'w', encoding='utf-8') as f:
            json.dump(self.data_dictionary, f, ensure_ascii=False, indent=2)
        print(f"\n数据字典已保存到: {dict_path}")
    
    def process_single_file(self, file_path):
        print(f"\n正在处理: {os.path.basename(file_path)}")
        
        df_original = self.load_data(file_path)
        df = df_original.copy()
        
        df = self.clean_missing_values(df)
        df = self.remove_duplicates(df)
        df = self.handle_outliers(df, method='iqr')
        df = self.standardize_formats(df)
        df = self.min_max_normalization(df)
        
        self.create_data_dictionary(df, os.path.basename(file_path))
        excel_path, csv_path = self.save_data(df, file_path)
        
        print(f"  完成: {os.path.basename(excel_path)}")
        return df
    
    def process_all_files(self):
        print("="*100)
        print("Excel数据系统化处理工具 - 快速模式")
        print("="*100)
        
        excel_files = self.find_excel_files()
        
        if not excel_files:
            print("未找到Excel文件！")
            return
        
        print(f"\n找到 {len(excel_files)} 个Excel文件")
        print(f"\n开始处理...\n")
        
        for file_path in excel_files:
            try:
                self.process_single_file(file_path)
            except Exception as e:
                print(f"\n处理文件 {file_path} 时出错: {e}")
                import traceback
                traceback.print_exc()
        
        self.save_data_dictionary()
        
        print("\n" + "="*100)
        print("所有文件处理完成！")
        print(f"共处理 {len(self.processed_files)} 个文件")
        print("="*100)

if __name__ == "__main__":
    processor = DataProcessorFast()
    processor.process_all_files()
