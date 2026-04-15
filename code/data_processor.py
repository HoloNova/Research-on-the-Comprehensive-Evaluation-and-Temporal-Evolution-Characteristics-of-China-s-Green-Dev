import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import json
import warnings
warnings.filterwarnings('ignore')

sns.set_style('whitegrid')
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

class DataProcessor:
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
    
    def pause_for_confirmation(self, message):
        print(f"\n{message}")
        input("按Enter键继续...")
    
    def visualize_data_distribution(self, df, title, save_path=None):
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            fig, axes = plt.subplots(1, min(3, len(numeric_cols)), figsize=(15, 5))
            if len(numeric_cols) == 1:
                axes = [axes]
            for i, col in enumerate(numeric_cols[:3]):
                sns.histplot(df[col].dropna(), kde=True, ax=axes[i])
                axes[i].set_title(f'{col} 分布')
                axes[i].set_xlabel(col)
            plt.tight_layout()
            if save_path:
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.show()
            plt.close()
    
    def visualize_outliers(self, df, title, save_path=None):
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            fig, ax = plt.subplots(figsize=(12, 6))
            df[numeric_cols].boxplot(rot=45, ax=ax)
            ax.set_title(title)
            plt.tight_layout()
            if save_path:
                plt.savefig(save_path, dpi=300, bbox_inches='tight')
            plt.show()
            plt.close()
    
    def parse_and_transform_data(self, df):
        print(f"\n{'='*80}")
        print("数据格式转换")
        print('='*80)
        
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
        
        print(f"\n转换后数据形状: {df_transformed.shape}")
        print(f"\n转换后列名: {list(df_transformed.columns)}")
        print(f"\n转换后前5行数据:")
        print(df_transformed.head())
        
        self.pause_for_confirmation("数据格式转换完成，请检查后继续...")
        return df_transformed
    
    def load_data(self, file_path):
        print(f"\n{'='*80}")
        print(f"正在加载文件: {file_path}")
        print('='*80)
        
        df = pd.read_excel(file_path, header=None)
        print(f"\n原始数据形状: {df.shape}")
        print(f"\n原始前10行数据:")
        print(df.head(10))
        
        df_transformed = self.parse_and_transform_data(df)
        
        return df_transformed
    
    def clean_missing_values(self, df):
        print(f"\n{'='*80}")
        print("步骤1: 处理缺失值")
        print('='*80)
        
        missing_before = df.isnull().sum().sum()
        print(f"\n总缺失值数量（处理前）: {missing_before}")
        
        for col in df.columns[1:]:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(df[col].median())
        
        missing_after = df.isnull().sum().sum()
        print(f"总缺失值数量（处理后）: {missing_after}")
        
        self.pause_for_confirmation("缺失值处理完成，请检查后继续...")
        return df
    
    def remove_duplicates(self, df):
        print(f"\n{'='*80}")
        print("步骤2: 处理重复值")
        print('='*80)
        
        duplicates_before = df.duplicated(subset=['指标']).sum()
        print(f"\n重复指标数（处理前）: {duplicates_before}")
        
        df = df.drop_duplicates(subset=['指标'], keep='first')
        
        duplicates_after = df.duplicated(subset=['指标']).sum()
        print(f"重复指标数（处理后）: {duplicates_after}")
        print(f"当前数据形状: {df.shape}")
        
        self.pause_for_confirmation("重复值处理完成，请检查后继续...")
        return df
    
    def handle_outliers(self, df, method='iqr'):
        print(f"\n{'='*80}")
        print("步骤3: 处理异常值")
        print('='*80)
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            print("没有数值型列，跳过异常值处理")
            return df
        
        print(f"\n使用{method.upper()}方法检测异常值")
        
        try:
            self.visualize_outliers(df, "异常值检测（处理前）", "outliers_before.png")
        except:
            print("可视化跳过")
        
        outliers_count = {}
        for col in numeric_cols:
            if method == 'iqr':
                Q1 = df[col].quantile(0.25)
                Q3 = df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                outliers = ((df[col] < lower_bound) | (df[col] > upper_bound)).sum()
                df[col] = np.where(df[col] < lower_bound, lower_bound, df[col])
                df[col] = np.where(df[col] > upper_bound, upper_bound, df[col])
            elif method == 'zscore':
                z_scores = np.abs((df[col] - df[col].mean()) / df[col].std())
                outliers = (z_scores > 3).sum()
                df[col] = np.where(z_scores > 3, df[col].median(), df[col])
            outliers_count[col] = outliers
        
        print(f"\n各列异常值数量:")
        for col, count in outliers_count.items():
            print(f"  {col}: {count}")
        
        try:
            self.visualize_outliers(df, "异常值处理（处理后）", "outliers_after.png")
        except:
            print("可视化跳过")
        
        self.pause_for_confirmation("异常值处理完成，请检查后继续...")
        return df
    
    def standardize_formats(self, df):
        print(f"\n{'='*80}")
        print("步骤4: 数据格式标准化")
        print('='*80)
        
        df['指标'] = df['指标'].astype(str).str.strip()
        
        print(f"\n处理后数据类型:")
        print(df.dtypes)
        
        self.pause_for_confirmation("格式标准化完成，请检查后继续...")
        return df
    
    def min_max_normalization(self, df):
        print(f"\n{'='*80}")
        print("步骤5: 极差标准化 [0,1]")
        print('='*80)
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) == 0:
            print("没有数值型列，跳过标准化")
            return df
        
        print(f"\n将标准化以下列: {list(numeric_cols)}")
        
        df_original = df.copy()
        
        for col in numeric_cols:
            min_val = df[col].min()
            max_val = df[col].max()
            if max_val - min_val != 0:
                df[f"{col}_标准化"] = (df[col] - min_val) / (max_val - min_val)
            else:
                df[f"{col}_标准化"] = 0
        
        print(f"\n标准化完成，新增列: {[f'{col}_标准化' for col in numeric_cols]}")
        
        try:
            self.visualize_data_distribution(df_original, "原始数据分布", "original_dist.png")
            self.visualize_data_distribution(df[[f'{col}_标准化' for col in numeric_cols]], "标准化数据分布", "normalized_dist.png")
        except:
            print("可视化跳过")
        
        self.pause_for_confirmation("标准化完成，请检查后继续...")
        return df
    
    def create_data_dictionary(self, df, file_name):
        print(f"\n{'='*80}")
        print("创建数据字典")
        print('='*80)
        
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
        print(f"\n数据字典已创建")
        return data_dict
    
    def compare_before_after(self, df_before, df_after):
        print(f"\n{'='*80}")
        print("数据处理前后对比")
        print('='*80)
        
        print(f"\n形状对比:")
        print(f"  处理前: {df_before.shape}")
        print(f"  处理后: {df_after.shape}")
        
        numeric_cols_before = df_before.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols_before) > 0:
            print(f"\n数值型列统计对比（前3列）:")
            for col in numeric_cols_before[:3]:
                if col in df_after.columns:
                    print(f"\n  {col}:")
                    print(f"    处理前 - 均值: {df_before[col].mean():.4f}, 标准差: {df_before[col].std():.4f}")
                    print(f"    处理后 - 均值: {df_after[col].mean():.4f}, 标准差: {df_after[col].std():.4f}")
        
        print(f"\n处理后数据预览（前5行）:")
        print(df_after.head())
        
        self.pause_for_confirmation("对比完成，请检查后继续...")
    
    def save_data(self, df, original_path):
        print(f"\n{'='*80}")
        print("保存处理后的数据")
        print('='*80)
        
        directory = os.path.dirname(original_path)
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        folder_name = os.path.basename(directory) if directory else '数据'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        new_excel_name = f"{folder_name}_{base_name}_已处理_{timestamp}.xlsx"
        new_csv_name = f"{folder_name}_{base_name}_已处理_{timestamp}.csv"
        
        excel_path = os.path.join(directory, new_excel_name)
        csv_path = os.path.join(directory, new_csv_name)
        
        df.to_excel(excel_path, index=False)
        print(f"\nExcel文件已保存: {excel_path}")
        
        df.to_csv(csv_path, index=False, encoding='utf-8-sig', sep=',', quotechar='"')
        print(f"CSV文件已保存: {csv_path}")
        
        self.processed_files.append({
            'original': original_path,
            'excel': excel_path,
            'csv': csv_path
        })
        
        self.pause_for_confirmation("文件保存完成！")
        return excel_path, csv_path
    
    def save_data_dictionary(self):
        dict_path = "数据字典.json"
        with open(dict_path, 'w', encoding='utf-8') as f:
            json.dump(self.data_dictionary, f, ensure_ascii=False, indent=2)
        print(f"\n数据字典已保存到: {dict_path}")
    
    def process_single_file(self, file_path):
        print(f"\n" + "="*100)
        print(f"开始处理文件: {os.path.basename(file_path)}")
        print("="*100)
        
        df_original = self.load_data(file_path)
        df = df_original.copy()
        
        df = self.clean_missing_values(df)
        df = self.remove_duplicates(df)
        df = self.handle_outliers(df, method='iqr')
        df = self.standardize_formats(df)
        df = self.min_max_normalization(df)
        
        self.compare_before_after(df_original, df)
        self.create_data_dictionary(df, os.path.basename(file_path))
        excel_path, csv_path = self.save_data(df, file_path)
        
        return df
    
    def process_all_files(self):
        print("="*100)
        print("Excel数据系统化处理工具")
        print("="*100)
        
        excel_files = self.find_excel_files()
        
        if not excel_files:
            print("未找到Excel文件！")
            return
        
        print(f"\n找到 {len(excel_files)} 个Excel文件:")
        for i, f in enumerate(excel_files, 1):
            print(f"  {i}. {f}")
        
        self.pause_for_confirmation("准备开始处理文件...")
        
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
    processor = DataProcessor()
    processor.process_all_files()
