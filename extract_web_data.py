"""
网页数据提取脚本
从保存的统计公报HTML文件中提取表格数据并保存为CSV
"""

import os
import re
import pandas as pd
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings('ignore')


def extract_tables_from_html(html_file):
    """从HTML文件中提取所有表格"""
    with open(html_file, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    
    return tables


def parse_table(table):
    """解析单个表格为DataFrame"""
    rows = table.find_all('tr')
    data = []
    
    for row in rows:
        cols = row.find_all(['td', 'th'])
        cols = [col.get_text(strip=True) for col in cols]
        if cols:
            data.append(cols)
    
    if not data:
        return None
    
    df = pd.DataFrame(data)
    
    if len(df) > 0:
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    
    return df


def clean_numeric_value(value):
    """清洗数值数据"""
    if pd.isna(value) or value == '':
        return None
    
    value_str = str(value)
    
    value_str = re.sub(r'[^\d.,\-]', '', value_str)
    
    if ',' in value_str and '.' in value_str:
        if value_str.index(',') > value_str.index('.'):
            value_str = value_str.replace('.', '')
            value_str = value_str.replace(',', '.')
        else:
            value_str = value_str.replace(',', '')
    elif ',' in value_str:
        if len(value_str.split(',')[-1]) == 3:
            value_str = value_str.replace(',', '')
        else:
            value_str = value_str.replace(',', '.')
    
    try:
        return float(value_str)
    except:
        return value


def extract_year_from_filename(filename):
    """从文件名中提取年份"""
    match = re.search(r'(\d{4})', filename)
    if match:
        return int(match.group(1))
    return None


def process_all_html_files(input_dir, output_dir):
    """处理目录下所有HTML文件"""
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    html_files = [f for f in os.listdir(input_dir) if f.endswith('.html')]
    
    all_data = []
    
    for html_file in html_files:
        print(f"正在处理: {html_file}")
        
        file_path = os.path.join(input_dir, html_file)
        year = extract_year_from_filename(html_file)
        
        tables = extract_tables_from_html(file_path)
        
        for i, table in enumerate(tables):
            df = parse_table(table)
            
            if df is not None and len(df) > 0:
                df['来源文件'] = html_file
                if year:
                    df['年份'] = year
                df['表格编号'] = i + 1
                
                all_data.append(df)
                
                output_file = os.path.join(output_dir, f"{os.path.splitext(html_file)[0]}_表{i+1}.csv")
                df.to_csv(output_file, index=False, encoding='utf-8-sig')
                print(f"  保存表格 {i+1}: {output_file}")
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        combined_file = os.path.join(output_dir, '所有表格合并数据.csv')
        combined_df.to_csv(combined_file, index=False, encoding='utf-8-sig')
        print(f"\n合并数据已保存: {combined_file}")
    
    return all_data


def main():
    print("="*80)
    print("网页数据提取工具")
    print("="*80)
    
    input_dir = '网页页面'
    output_dir = '网页提取数据'
    
    if not os.path.exists(input_dir):
        print(f"错误: 输入目录 '{input_dir}' 不存在!")
        return
    
    print(f"\n输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    print("\n开始提取数据...\n")
    
    process_all_html_files(input_dir, output_dir)
    
    print("\n" + "="*80)
    print("数据提取完成!")
    print("="*80)


if __name__ == '__main__':
    main()
